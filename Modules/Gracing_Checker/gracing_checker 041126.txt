"""
==================================================
GRACING CHECKER + RESULT VERIFICATION & AUDIT ENGINE
==================================================

PURPOSE
--------------------------------------------------
This Streamlit tool is an Intelligent Result Verification and Audit Tool for
examination gracing sheets. Before anything is formatted it independently
re-validates every student, subject and calculated field (see RULES below):

  * Range checks   - Internal/Sem Actual <= their Max, Composite <= Composite Max,
                     negative marks / grace, invalid maximum marks
  * Recalculation  - Composite Max = Internal Max + Sem Max;
                     Composite = Internal + Sem + Sem Grace;
                     Grade from Composite % (O 90-100, A+ 80-<90, A 70-<80, B+ 60-<70,
                     B 55-<60, C 50-<55, P 40-<50, F <40);
                     Total = sum of subject Composites; Percentage = Total / Total Max
  * Business rules - Grade other than F => "Completed with Success", F => unsuccessful,
                     special codes (Z1 = AB, Z3 = UFM), abnormal / unjustified grace
  * Data integrity - duplicates, missing mandatory student details, missing or duplicate
                     appraisal entries, incomplete subject records, cohort inconsistencies

Every discrepancy is painted RED fill + BOLD YELLOW font on the exact cell, with a plain
language remark in the "Audit Remarks" column. Warnings use an amber fill. A full
"Audit Summary" / "Audit Findings" / "Student Audit Status" / "Recalculation Trail" report
is exported with the workbook (thresholds are configurable under AUDIT CONFIGURATION).

Formatting / reporting features (unchanged):

1. Loads uploaded Excel files containing pre/post gracing records.
2. Restructures student records into a consolidated tabular layout with exact column order:
   - Module Code, Module Desc, Credit, Internal Actual (Max), Internal Actual,
     Sem Actual (Max), Sem Actual, Sem Grace, Composite Score (Max), 
     Composite Score, Overall Grade, Completion Status after Gracing, Remarks & Scale.
   - Merged student header row across all columns:
     "Student Number | Student Name | Additional ID | Gender"
   - Converts marks to numeric values; replaces Z1 with "AB" and Z3 with "UFM" (0.0 calculation).
   - Highlights Z1 ("AB"), Z3 ("UFM"), marks < 40% of Max, and "F" grades in BOLD RED.
   - Summary row showing Total Composite Score, Total Composite Score (Max), 
     and Percentage in Overall Grade with thick bottom border for student separation.
   - Center alignment for columns Credit through Overall Grade.
   - Highlights 'Completed Unsuccessfuly' statuses in BOLD RED.
3. Identifies all instances where grace marks were awarded and students who failed.
4. Computes high-level statistics & summary reports.
5. Displays tab-based previews with 1-based indexing:
   - 📊 All Students (First Tab)
   - ⚠️ Failed Students
   - ✅ Graced Students
   - 📋 Graced Cases Details
   - 📚 Subject Summary
   - 👤 Student Summary
6. Generates a formatted multi-sheet Excel report ("Audit Summary", "Audit Findings", "All Students",
   "Failed Students", "Graced Students", etc.).
==================================================
"""

from collections import Counter, defaultdict
from datetime import datetime
import io
import os
import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st


# ==================================================
# HELPER FUNCTIONS
# ==================================================
def parse_scale_num(scale_val):
    """Parses maximum mark numeric float from scale strings (e.g., 'R050' -> 50.0)."""
    s = str(scale_val).strip()
    digits = re.findall(r'\d+', s)
    if digits:
        return float(digits[0])
    return 0.0


def parse_scale_str(scale_val):
    """Parses maximum mark integer display from scale strings (e.g., 'R050' -> '50')."""
    s = str(scale_val).strip()
    digits = re.findall(r'\d+', s)
    if digits:
        return str(int(digits[0]))
    return s if s.upper() not in ['NAN', 'NONE', 'NA'] else ''


def parse_mark(val):
    """
    Converts mark values to numeric or replaced text strings:
    - 'Z1' -> ('AB', 0.0, True)
    - 'Z3' -> ('UFM', 0.0, True)
    - Numeric values -> (int/float display, float calculation, False)
    """
    s = str(val).strip().upper()
    if s == 'Z1':
        return 'AB', 0.0, True
    elif s == 'Z3':
        return 'UFM', 0.0, True
    elif s in ['', 'NAN', 'NONE', 'NA']:
        return '', 0.0, False
    else:
        try:
            num = float(s)
            out_val = int(num) if num.is_integer() else num
            return out_val, float(num), False
        except (TypeError, ValueError):
            return val, 0.0, False


INTERNAL_FLAGS = [
    'Row_Type', 'Comp_Mismatch', 'Fail_Internal', 'Fail_Sem', 'Fail_Composite',
    'Red_Internal', 'Red_Sem', 'Red_Grace', 'Red_Composite', 'Audit_Cells'
]


# ==================================================
# RESULT VERIFICATION & AUDIT ENGINE
# ==================================================
# ==================================================
# AUDIT CONFIGURATION  (edit these to match college policy)
# ==================================================
GRADE_SCALE = [(90, 'O'), (80, 'A+'), (70, 'A'), (60, 'B+'), (55, 'B'), (50, 'C'), (40, 'P')]  # else 'F'
VALID_GRADES = {'O', 'A+', 'A', 'B+', 'B', 'C', 'P', 'F'}
SPECIAL_CODES = {'Z1': 'AB', 'Z3': 'UFM'}          # raw code -> display text
KNOWN_SPECIAL = set(SPECIAL_CODES) | set(SPECIAL_CODES.values())
SPECIAL_CODE_FORCES_FAIL = True                     # AB / UFM in any component => subject must be F
STATUS_SUCCESS = 'completed with success'
STATUS_FAIL = ('completed unsuccessfuly', 'completed unsuccessfully')   # file spelling + correct spelling
VALID_GENDERS = {'male', 'female', 'other', 'transgender'}
SUMMARY_NAMES = ('TOTAL', 'PERCENTAGE')

MAX_GRACE_ABS = 5            # grace in one subject above this many marks  -> warning
MAX_GRACE_PCT_OF_SEM = 10.0  # ... or above this % of Sem max             -> warning
MAX_GRACE_PER_STUDENT = 15   # total grace across all subjects of a student -> warning
MAX_GRACED_SUBJECTS = 3      # number of subjects graced for one student   -> warning
PERCENT_TOLERANCE = 0.011    # percentage in the file is rounded to 2 decimals
MARK_TOLERANCE = 0.01

EXPECTED_APPRAISALS = [
    'Overall Grade', 'Composite Score', 'Internal Actual', 'Internal Total Marks',
    'Sem Actual', 'Sem Grace', 'Semester Total Marks', 'Grace on Total marks',
]
MANDATORY_VALUE_APPRAISALS = ['Overall Grade', 'Composite Score', 'Internal Actual', 'Sem Actual']

ERR, WARN = 'ERROR', 'WARNING'

# rule id: (severity, category, description)
RULES = {
    'DI-01': (ERR, 'Data Integrity', 'Mandatory student details present (Student Number, Name, Additional ID, Gender)'),
    'DI-02': (ERR, 'Data Integrity', 'Student details identical on every row of the student'),
    'DI-03': (ERR, 'Data Integrity', 'No duplicate student (same Additional ID under different Student Numbers)'),
    'DI-04': (ERR, 'Data Integrity', 'No fully duplicated records (identical rows)'),
    'DI-05': (ERR, 'Data Integrity', 'No duplicate appraisal entries with conflicting values'),
    'DI-06': (ERR, 'Data Integrity', 'All 8 appraisal entries present for every subject'),
    'DI-07': (ERR, 'Data Integrity', 'Appraisal Type Code is a recognised value'),
    'DI-08': (ERR, 'Data Integrity', 'Module Code and Module Description present'),
    'DI-09': (ERR, 'Data Integrity', 'Module Description / Credit consistent for the same Module Code across students'),
    'DI-10': (ERR, 'Data Integrity', 'Subject record complete (Internal, Sem, Composite, Grade, Status, Credit not blank)'),
    'DI-11': (ERR, 'Data Integrity', 'Marks are numeric or a recognised result code'),
    'DI-12': (WARN, 'Data Integrity', 'Subject set of the student matches the cohort norm'),
    'DI-13': (ERR, 'Data Integrity', 'Exactly one TOTAL and one PERCENTAGE row per student'),
    'DI-14': (WARN, 'Data Integrity', 'Student Number format consistent; Gender value recognised; Credit > 0'),
    'RG-01': (ERR, 'Range & Limits', 'No negative marks or credits'),
    'RG-02': (ERR, 'Range & Limits', 'No negative grace marks'),
    'RG-03': (ERR, 'Range & Limits', 'Maximum marks (scale) valid and > 0'),
    'RG-04': (WARN, 'Range & Limits', 'Maximum marks consistent with cohort for the same Module Code'),
    'RG-05': (ERR, 'Range & Limits', 'Internal Actual <= Internal Actual (Max)'),
    'RG-06': (ERR, 'Range & Limits', 'Sem Actual <= Sem Actual (Max)'),
    'RG-07': (ERR, 'Range & Limits', 'Composite Score <= Composite Score (Max)'),
    'RG-08': (ERR, 'Range & Limits', 'Sem Actual + Sem Grace <= Sem Actual (Max)'),
    'CA-01': (ERR, 'Calculation', 'Composite Score (Max) = Internal (Max) + Sem (Max)'),
    'CA-02': (ERR, 'Calculation', 'Composite Score = Internal Actual + Sem Actual + Sem Grace'),
    'CA-03': (ERR, 'Calculation', 'Internal Total Marks = Internal Actual (absent status must match)'),
    'CA-04': (ERR, 'Calculation', 'Semester Total Marks = Sem Actual + Sem Grace (absent status must match)'),
    'CA-05': (ERR, 'Calculation', 'Overall Grade matches grade scale applied to reported Composite %'),
    'CA-06': (ERR, 'Calculation', 'Overall Grade matches grade scale applied to recalculated Composite %'),
    'CA-07': (ERR, 'Calculation', 'Overall Grade is a valid grade (O, A+, A, B+, B, C, P, F)'),
    'ST-01': (ERR, 'Student Totals', 'Total = sum of all subject Composite Scores'),
    'ST-02': (ERR, 'Student Totals', 'Percentage = Total / Total Max x 100'),
    'ST-03': (ERR, 'Student Totals', 'Percentage within 0 - 100'),
    'ST-04': (ERR, 'Student Totals', 'Total = sum of recalculated (Internal + Sem + Grace) scores'),
    'BR-01': (ERR, 'Business Rules', 'Grade other than F => "Completed with Success"'),
    'BR-02': (ERR, 'Business Rules', 'Grade F => "Completed Unsuccessfuly"'),
    'BR-03': (ERR, 'Business Rules', 'Completion Status is a recognised value'),
    'BR-04': (ERR, 'Business Rules', 'Special code (AB/UFM) in a component => subject graded F / unsuccessful'),
    'BR-05': (ERR, 'Business Rules', 'Special result codes recognised (Z1, Z3)'),
    'BR-06': (WARN, 'Business Rules', 'Special code not replaced by marks after gracing'),
    'BR-07': (ERR, 'Business Rules', 'No grace awarded to AB / UFM candidates'),
    'BR-08': (ERR, 'Business Rules', 'Scale NABS only on Internal/Sem rows and only with absent code Z1'),
    'BR-09': (WARN, 'Business Rules', 'Absent code Z1 on Internal/Sem rows carries scale NABS'),
    'BR-10': (ERR, 'Business Rules', '40% pass mark in EACH component (Internal; Sem Actual + Sem Grace) else grade F'),
    'GR-01': (WARN, 'Gracing', 'Grace marks within normal limits (per subject)'),
    'GR-02': (WARN, 'Gracing', 'Total grace per student and graced-subject count within normal limits'),
    'GR-03': (WARN, 'Gracing', 'Grace awarded only where Sem (or the subject) would otherwise fail'),
    'GR-04': (WARN, 'Gracing', 'Grace awarded but subject still fails'),
    'GR-05': (WARN, 'Gracing', '"Grace on Total marks" is zero / reflected in Composite'),
    'GR-06': (WARN, 'Gracing', 'Marks unchanged between before/after gracing unless grace awarded'),
}


# ==================================================
# AUDIT HELPERS
# ==================================================
def _sv(x):
    """Safe stripped string."""
    if x is None:
        return ''
    s = str(x).strip()
    return '' if s.upper() in ('NAN', 'NONE') else s


def _fmt(x):
    """Compact number formatting for remarks."""
    try:
        f = float(x)
        return str(int(f)) if f.is_integer() else f"{f:g}"
    except (TypeError, ValueError):
        return str(x)


def parse_audit_value(raw):
    """Classifies a marks cell: blank / num / code / unknown_code / invalid."""
    s = _sv(raw)
    if s == '' or s.upper() == 'NA':
        return {'kind': 'blank', 'num': 0.0, 'code': None, 'raw': s}
    u = s.upper()
    if u in KNOWN_SPECIAL:
        return {'kind': 'code', 'num': 0.0, 'code': SPECIAL_CODES.get(u, u), 'raw': s}
    if re.fullmatch(r'Z\d+', u):
        return {'kind': 'unknown_code', 'num': 0.0, 'code': u, 'raw': s}
    try:
        num = float(s.replace(',', ''))
        if num != num or num in (float('inf'), float('-inf')):
            raise ValueError
        return {'kind': 'num', 'num': num, 'code': None, 'raw': s}
    except ValueError:
        return {'kind': 'invalid', 'num': 0.0, 'code': None, 'raw': s}


def scale_max(scale_raw):
    """Maximum marks from scale code (e.g. 'S020' -> 20.0). None if unparsable."""
    digits = re.findall(r'\d+', _sv(scale_raw))
    return float(digits[0]) if digits else None


def grade_for_percentage(pct):
    pct = round(pct, 6)
    for lo, g in GRADE_SCALE:
        if pct + 1e-9 >= lo:
            return g
    return 'F'


class AuditResult:
    """Container for everything the audit produces."""

    def __init__(self):
        self.findings = []          # list of dict
        self.checked = Counter()    # rule -> number of checks executed
        self.failed = Counter()     # rule -> number of failures
        self.trail = []             # per subject recalculation rows
        self.student_rows = []      # per student audit status rows
        self.stats = {}
        self.ref_max = {}
        self._flags = defaultdict(lambda: {'cols': {}, 'remarks': []})

    # ---- recording ------------------------------------------------------
    def check(self, rule, ok, **kw):
        self.checked[rule] += 1
        if not ok:
            self.flag(rule, **kw)
        return ok

    def flag(self, rule, student='', name='', module='', desc='', field='', rows=(),
             expected='', actual='', message='', cols=(), target='module'):
        sev, cat, _ = RULES[rule]
        self.failed[rule] += 1
        self.findings.append({
            'Severity': sev, 'Rule ID': rule, 'Category': cat,
            'Student Number': student, 'Student Name': name,
            'Module Code': module, 'Module Desc': desc, 'Field': field,
            'Source Excel Row(s)': ', '.join(str(r) for r in sorted(set(rows))),
            'Expected': expected, 'Actual': actual, 'Remark': message,
        })
        key = (student, module if target == 'module' else f'#{target.upper()}')
        slot = self._flags[key]
        for c in cols:
            if slot['cols'].get(c) != ERR:
                slot['cols'][c] = sev
        slot['remarks'].append(f"[{sev}] {rule}: {message}")

    # ---- lookups used by the processed-sheet builder --------------------
    def cell_flags(self, student, module=None, target='module'):
        stu = str(student).strip() or '(blank)'
        key = (stu, str(module).strip() if target == 'module' else f'#{target.upper()}')
        return self._flags[key] if key in self._flags else {'cols': {}, 'remarks': []}

    def student_counts(self, student):
        """(errors, warnings) for one student - cached, rebuilt whenever findings change."""
        if getattr(self, '_cnt_len', -1) != len(self.findings):
            self._cnt = defaultdict(lambda: [0, 0])
            for f in self.findings:
                self._cnt[f['Student Number']][0 if f['Severity'] == ERR else 1] += 1
            self._cnt_len = len(self.findings)
        e, w = self._cnt[str(student).strip() or '(blank)'] if (str(student).strip() or '(blank)') in self._cnt else (0, 0)
        return e, w

    # ---- derived tables -------------------------------------------------
    def findings_df(self):
        cols = ['Severity', 'Rule ID', 'Category', 'Student Number', 'Student Name', 'Module Code',
                'Module Desc', 'Field', 'Source Excel Row(s)', 'Expected', 'Actual', 'Remark']
        df = pd.DataFrame(self.findings, columns=cols)
        if not df.empty:
            df['_o'] = df['Severity'].map({ERR: 0, WARN: 1})
            df = df.sort_values(['_o', 'Rule ID', 'Student Number', 'Module Code'], kind='stable').drop(columns='_o')
        df.insert(0, 'Sr No', range(1, len(df) + 1))
        return df.reset_index(drop=True)

    def rule_table(self):
        rows = []
        for rid, (sev, cat, desc) in RULES.items():
            chk, bad = self.checked.get(rid, 0), self.failed.get(rid, 0)
            if chk == 0 and bad == 0:
                status = 'N/A (not applicable)'
            elif bad == 0:
                status = 'PASS'
            else:
                status = 'FAIL' if sev == ERR else 'REVIEW'
            rows.append({'Rule ID': rid, 'Category': cat, 'Severity if Failed': sev, 'Validation Rule': desc,
                         'Checks Executed': chk, 'Failures': bad, 'Result': status})
        return pd.DataFrame(rows)

    @property
    def n_errors(self):
        return sum(1 for f in self.findings if f['Severity'] == ERR)

    @property
    def n_warnings(self):
        return sum(1 for f in self.findings if f['Severity'] == WARN)


# ==================================================
# AUDIT ENGINE
# ==================================================
def run_audit(raw_df):
    """
    Independent re-validation of every student, subject and calculated field.
    Works on the raw uploaded sheet; never trusts a derived value in the file.
    """
    A = AuditResult()
    df = raw_df.copy()
    df.columns = df.columns.astype(str).str.strip()
    df['_xl'] = range(2, len(df) + 2)                       # Excel row number (header = row 1)
    for c in ['Student Number', 'Student Name', 'Additional ID', 'Gender', 'Module Code', 'Module Desc',
              'Appraisal Type Code', 'Scale', 'Marks before Gracing', 'Marks after Gracing',
              'Completion Status after Gracing', 'Credit']:
        if c not in df.columns:
            df[c] = ''
    S = lambda row, col: _sv(row.get(col, ''))             # noqa: E731

    records = df.to_dict('records')
    content_cols = [c for c in df.columns if c != '_xl']

    # ---------- 1. fully duplicated rows -----------------------------------
    sig_groups = defaultdict(list)
    for r in records:
        sig_groups[tuple(_sv(r.get(c, '')) for c in content_cols)].append(r)
    dup_skip = set()
    for sig, grp in sig_groups.items():
        A.checked['DI-04'] += len(grp) if len(grp) > 1 else 1
        if len(grp) > 1:
            for extra in grp[1:]:
                dup_skip.add(extra['_xl'])
            first = grp[0]
            A.flag('DI-04', student=S(first, 'Student Number'), name=S(first, 'Student Name'),
                   module=S(first, 'Module Code'), desc=S(first, 'Module Desc'),
                   field=S(first, 'Appraisal Type Code') or S(first, 'Student Name'),
                   rows=[g['_xl'] for g in grp], actual=f'{len(grp)} identical rows', expected='1 row',
                   message=f"Identical record appears {len(grp)} times (Excel rows "
                           f"{', '.join(str(g['_xl']) for g in grp)}); only the first is used.",
                   cols=['Module Code'],
                   target='module' if S(first, 'Module Code') else 'header')

    # ---------- 2. group by student ----------------------------------------
    students = defaultdict(list)
    for r in records:
        students[S(r, 'Student Number')].append(r)

    prn_lengths = Counter(len(k) for k in students if k)
    common_len = prn_lengths.most_common(1)[0][0] if prn_lengths else None

    # duplicate students via Additional ID
    addl_map = defaultdict(set)
    for r in records:
        a, p = S(r, 'Additional ID'), S(r, 'Student Number')
        if a and p and S(r, 'Student Name').upper() not in SUMMARY_NAMES:
            addl_map[a].add(p)

    # Reference component maxima per module (from rows whose scale is not NABS); used to resolve the
    # maximum of an absent (scale NABS) component.
    _rc = defaultdict(Counter)
    for r in records:
        code = S(r, 'Appraisal Type Code')
        if code in ('Internal Actual', 'Sem Actual') and S(r, 'Scale').upper() != 'NABS':
            mx = scale_max(S(r, 'Scale'))
            if mx:
                _rc[(S(r, 'Module Code'), code)][mx] += 1
    A.ref_max = {k: c.most_common(1)[0][0] for k, c in _rc.items()}

    module_obs = defaultdict(list)      # module -> list of (student, name, desc, maxes, rows)
    module_ref = defaultdict(lambda: {'desc': Counter(), 'credit': Counter()})
    student_modules = {}
    student_names = {}
    n_subject_records = 0

    for stu, rows in students.items():
        rows = [r for r in rows if r['_xl'] not in dup_skip]
        summary_rows = [r for r in rows if S(r, 'Student Name').upper() in SUMMARY_NAMES
                        and S(r, 'Appraisal Type Code') == '']
        mod_rows = [r for r in rows if r not in summary_rows]

        # ---- student details ------------------------------------------------
        det_src = mod_rows[0] if mod_rows else (rows[0] if rows else {})
        name = S(det_src, 'Student Name')
        addl = S(det_src, 'Additional ID')
        gender = S(det_src, 'Gender')
        student_names[stu] = name
        disp_stu = stu or '(blank)'
        xl_rows = [r['_xl'] for r in (mod_rows or rows)]

        for fld, val in (('Student Number', stu), ('Student Name', name),
                         ('Additional ID', addl), ('Gender', gender)):
            A.check('DI-01', val != '', student=disp_stu, name=name, field=fld, rows=xl_rows[:5],
                    expected='value present', actual='blank', cols=['Module Code'], target='header',
                    message=f"Mandatory student detail '{fld}' is missing.")

        for fld in ('Student Name', 'Additional ID', 'Gender'):
            vals = {S(r, fld) for r in mod_rows}
            A.check('DI-02', len(vals) <= 1, student=disp_stu, name=name, field=fld, rows=xl_rows,
                    expected='one consistent value', actual=' / '.join(sorted(v or '(blank)' for v in vals)),
                    cols=['Module Code'], target='header',
                    message=f"'{fld}' differs between rows of the same student: "
                            f"{', '.join(sorted(v or '(blank)' for v in vals))}.")

        if addl:
            others = addl_map.get(addl, set())
            A.check('DI-03', len(others) <= 1, student=disp_stu, name=name, field='Additional ID', rows=xl_rows[:5],
                    expected='unique per student', actual=', '.join(sorted(others)),
                    cols=['Module Code'], target='header',
                    message=f"Additional ID '{addl}' is used by multiple Student Numbers "
                            f"({', '.join(sorted(others))}) - possible duplicate student.")

        if stu:
            fmt_ok = stu.isdigit() and (common_len is None or len(stu) == common_len)
            A.check('DI-14', fmt_ok, student=disp_stu, name=name, field='Student Number', rows=xl_rows[:5],
                    expected=f"{common_len} digits" if common_len else 'numeric', actual=stu,
                    cols=['Module Code'], target='header',
                    message=f"Student Number '{stu}' is not in the usual format "
                            f"({common_len}-digit numeric).")
        if gender:
            A.check('DI-14', gender.lower() in VALID_GENDERS, student=disp_stu, name=name, field='Gender',
                    rows=xl_rows[:5], expected='Male / Female / Other', actual=gender,
                    cols=['Module Code'], target='header', message=f"Unrecognised Gender value '{gender}'.")

        # ---- TOTAL / PERCENTAGE rows ----------------------------------------
        tot_rows = [r for r in summary_rows if S(r, 'Student Name').upper() == 'TOTAL']
        pct_rows = [r for r in summary_rows if S(r, 'Student Name').upper() == 'PERCENTAGE']
        for lbl, lst in (('TOTAL', tot_rows), ('PERCENTAGE', pct_rows)):
            A.check('DI-13', len(lst) == 1, student=disp_stu, name=name, field=lbl,
                    rows=[r['_xl'] for r in lst], expected='1 row', actual=f'{len(lst)} rows',
                    cols=['Composite Score', 'Overall Grade'], target='summary',
                    message=f"{'Missing' if not lst else 'Duplicate'} {lbl} row ({len(lst)} found, 1 expected).")

        # ---- unrecognised module-level rows ---------------------------------
        by_mod = defaultdict(list)
        for r in mod_rows:
            by_mod[S(r, 'Module Code')].append(r)

        student_modules[stu] = set(m for m in by_mod if m)
        s_sum_comp = s_sum_comp_rec = s_sum_max = 0.0
        s_total_grace = 0.0
        s_graced_subjects = 0
        s_sum_valid = True

        for m_code, m_rows in by_mod.items():
            n_subject_records += 1
            m_desc = next((S(r, 'Module Desc') for r in m_rows if S(r, 'Module Desc')), '')
            m_xl = [r['_xl'] for r in m_rows]
            ctx = dict(student=disp_stu, name=name, module=m_code, desc=m_desc)

            A.check('DI-08', m_code != '' and m_desc != '', rows=m_xl[:5], field='Module Code / Desc',
                    expected='both present', actual=f"code='{m_code}' desc='{m_desc}'",
                    cols=['Module Code', 'Module Desc'],
                    message="Module Code or Module Description is missing.", **ctx)

            app_map = defaultdict(list)
            for r in m_rows:
                code = S(r, 'Appraisal Type Code')
                ok = code in EXPECTED_APPRAISALS
                A.check('DI-07', ok, rows=[r['_xl']], field='Appraisal Type Code',
                        expected=' / '.join(EXPECTED_APPRAISALS[:3]) + ' ...', actual=code or '(blank)',
                        cols=['Module Code'], message=f"Unrecognised Appraisal Type Code '{code or '(blank)'}'.",
                        **ctx)
                if ok:
                    app_map[code].append(r)

            def first(code):
                return app_map[code][0] if app_map.get(code) else None

            def after(code):
                r = first(code)
                return parse_audit_value(r.get('Marks after Gracing', '')) if r else {'kind': 'missing', 'num': 0.0, 'code': None, 'raw': ''}

            def before(code):
                r = first(code)
                return parse_audit_value(r.get('Marks before Gracing', '')) if r else {'kind': 'missing', 'num': 0.0, 'code': None, 'raw': ''}

            def xl(code):
                r = first(code)
                return [r['_xl']] if r else m_xl[:3]

            col_of = {'Internal Actual': 'Internal Actual', 'Internal Total Marks': 'Internal Actual',
                      'Sem Actual': 'Sem Actual', 'Sem Grace': 'Sem Grace', 'Semester Total Marks': 'Sem Actual',
                      'Composite Score': 'Composite Score', 'Grace on Total marks': 'Sem Grace'}

            v_int, v_sem, v_gr = after('Internal Actual'), after('Sem Actual'), after('Sem Grace')
            v_comp = after('Composite Score')
            v_itot, v_stot, v_gtot = after('Internal Total Marks'), after('Semester Total Marks'), after('Grace on Total marks')
            int_max = scale_max(first('Internal Actual').get('Scale', '')) if first('Internal Actual') else None
            sem_max = scale_max(first('Sem Actual').get('Scale', '')) if first('Sem Actual') else None
            comp_max = scale_max(first('Composite Score').get('Scale', '')) if first('Composite Score') else None
            # Scale NABS = candidate absent: component max is not stated, so resolve it from the module reference
            int_nabs = bool(first('Internal Actual')) and _sv(first('Internal Actual').get('Scale', '')).upper() == 'NABS'
            sem_nabs = bool(first('Sem Actual')) and _sv(first('Sem Actual').get('Scale', '')).upper() == 'NABS'
            if int_nabs:
                int_max = A.ref_max.get((m_code, 'Internal Actual'))
                if int_max is None and comp_max and sem_max and not sem_nabs and comp_max > sem_max:
                    int_max = comp_max - sem_max
            if sem_nabs:
                sem_max = A.ref_max.get((m_code, 'Sem Actual'))
                if sem_max is None and comp_max and int_max and not int_nabs and comp_max > int_max:
                    sem_max = comp_max - int_max

            # A subject may legitimately have no Internal (or no Sem) component: all its entries blank/missing
            # and max zero/blank. Such a component is "not applicable" - not an error.
            _empty = ('missing', 'blank')
            int_absent = (v_int['kind'] in _empty and v_itot['kind'] in _empty and not int_max)
            sem_absent = (v_sem['kind'] in _empty and v_stot['kind'] in _empty and v_gr['kind'] in _empty
                          and not sem_max)
            if int_absent and sem_absent:
                int_absent = sem_absent = False      # at least one component must exist
            skip_codes = set()
            if int_absent:
                int_max = 0.0
                skip_codes |= {'Internal Actual', 'Internal Total Marks'}
            if sem_absent:
                sem_max = 0.0
                skip_codes |= {'Sem Actual', 'Semester Total Marks', 'Sem Grace'}
            for code in EXPECTED_APPRAISALS:
                if code in skip_codes:
                    continue
                lst = app_map.get(code, [])
                A.check('DI-06', len(lst) >= 1, rows=m_xl[:5], field=code, expected='1 entry', actual='missing',
                        cols=['Composite Score'] if code != 'Overall Grade' else ['Overall Grade'],
                        message=f"Appraisal entry '{code}' is missing for this subject.", **ctx)
                if len(lst) > 1:
                    distinct = {tuple(_sv(r.get(c, '')) for c in content_cols) for r in lst}
                    A.check('DI-05', len(distinct) <= 1, rows=[r['_xl'] for r in lst], field=code,
                            expected='1 entry', actual=f'{len(lst)} entries, {len(distinct)} different',
                            cols=['Composite Score'],
                            message=f"Appraisal '{code}' appears {len(lst)} times with conflicting values "
                                    f"(Excel rows {', '.join(str(r['_xl']) for r in lst)}).", **ctx)


            # ---- value validity: blank / invalid / unknown code ------------
            vals = {'Internal Actual': v_int, 'Sem Actual': v_sem, 'Composite Score': v_comp,
                    'Sem Grace': v_gr, 'Internal Total Marks': v_itot, 'Semester Total Marks': v_stot,
                    'Grace on Total marks': v_gtot}
            for code, v in vals.items():
                if v['kind'] == 'missing':
                    continue
                if code in MANDATORY_VALUE_APPRAISALS and code not in skip_codes:
                    A.check('DI-10', v['kind'] != 'blank', rows=xl(code), field=code, expected='value',
                            actual='blank', cols=[col_of[code]],
                            message=f"'{code}' is blank - incomplete subject record.", **ctx)
                A.check('DI-11', v['kind'] not in ('invalid',), rows=xl(code), field=code,
                        expected='number / AB / UFM', actual=v['raw'], cols=[col_of[code]],
                        message=f"'{code}' contains a non-numeric value '{v['raw']}'.", **ctx)
                A.check('BR-05', v['kind'] != 'unknown_code', rows=xl(code), field=code,
                        expected='Z1 / Z3', actual=v['raw'], cols=[col_of[code]],
                        message=f"Unrecognised special result code '{v['raw']}' in '{code}'.", **ctx)

            # ---- grade / status / credit -----------------------------------
            g_row = first('Overall Grade')
            grade_raw = _sv(g_row.get('Marks after Gracing', '')) if g_row else ''
            status_raw = _sv(g_row.get('Completion Status after Gracing', '')) if g_row else ''
            credit_raw = _sv(g_row.get('Credit', '')) if g_row else ''
            grade = grade_raw.upper()
            status_l = status_raw.lower()
            if g_row is not None:
                A.check('DI-10', grade != '', rows=xl('Overall Grade'), field='Overall Grade', expected='grade',
                        actual='blank', cols=['Overall Grade'],
                        message="'Overall Grade' is blank - incomplete subject record.", **ctx)
                A.check('DI-10', status_raw != '', rows=xl('Overall Grade'), field='Completion Status',
                        expected='status', actual='blank', cols=['Completion Status after Gracing'],
                        message="'Completion Status after Gracing' is blank.", **ctx)
                cv = parse_audit_value(credit_raw)
                A.check('DI-10', cv['kind'] != 'blank', rows=xl('Overall Grade'), field='Credit', expected='value',
                        actual='blank', cols=['Credit'], message="Credit is blank for the graded subject.", **ctx)
                if cv['kind'] == 'num':
                    A.check('RG-01', cv['num'] >= 0, rows=xl('Overall Grade'), field='Credit', expected='>= 0',
                            actual=credit_raw, cols=['Credit'], message=f"Negative credit ({credit_raw}).", **ctx)
                    A.check('DI-14', cv['num'] > 0, rows=xl('Overall Grade'), field='Credit', expected='> 0',
                            actual=credit_raw, cols=['Credit'], message="Credit is zero for a graded subject.", **ctx)
                elif cv['kind'] in ('invalid', 'code', 'unknown_code'):
                    A.check('DI-11', False, rows=xl('Overall Grade'), field='Credit', expected='number',
                            actual=credit_raw, cols=['Credit'], message=f"Credit '{credit_raw}' is not numeric.", **ctx)
                if m_code:
                    module_ref[m_code]['desc'][m_desc] += 1
                    if cv['kind'] == 'num':
                        module_ref[m_code]['credit'][cv['num']] += 1

            # ---- maximum marks ---------------------------------------------
            for lbl, mx, code in (('Internal Actual (Max)', int_max, 'Internal Actual'),
                                  ('Sem Actual (Max)', sem_max, 'Sem Actual'),
                                  ('Composite Score (Max)', comp_max, 'Composite Score')):
                if first(code) is not None and not ((code == 'Internal Actual' and (int_absent or int_nabs)) or (code == 'Sem Actual' and (sem_absent or sem_nabs))):
                    scale_txt = _sv(first(code).get('Scale', ''))
                    A.check('RG-03', mx is not None and mx > 0, rows=xl(code), field=lbl, expected='> 0',
                            actual=scale_txt or '(blank)', cols=[lbl],
                            message=f"{lbl} is invalid (scale '{scale_txt or 'blank'}').", **ctx)
            # NABS scale / Z1 consistency
            for code in ('Internal Actual', 'Internal Total Marks', 'Sem Actual', 'Semester Total Marks'):
                r_ = first(code)
                if r_ is None:
                    continue
                sc = _sv(r_.get('Scale', '')).upper()
                v_ = vals[code]
                is_ab = v_['kind'] == 'code' and v_['code'] == 'AB'
                if sc == 'NABS':
                    A.check('BR-08', is_ab, rows=xl(code), field=code, expected='Z1 (absent)', actual=v_['raw'] or '(blank)',
                            cols=[col_of[code]],
                            message=f"'{code}' has scale NABS (absent) but marks are '{v_['raw'] or 'blank'}' - NABS requires Z1.", **ctx)
                if is_ab:
                    A.check('BR-09', sc == 'NABS', rows=xl(code), field=code, expected='scale NABS',
                            actual=sc or '(blank)', cols=[col_of[code]],
                            message=f"'{code}' is Z1 (absent) but its scale is '{sc or 'blank'}' instead of NABS.", **ctx)
            for code in ('Composite Score', 'Sem Grace', 'Grace on Total marks'):
                r_ = first(code)
                if r_ is not None and _sv(r_.get('Scale', '')).upper() == 'NABS':
                    A.check('BR-08', False, rows=xl(code), field=code, expected='numeric scale', actual='NABS',
                            cols=[col_of[code]], message=f"Scale NABS is not valid for '{code}'.", **ctx)
            valid_max = {'int': int_absent or (int_max is not None and int_max > 0),
                         'sem': sem_absent or (sem_max is not None and sem_max > 0),
                         'comp': comp_max is not None and comp_max > 0}
            module_obs[m_code].append((disp_stu, name, m_desc, (int_max, sem_max, comp_max), m_xl))

            # ---- numeric checks --------------------------------------------
            n = lambda v: v['num'] if v['kind'] in ('num', 'code') else 0.0           # noqa: E731
            i_num, s_num_v, g_num, c_num = n(v_int), n(v_sem), n(v_gr), n(v_comp)

            for code, v in vals.items():
                if v['kind'] == 'num':
                    rule = 'RG-02' if code == 'Sem Grace' else 'RG-01'
                    A.check(rule, v['num'] >= 0, rows=xl(code), field=code, expected='>= 0', actual=_fmt(v['num']),
                            cols=[col_of[code]], message=f"Negative {'grace marks' if rule == 'RG-02' else 'marks'} "
                                                         f"in '{code}' ({_fmt(v['num'])}).", **ctx)

            if v_int['kind'] in ('num', 'code') and valid_max['int']:
                A.check('RG-05', i_num <= int_max + MARK_TOLERANCE, rows=xl('Internal Actual'),
                        field='Internal Actual', expected=f'<= {_fmt(int_max)}', actual=_fmt(i_num),
                        cols=['Internal Actual', 'Internal Actual (Max)'],
                        message=f"Internal Actual {_fmt(i_num)} exceeds Internal Actual (Max) {_fmt(int_max)}.", **ctx)
            if v_sem['kind'] in ('num', 'code') and valid_max['sem']:
                A.check('RG-06', s_num_v <= sem_max + MARK_TOLERANCE, rows=xl('Sem Actual'),
                        field='Sem Actual', expected=f'<= {_fmt(sem_max)}', actual=_fmt(s_num_v),
                        cols=['Sem Actual', 'Sem Actual (Max)'],
                        message=f"Sem Actual {_fmt(s_num_v)} exceeds Sem Actual (Max) {_fmt(sem_max)}.", **ctx)
                if v_gr['kind'] == 'num':
                    A.check('RG-08', s_num_v + g_num <= sem_max + MARK_TOLERANCE, rows=xl('Sem Grace'),
                            field='Sem Actual + Sem Grace', expected=f'<= {_fmt(sem_max)}',
                            actual=_fmt(s_num_v + g_num), cols=['Sem Actual', 'Sem Grace'],
                            message=f"Sem Actual {_fmt(s_num_v)} + Sem Grace {_fmt(g_num)} = "
                                    f"{_fmt(s_num_v + g_num)} exceeds Sem Actual (Max) {_fmt(sem_max)}.", **ctx)
            if v_comp['kind'] in ('num', 'code') and valid_max['comp']:
                A.check('RG-07', c_num <= comp_max + MARK_TOLERANCE, rows=xl('Composite Score'),
                        field='Composite Score', expected=f'<= {_fmt(comp_max)}', actual=_fmt(c_num),
                        cols=['Composite Score', 'Composite Score (Max)'],
                        message=f"Composite Score {_fmt(c_num)} exceeds Composite Score (Max) {_fmt(comp_max)}.", **ctx)

            rec_comp_max = (int_max + sem_max) if (valid_max['int'] and valid_max['sem']) else None
            if rec_comp_max is not None and valid_max['comp']:
                A.check('CA-01', abs(comp_max - rec_comp_max) < MARK_TOLERANCE, rows=xl('Composite Score'),
                        field='Composite Score (Max)', expected=_fmt(rec_comp_max), actual=_fmt(comp_max),
                        cols=['Composite Score (Max)', 'Internal Actual (Max)', 'Sem Actual (Max)'],
                        message=f"Composite Score (Max) {_fmt(comp_max)} != Internal (Max) {_fmt(int_max)} + "
                                f"Sem (Max) {_fmt(sem_max)} = {_fmt(rec_comp_max)}.", **ctx)

            # components all usable?
            comps_ok = all(v['kind'] in ('num', 'code', 'blank', 'missing') for v in (v_int, v_sem, v_gr)) \
                and (int_absent or v_int['kind'] not in _empty) and (sem_absent or v_sem['kind'] not in _empty)
            rec_comp = (i_num + s_num_v + g_num) if comps_ok else None
            if rec_comp is not None and v_comp['kind'] in ('num', 'code'):
                A.check('CA-02', abs(rec_comp - c_num) < MARK_TOLERANCE, rows=xl('Composite Score'),
                        field='Composite Score', expected=_fmt(rec_comp), actual=_fmt(c_num),
                        cols=['Composite Score'],
                        message=f"Composite Score {_fmt(c_num)} != Internal {_fmt(i_num)} + Sem {_fmt(s_num_v)} + "
                                f"Grace {_fmt(g_num)} = {_fmt(rec_comp)}.", **ctx)
            if v_int['kind'] == 'code' or v_itot['kind'] == 'code':
                A.check('CA-03', v_int['kind'] == v_itot['kind'] == 'code' and v_int['code'] == v_itot['code'],
                        rows=xl('Internal Total Marks'), field='Internal Total Marks',
                        expected=v_int['raw'] or '(blank)', actual=v_itot['raw'] or '(blank)', cols=['Internal Actual'],
                        message=f"Internal Total Marks '{v_itot['raw'] or 'blank'}' does not match Internal Actual "
                                f"'{v_int['raw'] or 'blank'}' (absent status must be the same).", **ctx)
            elif v_itot['kind'] == 'num' and v_int['kind'] in ('num', 'code'):
                A.check('CA-03', abs(v_itot['num'] - i_num) < MARK_TOLERANCE, rows=xl('Internal Total Marks'),
                        field='Internal Total Marks', expected=_fmt(i_num), actual=_fmt(v_itot['num']),
                        cols=['Internal Actual'],
                        message=f"Internal Total Marks {_fmt(v_itot['num'])} != Internal Actual {_fmt(i_num)}.", **ctx)
            if v_sem['kind'] == 'code' or v_stot['kind'] == 'code':
                A.check('CA-04', v_sem['kind'] == v_stot['kind'] == 'code' and v_sem['code'] == v_stot['code'],
                        rows=xl('Semester Total Marks'), field='Semester Total Marks',
                        expected=v_sem['raw'] or '(blank)', actual=v_stot['raw'] or '(blank)', cols=['Sem Actual'],
                        message=f"Semester Total Marks '{v_stot['raw'] or 'blank'}' does not match Sem Actual "
                                f"'{v_sem['raw'] or 'blank'}' (absent status must be the same).", **ctx)
            elif v_stot['kind'] == 'num' and v_sem['kind'] in ('num', 'code') and v_gr['kind'] in ('num', 'blank', 'code'):
                A.check('CA-04', abs(v_stot['num'] - (s_num_v + g_num)) < MARK_TOLERANCE,
                        rows=xl('Semester Total Marks'), field='Semester Total Marks',
                        expected=_fmt(s_num_v + g_num), actual=_fmt(v_stot['num']),
                        cols=['Sem Actual', 'Sem Grace'],
                        message=f"Semester Total Marks {_fmt(v_stot['num'])} != Sem Actual {_fmt(s_num_v)} + "
                                f"Sem Grace {_fmt(g_num)} = {_fmt(s_num_v + g_num)}.", **ctx)

            # ---- grade recalculation ---------------------------------------
            codes = {v['code'] for v in (v_int, v_sem, v_gr, v_comp) if v['kind'] in ('code', 'unknown_code')}
            forced_fail = SPECIAL_CODE_FORCES_FAIL and any(c in ('AB', 'UFM') for c in codes)
            exp_rep = exp_rec = None
            pct_rep = pct_rec = None
            if v_comp['kind'] in ('num', 'code') and valid_max['comp']:
                pct_rep = c_num * 100.0 / comp_max
                exp_rep = grade_for_percentage(pct_rep)
            if rec_comp is not None and rec_comp_max:
                pct_rec = rec_comp * 100.0 / rec_comp_max
                exp_rec = grade_for_percentage(pct_rec)

            # 40% pass mark applies to EACH component. Sem % always uses Sem Actual + Sem Grace.
            int_pct = sem_pct = None
            reasons, fail_cols = [], []
            if v_int['kind'] == 'num' and not int_absent and int_max and int_max > 0:
                int_pct = i_num * 100.0 / int_max
                if int_pct < 40 - 1e-9:
                    reasons.append(f"Internal {_fmt(i_num)}/{_fmt(int_max)} = {int_pct:.2f}% (< 40%)")
                    fail_cols.append('Internal Actual')
            if v_sem['kind'] == 'num' and not sem_absent and sem_max and sem_max > 0:
                sem_pct = (s_num_v + g_num) * 100.0 / sem_max
                if sem_pct < 40 - 1e-9:
                    reasons.append(f"Sem (Actual {_fmt(s_num_v)} + Grace {_fmt(g_num)} = {_fmt(s_num_v + g_num)})"
                                   f"/{_fmt(sem_max)} = {sem_pct:.2f}% (< 40%)")
                    fail_cols.append('Sem Actual')
            comp_fail = bool(reasons)
            if comp_fail:
                exp_rep = exp_rec = 'F'

            if grade:
                A.check('CA-07', grade in VALID_GRADES or (forced_fail and grade in KNOWN_SPECIAL),
                        rows=xl('Overall Grade'), field='Overall Grade', expected=', '.join(sorted(VALID_GRADES)),
                        actual=grade_raw, cols=['Overall Grade'], message=f"Invalid Overall Grade '{grade_raw}'.", **ctx)

            if forced_fail:
                acceptable = {'F'} | KNOWN_SPECIAL
                if grade:
                    A.check('BR-04', grade in acceptable, rows=xl('Overall Grade'), field='Overall Grade',
                            expected='F (candidate AB/UFM)', actual=grade_raw, cols=['Overall Grade'],
                            message=f"Component carries {'/'.join(sorted(codes))} but Overall Grade is "
                                    f"'{grade_raw}' (expected F).", **ctx)
                if status_raw:
                    A.check('BR-04', status_l in STATUS_FAIL, rows=xl('Overall Grade'), field='Completion Status',
                            expected='Completed Unsuccessfuly', actual=status_raw,
                            cols=['Completion Status after Gracing'],
                            message=f"Component carries {'/'.join(sorted(codes))} but status is '{status_raw}'.", **ctx)
                if v_gr['kind'] == 'num' and g_num > 0:
                    A.check('BR-07', False, rows=xl('Sem Grace'), field='Sem Grace', expected='0',
                            actual=_fmt(g_num), cols=['Sem Grace'],
                            message=f"Grace {_fmt(g_num)} awarded to a candidate marked {'/'.join(sorted(codes))}.", **ctx)
            elif grade in VALID_GRADES and comp_fail:
                A.check('BR-10', grade == 'F', rows=xl('Overall Grade'), field='Overall Grade', expected='F',
                        actual=grade_raw, cols=['Overall Grade'] + fail_cols,
                        message=f"Component below the 40% pass mark: {'; '.join(reasons)} - subject must be F, "
                                f"but Overall Grade is '{grade_raw}'.", **ctx)
            elif grade in VALID_GRADES:
                if int_pct is not None or sem_pct is not None:
                    A.check('BR-10', True)
                if exp_rep is not None:
                    A.check('CA-05', grade == exp_rep, rows=xl('Overall Grade'), field='Overall Grade',
                            expected=exp_rep, actual=grade_raw, cols=['Overall Grade'],
                            message=f"Overall Grade '{grade_raw}' != '{exp_rep}' expected from Composite "
                                    f"{_fmt(c_num)}/{_fmt(comp_max)} = {pct_rep:.2f}%.", **ctx)
                if exp_rec is not None and exp_rep is not None and exp_rec != exp_rep:
                    A.check('CA-06', grade == exp_rec, rows=xl('Overall Grade'), field='Overall Grade',
                            expected=exp_rec, actual=grade_raw, cols=['Overall Grade', 'Composite Score'],
                            message=f"Using the recalculated Composite {_fmt(rec_comp)}/{_fmt(rec_comp_max)} = "
                                    f"{pct_rec:.2f}% the grade should be '{exp_rec}', not '{grade_raw}'.", **ctx)
                elif exp_rec is not None and exp_rep is None:
                    A.check('CA-06', grade == exp_rec, rows=xl('Overall Grade'), field='Overall Grade',
                            expected=exp_rec, actual=grade_raw, cols=['Overall Grade'],
                            message=f"Recalculated Composite {_fmt(rec_comp)}/{_fmt(rec_comp_max)} = "
                                    f"{pct_rec:.2f}% gives grade '{exp_rec}', not '{grade_raw}'.", **ctx)

            if grade:
                is_fail_grade = grade == 'F' or grade in KNOWN_SPECIAL
                if status_raw:
                    known_status = status_l == STATUS_SUCCESS or status_l in STATUS_FAIL
                    A.check('BR-03', known_status, rows=xl('Overall Grade'), field='Completion Status',
                            expected='Completed with Success / Completed Unsuccessfuly', actual=status_raw,
                            cols=['Completion Status after Gracing'],
                            message=f"Unrecognised Completion Status '{status_raw}'.", **ctx)
                    if known_status and not forced_fail:
                        if is_fail_grade:
                            A.check('BR-02', status_l in STATUS_FAIL, rows=xl('Overall Grade'),
                                    field='Completion Status', expected='Completed Unsuccessfuly', actual=status_raw,
                                    cols=['Overall Grade', 'Completion Status after Gracing'],
                                    message=f"Grade is F but status is '{status_raw}'.", **ctx)
                        elif grade in VALID_GRADES:
                            A.check('BR-01', status_l == STATUS_SUCCESS, rows=xl('Overall Grade'),
                                    field='Completion Status', expected='Completed with Success', actual=status_raw,
                                    cols=['Overall Grade', 'Completion Status after Gracing'],
                                    message=f"Grade is '{grade_raw}' (not F) but status is '{status_raw}' - "
                                            f"must be 'Completed with Success'.", **ctx)

            # ---- before / after gracing special-code & marks consistency ----
            for code in ('Internal Actual', 'Sem Actual'):
                b, a = before(code), after(code)
                if b['kind'] == 'code':
                    A.check('BR-06', a['kind'] == 'code', rows=xl(code), field=code, expected=b['code'],
                            actual=a['raw'] or '(blank)', cols=[col_of[code]],
                            message=f"'{code}' was {b['code']} before gracing but is '{a['raw'] or 'blank'}' after.", **ctx)
                if b['kind'] == 'num' and a['kind'] == 'num':
                    changed = abs(b['num'] - a['num']) >= MARK_TOLERANCE
                    A.check('GR-06', not changed, rows=xl(code), field=code, expected=_fmt(b['num']),
                            actual=_fmt(a['num']), cols=[col_of[code]],
                            message=f"'{code}' changed from {_fmt(b['num'])} (before) to {_fmt(a['num'])} "
                                    f"(after gracing) - actual marks should not change; grace belongs in Sem Grace.", **ctx)

            # ---- gracing sanity --------------------------------------------
            if v_gtot['kind'] == 'num':
                A.check('GR-05', v_gtot['num'] == 0, rows=xl('Grace on Total marks'), field='Grace on Total marks',
                        expected='0 (not part of composite formula)', actual=_fmt(v_gtot['num']),
                        cols=['Sem Grace'],
                        message=f"'Grace on Total marks' = {_fmt(v_gtot['num'])} is not part of the Composite formula "
                                f"(Internal + Sem + Sem Grace) - confirm it is correctly reflected.", **ctx)
            if v_gr['kind'] == 'num' and g_num > 0:
                s_total_grace += g_num
                s_graced_subjects += 1
                lim_pct = (sem_max * MAX_GRACE_PCT_OF_SEM / 100.0) if valid_max['sem'] else MAX_GRACE_ABS
                limit = min(MAX_GRACE_ABS, lim_pct)
                A.check('GR-01', g_num <= limit + 1e-9, rows=xl('Sem Grace'), field='Sem Grace',
                        expected=f'<= {_fmt(limit)}', actual=_fmt(g_num), cols=['Sem Grace'],
                        message=f"Abnormal grace of {_fmt(g_num)} marks (normal limit {_fmt(limit)}: "
                                f"{MAX_GRACE_ABS} marks / {MAX_GRACE_PCT_OF_SEM:g}% of Sem Max).", **ctx)
                pre_sem = (s_num_v * 100.0 / sem_max) if (valid_max['sem'] and sem_max) else None
                pre_comp = ((i_num + s_num_v) * 100.0 / comp_max) if (valid_max['comp'] and comp_max) else None
                if pre_sem is not None or pre_comp is not None:
                    needed = (pre_sem is not None and pre_sem < 40 - 1e-9) or (pre_comp is not None and pre_comp < 40 - 1e-9)
                    txt = []
                    if pre_sem is not None:
                        txt.append(f"Sem {_fmt(s_num_v)}/{_fmt(sem_max)} = {pre_sem:.2f}%")
                    if pre_comp is not None:
                        txt.append(f"Composite {pre_comp:.2f}%")
                    A.check('GR-03', needed, rows=xl('Sem Grace'), field='Sem Grace',
                            expected='grace only when Sem or subject < 40% before grace',
                            actual=' / '.join(txt) + ' before grace', cols=['Sem Grace'],
                            message=f"Grace {_fmt(g_num)} awarded although the subject already clears 40% in Sem and "
                                    f"overall without it ({' / '.join(txt)}).", **ctx)
                post_grade = exp_rec if exp_rec is not None else exp_rep
                if post_grade is not None and not forced_fail:
                    post_txt = '; '.join(reasons) if reasons else (
                        f"{(pct_rec if pct_rec is not None else pct_rep):.2f}% after grace")
                    A.check('GR-04', post_grade != 'F', rows=xl('Sem Grace'), field='Sem Grace',
                            expected='subject clears pass mark', actual='still F',
                            cols=['Sem Grace', 'Overall Grade'],
                            message=f"Grace {_fmt(g_num)} awarded but subject still fails ({post_txt}).", **ctx)

            # ---- accumulate student-level numbers & trail -------------------
            if v_comp['kind'] in ('num', 'code'):
                s_sum_comp += c_num
            else:
                s_sum_valid = False
            s_sum_comp_rec += rec_comp if rec_comp is not None else c_num
            if valid_max['comp']:
                s_sum_max += comp_max

            A.trail.append({
                'Student Number': disp_stu, 'Student Name': name, 'Module Code': m_code, 'Module Desc': m_desc,
                'Internal Actual': v_int['raw'], 'Internal Max': int_max, 'Sem Actual': v_sem['raw'],
                'Sem Grace': v_gr['raw'], 'Sem Max': sem_max,
                'Reported Composite': v_comp['raw'], 'Recalculated Composite': rec_comp,
                'Reported Composite Max': comp_max, 'Recalculated Composite Max': rec_comp_max,
                'Recalculated %': round(pct_rec, 2) if pct_rec is not None else None,
                'Internal %': round(int_pct, 2) if int_pct is not None else None,
                'Sem % (Actual+Grace)': round(sem_pct, 2) if sem_pct is not None else None,
                'Reported Grade': grade_raw,
                'Recalculated Grade': ('F' if forced_fail else (exp_rec or exp_rep or '')),
                'Reported Status': status_raw,
                'Expected Status': ('' if not (exp_rec or exp_rep or forced_fail) else
                                    ('Completed Unsuccessfuly' if (forced_fail or (exp_rec or exp_rep) == 'F')
                                     else 'Completed with Success')),
            })

        # ---- student-level totals --------------------------------------------
        s_pct_exp = (s_sum_comp * 100.0 / s_sum_max) if s_sum_max > 0 else None
        rep_total = parse_audit_value(tot_rows[0].get('Marks after Gracing', '')) if tot_rows else None
        rep_pct = parse_audit_value(pct_rows[0].get('Marks after Gracing', '')) if pct_rows else None
        tot_xl = [tot_rows[0]['_xl']] if tot_rows else xl_rows[:2]
        pct_xl = [pct_rows[0]['_xl']] if pct_rows else xl_rows[:2]
        sctx = dict(student=disp_stu, name=name, target='summary')

        if rep_total is not None:
            if rep_total['kind'] != 'num':
                A.check('DI-11', False, rows=tot_xl, field='Total',
                        expected='number', actual=rep_total['raw'] or '(blank)', cols=['Composite Score'],
                        message=f"TOTAL value '{rep_total['raw'] or 'blank'}' is not numeric.", **sctx)
            else:
                A.check('ST-01', abs(rep_total['num'] - s_sum_comp) < MARK_TOLERANCE, rows=tot_xl, field='Total',
                        expected=_fmt(s_sum_comp), actual=_fmt(rep_total['num']), cols=['Composite Score'],
                        message=f"Total {_fmt(rep_total['num'])} != sum of subject Composite Scores "
                                f"{_fmt(s_sum_comp)}.", **sctx)
                if abs(s_sum_comp_rec - s_sum_comp) >= MARK_TOLERANCE:
                    A.check('ST-04', abs(rep_total['num'] - s_sum_comp_rec) < MARK_TOLERANCE, rows=tot_xl,
                            field='Total', expected=_fmt(s_sum_comp_rec), actual=_fmt(rep_total['num']),
                            cols=['Composite Score'],
                            message=f"Total {_fmt(rep_total['num'])} != sum of recalculated subject scores "
                                    f"{_fmt(s_sum_comp_rec)} (Internal + Sem + Grace).", **sctx)
        if rep_pct is not None:
            if rep_pct['kind'] != 'num':
                A.check('DI-11', False, rows=pct_xl, field='Percentage', expected='number',
                        actual=rep_pct['raw'] or '(blank)', cols=['Overall Grade'],
                        message=f"PERCENTAGE value '{rep_pct['raw'] or 'blank'}' is not numeric.", **sctx)
            else:
                A.check('ST-03', 0 <= rep_pct['num'] <= 100, rows=pct_xl, field='Percentage', expected='0 - 100',
                        actual=_fmt(rep_pct['num']), cols=['Overall Grade'],
                        message=f"Percentage {_fmt(rep_pct['num'])} is outside 0 - 100.", **sctx)
                if s_pct_exp is not None:
                    A.check('ST-02', abs(rep_pct['num'] - s_pct_exp) <= PERCENT_TOLERANCE, rows=pct_xl,
                            field='Percentage', expected=f"{s_pct_exp:.2f}", actual=_fmt(rep_pct['num']),
                            cols=['Overall Grade'],
                            message=f"Percentage {_fmt(rep_pct['num'])} != Total {_fmt(s_sum_comp)} / Max "
                                    f"{_fmt(s_sum_max)} x 100 = {s_pct_exp:.2f}.", **sctx)

        # ---- student-level grace ---------------------------------------------
        if s_graced_subjects:
            A.check('GR-02', s_total_grace <= MAX_GRACE_PER_STUDENT and s_graced_subjects <= MAX_GRACED_SUBJECTS,
                    student=disp_stu, name=name, field='Sem Grace (all subjects)', rows=xl_rows[:3],
                    expected=f"<= {MAX_GRACE_PER_STUDENT} marks in <= {MAX_GRACED_SUBJECTS} subjects",
                    actual=f"{_fmt(s_total_grace)} marks in {s_graced_subjects} subjects",
                    cols=['Module Code'], target='header',
                    message=f"Total grace {_fmt(s_total_grace)} marks across {s_graced_subjects} subject(s) is "
                            f"above normal limits ({MAX_GRACE_PER_STUDENT} marks / {MAX_GRACED_SUBJECTS} subjects).")

        e_cnt, w_cnt = A.student_counts(disp_stu)
        A.student_rows.append({
            'Student Number': disp_stu, 'Student Name': name, 'Subjects': len(by_mod),
            'Reported Total': rep_total['raw'] if rep_total else '', 'Recalculated Total': round(s_sum_comp, 2),
            'Total Max': s_sum_max, 'Reported %': rep_pct['raw'] if rep_pct else '',
            'Recalculated %': round(s_pct_exp, 2) if s_pct_exp is not None else None,
            'Errors': e_cnt, 'Warnings': w_cnt,
            'Audit Status': 'ERRORS' if e_cnt else ('WARNINGS' if w_cnt else 'CLEAN'),
        })

    # ---------- 3. cohort-level checks --------------------------------------
    for m_code, ref in module_ref.items():
        if len(ref['desc']) > 1:
            A.check('DI-09', False, module=m_code, desc=' / '.join(ref['desc']), field='Module Desc',
                    expected=ref['desc'].most_common(1)[0][0], actual=' / '.join(ref['desc']),
                    cols=['Module Desc'],
                    message=f"Module {m_code} has {len(ref['desc'])} different descriptions across students: "
                            f"{' / '.join(ref['desc'])}.")
        else:
            A.check('DI-09', True)
        if len(ref['credit']) > 1:
            A.check('DI-09', False, module=m_code, field='Credit', expected=_fmt(ref['credit'].most_common(1)[0][0]),
                    actual=' / '.join(_fmt(c) for c in ref['credit']), cols=['Credit'],
                    message=f"Module {m_code} has different Credit values across students: "
                            f"{' / '.join(_fmt(c) for c in ref['credit'])}.")
        else:
            A.check('DI-09', True)

    for m_code, obs in module_obs.items():
        if not m_code:
            continue
        counts = Counter(o[3] for o in obs if None not in o[3])
        if not counts:
            continue
        mode, mode_n = counts.most_common(1)[0]
        for stu, name, desc, mx, xl_rows in obs:
            if None in mx:
                continue
            A.check('RG-04', mx == mode, student=stu, name=name, module=m_code, desc=desc, rows=xl_rows[:3],
                    field='Maximum marks', expected=f"Int {_fmt(mode[0])} / Sem {_fmt(mode[1])} / Comp {_fmt(mode[2])}",
                    actual=f"Int {_fmt(mx[0])} / Sem {_fmt(mx[1])} / Comp {_fmt(mx[2])}",
                    cols=['Internal Actual (Max)', 'Sem Actual (Max)', 'Composite Score (Max)'],
                    message=f"Maximum marks (Int {_fmt(mx[0])}/Sem {_fmt(mx[1])}/Comp {_fmt(mx[2])}) differ from "
                            f"the cohort norm (Int {_fmt(mode[0])}/Sem {_fmt(mode[1])}/Comp {_fmt(mode[2])}) "
                            f"used by {mode_n} of {len(obs)} students.")

    sets = Counter(frozenset(v) for k, v in student_modules.items() if k)
    if len(student_modules) >= 3 and sets:
        norm, norm_n = sets.most_common(1)[0]
        if norm_n / sum(sets.values()) >= 0.5:
            for stu, mods in student_modules.items():
                if not stu:
                    continue
                missing, extra = sorted(norm - mods), sorted(mods - norm)
                A.check('DI-12', not missing and not extra, student=stu, name=student_names.get(stu, ''),
                        field='Subject set', expected=f"{len(norm)} subjects", actual=f"{len(mods)} subjects",
                        cols=['Module Code'], target='header',
                        message=f"Subject list differs from the cohort norm ({len(norm)} subjects): "
                                f"missing {', '.join(missing) or 'none'}; extra {', '.join(extra) or 'none'}.")

    # refresh error/warning counts after cohort checks
    for row in A.student_rows:
        key = row['Student Number']
        e, w = A.student_counts(key)
        row['Errors'], row['Warnings'] = e, w
        row['Audit Status'] = 'ERRORS' if e else ('WARNINGS' if w else 'CLEAN')

    # mark trail rows that picked up errors
    err_mods = {(f['Student Number'], f['Module Code']) for f in A.findings if f['Severity'] == ERR and f['Module Code']}
    for t in A.trail:
        t['Audit Result'] = 'DISCREPANCY' if (t['Student Number'], t['Module Code']) in err_mods else 'OK'

    A.stats = {
        'raw_records': len(records),
        'students': len([k for k in students if k]) + (1 if '' in students else 0),
        'subject_records': n_subject_records,
        'checks_run': int(sum(A.checked.values())),
        'errors': A.n_errors,
        'warnings': A.n_warnings,
        'students_error': sum(1 for r in A.student_rows if r['Audit Status'] == 'ERRORS'),
        'students_warning': sum(1 for r in A.student_rows if r['Audit Status'] == 'WARNINGS'),
        'students_clean': sum(1 for r in A.student_rows if r['Audit Status'] == 'CLEAN'),
    }
    return A


# ==================================================
# AUDIT REPORT: STYLES + EXCEL SHEETS
# ==================================================
AUDIT_ERR_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
AUDIT_ERR_FONT = Font(color="FFFF00", bold=True)
AUDIT_WARN_FILL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
AUDIT_WARN_FONT = Font(color="000000", bold=True)
AUDIT_ERR_CSS = "background-color: #FF0000; color: #FFFF00; font-weight: bold;"
AUDIT_WARN_CSS = "background-color: #FFC000; color: #000000; font-weight: bold;"

_THIN = Side(style="thin", color="000000")
_BOX = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
_TITLE_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
_SEC_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_OK_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
_OK_FONT = Font(color="006100", bold=True)


def _style_table(ws, first_row, last_row, ncols, widths=None, wrap_cols=()):
    """Header + borders + optional widths for a plain table already written in ws."""
    for c in range(1, ncols + 1):
        h = ws.cell(row=first_row, column=c)
        h.font, h.fill, h.border = Font(bold=True), _HDR_FILL, _BOX
        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(first_row + 1, last_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = _BOX
            cell.alignment = Alignment(vertical="top", wrap_text=(c in wrap_cols))
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def _sev_style(cell):
    v = str(cell.value)
    if v in ('ERROR', 'FAIL', 'ERRORS', 'DISCREPANCY'):
        cell.fill, cell.font = AUDIT_ERR_FILL, AUDIT_ERR_FONT
    elif v in ('WARNING', 'REVIEW', 'WARNINGS'):
        cell.fill, cell.font = AUDIT_WARN_FILL, AUDIT_WARN_FONT
    elif v in ('PASS', 'CLEAN', 'OK'):
        cell.fill, cell.font = _OK_FILL, _OK_FONT
    cell.alignment = Alignment(horizontal="center", vertical="top")


def write_audit_sheets(writer, audit, source_name=""):
    """
    Writes: Audit Summary (first tab), Audit Findings, and - at the end -
    Student Audit Status and Recalculation Trail. Call BEFORE and AFTER the other sheets
    via the two helpers below.
    """
    findings = audit.findings_df()
    findings.to_excel(writer, index=False, sheet_name="Audit Findings")
    ws = writer.book["Audit Findings"]
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = "A2"
    n = len(findings) + 1
    _style_table(ws, 1, n, findings.shape[1],
                 widths=[7, 11, 9, 16, 15, 28, 13, 28, 24, 14, 22, 22, 85], wrap_cols=(13, 8, 11, 12))
    for r in range(2, n + 1):
        _sev_style(ws.cell(row=r, column=2))
    ws.auto_filter.ref = f"A1:{get_column_letter(findings.shape[1])}{max(n, 2)}"

    _write_summary_sheet(writer.book, audit, findings, source_name)


def write_audit_detail_sheets(writer, audit):
    """Student Audit Status + Recalculation Trail (placed after the other sheets)."""
    stu = pd.DataFrame(audit.student_rows)
    stu.to_excel(writer, index=False, sheet_name="Student Audit Status")
    ws = writer.book["Student Audit Status"]
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = "A2"
    _style_table(ws, 1, len(stu) + 1, stu.shape[1], widths=[16, 28, 10, 14, 14, 10, 12, 14, 9, 10, 14])
    for r in range(2, len(stu) + 2):
        _sev_style(ws.cell(row=r, column=stu.shape[1]))
    ws.auto_filter.ref = f"A1:{get_column_letter(stu.shape[1])}{max(len(stu) + 1, 2)}"

    trail = pd.DataFrame(audit.trail)
    trail.to_excel(writer, index=False, sheet_name="Recalculation Trail")
    ws = writer.book["Recalculation Trail"]
    ws.views.sheetView[0].showGridLines = False
    ws.freeze_panes = "A2"
    _style_table(ws, 1, len(trail) + 1, trail.shape[1],
                 widths=[16, 26, 13, 34, 10, 9, 9, 9, 9, 12, 14, 12, 14, 12, 10, 12, 24, 24, 14])
    for r in range(2, len(trail) + 2):
        _sev_style(ws.cell(row=r, column=trail.shape[1]))
    ws.auto_filter.ref = f"A1:{get_column_letter(trail.shape[1])}{max(len(trail) + 1, 2)}"


def _write_summary_sheet(wb, audit, findings, source_name):
    ws = wb.create_sheet("Audit Summary", 0)
    ws.views.sheetView[0].showGridLines = False
    st_ = audit.stats
    widths = [14, 22, 18, 70, 14, 12, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:G1")
    ws["A1"] = "EXAMINATION RESULT - VALIDATION AUDIT REPORT"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = _TITLE_FILL
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A3"], ws["B3"] = "Source file", source_name or "(uploaded workbook)"
    ws["A4"], ws["B4"] = "Audit run on", datetime.now().strftime("%d-%m-%Y %H:%M")
    for r in (3, 4):
        ws.cell(row=r, column=1).font = Font(bold=True)

    ws.merge_cells("A6:G6")
    if st_['errors'] == 0:
        txt = "PASS - no errors found. " + (f"{st_['warnings']} warning(s) to review before publication."
                                           if st_['warnings'] else "Result data is clean.")
        ws["A6"].fill, ws["A6"].font = _OK_FILL, Font(size=13, bold=True, color="006100")
    else:
        txt = (f"NOT READY FOR PUBLICATION - {st_['errors']} error(s) and {st_['warnings']} warning(s) "
               f"must be resolved/reviewed.")
        ws["A6"].fill, ws["A6"].font = AUDIT_ERR_FILL, Font(size=13, bold=True, color="FFFF00")
    ws["A6"] = txt
    ws["A6"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[6].height = 26

    r = 8
    ws.cell(row=r, column=1, value="1. AUDIT SCOPE & RESULT").font = Font(bold=True, size=12)
    r += 1
    metrics = [
        ("Raw records read", st_['raw_records']),
        ("Students audited", st_['students']),
        ("Subject records audited", st_['subject_records']),
        ("Validation checks executed", st_['checks_run']),
        ("Errors (discrepancies)", st_['errors']),
        ("Warnings (review)", st_['warnings']),
        ("Students - clean", st_['students_clean']),
        ("Students - warnings only", st_['students_warning']),
        ("Students - with errors", st_['students_error']),
    ]
    top = r
    for k, v in metrics:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=3, value=v).alignment = Alignment(horizontal="center")
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).border = _BOX
        ws.cell(row=r, column=1).font = Font(bold=True)
        if k.startswith("Errors") and v:
            ws.cell(row=r, column=3).fill, ws.cell(row=r, column=3).font = AUDIT_ERR_FILL, AUDIT_ERR_FONT
        if k.startswith("Warnings") and v:
            ws.cell(row=r, column=3).fill, ws.cell(row=r, column=3).font = AUDIT_WARN_FILL, AUDIT_WARN_FONT
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="2. FINDINGS BY CATEGORY").font = Font(bold=True, size=12)
    r += 1
    cat_rows = []
    for cat in dict.fromkeys(v[1] for v in RULES.values()):
        e = sum(1 for f in audit.findings if f['Category'] == cat and f['Severity'] == ERR)
        w = sum(1 for f in audit.findings if f['Category'] == cat and f['Severity'] == WARN)
        cat_rows.append((cat, e, w))
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1, value="Category")
    ws.cell(row=r, column=5, value="Errors")
    ws.cell(row=r, column=6, value="Warnings")
    for c in (1, 2, 3, 4, 5, 6):
        h = ws.cell(row=r, column=c)
        h.font, h.fill, h.border = Font(bold=True), _HDR_FILL, _BOX
        h.alignment = Alignment(horizontal="center")
    r += 1
    for cat, e, w in cat_rows:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=5, value=e)
        ws.cell(row=r, column=6, value=w)
        for c in (1, 2, 3, 4, 5, 6):
            ws.cell(row=r, column=c).border = _BOX
        for c in (5, 6):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="3. RULE-WISE RESULTS (every rule is listed, including those that passed)").font = Font(bold=True, size=12)
    r += 1
    rt = audit.rule_table()
    for j, col in enumerate(rt.columns, start=1):
        ws.cell(row=r, column=j, value=col)
    hdr_row = r
    for i, rec in enumerate(rt.itertuples(index=False), start=1):
        for j, v in enumerate(rec, start=1):
            ws.cell(row=r + i, column=j, value=v)
    last = r + len(rt)
    _style_table(ws, hdr_row, last, rt.shape[1], wrap_cols=(4,))
    for rr in range(hdr_row + 1, last + 1):
        _sev_style(ws.cell(row=rr, column=7))
        ws.cell(row=rr, column=5).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=rr, column=6).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=rr, column=3).alignment = Alignment(horizontal="center", vertical="top")
    r = last + 2

    ws.cell(row=r, column=1, value="4. SUBJECT-WISE FINDINGS").font = Font(bold=True, size=12)
    r += 1
    subj = defaultdict(lambda: [0, 0, ''])
    for f in audit.findings:
        if f['Module Code']:
            s = subj[f['Module Code']]
            s[0 if f['Severity'] == ERR else 1] += 1
            s[2] = s[2] or f['Module Desc']
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.cell(row=r, column=1, value="Module Code")
    ws.cell(row=r, column=2, value="Module Description")
    ws.cell(row=r, column=5, value="Errors")
    ws.cell(row=r, column=6, value="Warnings")
    for c in (1, 2, 3, 4, 5, 6):
        h = ws.cell(row=r, column=c)
        h.font, h.fill, h.border = Font(bold=True), _HDR_FILL, _BOX
        h.alignment = Alignment(horizontal="center")
    r += 1
    if not subj:
        ws.cell(row=r, column=1, value="No subject-level findings.")
        r += 1
    for code, (e, w, d) in sorted(subj.items(), key=lambda kv: -(kv[1][0] * 1000 + kv[1][1])):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=d)
        ws.cell(row=r, column=5, value=e)
        ws.cell(row=r, column=6, value=w)
        for c in (1, 2, 3, 4, 5, 6):
            ws.cell(row=r, column=c).border = _BOX
        for c in (5, 6):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Highlight legend:").font = Font(bold=True)
    c = ws.cell(row=r, column=2, value="Error / discrepancy")
    c.fill, c.font = AUDIT_ERR_FILL, AUDIT_ERR_FONT
    c = ws.cell(row=r, column=3, value="Warning")
    c.fill, c.font = AUDIT_WARN_FILL, AUDIT_WARN_FONT
    r += 1
    ws.cell(row=r, column=1, value="Grade scale used:").font = Font(bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="O 90-100 | A+ 80-<90 | A 70-<80 | B+ 60-<70 | B 55-<60 | C 50-<55 | P 40-<50 | F <40 "
                                   "(percentage = Composite Score / Composite Score (Max) x 100, not rounded)")


def audit_only_excel_bytes(audit, source_name=""):
    """Standalone workbook containing only the audit sheets."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        write_audit_sheets(writer, audit, source_name)
        write_audit_detail_sheets(writer, audit)
    return buffer.getvalue()


# ==================================================
# STEP 1: PROCESSED DATA GENERATION
# ==================================================
def _resolve_max(audit, m_code, code, scale_raw):
    """(max number, max display) for a scale code. Scale NABS (absent) is resolved from the module reference."""
    if str(scale_raw).strip().upper() == 'NABS' and audit is not None:
        ref = getattr(audit, 'ref_max', {}).get((str(m_code).strip(), code))
        if ref:
            return float(ref), str(int(ref))
    return parse_scale_num(scale_raw), parse_scale_str(scale_raw)


def create_processed_data(df, target_student_ids=None, audit=None):
    """
    Restructures raw examination records into a clean student-wise dataset:
    - Merged Student Info Row across all columns: "Student Number | Name | Add ID | Gender"
    - Replaces Z1 with "AB" and Z3 with "UFM"
    - Checks 40% minimum passing criteria on max marks
    - Formats numbers, totals, and percentages
    - If target_student_ids is provided, filters strictly for those students.
    - If an AuditResult is supplied, each row carries its audit flags (cells to highlight)
      and an 'Audit Remarks' text.
    """
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    if target_student_ids is not None:
        target_set = set(str(sid).strip() for sid in target_student_ids)
        student_ids = [sid for sid in df['Student Number'].dropna().unique() if str(sid).strip() in target_set]
    else:
        student_ids = df['Student Number'].dropna().unique()

    all_processed_rows = []

    def _emit(row, s_key, m_key=None):
        """Appends a processed row and attaches its audit flags / remarks."""
        target = {'HEADER': 'header', 'SUMMARY': 'summary'}.get(row['Row_Type'], 'module')
        info = audit.cell_flags(s_key, m_key, target) if audit is not None else {'cols': {}, 'remarks': []}
        remarks = list(info['remarks'])
        if row['Row_Type'] == 'HEADER' and audit is not None:
            e_cnt, w_cnt = audit.student_counts(s_key)
            if e_cnt or w_cnt:
                remarks.insert(0, f"AUDIT: {e_cnt} error(s), {w_cnt} warning(s) in this student's record.")
        row['Audit_Cells'] = dict(info['cols'])
        row['Audit Remarks'] = "\n".join(remarks)
        all_processed_rows.append(row)

    for s_num in student_ids:
        s_df = df[df['Student Number'] == s_num]

        # Non-summary module rows
        module_df = s_df[
            s_df['Module Desc'].notna() & 
            (~s_df['Student Name'].astype(str).str.upper().isin(['TOTAL', 'PERCENTAGE']))
        ]

        # Summary rows
        total_row = s_df[s_df['Student Name'].astype(str).str.upper() == 'TOTAL']
        perc_row = s_df[s_df['Student Name'].astype(str).str.upper() == 'PERCENTAGE']

        std_name = module_df.iloc[0]['Student Name'] if not module_df.empty else (s_df.iloc[0]['Student Name'] if not s_df.empty else '')
        add_id = module_df.iloc[0]['Additional ID'] if not module_df.empty else (s_df.iloc[0]['Additional ID'] if not s_df.empty else '')
        gender = module_df.iloc[0]['Gender'] if not module_df.empty else (s_df.iloc[0]['Gender'] if not s_df.empty else '')

        # 1. Clean Merged Student Header Row
        student_header_str = f"{s_num} | {std_name} | {add_id} | {gender}"
        _emit({
            'Module Code': student_header_str,
            'Module Desc': '',
            'Credit': '',
            'Internal Actual (Max)': '',
            'Internal Actual': '',
            'Sem Actual (Max)': '',
            'Sem Actual': '',
            'Sem Grace': '',
            'Composite Score (Max)': '',
            'Composite Score': '',
            'Overall Grade': '',
            'Completion Status after Gracing': '',
            'Remarks & Scale': '',
            'Row_Type': 'HEADER',
            'Comp_Mismatch': False,
            'Red_Internal': False,
            'Red_Sem': False,
            'Red_Grace': False,
            'Red_Composite': False
        }, s_num)

        modules = module_df['Module Code'].dropna().unique()
        calc_total_comp = 0.0
        calc_total_max_comp = 0.0

        for m_code in modules:
            m_sub = module_df[module_df['Module Code'] == m_code]
            m_desc = m_sub.iloc[0]['Module Desc']

            # Credit parsing
            credit_row = m_sub[m_sub['Appraisal Type Code'].astype(str).str.strip() == 'Overall Grade']
            credit_raw = credit_row.iloc[0]['Credit'] if not credit_row.empty else m_sub.iloc[0]['Credit']
            credit_val, _, _ = parse_mark(credit_raw)

            # Helper for extracting appraisal type values and scale max marks
            def get_val_and_scale(code):
                sub = m_sub[m_sub['Appraisal Type Code'].astype(str).str.strip() == code]
                if not sub.empty:
                    val = sub['Marks after Gracing'].iloc[0]
                    scale = sub['Scale'].iloc[0] if 'Scale' in sub.columns else ''
                    return val, scale
                return '', ''

            int_raw, int_scale_raw = get_val_and_scale('Internal Actual')
            sem_raw, sem_scale_raw = get_val_and_scale('Sem Actual')
            grace_raw, _ = get_val_and_scale('Sem Grace')
            comp_raw, comp_scale_raw = get_val_and_scale('Composite Score')
            overall_grade, _ = get_val_and_scale('Overall Grade')

            int_display, i_num, int_is_abs_ufm = parse_mark(int_raw)
            sem_display, s_num_val, sem_is_abs_ufm = parse_mark(sem_raw)
            grace_display, g_num_val, grace_is_abs_ufm = parse_mark(grace_raw)
            comp_display, c_num_val, comp_is_abs_ufm = parse_mark(comp_raw)

            int_max_num, int_max_display = _resolve_max(audit, m_code, 'Internal Actual', int_scale_raw)

            sem_max_num, sem_max_display = _resolve_max(audit, m_code, 'Sem Actual', sem_scale_raw)

            comp_max_num = parse_scale_num(comp_scale_raw)
            comp_max_display = parse_scale_str(comp_scale_raw)

            sub_overall = m_sub[m_sub['Appraisal Type Code'].astype(str).str.strip() == 'Overall Grade']
            completion_status = sub_overall['Completion Status after Gracing'].iloc[0] if not sub_overall.empty else ''

            # Merge Appraisal Remarks, General Remarks, Scale Type
            app_rem = str(sub_overall['Appraisal Remarks'].iloc[0]).strip() if not sub_overall.empty else ''
            gen_rem = str(sub_overall['General Remarks'].iloc[0]).strip() if not sub_overall.empty else ''
            scale_t = str(sub_overall['Scale Type'].iloc[0]).strip() if not sub_overall.empty else ''

            remarks_list = [r for r in [app_rem, gen_rem, scale_t] if r and r.upper() not in ['NAN', 'NONE', 'NA']]
            remarks_combined = " | ".join(remarks_list)

            # Verification: Internal Actual + Sem Actual + Sem Grace == Composite Score
            sum_components = i_num + s_num_val + g_num_val
            calc_total_comp += c_num_val
            calc_total_max_comp += comp_max_num

            is_mismatch = False
            if str(comp_raw).strip() not in ['', 'nan', 'NA']:
                if abs(sum_components - c_num_val) >= 0.01:
                    is_mismatch = True

            # 40% Minimum Passing Criteria Checks
            fail_int = (int_max_num > 0) and (i_num < 0.40 * int_max_num or int_is_abs_ufm)
            # 40% on Sem uses Sem Actual + Sem Grace
            fail_sem = (sem_max_num > 0) and ((s_num_val + g_num_val) < 0.40 * sem_max_num or sem_is_abs_ufm)
            fail_comp = (comp_max_num > 0) and (c_num_val < 0.40 * comp_max_num or comp_is_abs_ufm)

            _emit({
                'Module Code': m_code,
                'Module Desc': m_desc,
                'Credit': credit_val,
                'Internal Actual (Max)': int_max_display,
                'Internal Actual': int_display,
                'Sem Actual (Max)': sem_max_display,
                'Sem Actual': sem_display,
                'Sem Grace': grace_display,
                'Composite Score (Max)': comp_max_display,
                'Composite Score': comp_display,
                'Overall Grade': overall_grade,
                'Completion Status after Gracing': completion_status,
                'Remarks & Scale': remarks_combined,
                'Row_Type': 'DATA',
                'Comp_Mismatch': is_mismatch,
                'Red_Internal': int_is_abs_ufm or fail_int,
                'Red_Sem': sem_is_abs_ufm or fail_sem,
                'Red_Grace': grace_is_abs_ufm,
                'Red_Composite': comp_is_abs_ufm or fail_comp or is_mismatch
            }, s_num, m_code)

        # 2. Total & Percentage Summary Row
        tot_val_raw = total_row.iloc[0]['Marks after Gracing'] if not total_row.empty else (total_row.iloc[0]['Marks before Gracing'] if not total_row.empty else '')
        perc_val_raw = perc_row.iloc[0]['Marks after Gracing'] if not perc_row.empty else (perc_row.iloc[0]['Marks before Gracing'] if not perc_row.empty else '')

        tot_display, tot_num, _ = parse_mark(tot_val_raw)

        total_mismatch = False
        if str(tot_val_raw).strip() not in ['', 'nan', 'NA']:
            if abs(calc_total_comp - tot_num) >= 0.01:
                total_mismatch = True

        try:
            p_num = float(perc_val_raw)
            perc_display = f"{p_num:g}%"
        except (TypeError, ValueError):
            perc_display = f"{perc_val_raw}%" if str(perc_val_raw).strip() != '' else ''

        tot_max_display = int(calc_total_max_comp) if calc_total_max_comp.is_integer() else calc_total_max_comp

        _emit({
            'Module Code': '',
            'Module Desc': '',
            'Credit': '',
            'Internal Actual (Max)': '',
            'Internal Actual': '',
            'Sem Actual (Max)': '',
            'Sem Actual': '',
            'Sem Grace': '',
            'Composite Score (Max)': tot_max_display,
            'Composite Score': tot_display,
            'Overall Grade': perc_display,
            'Completion Status after Gracing': '',
            'Remarks & Scale': '',
            'Row_Type': 'SUMMARY',
            'Comp_Mismatch': total_mismatch,
            'Red_Internal': False,
            'Red_Sem': False,
            'Red_Grace': False,
            'Red_Composite': total_mismatch
        }, s_num)

    return pd.DataFrame(all_processed_rows)


# ==================================================
# STEP 2: GRACED CASES & SUMMARIES
# ==================================================
def process_gracing_data(df):
    """Extracts cases strictly where grace marks (>0) were awarded."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    non_summary = df[
        df['Module Desc'].notna() & 
        (~df['Student Name'].astype(str).str.upper().isin(['TOTAL', 'PERCENTAGE']))
    ].copy()

    grace_rows = non_summary[
        non_summary['Appraisal Type Code'].astype(str).str.strip().isin(['Sem Grace', 'Grace on Total marks'])
    ].copy()

    grace_rows['Grace_Marks_Val'] = pd.to_numeric(grace_rows['Marks after Gracing'], errors='coerce').fillna(0)
    awarded_grace = grace_rows[grace_rows['Grace_Marks_Val'] > 0].copy()

    records = []
    for _, g_row in awarded_grace.iterrows():
        s_num = g_row['Student Number']
        m_code = g_row['Module Code']

        sub_df = non_summary[
            (non_summary['Student Number'] == s_num) & 
            (non_summary['Module Code'] == m_code)
        ]

        sem_actual = sub_df[sub_df['Appraisal Type Code'].astype(str).str.strip() == 'Sem Actual']
        sem_tot = sub_df[sub_df['Appraisal Type Code'].astype(str).str.strip() == 'Semester Total Marks']
        comp_score = sub_df[sub_df['Appraisal Type Code'].astype(str).str.strip() == 'Composite Score']
        grade_row = sub_df[sub_df['Appraisal Type Code'].astype(str).str.strip() == 'Overall Grade']

        marks_before = sem_actual.iloc[0]['Marks before Gracing'] if not sem_actual.empty else g_row.get('Marks before Gracing', '')
        marks_after = sem_tot.iloc[0]['Marks after Gracing'] if not sem_tot.empty else ''
        comp_before = comp_score.iloc[0]['Marks before Gracing'] if not comp_score.empty else ''
        comp_after = comp_score.iloc[0]['Marks after Gracing'] if not comp_score.empty else ''
        grade = grade_row.iloc[0]['Marks after Gracing'] if not grade_row.empty else ''
        status = grade_row.iloc[0]['Completion Status after Gracing'] if not grade_row.empty else ''

        records.append({
            'Student Number': s_num,
            'Student Name': g_row.get('Student Name', ''),
            'Additional ID': g_row.get('Additional ID', ''),
            'Module Code': m_code,
            'Module Desc': g_row.get('Module Desc', ''),
            'Appraisal Type': g_row.get('Appraisal Type Code', ''),
            'Marks Before Gracing': marks_before,
            'Grace Marks': g_row['Grace_Marks_Val'],
            'Marks After Gracing': marks_after,
            'Composite Score Before': comp_before,
            'Composite Score After': comp_after,
            'Grade': grade,
            'Status': status
        })

    return pd.DataFrame(records)


def get_subject_summary(graced_df):
    """Generates subject-wise summary for graced marks."""
    if graced_df.empty:
        return pd.DataFrame(columns=[
            'Subject Code', 'Subject Name', 'Graced Students Count', 
            'Total Grace Marks', 'Max Grace Marks', 'Avg Grace Marks'
        ])

    subj_summary = graced_df.groupby(['Module Code', 'Module Desc']).agg(
        Graced_Students=('Student Number', 'nunique'),
        Total_Grace=('Grace Marks', 'sum'),
        Max_Grace=('Grace Marks', 'max'),
        Avg_Grace=('Grace Marks', 'mean')
    ).reset_index()

    subj_summary['Avg_Grace'] = subj_summary['Avg_Grace'].round(2)

    return subj_summary.rename(columns={
        'Module Code': 'Subject Code',
        'Module Desc': 'Subject Name',
        'Graced_Students': 'Graced Students Count',
        'Total_Grace': 'Total Grace Marks',
        'Max_Grace': 'Max Grace Marks',
        'Avg_Grace': 'Avg Grace Marks'
    })


def get_student_summary(graced_df):
    """Generates student-wise summary for graced marks."""
    if graced_df.empty:
        return pd.DataFrame(columns=[
            'Student Number', 'Student Name', 'Additional ID', 
            'Graced Subjects Count', 'Total Grace Marks', 'Graced Subjects & Marks'
        ])

    student_summary = graced_df.groupby(['Student Number', 'Student Name', 'Additional ID']).agg(
        Graced_Subjects_Count=('Module Code', 'count'),
        Total_Grace_Marks=('Grace Marks', 'sum'),
        Subjects_List=('Module Desc', lambda x: ', '.join([
            f"{subj} (+{int(g) if float(g).is_integer() else g})" 
            for subj, g in zip(x, graced_df.loc[x.index, 'Grace Marks'])
        ]))
    ).reset_index()

    return student_summary.rename(columns={
        'Graced_Subjects_Count': 'Graced Subjects Count',
        'Total_Grace_Marks': 'Total Grace Marks',
        'Subjects_List': 'Graced Subjects & Marks'
    })


# ==================================================
# STEP 3: DISPLAY & STYLING HELPERS
# ==================================================
def format_value(value):
    """Formats cell values cleanly for Streamlit interactive preview."""
    text = str(value).strip()
    if text == "" or text.upper() in ["NA", "NAN", "NONE"]:
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
        return f"{number:g}"
    except (TypeError, ValueError):
        return text


def style_processed_preview(proc_df):
    """
    Styles the Processed Data table:
    - Bold max columns
    - Center alignment for Credit through Overall Grade
    - BOLD Red font for marks < 40%, Z1 ("AB"), Z3 ("UFM"), or mismatched Composite Scores
    - BOLD Red font for 'F' grades
    - BOLD Red font for 'Completed Unsuccessfuly'
    - Thick bottom border on student summary rows
    """
    internal_flags = INTERNAL_FLAGS
    display_df = proc_df.drop(columns=internal_flags, errors='ignore').copy()
    display_df.index = range(1, len(display_df) + 1)

    def apply_custom_styles(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)

        center_cols = [
            'Credit', 'Internal Actual (Max)', 'Internal Actual',
            'Sem Actual (Max)', 'Sem Actual', 'Sem Grace', 
            'Composite Score (Max)', 'Composite Score', 'Overall Grade'
        ]

        max_cols = ['Internal Actual (Max)', 'Sem Actual (Max)', 'Composite Score (Max)']

        for idx in data.index:
            orig_idx = idx - 1
            p_row = proc_df.iloc[orig_idx]
            row_type = p_row.get('Row_Type', '')

            status_val = str(data.loc[idx, 'Completion Status after Gracing']).strip() if 'Completion Status after Gracing' in data.columns else ''

            # Center align specified columns
            for col in center_cols:
                if col in data.columns:
                    styles.loc[idx, col] += "text-align: center;"

            # Bold max columns
            for col in max_cols:
                if col in data.columns:
                    styles.loc[idx, col] += "font-weight: bold;"

            # 1. Student Header Row (Soft Blue Background + Bold text)
            if row_type == 'HEADER':
                styles.loc[idx, :] += "background-color: #D9E1F2; font-weight: bold; color: #1F4E78;"
                continue

            # 2. Total Summary Row (Soft Yellow Background + Bold text + Thick Bottom Border)
            if row_type == 'SUMMARY':
                styles.loc[idx, :] += "background-color: #FFF2CC; font-weight: bold; color: #000000; border-bottom: 3px solid #000000;"
                if p_row.get('Red_Composite', False) and 'Composite Score' in data.columns:
                    styles.loc[idx, 'Composite Score'] += "color: red; font-weight: bold;"
                continue

            # 3. BOLD Red font for < 40%, Z1 ("AB"), Z3 ("UFM")
            if p_row.get('Red_Internal', False) and 'Internal Actual' in data.columns:
                styles.loc[idx, 'Internal Actual'] += "color: red; font-weight: bold;"
            if p_row.get('Red_Sem', False) and 'Sem Actual' in data.columns:
                styles.loc[idx, 'Sem Actual'] += "color: red; font-weight: bold;"
            if p_row.get('Red_Grace', False) and 'Sem Grace' in data.columns:
                styles.loc[idx, 'Sem Grace'] += "color: red; font-weight: bold;"
            if p_row.get('Red_Composite', False) and 'Composite Score' in data.columns:
                styles.loc[idx, 'Composite Score'] += "color: red; font-weight: bold;"

            # 4. BOLD Red font for Grade 'F'
            grade_val = str(data.loc[idx, 'Overall Grade']).strip().upper() if 'Overall Grade' in data.columns else ''
            if grade_val == 'F' and 'Overall Grade' in data.columns:
                styles.loc[idx, 'Overall Grade'] += "color: red; font-weight: bold;"

            # 5. BOLD Red font for 'Completed Unsuccessfuly'
            if status_val.lower() == 'completed unsuccessfuly' and 'Completion Status after Gracing' in data.columns:
                styles.loc[idx, 'Completion Status after Gracing'] += "color: red; font-weight: bold;"

        # 6. AUDIT highlights: ERROR = red fill + bold yellow font, WARNING = amber fill
        for idx in data.index:
            cells = proc_df.iloc[idx - 1].get('Audit_Cells') or {}
            for col_name, sev in cells.items():
                if col_name in data.columns:
                    styles.loc[idx, col_name] += AUDIT_ERR_CSS if sev == ERR else AUDIT_WARN_CSS
        if 'Audit Remarks' in data.columns:
            styles['Audit Remarks'] += "white-space: pre-wrap; color: #C00000; min-width: 380px;"

        return styles

    return (
        display_df.style
        .format(format_value)
        .apply(apply_custom_styles, axis=None)
    )


def style_preview(df):
    """Prepares dataframe with 1-based indexing for preview display."""
    display_df = df.copy()
    display_df.index = range(1, len(display_df) + 1)
    return display_df.style.format(format_value)


def style_audit_table(df, status_col):
    """Colours a status column: ERROR/FAIL = red fill + bold yellow, WARNING/REVIEW = amber, PASS/CLEAN/OK = green."""
    def _css(v):
        v = str(v)
        if v in ('ERROR', 'FAIL', 'ERRORS', 'DISCREPANCY'):
            return AUDIT_ERR_CSS
        if v in ('WARNING', 'REVIEW', 'WARNINGS'):
            return AUDIT_WARN_CSS
        if v in ('PASS', 'CLEAN', 'OK'):
            return "background-color: #C6EFCE; color: #006100; font-weight: bold;"
        return ""
    styler = df.style.format(format_value)
    if status_col in df.columns:
        styler = styler.map(_css, subset=[status_col]) if hasattr(styler, 'map') else styler.applymap(_css, subset=[status_col])
    return styler


# ==================================================
# STEP 4: EXCEL REPORT GENERATION
# ==================================================
def format_excel_processed_sheet(worksheet, clean_proc, proc_df):
    """Helper to apply formatting, borders, bold fonts, and red highlights to processed student sheets."""
    cols = clean_proc.columns.tolist()
    n_data_cols = len(cols) - (1 if 'Audit Remarks' in cols else 0)
    audit_idx = cols.index('Audit Remarks') + 1 if 'Audit Remarks' in cols else None

    column_widths = {
        'Module Code': 15,
        'Module Desc': 40,
        'Credit': 7,
        'Internal Actual (Max)': 11,
        'Internal Actual': 11,
        'Sem Actual (Max)': 11,
        'Sem Actual': 11,
        'Sem Grace': 11,
        'Composite Score (Max)': 11,
        'Composite Score': 11,
        'Overall Grade': 11,
        'Completion Status after Gracing': 25,
        'Remarks & Scale': 25,
        'Audit Remarks': 75
    }

    center_cols_indices = [
        i + 1 for i, c in enumerate(cols) if c in [
            'Credit', 'Internal Actual (Max)', 'Internal Actual',
            'Sem Actual (Max)', 'Sem Actual', 'Sem Grace', 
            'Composite Score (Max)', 'Composite Score', 'Overall Grade'
        ]
    ]
    max_cols_indices = [
        i + 1 for i, c in enumerate(cols) if c in [
            'Internal Actual (Max)', 'Sem Actual (Max)', 'Composite Score (Max)'
        ]
    ]

    int_act_idx = cols.index('Internal Actual') + 1 if 'Internal Actual' in cols else None
    sem_act_idx = cols.index('Sem Actual') + 1 if 'Sem Actual' in cols else None
    sem_grace_idx = cols.index('Sem Grace') + 1 if 'Sem Grace' in cols else None
    comp_score_idx = cols.index('Composite Score') + 1 if 'Composite Score' in cols else None
    grade_idx = cols.index('Overall Grade') + 1 if 'Overall Grade' in cols else None
    status_col_idx = cols.index('Completion Status after Gracing') + 1 if 'Completion Status after Gracing' in cols else None

    thin_side = Side(style="thin", color="000000")
    thick_side = Side(style="medium", color="000000")

    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    summary_thick_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_side)

    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center")

    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    student_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    summary_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    red_bold_font = Font(color="FF0000", bold=True)
    bold_font = Font(bold=True)
    bold_blue_font = Font(color="1F4E78", bold=True)

    worksheet.views.sheetView[0].showGridLines = False
    worksheet.freeze_panes = "A2"

    for row_idx in range(1, len(clean_proc) + 2):
        if row_idx == 1:
            for col_idx in range(1, len(cols) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
                cell.font = bold_font
                cell.fill = header_fill
                cell.alignment = header_align
            continue

        proc_idx = row_idx - 2
        p_row = proc_df.iloc[proc_idx]
        row_type_val = p_row.get('Row_Type', '')

        is_student_header = (row_type_val == 'HEADER')
        is_summary_row = (row_type_val == 'SUMMARY')

        if is_student_header:
            worksheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=n_data_cols)
            for col_idx in range(1, len(cols) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.border = thin_border
                cell.font = bold_blue_font
                cell.fill = student_header_fill
        else:
            for col_idx in range(1, len(cols) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)

                if is_summary_row:
                    cell.border = summary_thick_border
                    cell.font = bold_font
                    cell.fill = summary_fill
                else:
                    cell.border = thin_border
                    if col_idx in max_cols_indices:
                        cell.font = bold_font

                if col_idx in center_cols_indices:
                    cell.alignment = center_align

                # BOLD RED font for marks < 40%, Z1 ("AB"), Z3 ("UFM")
                if p_row.get('Red_Internal', False) and col_idx == int_act_idx:
                    cell.font = red_bold_font
                if p_row.get('Red_Sem', False) and col_idx == sem_act_idx:
                    cell.font = red_bold_font
                if p_row.get('Red_Grace', False) and col_idx == sem_grace_idx:
                    cell.font = red_bold_font
                if p_row.get('Red_Composite', False) and col_idx == comp_score_idx:
                    cell.font = red_bold_font

                # BOLD RED font for Overall Grade 'F'
                grade_val = str(cell.value).strip().upper() if col_idx == grade_idx else ''
                if col_idx == grade_idx and grade_val == 'F':
                    cell.font = red_bold_font

                # BOLD RED font for unsuccessful completion status
                status_val = str(worksheet.cell(row=row_idx, column=status_col_idx).value).strip() if status_col_idx else ''
                if status_col_idx and col_idx == status_col_idx and status_val.lower() == 'completed unsuccessfuly':
                    cell.font = red_bold_font

        # AUDIT highlights: ERROR = red fill + bold yellow font; WARNING = amber fill + bold black font
        audit_flags = p_row.get('Audit_Cells') or {}
        for col_name, sev in audit_flags.items():
            if is_student_header:
                targets = range(1, n_data_cols + 1)
            elif col_name in cols:
                targets = [cols.index(col_name) + 1]
            else:
                continue
            for t_idx in targets:
                t_cell = worksheet.cell(row=row_idx, column=t_idx)
                if sev == ERR:
                    t_cell.fill, t_cell.font = AUDIT_ERR_FILL, AUDIT_ERR_FONT
                else:
                    t_cell.fill, t_cell.font = AUDIT_WARN_FILL, AUDIT_WARN_FONT
        if not is_student_header:
            for v_idx in (1, 2):
                v_cell = worksheet.cell(row=row_idx, column=v_idx)
                v_cell.alignment = Alignment(horizontal=v_cell.alignment.horizontal, vertical="center")
        if audit_idx:
            a_cell = worksheet.cell(row=row_idx, column=audit_idx)
            a_cell.alignment = Alignment(vertical="center", wrap_text=True)
            has_err = ERR in audit_flags.values()
            if a_cell.value:
                a_cell.font = Font(color="C00000" if has_err or "[ERROR]" in str(a_cell.value) else "9C5700", bold=False)

    # Apply specified fixed column widths
    for col_name, width in column_widths.items():
        if col_name in cols:
            c_idx = cols.index(col_name) + 1
            col_letter = get_column_letter(c_idx)
            worksheet.column_dimensions[col_letter].width = width


def to_excel_bytes(all_proc_df, failed_proc_df, graced_proc_df, graced_df, subject_summary, student_summary,
                   audit=None, source_name=""):
    """
    Generates styled multi-sheet Excel output file:
    - Sheets: "Audit Summary", "Audit Findings" (when audit supplied), "All Students", "Failed Students", "Graced Students", 
              "Graced Cases Details", "Subject Summary", "Student Summary"
    """
    buffer = io.BytesIO()
    internal_flags = INTERNAL_FLAGS

    clean_all = all_proc_df.drop(columns=internal_flags, errors='ignore')
    clean_failed = failed_proc_df.drop(columns=internal_flags, errors='ignore')
    clean_graced_students = graced_proc_df.drop(columns=internal_flags, errors='ignore')

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if audit is not None:
            write_audit_sheets(writer, audit, source_name)
        clean_all.to_excel(writer, index=False, sheet_name="All Students")
        clean_failed.to_excel(writer, index=False, sheet_name="Failed Students")
        clean_graced_students.to_excel(writer, index=False, sheet_name="Graced Students")

        if not graced_df.empty:
            graced_df.to_excel(writer, index=False, sheet_name="Graced Cases Details")
            subject_summary.to_excel(writer, index=False, sheet_name="Subject Summary")
            student_summary.to_excel(writer, index=False, sheet_name="Student Summary")

        workbook = writer.book

        # Format All Students, Failed Students, Graced Students sheets
        format_excel_processed_sheet(workbook["All Students"], clean_all, all_proc_df)
        format_excel_processed_sheet(workbook["Failed Students"], clean_failed, failed_proc_df)
        format_excel_processed_sheet(workbook["Graced Students"], clean_graced_students, graced_proc_df)

        # Format remaining summary sheets
        thin_side = Side(style="thin", color="000000")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        bold_font = Font(bold=True)

        other_sheets = ["Graced Cases Details", "Subject Summary", "Student Summary"] if not graced_df.empty else []
        for sheet_name in other_sheets:
            worksheet = workbook[sheet_name]
            worksheet.views.sheetView[0].showGridLines = False
            worksheet.freeze_panes = "A2"

            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.border = thin_border
                    if cell.row == 1:
                        cell.font = bold_font
                        cell.fill = header_fill
                        cell.alignment = header_align

            for col in worksheet.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 14)

        if audit is not None:
            write_audit_detail_sheets(writer, audit)

    return buffer.getvalue()


# ==================================================
# MAIN ENTRY POINT
# ==================================================
def show():
    """Renders the Gracing Checker tool page in Streamlit."""
    st.title("✅ Gracing Checker & Result Audit Engine")
    st.caption("Upload an examination result Excel file to verify every record and analyze awarded grace marks.")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file is None:
        return

    try:
        # Load raw file
        with st.spinner("Loading and processing gracing data..."):
            df = pd.read_excel(uploaded_file, dtype=object, keep_default_na=False)

        if df.empty:
            st.warning("The uploaded file is empty.")
            return

        # Clean column headers
        df.columns = df.columns.astype(str).str.strip()

        # Check required columns
        required_cols = ['Student Number', 'Module Desc', 'Marks after Gracing', 'Appraisal Type Code']
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            st.error(f"The uploaded file is missing expected columns: {', '.join(missing_cols)}")
            return

        # Filter student lists
        failed_prns = df[
            df['Completion Status after Gracing'].astype(str).str.strip().str.lower() == 'completed unsuccessfuly'
        ]['Student Number'].dropna().unique().tolist()

        graced_prns = df[
            (df['Appraisal Type Code'].astype(str).str.strip().isin(['Sem Grace', 'Grace on Total marks'])) &
            (pd.to_numeric(df['Marks after Gracing'], errors='coerce').fillna(0) > 0)
        ]['Student Number'].dropna().unique().tolist()

        # Generate Processed Datasets
        with st.spinner("Running validation audit (recalculating every record)..."):
            audit = run_audit(df)

        all_proc_df = create_processed_data(df, audit=audit)
        failed_proc_df = create_processed_data(df, target_student_ids=failed_prns, audit=audit)
        graced_proc_df = create_processed_data(df, target_student_ids=graced_prns, audit=audit)

        # Generate Summaries
        graced_df = process_gracing_data(df)
        subject_summary = get_subject_summary(graced_df)
        student_summary = get_student_summary(graced_df)

        # Statistics
        total_students = df['Student Number'].nunique()
        graced_students_count = len(graced_prns)
        total_graced_cases = len(graced_df)
        total_grace_marks = graced_df['Grace Marks'].sum() if not graced_df.empty else 0

        # Update session state
        st.session_state.students_processed = int(total_students)

        st.markdown("---")

        # Metric Cards
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total Candidates Evaluated", total_students)

        with col2:
            st.metric("Candidates Receiving Grace", graced_students_count)

        with col3:
            st.metric("Total Gracing Cases", total_graced_cases)

        with col4:
            st.metric("Total Grace Marks Awarded", f"{total_grace_marks:g}")

        # Audit verdict banner + metrics
        if audit.n_errors == 0:
            if audit.n_warnings == 0:
                st.success("✅ AUDIT PASSED - no errors or warnings. Result data is clean.")
            else:
                st.warning(f"⚠️ AUDIT PASSED WITH WARNINGS - {audit.n_warnings} warning(s) to review before publication.")
        else:
            st.error(f"❌ NOT READY FOR PUBLICATION - {audit.n_errors} error(s) and "
                     f"{audit.n_warnings} warning(s) found. Affected cells are marked RED / YELLOW in the report.")

        a1, a2, a3, a4 = st.columns(4)
        a1.metric("Validation Checks Run", f"{audit.stats['checks_run']:,}")
        a2.metric("Errors", audit.n_errors)
        a3.metric("Warnings", audit.n_warnings)
        a4.metric("Students Needing Attention", audit.stats['students_error'] + audit.stats['students_warning'])

        st.markdown("---")

        # Tab-based Preview Windows
        tab_audit, tab_find, tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "🛡️ Audit Summary",
            "🚨 Audit Findings",
            "📊 All Students", 
            "⚠️ Failed Students", 
            "✅ Graced Students",
            "📋 Graced Cases Details", 
            "📚 Subject Summary", 
            "👤 Student Summary"
        ])

        with tab_audit:
            st.subheader("Validation Audit Summary")
            st.caption("Every rule below was executed against the uploaded file; rules that passed are listed too.")
            st.write(style_audit_table(audit.rule_table(), 'Result'))
            with st.expander("👤 Student-wise audit status", expanded=(audit.n_errors + audit.n_warnings) > 0):
                stu_df = pd.DataFrame(audit.student_rows)
                st.write(style_audit_table(stu_df.set_index('Student Number'), 'Audit Status'))
            with st.expander("🧮 Recalculation trail (reported vs recalculated, every subject)"):
                st.write(style_audit_table(pd.DataFrame(audit.trail), 'Audit Result'))

        with tab_find:
            st.subheader("Audit Findings")
            findings_df = audit.findings_df()
            if findings_df.empty:
                st.success("No discrepancies or warnings found.")
            else:
                f1, f2, f3 = st.columns(3)
                sev_sel = f1.multiselect("Severity", [ERR, WARN], default=[ERR, WARN])
                cat_opts = sorted(findings_df['Category'].unique())
                cat_sel = f2.multiselect("Category", cat_opts, default=cat_opts)
                query = f3.text_input("Search student / module / remark")
                view = findings_df[findings_df['Severity'].isin(sev_sel) & findings_df['Category'].isin(cat_sel)]
                if query.strip():
                    hay = view[['Student Number', 'Student Name', 'Module Code', 'Module Desc', 'Remark']] \
                        .astype(str).agg(' '.join, axis=1).str.lower()
                    view = view[hay.str.contains(query.strip().lower(), regex=False)]
                st.caption(f"Showing {len(view)} of {len(findings_df)} findings.")
                st.write(style_audit_table(view.set_index('Sr No'), 'Severity'))

        with tab1:
            st.subheader("All Students Preview")
            if not all_proc_df.empty:
                st.write(style_processed_preview(all_proc_df))
            else:
                st.info("No student data generated.")

        with tab2:
            st.subheader("Failed Students Preview")
            if not failed_proc_df.empty:
                st.write(style_processed_preview(failed_proc_df))
            else:
                st.info("No failed students found in this file.")

        with tab3:
            st.subheader("Graced Students Preview")
            if not graced_proc_df.empty:
                st.write(style_processed_preview(graced_proc_df))
            else:
                st.info("No graced students found in this file.")

        with tab4:
            st.subheader("Graced Cases Details")
            if not graced_df.empty:
                st.write(style_preview(graced_df))
            else:
                st.info("No grace marks were awarded in this file.")

        with tab5:
            st.subheader("Subject-Wise Gracing Summary")
            if not subject_summary.empty:
                st.write(style_preview(subject_summary))
            else:
                st.info("No subject-wise gracing data available.")

        with tab6:
            st.subheader("Student-Wise Gracing Summary")
            if not student_summary.empty:
                st.write(style_preview(student_summary))
            else:
                st.info("No student-wise gracing data available.")

        # Export Section
        st.markdown("---")
        st.subheader("📥 Export Gracing Report")

        source_name = getattr(uploaded_file, 'name', '')
        excel_bytes = to_excel_bytes(all_proc_df, failed_proc_df, graced_proc_df, graced_df, subject_summary,
                                     student_summary, audit=audit, source_name=source_name)
        date_str = datetime.now().strftime("%d-%m-%y")
        base_name = os.path.splitext(os.path.basename(source_name))[0] or "Result"
        output_filename = f"Gracing Analysis - {base_name}_{int(total_students)}_{date_str}.xlsx"

        st.download_button(
            label=f"📥 Download Complete Gracing Report ({output_filename})",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        audit_filename = f"Gracing Audit Only - {base_name}_{int(total_students)}_{date_str}.xlsx"
        st.download_button(
            label=f"🛡️ Download Validation Audit Report only ({audit_filename})",
            data=audit_only_excel_bytes(audit, source_name),
            file_name=audit_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as error:
        st.error(f"Something went wrong while processing the file: {error}")