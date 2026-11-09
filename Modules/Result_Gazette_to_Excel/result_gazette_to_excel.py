import colorsys
import io
import numbers
import re
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ==============================================================================
# ACADEMIC VALIDATION CONSTANTS & GRADE / GP LOOKUP HELPERS
# ==============================================================================

GRADE_POINTS_MAP = {
    "O": 10,
    "A+": 9,
    "A": 8,
    "B+": 7,
    "B": 6,
    "C": 5,
    "P": 4,
    "F": 0,
}

VALID_GRADES = set(GRADE_POINTS_MAP.keys())

ROMAN_ORDER = [
    "I",
    "II",
    "III",
    "IV",
    "V",
    "VI",
    "VII",
    "VIII",
    "IX",
    "X",
]


def sem_rank(sem_str):
    """Return numeric sort rank for a Roman or integer semester token."""
    s = str(sem_str or "").strip().upper()
    if s in ROMAN_ORDER:
        return ROMAN_ORDER.index(s)
    if s.isdigit():
        return int(s) - 1
    return 99


def to_decimal(val):
    """Safely convert a numeric value or string to Decimal (returns None for '--', 'AB', etc.)."""
    if val is None:
        return None
    s = str(val).strip().replace("%", "").replace(",", ".")
    if not s or s in ["--", "-", "AB", "NV", "nan", "NaN", "None", "AD"]:
        return None
    s = (
        s.replace("$", "")
        .replace("#", "")
        .replace("®", "")
        .replace("*", "")
        .replace("~", "")
        .strip()
    )
    if not s or s in ["--", "-"]:
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


# --- Special gazette notations (informational, never errors) ----------------------
#   '#'  after a semester SGPA : candidate eligible for O.229 grace; the unused part of
#        the 10 grace marks is added to the SGPA (0.01 per unused mark, e.g. 10 unused -> +0.10)
#   'AD' in a semester SGPA     : Attendance Defaulter in that semester
#   '$$' after Result Status    : candidate has not cleared a lower semester
GRACE_MAX_MARKS = 10
GRACE_SGPA_PER_UNUSED_MARK = Decimal("0.01")
_SEM_SGPA_COL_RE = re.compile(r"^Sem_[IVXLCDM\d]+_SGPA$")


def grace_sgpa_info(row):
    """Return (used_marks, unused_marks, sgpa_bonus) for the O.229 grace rule of a student row."""
    used = 0
    for col in row.index:
        if re.match(r"^Sub\d+ (IA|TH) Marks$", str(col)):
            for m in re.finditer(r"\+\s*(\d+)", str(row[col])):
                used += int(m.group(1))
    used = min(used, GRACE_MAX_MARKS)
    unused = GRACE_MAX_MARKS - used
    return used, unused, Decimal(unused) * GRACE_SGPA_PER_UNUSED_MARK


def classify_special_notation(col_name, raw_value):
    """Return 'grace' ('#'), 'ad' (Attendance Defaulter) or 'lower_unclear' ('$$'), else None."""
    txt = str(raw_value).strip()
    if _SEM_SGPA_COL_RE.match(str(col_name)):
        if "#" in txt:
            return "grace"
        if txt.upper() == "AD":
            return "ad"
    elif str(col_name) == "Result Status" and "$$" in txt.replace(" ", ""):
        return "lower_unclear"
    return None


def special_notation_note(kind, row):
    if kind == "grace":
        used, unused, bonus = grace_sgpa_info(row)
        return (
            "'#' - Eligible for O.229 grace marks (max 10).\n"
            f"Grace marks used in subjects: {used}; unused: {unused}.\n"
            f"Unused grace converted to SGPA bonus: +{bonus:.2f} (already included in the printed SGPA)."
        )
    if kind == "ad":
        return "AD - Attendance Defaulter in this semester (SGPA not declared)."
    if kind == "lower_unclear":
        return "$$ - Candidate has not cleared a lower semester examination (e.g. AD or failure in an earlier semester)."
    return None


def derive_grade_from_unrounded_percentage(marks_obt, max_marks):
    """
    Calculate percentage as (Total Marks Obt. / Max Total Marks * 100)
    WITHOUT rounding before grade determination, then map to Grade Scale:
    O = 90-100, A+ = 80 to <90, A = 70 to <80, B+ = 60 to <70,
    B = 55 to <60, C = 50 to <55, P = 40 to <50, F = <40.
    """
    d_obt = to_decimal(marks_obt)
    d_max = to_decimal(max_marks)
    if d_obt is None or d_max is None or d_max <= Decimal("0"):
        return None, None

    unrounded_pct = (d_obt / d_max) * Decimal("100")

    if unrounded_pct < Decimal("0") or unrounded_pct > Decimal("100"):
        return None, unrounded_pct

    if unrounded_pct >= Decimal("90"):
        grade = "O"
    elif unrounded_pct >= Decimal("80"):
        grade = "A+"
    elif unrounded_pct >= Decimal("70"):
        grade = "A"
    elif unrounded_pct >= Decimal("60"):
        grade = "B+"
    elif unrounded_pct >= Decimal("55"):
        grade = "B"
    elif unrounded_pct >= Decimal("50"):
        grade = "C"
    elif unrounded_pct >= Decimal("40"):
        grade = "P"
    else:
        grade = "F"

    return grade, unrounded_pct


def derive_grade_from_gpa(gpa_val):
    """Derive Semester Grade or Final Grade from SGPA / CGPA."""
    d_gpa = to_decimal(gpa_val)
    if d_gpa is None:
        return None
    if d_gpa < Decimal("0") or d_gpa > Decimal("10"):
        return None

    if d_gpa == Decimal("10"):
        return "O"
    elif d_gpa >= Decimal("9"):
        return "A+"
    elif d_gpa >= Decimal("8"):
        return "A"
    elif d_gpa >= Decimal("7"):
        return "B+"
    elif d_gpa >= Decimal("6"):
        return "B"
    elif d_gpa >= Decimal("5"):
        return "C"
    elif d_gpa >= Decimal("4"):
        return "P"
    else:
        return "F"


def normalize_grade_token(grade_str):
    """Strip special symbols ($, #, ®) from a grade string while returning base grade and flags."""
    raw = str(grade_str or "").strip()
    has_carry = "$" in raw
    has_hash = "#" in raw
    has_reval = "®" in raw
    base = (
        raw.replace("$", "")
        .replace("#", "")
        .replace("®", "")
        .replace("*", "")
        .strip()
        .upper()
    )
    if base in ["--", "-"]:
        base = ""
    elif base == "0":
        base = "O"
    return base, {
        "has_carry": has_carry,
        "has_hash": has_hash,
        "has_reval": has_reval,
    }


def parse_mark_component(raw_val):
    """
    Parse an IA, TH, or Total mark cell in the background and detect special cases:
    - AB (Absent - valid entry, treated as 0 numeric marks and F grade)
    - NV (Null & Void)
    - $ (Carried forward)
    - # or + (Ordinance 0.229 / Grace marks e.g. '17 +7', '10 +2', '8 +4' - valid entry)
    - ® (Revaluation)
    - -- (Not Applicable / Blank)
    """
    s = str(raw_val if raw_val is not None else "").strip()
    info = {
        "raw": s,
        "is_empty": s in ["", "--", "-", "nan", "NaN", "None"],
        "is_absent": False,
        "is_nv": False,
        "has_carry": "$" in s,
        "has_hash": "#" in s,
        "has_reval": "®" in s,
        "has_grace": False,
        "base_marks": None,
        "grace_marks": Decimal("0"),
        "numeric_marks": None,
        "is_valid_number": False,
    }

    if info["is_empty"]:
        return info

    cleaned = (
        s.replace("$", "")
        .replace("#", "")
        .replace("®", "")
        .replace("~", "")
        .replace("*", "")
        .strip()
    )

    if re.search(r"\bAB\b", cleaned, re.IGNORECASE):
        info["is_absent"] = True
        info["numeric_marks"] = Decimal("0")
        return info

    if re.search(r"\bNV\b", cleaned, re.IGNORECASE):
        info["is_nv"] = True
        info["numeric_marks"] = Decimal("0")
        return info

    # Check for grace marks syntax: e.g. "17 +7", "10 +2", "14 + 10", "21+3"
    m_grace = re.match(r"^(-?\d+(?:\.\d+)?)\s*\+\s*(\d+(?:\.\d+)?)$", cleaned)
    if m_grace:
        base_d = Decimal(m_grace.group(1))
        grace_d = Decimal(m_grace.group(2))
        info["has_grace"] = True
        info["base_marks"] = base_d
        info["grace_marks"] = grace_d
        info["numeric_marks"] = base_d + grace_d
        info["is_valid_number"] = True
        return info

    # Standard numeric value
    m_num = re.match(r"^(-?\d+(?:\.\d+)?)$", cleaned)
    if m_num:
        val_d = Decimal(m_num.group(1))
        info["base_marks"] = val_d
        info["numeric_marks"] = val_d
        info["is_valid_number"] = True
        return info

    return info


# ==============================================================================
# PAGE 1 & PAGE 2 PARSERS (WITH MULTI-LINE & TABLE EXTRACTION SUPPORT)
# ==============================================================================


def clean_text_preserve_symbols(val):
    """Clean whitespace while preserving exact PDF academic notations ($, ~, /, #, +, ®, --)."""
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def clean_dashes_and_artifacts(text):
    """Remove long separator lines (4+ dashes/underscores) while preserving '--' tokens."""
    if not text:
        return ""
    text = re.sub(r"[-_=]{4,}", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_page_1(text):
    """Extract metadata key-values and accurately identify Regular vs Re-examination."""
    metadata = []
    text_clean = re.sub(r"\s+", " ", text or "")

    inst_match = re.search(
        r"(Shri Vile Parle Kelavani Mandal[^\n\r]*[\s\S]*?(?:AUTONOMOUS|\(Empowered Autonomous\)|\(Autonomous\)))",
        text or "",
        re.IGNORECASE,
    )
    inst_name = (
        re.sub(r"\s+", " ", inst_match.group(1)).strip()
        if inst_match
        else "SVKM Educational Institution (AUTONOMOUS)"
    )
    metadata.append({"Property": "Institution / Trust", "Value": inst_name})

    affil_match = re.search(r"^[ \t]*(AFFILIATED\s+TO[^\n\r]*)", text or "", re.IGNORECASE | re.MULTILINE)
    if affil_match:
        metadata.append(
            {"Property": "Affiliation", "Value": re.sub(r"\s+", " ", affil_match.group(1)).strip()}
        )
    naac_match = re.search(r"^[ \t]*(NAAC[^\n\r]*)", text or "", re.IGNORECASE | re.MULTILINE)
    if naac_match:
        metadata.append(
            {"Property": "NAAC Accreditation", "Value": re.sub(r"\s+", " ", naac_match.group(1)).strip()}
        )

    prog_match = re.search(
        r"(Bachelor of [^\n\r]+|Master of [^\n\r]+|M\.Com[^\n\r]*|B\.Sc[^\n\r]*|B\.Com[^\n\r]*|B\.A\.[^\n\r]*|B\.M\.S[^\n\r]*)",
        text or "",
        re.IGNORECASE,
    )
    programme_val = "N/A"
    if prog_match:
        # Read the whole programme line exactly as printed. A comma only ends the
        # programme name when it is outside brackets, so
        # "Bachelor of Arts (Film, Television and New Media Production)" stays intact.
        depth, end = 0, len(prog_match.group(1))
        for i, ch in enumerate(prog_match.group(1)):
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth = max(depth - 1, 0)
            elif ch == "," and depth == 0:
                end = i
                break
        programme_val = re.sub(r"\s+", " ", prog_match.group(1)[:end]).strip()
    metadata.append({"Property": "Programme", "Value": programme_val})

    sem_match = re.search(r"Semester\s+([IVXLCDM\d]+)", text or "", re.IGNORECASE)
    metadata.append(
        {
            "Property": "Semester",
            "Value": f"Semester {sem_match.group(1).upper()}" if sem_match else "N/A",
        }
    )

    ncrf_match = re.search(r"NCr?F\s*Level[:\s]*([0-9.]+)", text or "", re.IGNORECASE)
    metadata.append(
        {
            "Property": "NCrF Level",
            "Value": ncrf_match.group(1) if ncrf_match else "N/A",
        }
    )

    ay_match = re.search(r"A\.Y\.?\s*([0-9]{4}\s*-\s*[0-9]{4})", text or "", re.IGNORECASE)
    metadata.append(
        {
            "Property": "Academic Year",
            "Value": ay_match.group(1) if ay_match else "N/A",
        }
    )

    if re.search(r"Re\s*[-–]?\s*examination", text_clean, re.IGNORECASE):
        exam_type = "Re-examination"
    elif re.search(r"Regular\s+examination", text_clean, re.IGNORECASE):
        exam_type = "Regular Examination"
    elif re.search(r"ATKT\s+examination", text_clean, re.IGNORECASE):
        exam_type = "ATKT Examination"
    else:
        exam_type = (
            "Re-examination"
            if re.search(r"\bRe\b", text_clean, re.I)
            else "Regular Examination"
        )

    exam_session_match = re.search(
        r"(?:examination\s+held\s+in|held\s+in)\s+([A-Za-z]+[\s,]*(?:\d{4})?)",
        text_clean,
        re.IGNORECASE,
    )
    exam_session = (
        f"{exam_type} held in {exam_session_match.group(1).strip()}"
        if exam_session_match
        else exam_type
    )
    exam_session = re.sub(r"\s+", " ", exam_session).strip()
    # Prefer the examination line exactly as it is printed on the cover page
    printed_session = re.search(
        r"^[ \t]*((?:Regular|Re\s*-?\s*|ATKT|Supplementary)\s*examination\s+held\s+in[^\n\r]*)",
        text or "",
        re.IGNORECASE | re.MULTILINE,
    )
    if printed_session:
        exam_session = re.sub(r"\s+", " ", printed_session.group(1)).strip()

    metadata.append({"Property": "Examination Type", "Value": exam_type})
    metadata.append({"Property": "Examination Session", "Value": exam_session})

    sign_match = re.search(
        r"^[ \t]*(Controller\s+of\s+Examinations[ \t]+Principal)[ \t]*$",
        text or "",
        re.IGNORECASE | re.MULTILINE,
    )
    if sign_match:
        metadata.append(
            {"Property": "Signatories", "Value": re.sub(r"\s+", " ", sign_match.group(1)).strip()}
        )

    return pd.DataFrame(metadata), exam_type


def _cast_pdf_number_or_raw(val):
    """Preserve exact PDF representation ('--', '', int, or float) without fabricating 0s."""
    s = clean_text_preserve_symbols(val)
    if not s:
        return ""
    if s in ["--", "-"]:
        return s
    if re.fullmatch(r"\d+", s):
        return int(s)
    if re.fullmatch(r"\d+\.\d+", s):
        return float(s)
    return s


def parse_page_2(page):
    """
    Extract complete course catalogue and schema from Page 2.
    Uses table extraction first (to capture multi-line wrapped course titles such as
    UFLRF302, UFMAD306, NHCMU269, NHCDR268, NHCDN267 and blank/'--' Sem End Max Marks
    such as NHCEP264, NHACW279X), with a multi-line stitch text parser fallback.
    """
    courses = []
    seen_codes = set()

    # 1. Primary Method: pdfplumber table extraction (handles multi-line wrapped titles natively)
    try:
        tables = page.extract_tables() or []
    except Exception:
        tables = []

    for tbl in tables:
        for row in tbl:
            if not row:
                continue
            cleaned_cells = [
                re.sub(r"\s+", " ", (c or "")).strip() for c in row
            ]
            non_empty = [c for c in cleaned_cells if c]
            if len(non_empty) < 4:
                continue

            # Find cell matching Course Code pattern (e.g. UFIKS301, NHDBM278P, NHACW279X)
            code_idx = None
            for idx, cell_val in enumerate(cleaned_cells):
                if re.fullmatch(r"[A-Z]{2,8}\d{2,4}[A-Z0-9]{0,3}", cell_val):
                    code_idx = idx
                    break
            if code_idx is None:
                continue

            code = cleaned_cells[code_idx]
            rem = cleaned_cells[code_idx + 1 :]
            # Find the Credits cell (matches \d+\.\d{2} or integer credit)
            cred_rel_idx = None
            for r_idx, c_val in enumerate(rem):
                if re.fullmatch(r"\d+(?:\.\d{1,2})?", c_val) and r_idx >= 1:
                    cred_rel_idx = r_idx
                    break
            if cred_rel_idx is None:
                continue

            title = " ".join(c for c in rem[:cred_rel_idx] if c).strip()
            creds_raw = rem[cred_rel_idx]
            marks_Tail = rem[cred_rel_idx + 1 :]

            # Extract Int, Sem, Total columns preserving blank / '--' as-is
            int_raw = marks_Tail[0] if len(marks_Tail) > 0 else ""
            sem_raw = marks_Tail[1] if len(marks_Tail) > 1 else ""
            tot_raw = marks_Tail[2] if len(marks_Tail) > 2 else ""

            # If Total landed in marks_Tail[1] because Sem was merged/omitted
            if tot_raw == "" and sem_raw != "":
                tot_raw = sem_raw
                sem_raw = ""

            if code not in seen_codes:
                seen_codes.add(code)
                courses.append(
                    {
                        "Course Code": code,
                        "Course Title": title,
                        "Credits": _cast_pdf_number_or_raw(creds_raw),
                        "Internal Max Marks": _cast_pdf_number_or_raw(int_raw),
                        "Sem End Max Marks": _cast_pdf_number_or_raw(sem_raw),
                        "Total Max Marks": _cast_pdf_number_or_raw(tot_raw),
                    }
                )

    # 2. Fallback / Supplement: Multi-line text stitching parser
    text = page.extract_text() or ""
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    course_line_re = re.compile(
        r"^([A-Z]{2,8}\d{2,4}[A-Z0-9]{0,3})\s+(.+?)\s+(\d+\.\d{2})\s+(\d+|--)(?:\s+(\d+|--))?\s+(\d+)\b"
    )
    code_only_start_re = re.compile(r"^([A-Z]{2,8}\d{2,4}[A-Z0-9]{0,3})\b")

    i = 0
    while i < len(lines):
        line = lines[i]
        m = course_line_re.match(line)
        if not m and code_only_start_re.match(line) and i + 1 < len(lines):
            # Try stitching wrapped course title line(s)
            combined = line + " " + lines[i + 1]
            m = course_line_re.match(combined)
            if m:
                i += 1
        if m:
            code, title, creds, int_m, sem_m, tot_m = m.groups()
            code = code.strip()
            # Check if the next line is a trailing wrapped title fragment (e.g. "TV & OTT", "New Media", "Music Level III")
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if (
                    nxt
                    and not code_only_start_re.match(nxt)
                    and not re.search(
                        r"Shri Vile Parle|SVKM|AFFILIATED|NAAC|Module Credits|COURSE DETAILS|Maximum Marks|Controller|Principal|Page:",
                        nxt,
                        re.I,
                    )
                    and not re.search(r"\d+\.\d{2}", nxt)
                    and len(nxt) < 45
                ):
                    title = f"{title.strip()} {nxt}"
                    i += 1
            if code not in seen_codes:
                seen_codes.add(code)
                courses.append(
                    {
                        "Course Code": code,
                        "Course Title": re.sub(r"\s+", " ", title).strip(),
                        "Credits": _cast_pdf_number_or_raw(creds),
                        "Internal Max Marks": _cast_pdf_number_or_raw(int_m),
                        "Sem End Max Marks": _cast_pdf_number_or_raw(
                            sem_m if sem_m is not None else ""
                        ),
                        "Total Max Marks": _cast_pdf_number_or_raw(tot_m),
                    }
                )
        i += 1

    df = pd.DataFrame(courses)
    if df.empty:
        df = pd.DataFrame(
            columns=[
                "Course Code",
                "Course Title",
                "Credits",
                "Internal Max Marks",
                "Sem End Max Marks",
                "Total Max Marks",
            ]
        )
    return df


# ==============================================================================
# RESULT PAGES PARSER (COORDINATE-AWARE + PIPE-AWARE, AS-IS PDF FIDELITY)
# ==============================================================================

_HEADER_SKIP_PATTERNS = [
    "Shri Vile Parle Kelavani Mandal",
    "SVKM's",
    "SVKM'S",
    "Narsee Monjee College",
    "UPG College",
    "MITHIBAI COLLEGE",
    "AFFILIATED TO UNIVERSITY",
    "Affiliated to University",
    "NAAC RE-ACCREDITED",
    "NAAC Reaccredited",
    "Module Credits",
    "Consolidated Result",
    "COURSE DETAILS",
    "Candidate's Full Name",
    "MRK G GP C",
    "CG ; CG ;",
]

_COURSE_CODE_RE = re.compile(r"^[A-Z]{2,8}\d{2,4}[A-Z0-9]{0,3}$")


def clean_page_lines(raw_text):
    """Filter out non-student header/footer lines, separator dashes, and execution logs."""
    lines = (raw_text or "").split("\n")
    cleaned_lines = []

    for line in lines:
        l_str = line.strip()
        if not l_str:
            continue

        if re.fullmatch(r"[-_=\s\|]+", l_str):
            continue

        if any(h in l_str for h in _HEADER_SKIP_PATTERNS):
            continue

        if re.search(
            r"^\s*/\s*:\s*Female|Grade\s+Marks\s+GP|Date:\s*\(Execution|Page:\s*\d+/\d+|CONTROLLER OF EXAMINATIONS|PRINCIPAL",
            l_str,
            re.IGNORECASE,
        ):
            continue

        # Page legend row 2 ("O >=90 10 A+ 80-89.99 9 ..."), the "Sub1 Sub2 ... G.T."
        # column-header row and the wrapped title tail (",2026") repeat on every page.
        # They must never leak into the last student's chunk (name contamination).
        if re.search(
            r"^\s*O\s*>=\s*90\b|\bSub1\s+Sub2\s+Sub3\b|^\s*,\s*\d{4}\s*$|\bG\.T\.\s*$",
            l_str,
        ):
            continue

        cleaned_lines.append(l_str)

    return cleaned_lines


def split_pipe_row(raw_line, label):
    """Split row on '|' preserving empty cells for internal-only / practicals alignment."""
    if raw_line is None:
        return []
    body = raw_line.strip()
    if label and body.startswith(label):
        body = body[len(label) :]
    segs = body.split("|")
    if segs and segs[-1].strip() == "":
        segs = segs[:-1]
    return [s.strip() for s in segs]


def robust_parse_total_segment(seg):
    """
    Parse a single subject's Total-line segment preserving exact marks, grades, GP,
    credits, and special symbols ($, #, ®) as printed in the PDF.
    """
    seg = (seg or "").strip()
    if not seg:
        return "", "", "", ""

    # Normalize glued tokens like F00 -> F 0 0, P$4 -> P$ 4, Grade$ -> Grade$
    seg_norm = re.sub(r"\$([0-9]+)", r"$ \1", seg)
    seg_norm = re.sub(r"([A-Za-z0-9\+\-]+)\s+\$", r"\1$", seg_norm)
    seg_norm = re.sub(r"\bF00\b", "F 0 0", seg_norm)
    seg_norm = re.sub(r"\bF\s+00\b", "F 0 0", seg_norm)
    seg_norm = re.sub(r"\bDO\b", "0 0", seg_norm)
    seg_norm = re.sub(r"\bPS(\d+)\b", r"P$ \1", seg_norm)

    m_start = re.match(
        r"^([0-9]+(?:\s*\+\s*[0-9]+)?|[0-9\+ABNV®#\$]+|\-\-)\s+(.*)$",
        seg_norm,
    )
    if not m_start:
        tokens = seg_norm.split()
        if len(tokens) == 1:
            return tokens[0], "", "", ""
        return "", "", "", ""

    marks = m_start.group(1).strip()
    rest = m_start.group(2).strip()
    tokens = rest.split()

    grade, gp, creds = "", "", ""
    if len(tokens) == 1:
        grade = tokens[0]
    elif len(tokens) == 2:
        grade = tokens[0]
        if len(tokens[1]) == 2 and tokens[1].isdigit():
            gp, creds = tokens[1][0], tokens[1][1]
        else:
            gp = tokens[1]
    elif len(tokens) == 3:
        grade, gp, creds = tokens[0], tokens[1], tokens[2]
    elif len(tokens) >= 4:
        if tokens[1] == "$":
            grade = tokens[0] + "$"
            gp, creds = tokens[2], tokens[3]
        else:
            grade, gp, creds = tokens[0], tokens[-2], tokens[-1]

    if grade == "0" and gp:
        grade = "O"
    elif grade == "0$" and gp:
        grade = "O$"

    return marks, grade, gp, creds


_NAME_TOKEN_RE = re.compile(r"^[~/A-Za-z.'\-]+$")


def sanitize_candidate_name(name):
    """
    Keep only the genuine name part of a candidate name. Parsing stops at the first
    token that cannot belong to a name (digits, '>=', ',', ';', 'Sub1' ...), so
    stray page header / legend text can never be concatenated into the name.
    """
    kept = []
    truncated = False
    for tok in str(name or "").split():
        if _NAME_TOKEN_RE.match(tok):
            kept.append(tok)
        else:
            truncated = True
            break
    # If text had to be cut, a trailing lone grade letter (O/A/B/C/P/F) is a remnant of
    # the grade legend ("O >=90 ...") and is removed too. Untouched names are never altered.
    if truncated:
        while kept and kept[-1] in ("O", "A", "B", "C", "P", "F"):
            kept.pop()
    return re.sub(r"^/\s+", "/", " ".join(kept).strip())


def parse_student_demographics_from_chunk(chunk_lines):
    """
    Extract SNo, SAP No, Roll No, PRN No, Candidate Full Name, Gender, ABC ID,
    Is_PwD, and Is_Abeyance across ALL demographic lines of a student chunk.
    Handles multi-line wrapped names (e.g., NM College where Surname or Father/Mother
    name wraps to a second line or PRN/Name/ABC ID appear on a separate line).
    """
    if not chunk_lines:
        return None

    full_chunk_text = " ".join(chunk_lines)
    is_abeyance = bool(
        re.search(r"Result\s+in\s+abeyance", full_chunk_text, re.I)
    )

    # Collect candidate header lines (lines before/outside IA, TH, Total, CG, Semester summary)
    header_fragments = []
    for idx, l in enumerate(chunk_lines):
        l_clean = clean_dashes_and_artifacts(l).replace("|", " ").strip()
        if not l_clean:
            continue
        if re.match(r"^(IA|TH|Total)\b", l_clean):
            break
        if re.search(r"\[?\bSemester\s+[IVXLCDM\d]+", l_clean, re.I):
            continue
        # If line contains 3+ course codes, stop collecting header lines
        codes_in_line = [
            tok for tok in l_clean.split() if _COURSE_CODE_RE.match(tok)
        ]
        if len(codes_in_line) >= 3:
            # Check if any demographic token (PRN or 12-digit ABC ID) was mixed into this line
            continue
        if re.fullmatch(r"[\d\-\s\.\$#]+", l_clean) and idx > 0:
            continue
        header_fragments.append(l_clean)

    # Also check if PRN (MU... or 16-digit) or 12-digit ABC ID got pushed to a lower line in text stream
    for l in chunk_lines[len(header_fragments) :]:
        l_clean = clean_dashes_and_artifacts(l).replace("|", " ").strip()
        if re.search(r"\[?\bSemester\s+[IVXLCDM\d]+", l_clean, re.I):
            continue
        if re.search(r"\b(MU\d{10,16}|\d{12}|\d{16})\b", l_clean) and not re.match(
            r"^(IA|TH|Total)\b", l_clean
        ):
            # Line containing PRN / Name / ABC ID displaced in stream (e.g. NM Student 004, 016)
            codes_in_line = [
                tok for tok in l_clean.split() if _COURSE_CODE_RE.match(tok)
            ]
            if len(codes_in_line) < 3:
                header_fragments.append(l_clean)

    combined_header = " ".join(header_fragments)
    combined_header = re.sub(
        r"Result\s+in\s+abeyance", "", combined_header, flags=re.I
    ).strip()
    # Remove any accidental G.T. / Percentage tokens from header fragments
    combined_header = re.sub(r"\b\d{2,3}/\d{3}\b", "", combined_header)
    combined_header = re.sub(r"\b\d{2}\.\d{2}%?", "", combined_header)

    tokens = combined_header.split()
    if len(tokens) < 3:
        return None

    s_no = ""
    sap_no = ""
    roll_no = ""
    prn_no = ""
    abc_id = ""
    name_tokens = []

    # 1. Identify SNo (3-digit or 1-4 digit leading serial) & SAP No (9-12 digit number)
    used_indices = set()
    for i, tok in enumerate(tokens):
        if not s_no and re.fullmatch(r"\d{1,4}", tok):
            s_no = tok
            used_indices.add(i)
            break

    for i, tok in enumerate(tokens):
        if i in used_indices:
            continue
        if re.fullmatch(r"\d{10,11}", tok):
            sap_no = tok
            used_indices.add(i)
            break

    if not s_no or not sap_no:
        return None

    # 2. Identify Roll No (alphanumeric token like D001, A001 immediately following SAP No or SNo)
    for i, tok in enumerate(tokens):
        if i in used_indices:
            continue
        if re.fullmatch(r"[A-Z]\d{2,4}", tok):
            roll_no = tok
            used_indices.add(i)
            break

    # 3. Identify PRN No (starts with MU or 15-16 digit university PRN)
    for i, tok in enumerate(tokens):
        if i in used_indices:
            continue
        if tok.startswith("MU") or (tok.isdigit() and len(tok) >= 14):
            prn_no = tok
            used_indices.add(i)
            break

    # 4. Identify 12-digit ABC / APAAR ID (search from right to left)
    for i in range(len(tokens) - 1, -1, -1):
        if i in used_indices:
            continue
        if tokens[i].isdigit() and len(tokens[i]) == 12:
            abc_id = tokens[i]
            used_indices.add(i)
            break

    # 5. Remaining tokens form Candidate Full Name (filtering out any stray course codes)
    for i, tok in enumerate(tokens):
        if i in used_indices:
            continue
        if _COURSE_CODE_RE.match(tok):
            continue
        name_tokens.append(tok)

    full_name = " ".join(name_tokens).strip()
    # Normalize "/ SURNAME" -> "/SURNAME" and drop any non-name trailing text
    full_name = re.sub(r"^/\s+", "/", full_name)
    full_name = sanitize_candidate_name(full_name)
    gender = "Female" if "/" in full_name else "Male"
    is_pwd = bool(re.search(r"~", full_name))

    return {
        "SNo": s_no,
        "SAP No": sap_no,
        "Roll No": roll_no,
        "PRN No": prn_no,
        "Candidate Full Name": full_name,
        "Gender": gender,
        "ABC ID": abc_id,
        "Is_PwD": is_pwd,
        "Is_Abeyance": is_abeyance,
    }


def extract_subjects_universal(chunk_lines, known_course_codes=None):
    """
    Universally extracts Subject Codes, IA Marks, TH Marks, Total details,
    CG values, Grand Total, and Percentage across all SVKM exam gazette formats.
    Strictly preserves raw PDF values without auto-calculating missing entries.
    """
    subj_codes = []
    ia_scores = []
    th_scores = []
    tot_info = []
    cg_scores = []
    gt_score = ""
    percentage = ""

    # 1. Extract Grand Total (e.g. 411/550, 638/700) and Percentage (e.g. 74.73%)
    for l in chunk_lines:
        if re.search(r"\[?\bSemester\s+[IVXLCDM\d]+", l, re.I):
            continue
        gt_match = re.search(
            r"\b(\d{2,3}/\d{3})\b(?:\s*([\d\.,]+%?))?", l
        )
        if gt_match:
            gt_score = gt_match.group(1)
            if gt_match.group(2):
                percentage = gt_match.group(2)
            break

    if gt_score and not percentage:
        for l in chunk_lines:
            if re.search(r"\[?\bSemester\s+[IVXLCDM\d]+", l, re.I):
                continue
            pct_match = re.search(r"\b(\d{2}[.,]\d{2})%", l)
            if pct_match:
                percentage = pct_match.group(1) + "%"
                break

    # 2. Check if IA line has embedded course codes
    ia_with_codes = None
    for l in chunk_lines:
        if l.strip().startswith("IA") and re.search(
            r"[A-Z]{2,8}\d{2,4}[A-Z0-9]{0,3}", l
        ):
            ia_with_codes = l
            break

    if ia_with_codes:
        segments = [
            s.strip()
            for s in ia_with_codes.split("|")
            if s.strip() and not s.strip().startswith("IA")
        ]
        for seg in segments:
            if "/" in seg and any(c.isdigit() for c in seg):
                m_gt = re.search(r"(\d+/\d+)\s*([\d\.,]+%?)?", seg)
                if m_gt and not gt_score:
                    gt_score = m_gt.group(1)
                    percentage = m_gt.group(2) or ""
                seg = seg.split(gt_score)[0].strip() if gt_score else seg
            if not seg:
                continue

            m_code_ia = re.search(
                r"([A-Z0-9]{5,12})\s*([0-9\$\+ABNV®#\s\-]+)?", seg
            )
            if m_code_ia:
                code = m_code_ia.group(1).strip()
                ia_val = m_code_ia.group(2).strip() if m_code_ia.group(2) else ""
                subj_codes.append(code)
                ia_scores.append(ia_val)
    else:
        # Collect all course codes in chunk (handling displaced course codes as in NM Student 054)
        for l in chunk_lines:
            if re.search(r"\[?\bSemester\s+[IVXLCDM\d]+", l, re.I):
                continue
            tokens = l.replace("|", " ").split()
            line_codes = [
                t
                for t in tokens
                if _COURSE_CODE_RE.match(t)
                and (not known_course_codes or t in known_course_codes or len(t) >= 7)
                and not t.startswith("MU0")
            ]
            if len(line_codes) >= 3:
                subj_codes.extend(line_codes)
                break

        # Check if any additional course codes were displaced onto another line (e.g. NM Student 054)
        if subj_codes and known_course_codes:
            for l in chunk_lines:
                if re.search(r"\[?\bSemester\s+[IVXLCDM\d]+", l, re.I):
                    continue
                tokens = l.replace("|", " ").split()
                extra_codes = [
                    t
                    for t in tokens
                    if t in known_course_codes and t not in subj_codes
                ]
                if extra_codes:
                    # Re-collect from pipe positions if possible
                    pass

        for l in chunk_lines:
            if l.strip().startswith("IA"):
                ia_scores = split_pipe_row(l, "IA")
                break

    total_line_idx = None
    th_line_idx = None
    for i, l in enumerate(chunk_lines):
        l_str = l.strip()
        if l_str.startswith("TH"):
            th_scores = split_pipe_row(l_str, "TH")
            th_line_idx = i
        elif l_str.startswith("Total"):
            tot_info = split_pipe_row(l_str, "Total")
            total_line_idx = i

    # Handle wrapped grace marks on line between TH and Total (e.g., "+7" in NM Student 053)
    if (
        th_line_idx is not None
        and total_line_idx is not None
        and total_line_idx > th_line_idx + 1
    ):
        for mid_idx in range(th_line_idx + 1, total_line_idx):
            mid_line = chunk_lines[mid_idx].strip()
            if re.search(r"\+\s*\d+", mid_line):
                mid_segs = split_pipe_row(mid_line, "")
                for s_i, m_seg in enumerate(mid_segs):
                    if m_seg and s_i < len(th_scores):
                        th_scores[s_i] = f"{th_scores[s_i]} {m_seg}".strip()

    # Extract CG line immediately following Total line
    if total_line_idx is not None and total_line_idx + 1 < len(chunk_lines):
        candidate = chunk_lines[total_line_idx + 1]
        if candidate.strip() and re.fullmatch(
            r"[\d\-\s\|\.\$#]+", candidate.strip()
        ):
            cg_scores = split_pipe_row(candidate, "")

    return (
        subj_codes,
        ia_scores,
        th_scores,
        tot_info,
        cg_scores,
        gt_score,
        percentage,
    )


def extract_subjects_coordinate_aware(student_words, known_course_codes=None):
    """
    Coordinate-based subject extractor using pdfplumber word bounding boxes.
    Aligns every IA, TH, Total, and CG token to its exact Subject column (Sub1..SubN)
    by horizontal X-coordinate, preventing column shifts when TH is blank (e.g. NHCEP264),
    when grace marks wrap onto a new line (e.g. 17 +7 in Student 053), or when course
    codes drop below the Total row in the PDF stream (e.g. Student 054).
    """
    if not student_words:
        return None

    # Find all course code words in this student's vertical block
    code_words = []
    for w in student_words:
        txt = w["text"].strip().strip("|")
        if not txt or txt.startswith("MU0"):
            continue
        if known_course_codes and txt in known_course_codes:
            code_words.append((w["x0"], w["top"], txt))
        elif not known_course_codes and _COURSE_CODE_RE.match(txt) and len(txt) >= 7:
            if not re.match(r"^[AD]\d{3}$", txt):
                code_words.append((w["x0"], w["top"], txt))

    if len(code_words) < 3:
        return None

    # Sort course codes strictly left-to-right by X coordinate (deduplicating by X proximity)
    code_words.sort(key=lambda item: item[0])
    unique_codes = []
    for x0, top, code in code_words:
        if not unique_codes or abs(x0 - unique_codes[-1][0]) > 18:
            unique_codes.append((x0, top, code))

    n_subs = len(unique_codes)
    if n_subs < 3:
        return None

    subj_codes = [c[2] for c in unique_codes]
    col_x_starts = [c[0] for c in unique_codes]

    # Find G.T. token (e.g. 638/700 or 411/550) to bound the rightmost subject column
    gt_score = ""
    percentage = ""
    gt_x0 = 9999.0
    for w in student_words:
        txt = w["text"].strip().strip("|")
        m_gt = re.match(r"^(\d{2,3}/\d{3})$", txt)
        if m_gt:
            gt_score = m_gt.group(1)
            gt_x0 = min(gt_x0, w["x0"])
        elif re.match(r"^\d{2}[.,]\d{2}%$", txt):
            percentage = txt

    if not percentage:
        for w in student_words:
            txt = w["text"].strip().strip("|")
            if (
                gt_score
                and w["x0"] >= gt_x0 - 5
                and re.match(r"^\d{2}[.,]\d{2}%?$", txt)
            ):
                percentage = txt
                break

    # Compute horizontal X interval [left_b, right_b) for each subject column
    col_bounds = []
    for idx in range(n_subs):
        left_b = (
            col_x_starts[0] - 12
            if idx == 0
            else (col_x_starts[idx - 1] + col_x_starts[idx]) / 2.0
        )
        # Keep left boundary slightly left of the course code start
        if idx > 0:
            left_b = min(left_b, col_x_starts[idx] - 4)
        right_b = (
            gt_x0 - 4
            if idx == n_subs - 1
            else min(
                (col_x_starts[idx] + col_x_starts[idx + 1]) / 2.0,
                col_x_starts[idx + 1] - 4,
            )
        )
        col_bounds.append((left_b, right_b))

    # Locate vertical Y coordinates of IA, TH, Total, and Semester rows
    ia_tops = [
        w["top"]
        for w in student_words
        if w["text"].strip().strip("|") == "IA" and w["x0"] < col_x_starts[0] + 15
    ]
    th_tops = [
        w["top"]
        for w in student_words
        if w["text"].strip().strip("|") in ("TH", "HL")
        and w["x0"] < col_x_starts[0] + 15
    ]
    tot_tops = [
        w["top"]
        for w in student_words
        if w["text"].strip().strip("|") == "Total"
        and w["x0"] < col_x_starts[0] + 15
    ]
    sem_tops = [
        w["top"]
        for w in student_words
        if "Semester" in w["text"] or "Remark" in w["text"]
    ]

    if not ia_tops or not tot_tops:
        return None

    ia_y = ia_tops[0]
    tot_y = tot_tops[0]
    th_y = th_tops[0] if th_tops else (ia_y + tot_y) / 2.0
    sem_y = min(sem_tops) if sem_tops else tot_y + 25.0

    y_mid_ia_th = (ia_y + th_y) / 2.0
    y_mid_th_tot = (th_y + tot_y) / 2.0
    y_cg_min = tot_y + 4.5
    y_cg_max = min(sem_y - 2.0, tot_y + 20.0)

    ia_buckets = [[] for _ in range(n_subs)]
    th_buckets = [[] for _ in range(n_subs)]
    tot_buckets = [[] for _ in range(n_subs)]
    cg_buckets = [[] for _ in range(n_subs)]

    for w in student_words:
        txt = w["text"].strip().strip("|").strip()
        if not txt or txt in ("IA", "TH", "HL", "Total"):
            continue
        if txt in subj_codes or txt == gt_score or txt == percentage:
            continue
        wx = (w["x0"] + w["x1"]) / 2.0
        wy = w["top"]

        # Determine which subject column wx falls into
        sub_idx = None
        for idx, (lb, rb) in enumerate(col_bounds):
            if lb <= wx < rb:
                sub_idx = idx
                break
        if sub_idx is None:
            continue

        if ia_y - 5.0 <= wy < y_mid_ia_th:
            ia_buckets[sub_idx].append((w["x0"], wy, txt))
        elif y_mid_ia_th <= wy < y_mid_th_tot:
            th_buckets[sub_idx].append((w["x0"], wy, txt))
        elif y_mid_th_tot <= wy <= tot_y + 4.5:
            tot_buckets[sub_idx].append((w["x0"], wy, txt))
        elif y_cg_min < wy <= y_cg_max:
            if re.fullmatch(r"[\d\.\$#\-]+", txt):
                cg_buckets[sub_idx].append((w["x0"], wy, txt))

    ia_scores = [
        " ".join(t[2] for t in sorted(b, key=lambda x: (x[1], x[0])))
        for b in ia_buckets
    ]
    th_scores = [
        " ".join(t[2] for t in sorted(b, key=lambda x: (x[1], x[0])))
        for b in th_buckets
    ]
    tot_info = [
        " ".join(t[2] for t in sorted(b, key=lambda x: x[0]))
        for b in tot_buckets
    ]
    cg_scores = [
        " ".join(t[2] for t in sorted(b, key=lambda x: x[0]))
        for b in cg_buckets
    ]

    return (
        subj_codes,
        ia_scores,
        th_scores,
        tot_info,
        cg_scores,
        gt_score,
        percentage,
    )


def extract_semester_and_final_metrics(rem_str):
    """
    Dynamically parses semester progression metrics, individual grade value, and status
    while preserving raw PDF symbols (such as '--' or '$') as printed.
    """
    rem_clean = clean_dashes_and_artifacts(rem_str)

    sem_data = {}
    sem_tokens = re.split(r"\[?\bSemester\s+", rem_clean, flags=re.IGNORECASE)

    for token in sem_tokens[1:]:
        raw_body = token.split("]")[0]
        m_num = re.match(r"^([IVXLCDM\d]+)", raw_body, re.I)
        if not m_num:
            continue

        sem_raw = m_num.group(1).upper()
        if (
            sem_raw.endswith("C")
            and len(sem_raw) > 1
            and sem_raw[:-1] in ROMAN_ORDER
        ):
            sem_num = sem_raw[:-1]
        else:
            sem_num = sem_raw

        body = raw_body[len(m_num.group(1)) :]

        c_m = re.search(r"\bC\s*(?:[-:\.]\s*)?(--|\d+)", body)
        if not c_m:
            c_lead = re.match(r"^[-:\.]\s*(--|\d+)", body)
            c_val = c_lead.group(1) if c_lead else ""
        else:
            c_val = c_m.group(1)

        cg_m = re.search(r"\bCG\s*(?:[-:\.]\s*)?(--|\d+)", body)
        sgpa_m = re.search(
            r"\bSGPA\s*(?:[-:\.]\s*)?(--|[0-9]+[.,][0-9]+#?|AD)", body
        )
        att_m = re.search(r"Attempt[/s]*\s*[-:\.]\s*(\d{1,3})", body, re.I)
        fail_m = re.search(r"\bF\s*\(\s*(\d+)\s*\)", body, re.I)

        sem_key = f"Sem_{sem_num}"
        # '--' printed in the PDF is preserved as '--'; cells stay blank only when
        # nothing is printed for them.
        sem_data[f"{sem_key}_Credits"] = c_val
        sem_data[f"{sem_key}_CG"] = cg_m.group(1) if cg_m else ""
        sem_data[f"{sem_key}_SGPA"] = (
            sgpa_m.group(1).replace(",", ".") if sgpa_m else ""
        )
        sem_data[f"{sem_key}_Attempts"] = att_m.group(1) if att_m else ""
        sem_data[f"{sem_key}_Fail_Info"] = (
            f"F({fail_m.group(1)})" if fail_m else ""
        )

    tot_cred = re.search(r"Total\s+Credit\s*:\s*(--|\d+)?", rem_clean, re.I)
    tot_cg = re.search(r"Total\s+CG\s*:\s*(--|\d+)?", rem_clean, re.I)
    cgpa = re.search(r"\bCGPA\s*:\s*(--|[0-9\.]+)?", rem_clean, re.I)

    final_grade_m = re.search(
        r"Final\s+Grade\s*(?:[-:]\s*)?(--|[OABCDPF0][\+\-!]?)",
        rem_clean,
        re.I,
    )
    final_grade_val = final_grade_m.group(1).strip() if final_grade_m else ""
    if final_grade_val == "0":
        final_grade_val = "O"

    # Match Grade in Remark block (handles Grade-A, Grade-A+, Grade--, Grade - --)
    sem_grade_m = re.search(
        r"(?<!Final\s)Grade\s*(?:[-:\s]\s*)?(--|[OABCDPF0][\+\-!]?)(?=\s|$|\]|;|,|1|l|\)|Eligible|Not\s+Eligible|Program)",
        rem_clean,
        re.I,
    )
    extracted_grade = sem_grade_m.group(1).strip() if sem_grade_m else ""
    if extracted_grade == "0":
        extracted_grade = "O"

    cum_perc = re.search(
        r"Cum\.?\s*Perc\.?\s*:\s*(--|[0-9\.]+%?)?", rem_clean, re.I
    )

    remark_m = re.search(r"Remark\s*:\s*([^\];]+)", rem_clean, re.I)
    raw_remark = remark_m.group(1).strip() if remark_m else ""
    status_m = re.search(
        r"\b(Successful|Unsuccessful)(\s*\$+)?", raw_remark, re.I
    )
    if status_m:
        result_status = status_m.group(0).strip()
    elif (
        "Unsuccessful" in rem_clean
        or "Not Eligible" in rem_clean
        or "F(" in rem_clean
    ):
        result_status = "Unsuccessful"
    elif "Successful" in rem_clean or "Successfully" in rem_clean:
        result_status = "Successful"
    else:
        result_status = ""

    prog_m = re.search(
        r"(Program\s+Completed\s+Successfully|Eligible\s+for\s+next\s+semester|Not\s+Eligible\s+for\s+next\s+semester|Not\s+Eligible|Successfully\s+Completed)",
        rem_clean,
        re.I,
    )
    eligibility_status = prog_m.group(1).strip() if prog_m else ""

    return {
        "semesters": sem_data,
        "Total Credit": (tot_cred.group(1) if (tot_cred and tot_cred.group(1)) else ""),
        "Total CG": (tot_cg.group(1) if (tot_cg and tot_cg.group(1)) else ""),
        "CGPA": (cgpa.group(1) if (cgpa and cgpa.group(1)) else ""),
        "Final Grade": final_grade_val,
        "Cumulative Percentage": (
            cum_perc.group(1) if (cum_perc and cum_perc.group(1)) else ""
        ),
        "Result Status": result_status,
        "Grade": extracted_grade,
        "Eligibility Status": eligibility_status,
    }


def _extract_page_student_word_groups(page):
    """
    Slice a PDF result page into per-student word lists keyed by SNo or SAP No,
    enabling coordinate-aware subject column alignment.
    """
    try:
        words = page.extract_words(x_tolerance=2, y_tolerance=2) or []
    except Exception:
        return {}

    if not words:
        return {}

    # Filter out top page header and bottom page footer
    footer_top = 9999.0
    header_bottom = 0.0
    for w in words:
        txt = w["text"].strip()
        if txt in ("Female,", "Execution", "CONTROLLER") and w["top"] > 400:
            footer_top = min(footer_top, w["top"] - 4)
        if txt == "CG" and w["top"] < 175:
            header_bottom = max(header_bottom, w["bottom"] + 1)

    body_words = [
        w
        for w in words
        if w["top"] >= header_bottom and w["top"] < footer_top
    ]
    if not body_words:
        return {}

    # Identify student anchor Y-positions using SNo (leftmost 3-digit token) or 10-11 digit SAP No
    anchors = []
    for w in body_words:
        txt = w["text"].strip().strip("|")
        if w["x0"] < 52 and re.fullmatch(r"\d{3}", txt):
            anchors.append((w["top"], txt))

    anchors.sort(key=lambda x: x[0])
    # Deduplicate anchors within 15pt vertically
    dedup_anchors = []
    for top_y, sno in anchors:
        if not dedup_anchors or abs(top_y - dedup_anchors[-1][0]) > 18:
            dedup_anchors.append((top_y, sno))

    student_word_map = {}
    for idx, (top_y, sno) in enumerate(dedup_anchors):
        y_min = top_y - 8.0
        y_max = (
            dedup_anchors[idx + 1][0] - 6.0
            if idx + 1 < len(dedup_anchors)
            else footer_top
        )
        s_words = [w for w in body_words if y_min <= w["top"] < y_max]
        student_word_map[sno] = s_words

    return student_word_map


def parse_result_pages(
    pdf_pages, course_title_map=None, exam_type="Regular"
):
    """
    Extract all student records and generate the Result_Data dataset.
    Strictly preserves PDF values as-is without auto-calculating missing fields.
    Orders the highest (current) semester columns as:
    Credits, CG, SGPA, Attempts, Fail_Info.
    """
    if course_title_map is None:
        course_title_map = {}

    known_codes = set(course_title_map.keys())
    all_lines = []
    global_word_groups_by_sno = {}

    for p_idx, page in enumerate(pdf_pages):
        p_no = p_idx + 1
        p_text = page.extract_text() or ""

        if "Consolidated Result" not in p_text and not re.search(
            r"^\d{1,4}\s+\|?\s*\d{9,14}", p_text, re.M
        ):
            if len(pdf_pages) > 2 and p_no < 3:
                continue

        cleaned = clean_page_lines(p_text)
        all_lines.extend(cleaned)

        page_word_map = _extract_page_student_word_groups(page)
        global_word_groups_by_sno.update(page_word_map)

    # Group lines into student chunks
    student_chunks = []
    current_chunk = []

    new_student_line_re = re.compile(
        r"^\d{3}\b.*?\b\d{10,11}\b|^\d{1,4}\s+\d{9,14}\s+[A-Z0-9]+"
    )

    for line in all_lines:
        l_norm = re.sub(r"\s*\|\s*", " ", line).strip()
        # Check if this line starts a new student record
        is_start = bool(new_student_line_re.match(l_norm)) or bool(
            re.match(r"^\d{3}\s+(?:MU\d{10,16}|\d{10,16})\b", l_norm)
        )
        if is_start:
            if current_chunk:
                student_chunks.append(current_chunk)
                current_chunk = []
        if current_chunk or is_start:
            current_chunk.append(line)

    if current_chunk:
        student_chunks.append(current_chunk)

    detailed_records = []
    all_discovered_sem_keys = set()

    for chunk in student_chunks:
        if not chunk:
            continue

        header_info = parse_student_demographics_from_chunk(chunk)
        if not header_info:
            continue

        chunk_str = " ".join(chunk)
        is_abeyance = (
            header_info["Is_Abeyance"] or "abeyance" in chunk_str.lower()
        )

        if is_abeyance:
            detailed_row = {
                "SNo": header_info["SNo"],
                "SAP No": header_info["SAP No"],
                "Roll No": header_info["Roll No"],
                "PRN No": header_info["PRN No"],
                "Candidate Full Name": header_info["Candidate Full Name"],
                "Gender": header_info["Gender"],
                "ABC ID": header_info["ABC ID"],
                "Examination Type": exam_type,
                "Is_PwD": "Yes" if header_info["Is_PwD"] else "No",
                "Result Status": "Result in Abeyance",
                "Grade": "Result in Abeyance",
                "Eligibility Status": "Result in Abeyance",
            }
            detailed_records.append(detailed_row)
            continue

        # Primary: coordinate-aware extraction; Fallback/Complement: universal line/pipe extraction
        coord_res = extract_subjects_coordinate_aware(
            global_word_groups_by_sno.get(header_info["SNo"]),
            known_course_codes=known_codes,
        )
        pipe_res = extract_subjects_universal(
            chunk, known_course_codes=known_codes
        )

        if coord_res and len(coord_res[0]) >= len(pipe_res[0]):
            (
                subj_codes,
                ia_scores,
                th_scores,
                tot_info,
                cg_scores,
                gt_score,
                percentage,
            ) = coord_res
            if not gt_score and pipe_res[5]:
                gt_score = pipe_res[5]
            if not percentage and pipe_res[6]:
                percentage = pipe_res[6]
        else:
            (
                subj_codes,
                ia_scores,
                th_scores,
                tot_info,
                cg_scores,
                gt_score,
                percentage,
            ) = pipe_res

        rem_str = " ".join(chunk)
        eval_metrics = extract_semester_and_final_metrics(rem_str)
        all_discovered_sem_keys.update(eval_metrics["semesters"].keys())

        marks_obt, max_marks = "", ""
        if gt_score and "/" in gt_score:
            parts = gt_score.split("/")
            marks_obt = parts[0].strip()
            max_marks = parts[1].strip()

        raw_percentage_str = (
            percentage.replace("%", "").replace(",", ".").strip()
            if percentage
            else ""
        )

        detailed_row = {
            "SNo": header_info["SNo"],
            "SAP No": header_info["SAP No"],
            "Roll No": header_info["Roll No"],
            "PRN No": header_info["PRN No"],
            "Candidate Full Name": header_info["Candidate Full Name"],
            "Gender": header_info["Gender"],
            "ABC ID": header_info["ABC ID"],
            "Examination Type": exam_type,
            "Is_PwD": "Yes" if header_info["Is_PwD"] else "No",
        }

        for idx, code in enumerate(subj_codes):
            ia = clean_text_preserve_symbols(
                ia_scores[idx] if idx < len(ia_scores) else ""
            )
            th = clean_text_preserve_symbols(
                th_scores[idx] if idx < len(th_scores) else ""
            )

            t_entry = tot_info[idx] if idx < len(tot_info) else ""
            marks, grade, gp, creds = robust_parse_total_segment(t_entry)

            course_title = course_title_map.get(code, code)

            # STRICT AS-IS PDF EXTRACTION FOR CG (NO BACKGROUND CALCULATION IN PARSER)
            cg_entry = clean_text_preserve_symbols(
                cg_scores[idx] if idx < len(cg_scores) else ""
            )
            if cg_entry.isdigit():
                cg_val = int(cg_entry)
            else:
                cg_val = cg_entry

            sub_num = idx + 1
            detailed_row[f"Sub{sub_num} Code"] = code
            detailed_row[f"Sub{sub_num} Course Title"] = course_title
            detailed_row[f"Sub{sub_num} IA Marks"] = ia
            detailed_row[f"Sub{sub_num} TH Marks"] = th
            detailed_row[f"Sub{sub_num} Total Marks"] = marks
            detailed_row[f"Sub{sub_num} Grade"] = grade
            detailed_row[f"Sub{sub_num} GP"] = gp
            detailed_row[f"Sub{sub_num} Credits (C)"] = creds
            detailed_row[f"Sub{sub_num} CG"] = cg_val

        for sem_col, val in eval_metrics["semesters"].items():
            detailed_row[sem_col] = val

        detailed_row["Total Marks Obt."] = marks_obt
        detailed_row["Max Total Marks"] = max_marks
        detailed_row["Percentage"] = raw_percentage_str
        detailed_row["Total Credit"] = eval_metrics["Total Credit"]
        detailed_row["Total CG"] = eval_metrics["Total CG"]
        detailed_row["CGPA"] = eval_metrics["CGPA"]
        detailed_row["Final Grade"] = eval_metrics["Final Grade"]
        detailed_row["Cumulative Percentage"] = eval_metrics[
            "Cumulative Percentage"
        ]
        detailed_row["Result Status"] = eval_metrics["Result Status"]
        detailed_row["Grade"] = eval_metrics["Grade"]
        detailed_row["Eligibility Status"] = eval_metrics["Eligibility Status"]

        detailed_records.append(detailed_row)

    # Identify all discovered semesters and determine highest (current) semester
    discovered_sems = sorted(
        {
            m.group(1).upper()
            for k in all_discovered_sem_keys
            if (m := re.match(r"^Sem_([IVXLCDM\d]+)_", k))
        },
        key=sem_rank,
    )
    highest_sem = discovered_sems[-1] if discovered_sems else None

    # Ensure all 5 required columns exist for the highest (current) semester
    highest_sem_required_cols = []
    if highest_sem:
        for sfx in ["Credits", "CG", "SGPA", "Attempts", "Fail_Info"]:
            col_k = f"Sem_{highest_sem}_{sfx}"
            all_discovered_sem_keys.add(col_k)
            highest_sem_required_cols.append(col_k)

    # Column order inside semester groups:
    # Highest semester strictly ordered as: Credits, CG, SGPA, Attempts, Fail_Info
    highest_sem_suffix_order = [
        "Credits",
        "CG",
        "SGPA",
        "Attempts",
        "Fail_Info",
    ]
    lower_sem_suffix_order = [
        "Credits",
        "CG",
        "SGPA",
        "Attempts",
        "Fail_Info",
    ]

    def sem_col_sort_key(col_name):
        m = re.match(r"^Sem_([IVXLCDM\d]+)_(.*)$", col_name)
        sem_name = m.group(1).upper() if m else ""
        suffix = m.group(2) if m else ""
        order_list = (
            highest_sem_suffix_order
            if sem_name == highest_sem
            else lower_sem_suffix_order
        )
        suffix_rank = (
            order_list.index(suffix)
            if suffix in order_list
            else len(order_list)
        )
        return (sem_rank(sem_name), sem_name, suffix_rank, col_name)

    sorted_sem_cols = sorted(all_discovered_sem_keys, key=sem_col_sort_key)
    df_detailed = pd.DataFrame(detailed_records)

    if not df_detailed.empty:
        for req_c in highest_sem_required_cols:
            if req_c not in df_detailed.columns:
                df_detailed[req_c] = ""

        base_id_cols = [
            c
            for c in [
                "SNo",
                "SAP No",
                "Roll No",
                "PRN No",
                "Candidate Full Name",
                "Gender",
                "ABC ID",
                "Examination Type",
                "Is_PwD",
            ]
            if c in df_detailed.columns
        ]

        sub_cols = [
            c for c in df_detailed.columns if re.match(r"^Sub\d+\s", str(c))
        ]

        gt_cols = [
            c
            for c in [
                "Grade",
                "Total Marks Obt.",
                "Max Total Marks",
                "Percentage",
            ]
            if c in df_detailed.columns
        ]

        final_result_cols = [
            c
            for c in [
                "Total Credit",
                "Total CG",
                "CGPA",
                "Final Grade",
                "Cumulative Percentage",
                "Result Status",
                "Eligibility Status",
            ]
            if c in df_detailed.columns
        ]

        ordered_cols = (
            base_id_cols
            + sub_cols
            + [c for c in sorted_sem_cols if c in df_detailed.columns]
            + gt_cols
            + final_result_cols
        )
        all_ordered = ordered_cols + [
            c for c in df_detailed.columns if c not in ordered_cols
        ]
        df_detailed = df_detailed.reindex(columns=all_ordered).fillna("")

        # Keep all non-empty columns PLUS all 5 required columns of the highest semester
        non_empty_cols = [
            c
            for c in df_detailed.columns
            if c in highest_sem_required_cols
            or not df_detailed[c]
            .fillna("")
            .astype(str)
            .str.strip()
            .isin(["", "Result in Abeyance"])
            .all()
            or c in ["Result Status", "Grade"]
        ]
        df_detailed = df_detailed[non_empty_cols]

    return df_detailed


# ==============================================================================
# BACKGROUND RESULT VERIFICATION AND AUDIT ENGINE
# ==============================================================================


def verify_and_audit_results(df_result, df_p2):
    """
    Comprehensive Background Result Verification and Audit Engine.
    All expected marks, grades, GP, credits, CG, SGPA, CGPA, percentages, and
    statuses are calculated strictly in the background without altering the
    extracted PDF values in Result_Data. Only discrepancies and anomalies are
    flagged and returned for cell highlighting.
    """
    if df_result is None or df_result.empty:
        empty_report = pd.DataFrame(
            columns=[
                "Student SAP No",
                "Roll No",
                "Candidate Name",
                "Field Name",
                "Expected Value",
                "Actual Value",
                "Error Description",
                "Severity",
            ]
        )
        empty_summary = pd.DataFrame(columns=["Metric", "Count"])
        empty_cat = pd.DataFrame(
            columns=["Error Category", "Critical", "Warning", "Total"]
        )
        return df_result, empty_report, empty_summary, empty_cat, set(), {}

    df_audited = df_result.copy()

    # Build Page 2 course schema lookup dictionary
    scheme_map = {}
    if df_p2 is not None and not df_p2.empty and "Course Code" in df_p2.columns:
        for _, crow in df_p2.iterrows():
            c_code = str(crow["Course Code"]).strip()
            scheme_map[c_code] = {
                "title": str(crow.get("Course Title", c_code)).strip(),
                "credits": to_decimal(crow.get("Credits", 0)) or Decimal("0"),
                "int_max": to_decimal(crow.get("Internal Max Marks", 0))
                or Decimal("0"),
                "sem_max": to_decimal(crow.get("Sem End Max Marks", 0))
                or Decimal("0"),
                "tot_max": to_decimal(crow.get("Total Max Marks", 0))
                or Decimal("0"),
            }

    cols = list(df_audited.columns)
    sub_nums = sorted(
        {
            int(m.group(1))
            for c in cols
            if (m := re.match(r"^Sub(\d+)\s", str(c)))
        }
    )

    sem_prefixes = []
    for c in cols:
        m = re.match(r"^Sem_([IVXLCDM\d]+)_", str(c))
        if m and m.group(1) not in sem_prefixes:
            sem_prefixes.append(m.group(1))

    sem_prefixes.sort(key=sem_rank)
    highest_sem = sem_prefixes[-1] if sem_prefixes else None
    lower_sems = sem_prefixes[:-1] if len(sem_prefixes) > 1 else []

    has_cumulative_cols = any(
        c in df_audited.columns
        for c in ["Total Credit", "Total CG", "CGPA", "Final Grade"]
    )

    report_rows = []
    error_cells_set = set()
    error_messages_map = defaultdict(list)
    category_counter = Counter()
    category_severity_counter = Counter()
    verification_errors_col = []
    students_with_errors_count = 0

    for row_idx, row in df_audited.iterrows():
        sap_no = str(row.get("SAP No", "")).strip()
        roll_no = str(row.get("Roll No", "")).strip()
        cand_name = str(row.get("Candidate Full Name", "")).strip()
        row_error_messages = []

        def record_issue(
            field_name,
            expected_val,
            actual_val,
            description,
            severity="Critical",
            category="Validation Error",
        ):
            if field_name in df_audited.columns:
                error_cells_set.add((row_idx, field_name))
                error_messages_map[(row_idx, field_name)].append(
                    f"[{severity}] {description} (Expected: {expected_val}, Actual: {actual_val if actual_val != '' else 'Blank'})"
                )
            error_cells_set.add((row_idx, "Verification Errors"))
            msg_str = f"[{severity}] {field_name}: {description} (Expected: {expected_val}, Actual: {actual_val})"
            row_error_messages.append(msg_str)
            error_messages_map[(row_idx, "Verification Errors")].append(msg_str)
            report_rows.append(
                {
                    "Student SAP No": sap_no,
                    "Roll No": roll_no,
                    "Candidate Name": cand_name,
                    "Field Name": field_name,
                    "Expected Value": str(expected_val),
                    "Actual Value": str(
                        actual_val if actual_val != "" else "Missing/Blank"
                    ),
                    "Error Description": description,
                    "Severity": severity,
                }
            )
            category_counter[category] += 1
            category_severity_counter[(category, severity)] += 1

        # ------------------------------------------------------------------
        # 1. MANDATORY DEMOGRAPHIC / ID FIELD VALIDATIONS & SPECIAL CASES
        # ------------------------------------------------------------------
        for id_col in ["SAP No", "Roll No", "PRN No", "Candidate Full Name"]:
            if id_col in df_audited.columns and not str(row.get(id_col, "")).strip():
                record_issue(
                    id_col,
                    "Non-empty value",
                    "",
                    f"Missing mandatory student identifier '{id_col}'.",
                    severity="Critical",
                    category="Missing Mandatory Value",
                )

        if "Candidate Full Name" in df_audited.columns and cand_name:
            clean_nm = sanitize_candidate_name(cand_name)
            if clean_nm != cand_name:
                record_issue(
                    "Candidate Full Name",
                    clean_nm or "Alphabetic candidate name",
                    cand_name,
                    "Candidate Full Name contains non-name text (digits / page header / legend text) picked up from the PDF.",
                    severity="Critical",
                    category="Data Contamination / Format Anomaly",
                )

        if "ABC ID" in df_audited.columns:
            abc_val = str(row.get("ABC ID", "")).strip()
            if not abc_val:
                m_trailing_digits = re.search(r"\b(\d{10,15})\b$", cand_name)
                if m_trailing_digits:
                    bad_id = m_trailing_digits.group(1)
                    record_issue(
                        "ABC ID",
                        "12-digit APAAR/ABC ID",
                        bad_id,
                        f"Malformed {len(bad_id)}-digit APAAR/ABC ID '{bad_id}' in PDF caused ABC ID to be missing.",
                        severity="Critical",
                        category="Missing / Malformed Mandatory Value",
                    )
                    record_issue(
                        "Candidate Full Name",
                        cand_name[: m_trailing_digits.start()].strip(),
                        cand_name,
                        f"Unparsed {len(bad_id)}-digit APAAR/ABC ID '{bad_id}' concatenated into Candidate Full Name.",
                        severity="Critical",
                        category="Data Contamination / Format Anomaly",
                    )
                else:
                    record_issue(
                        "ABC ID",
                        "12-digit APAAR/ABC ID",
                        "",
                        "Missing mandatory 12-digit ABC / APAAR ID.",
                        severity="Warning",
                        category="Missing / Malformed Mandatory Value",
                    )

        # Result in Abeyance special case
        status_raw = str(row.get("Result Status", "")).strip()
        is_abeyance = "abeyance" in status_raw.lower()
        if is_abeyance:
            for ab_col in ["Result Status", "Grade", "Eligibility Status"]:
                if ab_col in df_audited.columns:
                    ab_val = str(row.get(ab_col, "")).strip()
                    if "abeyance" not in ab_val.lower():
                        record_issue(
                            ab_col,
                            "Result in Abeyance",
                            ab_val,
                            f"Student is in Result in Abeyance, but '{ab_col}' is inconsistent.",
                            severity="Critical",
                            category="Result in Abeyance Inconsistency",
                        )
            record_issue(
                "Result Status",
                "Standard Declared Result",
                status_raw,
                "Special Case: Candidate result is held in abeyance ('Result in Abeyance').",
                severity="Warning",
                category="Special Case: Result in Abeyance",
            )
            if row_error_messages:
                students_with_errors_count += 1
            verification_errors_col.append(" | ".join(row_error_messages))
            continue

        # ------------------------------------------------------------------
        # 2. SUBJECT-LEVEL INDEPENDENT BACKGROUND VALIDATIONS (Sub1..SubN)
        # ------------------------------------------------------------------
        sum_sub_total_marks = Decimal("0")
        sum_sub_max_marks = Decimal("0")
        sum_sub_scheme_credits = Decimal("0")
        sum_sub_actual_credits = Decimal("0")
        sum_sub_cg = Decimal("0")
        failed_subject_count = 0

        for n in sub_nums:
            col_code = f"Sub{n} Code"
            col_ia = f"Sub{n} IA Marks"
            col_th = f"Sub{n} TH Marks"
            col_tot = f"Sub{n} Total Marks"
            col_grd = f"Sub{n} Grade"
            col_gp = f"Sub{n} GP"
            col_cred = f"Sub{n} Credits (C)"
            col_cg = f"Sub{n} CG"

            sub_code = str(row.get(col_code, "")).strip()
            if not sub_code:
                record_issue(
                    col_code,
                    "Valid Course Code",
                    "",
                    f"Missing course code for Sub{n}.",
                    severity="Critical",
                    category="Missing Mandatory Value",
                )
                continue

            scheme = scheme_map.get(sub_code)
            if scheme is None:
                record_issue(
                    col_code,
                    "Course Code from Page 2",
                    sub_code,
                    f"Subject code '{sub_code}' does not exist in Page 2 Course Scheme.",
                    severity="Critical",
                    category="Course Scheme Mismatch",
                )
                scheme_credits = Decimal("0")
                int_max = Decimal("100")
                sem_max = Decimal("100")
                tot_max = Decimal("100")
            else:
                scheme_credits = scheme["credits"]
                int_max = scheme["int_max"]
                sem_max = scheme["sem_max"]
                tot_max = scheme["tot_max"]

            sum_sub_max_marks += tot_max
            sum_sub_scheme_credits += scheme_credits

            ia_info = parse_mark_component(row.get(col_ia, ""))
            th_info = parse_mark_component(row.get(col_th, ""))
            tot_info = parse_mark_component(row.get(col_tot, ""))

            # Validate IA Marks
            if ia_info["is_empty"] and int_max > Decimal("0"):
                record_issue(
                    col_ia,
                    f"0 to {int(int_max)} or AB",
                    ia_info["raw"],
                    f"Missing IA Marks for {sub_code}.",
                    severity="Critical",
                    category="Missing Mandatory Value",
                )
            elif ia_info["is_absent"]:
                pass
            elif ia_info["is_nv"]:
                record_issue(
                    col_ia,
                    f"Numeric Marks (0-{int(int_max)}) or AB",
                    ia_info["raw"],
                    f"Special Case NV (Null & Void) recorded in {sub_code} IA Marks.",
                    severity="Warning",
                    category="Special Case: NV",
                )
            elif ia_info["is_valid_number"]:
                if (
                    ia_info["numeric_marks"] < Decimal("0")
                    or ia_info["numeric_marks"] > int_max
                    or (ia_info["has_grace"] and ia_info["base_marks"] < Decimal("0"))
                ):
                    record_issue(
                        col_ia,
                        f"0 to {int(int_max)}",
                        ia_info["raw"],
                        f"Impossible IA Marks ({ia_info['raw']}) for {sub_code}; valid range is 0 to {int(int_max)}.",
                        severity="Critical",
                        category="Impossible Marks / Out of Range",
                    )
            else:
                record_issue(
                    col_ia,
                    f"0 to {int(int_max)} or AB",
                    ia_info["raw"],
                    f"Invalid non-numeric IA Marks '{ia_info['raw']}' in {sub_code}.",
                    severity="Critical",
                    category="Invalid Data Format",
                )

            # Validate TH Marks
            if sem_max == Decimal("0"):
                if not th_info["is_empty"]:
                    record_issue(
                        col_th,
                        "Blank (Internal-only course)",
                        th_info["raw"],
                        f"{sub_code} has 0 Sem End Max Marks on Page 2, but TH Marks contains '{th_info['raw']}'.",
                        severity="Critical",
                        category="Impossible Marks / Out of Range",
                    )
            else:
                if th_info["is_empty"]:
                    record_issue(
                        col_th,
                        f"0 to {int(sem_max)} or AB",
                        th_info["raw"],
                        f"Missing TH Marks for {sub_code}.",
                        severity="Critical",
                        category="Missing Mandatory Value",
                    )
                elif th_info["is_absent"]:
                    pass
                elif th_info["is_nv"]:
                    record_issue(
                        col_th,
                        f"Numeric Marks (0-{int(sem_max)}) or AB",
                        th_info["raw"],
                        f"Special Case NV (Null & Void) recorded in {sub_code} TH Marks.",
                        severity="Warning",
                        category="Special Case: NV",
                    )
                elif th_info["is_valid_number"]:
                    if (
                        th_info["numeric_marks"] < Decimal("0")
                        or th_info["numeric_marks"] > sem_max
                        or (th_info["has_grace"] and th_info["base_marks"] < Decimal("0"))
                    ):
                        record_issue(
                            col_th,
                            f"0 to {int(sem_max)}",
                            th_info["raw"],
                            f"Impossible TH Marks ({th_info['raw']}) for {sub_code}; valid range is 0 to {int(sem_max)}.",
                            severity="Critical",
                            category="Impossible Marks / Out of Range",
                        )
                else:
                    record_issue(
                        col_th,
                        f"0 to {int(sem_max)} or AB",
                        th_info["raw"],
                        f"Invalid non-numeric TH Marks '{th_info['raw']}' in {sub_code}.",
                        severity="Critical",
                        category="Invalid Data Format",
                    )

            # Compute Expected Subject Total Marks in background
            ia_val = (
                ia_info["numeric_marks"]
                if ia_info["numeric_marks"] is not None
                else Decimal("0")
            )
            th_val = (
                th_info["numeric_marks"]
                if th_info["numeric_marks"] is not None
                else Decimal("0")
            )
            expected_sub_total = ia_val + th_val

            if tot_info["is_empty"] or tot_info["numeric_marks"] is None:
                record_issue(
                    col_tot,
                    str(int(expected_sub_total)),
                    tot_info["raw"],
                    f"Missing or invalid Total Marks for {sub_code}.",
                    severity="Critical",
                    category="Subject Total Calculation Mismatch",
                )
                actual_sub_total = expected_sub_total
            else:
                actual_sub_total = tot_info["numeric_marks"]
                if actual_sub_total < Decimal("0") or actual_sub_total > tot_max:
                    record_issue(
                        col_tot,
                        f"0 to {int(tot_max)}",
                        tot_info["raw"],
                        f"Impossible Subject Total Marks ({tot_info['raw']}) for {sub_code}; exceeds Max Marks ({int(tot_max)}).",
                        severity="Critical",
                        category="Impossible Marks / Out of Range",
                    )
                if actual_sub_total != expected_sub_total:
                    record_issue(
                        col_tot,
                        str(int(expected_sub_total)),
                        tot_info["raw"],
                        f"Subject Total Marks mismatch in {sub_code}: IA ({ia_info['raw']}) + TH ({th_info['raw']}) = {int(expected_sub_total)}, but found {tot_info['raw']}.",
                        severity="Critical",
                        category="Subject Total Calculation Mismatch",
                    )

            sum_sub_total_marks += actual_sub_total

            # Validate Subject Grade (40% component passing + unrounded percentage mapping)
            raw_sub_grade = str(row.get(col_grd, "")).strip()
            actual_sub_grade, _ = normalize_grade_token(raw_sub_grade)

            ia_min_pass = (
                (int_max * Decimal("0.40"))
                if int_max > Decimal("0")
                else Decimal("0")
            )
            th_min_pass = (
                (sem_max * Decimal("0.40"))
                if sem_max > Decimal("0")
                else Decimal("0")
            )
            failed_ia_component = int_max > Decimal("0") and (
                ia_info["is_absent"] or ia_info["is_nv"] or ia_val < ia_min_pass
            )
            failed_th_component = sem_max > Decimal("0") and (
                th_info["is_absent"] or th_info["is_nv"] or th_val < th_min_pass
            )

            if failed_ia_component or failed_th_component:
                expected_sub_grade = "F"
                _, sub_unrounded_pct = derive_grade_from_unrounded_percentage(
                    actual_sub_total, tot_max
                )
            else:
                expected_sub_grade, sub_unrounded_pct = (
                    derive_grade_from_unrounded_percentage(
                        actual_sub_total, tot_max
                    )
                )

            if not actual_sub_grade:
                record_issue(
                    col_grd,
                    expected_sub_grade or "Valid Grade",
                    raw_sub_grade,
                    f"Missing Subject Grade for {sub_code}.",
                    severity="Critical",
                    category="Subject Grade Mismatch",
                )
            elif actual_sub_grade not in VALID_GRADES:
                record_issue(
                    col_grd,
                    expected_sub_grade or "O/A+/A/B+/B/C/P/F",
                    raw_sub_grade,
                    f"Invalid Subject Grade '{raw_sub_grade}' in {sub_code}.",
                    severity="Critical",
                    category="Subject Grade Mismatch",
                )
            elif expected_sub_grade and actual_sub_grade != expected_sub_grade:
                pct_str = (
                    f"{sub_unrounded_pct:.2f}%"
                    if sub_unrounded_pct is not None
                    else "N/A"
                )
                record_issue(
                    col_grd,
                    expected_sub_grade,
                    raw_sub_grade,
                    f"Subject Grade mismatch in {sub_code}: Marks {int(actual_sub_total)}/{int(tot_max)} ({pct_str}, IA={ia_info['raw']}, TH={th_info['raw']}) maps to '{expected_sub_grade}', but found '{raw_sub_grade}'.",
                    severity="Critical",
                    category="Subject Grade Mismatch",
                )

            is_sub_fail = actual_sub_grade == "F" or expected_sub_grade == "F"
            if is_sub_fail:
                failed_subject_count += 1

            # Validate Subject GP
            raw_gp = (
                str(row.get(col_gp, ""))
                .strip()
                .replace("$", "")
                .replace("#", "")
            )
            actual_gp = to_decimal(raw_gp)
            expected_gp = Decimal(
                str(
                    GRADE_POINTS_MAP.get(
                        actual_sub_grade or expected_sub_grade, 0
                    )
                )
            )

            if actual_gp is None:
                record_issue(
                    col_gp,
                    str(int(expected_gp)),
                    str(row.get(col_gp, "")),
                    f"Missing or non-numeric Subject GP for {sub_code}.",
                    severity="Critical",
                    category="Subject GP Mismatch",
                )
            else:
                if actual_gp < Decimal("0") or actual_gp > Decimal("10"):
                    record_issue(
                        col_gp,
                        "0 to 10",
                        str(row.get(col_gp, "")),
                        f"Impossible GP ({actual_gp}) in {sub_code}; GP must be between 0 and 10.",
                        severity="Critical",
                        category="Impossible Marks / Out of Range",
                    )
                if actual_sub_grade == "F" and actual_gp != Decimal("0"):
                    record_issue(
                        col_gp,
                        "0",
                        str(row.get(col_gp, "")),
                        f"Subject Grade is F in {sub_code}, so Subject GP must be 0 (found {actual_gp}).",
                        severity="Critical",
                        category="Subject GP Mismatch",
                    )
                elif (
                    actual_sub_grade
                    and actual_sub_grade != "F"
                    and actual_gp == Decimal("0")
                ):
                    record_issue(
                        col_gp,
                        str(int(expected_gp)),
                        str(row.get(col_gp, "")),
                        f"Subject Grade is '{actual_sub_grade}' (not F) in {sub_code}, so GP cannot be 0.",
                        severity="Critical",
                        category="Subject GP Mismatch",
                    )
                elif actual_gp != expected_gp:
                    record_issue(
                        col_gp,
                        str(int(expected_gp)),
                        str(row.get(col_gp, "")),
                        f"Subject GP ({int(actual_gp)}) does not match Grade '{actual_sub_grade}' (Expected GP = {int(expected_gp)}) in {sub_code}.",
                        severity="Critical",
                        category="Subject GP Mismatch",
                    )

            # Validate Subject Credits (C) against Page 2 Scheme
            raw_cred = (
                str(row.get(col_cred, ""))
                .strip()
                .replace("$", "")
                .replace("#", "")
            )
            actual_cred = to_decimal(raw_cred)
            if actual_cred is None:
                record_issue(
                    col_cred,
                    "0" if is_sub_fail else str(int(scheme_credits)),
                    str(row.get(col_cred, "")),
                    f"Missing Subject Credits (C) for {sub_code} (Page 2 defines {int(scheme_credits)} credits).",
                    severity="Critical",
                    category="Subject Credit Scheme Mismatch",
                )
            else:
                sum_sub_actual_credits += actual_cred
                if is_sub_fail:
                    if scheme is not None and actual_cred not in (
                        Decimal("0"),
                        scheme_credits,
                    ):
                        record_issue(
                            col_cred,
                            "0",
                            str(row.get(col_cred, "")),
                            f"Subject Grade is F in {sub_code}, so Subject Credits (C) must be 0 (found {actual_cred}).",
                            severity="Critical",
                            category="Subject Credit Scheme Mismatch",
                        )
                else:
                    if scheme is not None and actual_cred != scheme_credits:
                        record_issue(
                            col_cred,
                            str(int(scheme_credits)),
                            str(row.get(col_cred, "")),
                            f"Subject Credits ({actual_cred}) in {sub_code} does not match Page 2 course scheme credits ({int(scheme_credits)}).",
                            severity="Critical",
                            category="Subject Credit Scheme Mismatch",
                        )

            # Validate Subject CG = Subject GP * Subject Credit in background
            raw_cg = (
                str(row.get(col_cg, ""))
                .strip()
                .replace("$", "")
                .replace("#", "")
            )
            actual_cg = to_decimal(raw_cg)
            if is_sub_fail:
                expected_cg = Decimal("0")
            else:
                gp_for_cg = actual_gp if actual_gp is not None else expected_gp
                cred_for_cg = (
                    scheme_credits
                    if scheme is not None
                    else (actual_cred or Decimal("0"))
                )
                expected_cg = gp_for_cg * cred_for_cg

            if actual_cg is None:
                record_issue(
                    col_cg,
                    str(int(expected_cg)),
                    str(row.get(col_cg, "")),
                    f"Missing Subject CG in PDF for {sub_code} (Expected {int(expected_cg)}).",
                    severity="Critical",
                    category="Subject CG Calculation Mismatch",
                )
            else:
                sum_sub_cg += actual_cg
                if is_sub_fail and actual_cg != Decimal("0"):
                    record_issue(
                        col_cg,
                        "0",
                        str(row.get(col_cg, "")),
                        f"Subject Grade is F in {sub_code}, so Subject CG must be 0 (found {int(actual_cg)}).",
                        severity="Critical",
                        category="Subject CG Calculation Mismatch",
                    )
                elif (
                    actual_sub_grade
                    and actual_sub_grade != "F"
                    and actual_cg == Decimal("0")
                ):
                    record_issue(
                        col_cg,
                        str(int(expected_cg)),
                        str(row.get(col_cg, "")),
                        f"Subject Grade is '{actual_sub_grade}' (not F) in {sub_code}, so Subject CG cannot be 0.",
                        severity="Critical",
                        category="Subject CG Calculation Mismatch",
                    )
                elif actual_cg != expected_cg:
                    record_issue(
                        col_cg,
                        str(int(expected_cg)),
                        str(row.get(col_cg, "")),
                        f"Subject CG mismatch in {sub_code}: GP ({int(actual_gp or 0)}) x Credit ({int(scheme_credits)}) = {int(expected_cg)}, but found {int(actual_cg)}.",
                        severity="Critical",
                        category="Subject CG Calculation Mismatch",
                    )

        # ------------------------------------------------------------------
        # 3. SEMESTER-LEVEL BACKGROUND VALIDATIONS
        # ------------------------------------------------------------------
        lower_sem_fail_counts = {}
        for s in lower_sems:
            f_col = f"Sem_{s}_Fail_Info"
            if f_col in df_audited.columns:
                f_val = str(row.get(f_col, "")).strip()
                m_f = re.search(r"F\s*\(\s*(\d+)\s*\)", f_val, re.I)
                if m_f and int(m_f.group(1)) > 0:
                    lower_sem_fail_counts[s] = int(m_f.group(1))
        # Attendance Defaulter ('AD') in a lower semester also means that semester is not cleared
        lower_sem_ad_sems = []
        for s in lower_sems:
            sg_col = f"Sem_{s}_SGPA"
            if sg_col in df_audited.columns and str(row.get(sg_col, "")).strip().upper() == "AD":
                lower_sem_ad_sems.append(s)
        has_lower_sem_fail = len(lower_sem_fail_counts) > 0 or len(lower_sem_ad_sems) > 0

        if highest_sem:
            col_sem_cred = f"Sem_{highest_sem}_Credits"
            col_sem_cg = f"Sem_{highest_sem}_CG"
            col_sem_sgpa = f"Sem_{highest_sem}_SGPA"
            col_sem_fail = f"Sem_{highest_sem}_Fail_Info"

            actual_sem_cred = to_decimal(row.get(col_sem_cred, ""))
            actual_sem_cg = to_decimal(row.get(col_sem_cg, ""))
            actual_sem_sgpa = to_decimal(row.get(col_sem_sgpa, ""))
            actual_sem_fail = (
                str(row.get(col_sem_fail, "")).strip()
                if col_sem_fail in df_audited.columns
                else ""
            )

            if col_sem_cred in df_audited.columns:
                if failed_subject_count > 0:
                    if actual_sem_cred is not None and actual_sem_cred not in (
                        sum_sub_scheme_credits,
                        sum_sub_actual_credits,
                    ):
                        record_issue(
                            col_sem_cred,
                            "Blank (--) or Earned Credits",
                            str(row.get(col_sem_cred, "")),
                            f"Unexpected {col_sem_cred} ({int(actual_sem_cred)}) for student with {failed_subject_count} failed subject(s).",
                            severity="Critical",
                            category="Semester Credits Mismatch",
                        )
                else:
                    if actual_sem_cred is None:
                        record_issue(
                            col_sem_cred,
                            str(int(sum_sub_scheme_credits)),
                            str(row.get(col_sem_cred, "")),
                            f"Missing {col_sem_cred}; expected sum of subject credits = {int(sum_sub_scheme_credits)}.",
                            severity="Critical",
                            category="Missing Mandatory Semester Summary",
                        )
                    elif actual_sem_cred != sum_sub_scheme_credits:
                        record_issue(
                            col_sem_cred,
                            str(int(sum_sub_scheme_credits)),
                            str(row.get(col_sem_cred, "")),
                            f"{col_sem_cred} ({int(actual_sem_cred)}) does not equal sum of subject credits ({int(sum_sub_scheme_credits)}).",
                            severity="Critical",
                            category="Semester Credits Mismatch",
                        )

            if col_sem_cg in df_audited.columns:
                if failed_subject_count > 0:
                    if actual_sem_cg is not None and actual_sem_cg != sum_sub_cg:
                        record_issue(
                            col_sem_cg,
                            "Blank (--)",
                            str(row.get(col_sem_cg, "")),
                            f"{col_sem_cg} ({int(actual_sem_cg)}) does not match sum of subject CG values ({int(sum_sub_cg)}).",
                            severity="Critical",
                            category="Semester CG Mismatch",
                        )
                else:
                    if actual_sem_cg is None:
                        record_issue(
                            col_sem_cg,
                            str(int(sum_sub_cg)),
                            str(row.get(col_sem_cg, "")),
                            f"Missing {col_sem_cg}; expected sum of subject CG values = {int(sum_sub_cg)}.",
                            severity="Critical",
                            category="Missing Mandatory Semester Summary",
                        )
                    elif actual_sem_cg != sum_sub_cg:
                        record_issue(
                            col_sem_cg,
                            str(int(sum_sub_cg)),
                            str(row.get(col_sem_cg, "")),
                            f"{col_sem_cg} ({int(actual_sem_cg)}) does not equal sum of subject CG values ({int(sum_sub_cg)}).",
                            severity="Critical",
                            category="Semester CG Mismatch",
                        )

            expected_sgpa = (
                (sum_sub_cg / sum_sub_scheme_credits).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP
                )
                if sum_sub_scheme_credits > Decimal("0")
                else None
            )

            if col_sem_sgpa in df_audited.columns:
                if failed_subject_count > 0:
                    if actual_sem_sgpa is not None:
                        record_issue(
                            col_sem_sgpa,
                            "Blank (--)",
                            str(row.get(col_sem_sgpa, "")),
                            f"Student has {failed_subject_count} failed subject(s), so {col_sem_sgpa} should be blank ('--'), but found {actual_sem_sgpa}.",
                            severity="Critical",
                            category="Semester SGPA Mismatch",
                        )
                else:
                    if actual_sem_sgpa is None:
                        record_issue(
                            col_sem_sgpa,
                            (
                                f"{expected_sgpa:.2f}"
                                if expected_sgpa is not None
                                else "Valid SGPA"
                            ),
                            str(row.get(col_sem_sgpa, "")),
                            f"Missing {col_sem_sgpa}; recomputed CG ({int(sum_sub_cg)}) / Credits ({int(sum_sub_scheme_credits)}) = {expected_sgpa:.2f}.",
                            severity="Critical",
                            category="Missing Mandatory Semester Summary",
                        )
                    elif expected_sgpa is not None:
                        sgpa_has_hash = "#" in str(row.get(col_sem_sgpa, ""))
                        if sgpa_has_hash:
                            # '#': eligible for O.229 grace -> unused grace marks are added to SGPA
                            g_used, g_unused, g_bonus = grace_sgpa_info(row)
                            raw_sgpa = sum_sub_cg / sum_sub_scheme_credits
                            candidates = [
                                min(expected_sgpa + g_bonus, Decimal("10.00")),
                                min(
                                    (raw_sgpa + g_bonus).quantize(
                                        Decimal("0.01"), rounding=ROUND_HALF_UP
                                    ),
                                    Decimal("10.00"),
                                ),
                            ]
                            if all(
                                abs(actual_sem_sgpa - c) >= Decimal("0.005")
                                for c in candidates
                            ):
                                record_issue(
                                    col_sem_sgpa,
                                    f"{candidates[0]:.2f}",
                                    str(row.get(col_sem_sgpa, "")),
                                    f"{col_sem_sgpa} mismatch for O.229 ('#') candidate: {int(sum_sub_cg)} / {int(sum_sub_scheme_credits)} = {expected_sgpa:.2f}, plus unused grace bonus {g_bonus:.2f} ({g_unused} of {GRACE_MAX_MARKS} grace marks unused) = {candidates[0]:.2f}, but found {actual_sem_sgpa:.2f}.",
                                    severity="Critical",
                                    category="Semester SGPA Mismatch",
                                )
                        elif abs(actual_sem_sgpa - expected_sgpa) > Decimal("0.01"):
                            record_issue(
                                col_sem_sgpa,
                                f"{expected_sgpa:.2f}",
                                str(row.get(col_sem_sgpa, "")),
                                f"{col_sem_sgpa} mismatch: {int(sum_sub_cg)} / {int(sum_sub_scheme_credits)} = {expected_sgpa:.2f}, but found {actual_sem_sgpa:.2f}.",
                                severity="Critical",
                                category="Semester SGPA Mismatch",
                            )

            expected_fail_info = (
                f"F({failed_subject_count})" if failed_subject_count > 0 else ""
            )
            if col_sem_fail in df_audited.columns:
                if actual_sem_fail != expected_fail_info:
                    record_issue(
                        col_sem_fail,
                        (
                            expected_fail_info
                            if expected_fail_info
                            else "Blank (0 Failed Subjects)"
                        ),
                        actual_sem_fail,
                        f"Semester Fail_Info mismatch: Student has {failed_subject_count} failed subject(s), expected '{expected_fail_info}', found '{actual_sem_fail}'.",
                        severity="Critical",
                        category="Fail_Info Mismatch",
                    )

            if "Grade" in df_audited.columns:
                actual_sem_grade_raw = str(row.get("Grade", "")).strip()
                actual_sem_grade, _ = normalize_grade_token(
                    actual_sem_grade_raw
                )
                if failed_subject_count > 0:
                    if actual_sem_grade not in ("", "F"):
                        record_issue(
                            "Grade",
                            "Blank (--) or F",
                            actual_sem_grade_raw,
                            f"Student has {failed_subject_count} failed subject(s), so Semester Grade cannot be '{actual_sem_grade_raw}'.",
                            severity="Critical",
                            category="Semester Grade Mismatch",
                        )
                else:
                    expected_sem_grade = derive_grade_from_gpa(
                        actual_sem_sgpa or expected_sgpa
                    )
                    if not actual_sem_grade:
                        record_issue(
                            "Grade",
                            expected_sem_grade or "Valid Grade",
                            actual_sem_grade_raw,
                            f"Missing Semester Grade (Expected '{expected_sem_grade}').",
                            severity="Critical",
                            category="Semester Grade Mismatch",
                        )
                    elif (
                        expected_sem_grade
                        and actual_sem_grade != expected_sem_grade
                    ):
                        record_issue(
                            "Grade",
                            expected_sem_grade,
                            actual_sem_grade_raw,
                            f"Semester Grade ('{actual_sem_grade_raw}') does not match expected grade ('{expected_sem_grade}') derived from SGPA.",
                            severity="Critical",
                            category="Semester Grade Mismatch",
                        )

        # ------------------------------------------------------------------
        # 4. OVERALL RESULT BACKGROUND CALCULATIONS
        # ------------------------------------------------------------------
        if "Total Marks Obt." in df_audited.columns:
            actual_tot_obt = to_decimal(row.get("Total Marks Obt.", ""))
            if actual_tot_obt is None:
                record_issue(
                    "Total Marks Obt.",
                    str(int(sum_sub_total_marks)),
                    str(row.get("Total Marks Obt.", "")),
                    f"Missing Total Marks Obt. (Expected {int(sum_sub_total_marks)}).",
                    severity="Critical",
                    category="Overall Marks Mismatch",
                )
            elif actual_tot_obt != sum_sub_total_marks:
                record_issue(
                    "Total Marks Obt.",
                    str(int(sum_sub_total_marks)),
                    str(row.get("Total Marks Obt.", "")),
                    f"Total Marks Obt. ({int(actual_tot_obt)}) does not equal sum of Subject Total Marks ({int(sum_sub_total_marks)}).",
                    severity="Critical",
                    category="Overall Marks Mismatch",
                )

        if "Max Total Marks" in df_audited.columns:
            actual_max_tot = to_decimal(row.get("Max Total Marks", ""))
            if actual_max_tot is None:
                record_issue(
                    "Max Total Marks",
                    str(int(sum_sub_max_marks)),
                    str(row.get("Max Total Marks", "")),
                    f"Missing Max Total Marks (Expected {int(sum_sub_max_marks)} from Page 2).",
                    severity="Critical",
                    category="Overall Marks Mismatch",
                )
            elif actual_max_tot != sum_sub_max_marks:
                record_issue(
                    "Max Total Marks",
                    str(int(sum_sub_max_marks)),
                    str(row.get("Max Total Marks", "")),
                    f"Max Total Marks ({int(actual_max_tot)}) does not equal sum of Page 2 subject maximum marks ({int(sum_sub_max_marks)}).",
                    severity="Critical",
                    category="Overall Marks Mismatch",
                )

        if "Percentage" in df_audited.columns and sum_sub_max_marks > Decimal("0"):
            expected_pct_unrounded = (
                sum_sub_total_marks / sum_sub_max_marks
            ) * Decimal("100")
            expected_pct_rounded = expected_pct_unrounded.quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            actual_pct = to_decimal(row.get("Percentage", ""))
            if actual_pct is None:
                if failed_subject_count == 0:
                    record_issue(
                        "Percentage",
                        f"{expected_pct_rounded:.2f}",
                        str(row.get("Percentage", "")),
                        f"Missing Percentage in Gazette; expected ({int(sum_sub_total_marks)} / {int(sum_sub_max_marks)} x 100) = {expected_pct_rounded:.2f}%.",
                        severity="Warning",
                        category="Missing Percentage in Gazette",
                    )
            elif abs(actual_pct - expected_pct_rounded) > Decimal("0.01"):
                record_issue(
                    "Percentage",
                    f"{expected_pct_rounded:.2f}",
                    str(row.get("Percentage", "")),
                    f"Percentage mismatch: ({int(sum_sub_total_marks)} / {int(sum_sub_max_marks)} x 100) = {expected_pct_rounded:.2f}%, but found {actual_pct:.2f}%.",
                    severity="Critical",
                    category="Overall Percentage Mismatch",
                )

        if has_cumulative_cols:
            sem_credits_list = [
                to_decimal(row.get(f"Sem_{s}_Credits", ""))
                for s in sem_prefixes
                if f"Sem_{s}_Credits" in df_audited.columns
            ]
            sem_cg_list = [
                to_decimal(row.get(f"Sem_{s}_CG", ""))
                for s in sem_prefixes
                if f"Sem_{s}_CG" in df_audited.columns
            ]

            expected_tot_cred = (
                sum(c for c in sem_credits_list if c is not None)
                if sem_credits_list and all(c is not None for c in sem_credits_list)
                else None
            )
            expected_tot_cg = (
                sum(g for g in sem_cg_list if g is not None)
                if sem_cg_list and all(g is not None for g in sem_cg_list)
                else None
            )

            if "Total Credit" in df_audited.columns:
                actual_tot_cred = to_decimal(row.get("Total Credit", ""))
                if actual_tot_cred is None and failed_subject_count == 0 and not has_lower_sem_fail:
                    record_issue(
                        "Total Credit",
                        (
                            str(int(expected_tot_cred))
                            if expected_tot_cred is not None
                            else "Non-empty"
                        ),
                        str(row.get("Total Credit", "")),
                        "Missing Total Credit.",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )
                elif (
                    actual_tot_cred is not None
                    and expected_tot_cred is not None
                    and actual_tot_cred != expected_tot_cred
                ):
                    record_issue(
                        "Total Credit",
                        str(int(expected_tot_cred)),
                        str(row.get("Total Credit", "")),
                        f"Total Credit ({int(actual_tot_cred)}) does not match sum of semester credits ({int(expected_tot_cred)}).",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )

            if "Total CG" in df_audited.columns:
                actual_tot_cg = to_decimal(row.get("Total CG", ""))
                if actual_tot_cg is None and failed_subject_count == 0 and not has_lower_sem_fail:
                    record_issue(
                        "Total CG",
                        (
                            str(int(expected_tot_cg))
                            if expected_tot_cg is not None
                            else "Non-empty"
                        ),
                        str(row.get("Total CG", "")),
                        "Missing Total CG.",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )
                elif (
                    actual_tot_cg is not None
                    and expected_tot_cg is not None
                    and actual_tot_cg != expected_tot_cg
                ):
                    record_issue(
                        "Total CG",
                        str(int(expected_tot_cg)),
                        str(row.get("Total CG", "")),
                        f"Total CG ({int(actual_tot_cg)}) does not match sum of semester CG values ({int(expected_tot_cg)}).",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )

            if "CGPA" in df_audited.columns:
                actual_cgpa = to_decimal(row.get("CGPA", ""))
                tc_val = to_decimal(row.get("Total Credit", "")) or expected_tot_cred
                tcg_val = to_decimal(row.get("Total CG", "")) or expected_tot_cg
                expected_cgpa = (
                    (tcg_val / tc_val).quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    )
                    if (tc_val and tcg_val is not None and tc_val > Decimal("0"))
                    else None
                )
                if actual_cgpa is None and failed_subject_count == 0 and not has_lower_sem_fail:
                    record_issue(
                        "CGPA",
                        (
                            f"{expected_cgpa:.2f}"
                            if expected_cgpa is not None
                            else "Valid CGPA"
                        ),
                        str(row.get("CGPA", "")),
                        "Missing CGPA.",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )
                elif (
                    actual_cgpa is not None
                    and expected_cgpa is not None
                    and abs(actual_cgpa - expected_cgpa) > Decimal("0.01")
                ):
                    record_issue(
                        "CGPA",
                        f"{expected_cgpa:.2f}",
                        str(row.get("CGPA", "")),
                        f"CGPA mismatch: Total CG / Total Credit = {expected_cgpa:.2f}, but found {actual_cgpa:.2f}.",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )

            if "Final Grade" in df_audited.columns:
                actual_fg_raw = str(row.get("Final Grade", "")).strip()
                actual_fg, _ = normalize_grade_token(actual_fg_raw)
                cgpa_ref = to_decimal(row.get("CGPA", ""))
                expected_fg = (
                    "F"
                    if (failed_subject_count > 0 or has_lower_sem_fail)
                    else derive_grade_from_gpa(cgpa_ref)
                )
                if not actual_fg and failed_subject_count == 0 and not has_lower_sem_fail:
                    record_issue(
                        "Final Grade",
                        expected_fg or "Valid Final Grade",
                        actual_fg_raw,
                        "Missing Final Grade.",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )
                elif actual_fg and expected_fg and actual_fg != expected_fg:
                    record_issue(
                        "Final Grade",
                        expected_fg,
                        actual_fg_raw,
                        f"Final Grade ('{actual_fg_raw}') does not match CGPA-derived grade ('{expected_fg}').",
                        severity="Critical",
                        category="Cumulative Calculation Mismatch",
                    )

        # Result Status validation
        if "Result Status" in df_audited.columns:
            clean_status = status_raw.replace("$", "").strip()
            expected_status = (
                "Unsuccessful" if failed_subject_count > 0 else "Successful"
            )
            if not clean_status:
                record_issue(
                    "Result Status",
                    expected_status,
                    status_raw,
                    f"Missing Result Status (Expected '{expected_status}').",
                    severity="Critical",
                    category="Result Status / Fail Anomaly",
                )
            elif clean_status.lower() != expected_status.lower():
                record_issue(
                    "Result Status",
                    expected_status,
                    status_raw,
                    f"Result Status mismatch: Student has {failed_subject_count} failed subject(s) in current semester, so status must be '{expected_status}' (found '{status_raw}').",
                    severity="Critical",
                    category="Result Status / Fail Anomaly",
                )

        # Eligibility Status consistency check
        if "Eligibility Status" in df_audited.columns:
            elig_raw = str(row.get("Eligibility Status", "")).strip()
            if elig_raw:
                should_be_not_eligible = (
                    failed_subject_count > 0 or has_lower_sem_fail
                )
                if should_be_not_eligible and "not eligible" not in elig_raw.lower():
                    reason_str = (
                        f"{failed_subject_count} failed subject(s) in current semester"
                        if failed_subject_count > 0
                        else "lower semester not cleared ("
                        + ", ".join(
                            [f"Sem {k}: F({v})" for k, v in lower_sem_fail_counts.items()]
                            + [f"Sem {k}: AD" for k in lower_sem_ad_sems]
                        )
                        + ")"
                    )
                    record_issue(
                        "Eligibility Status",
                        "Not Eligible for next semester",
                        elig_raw,
                        f"Eligibility Status '{elig_raw}' conflicts with {reason_str}.",
                        severity="Critical",
                        category="Eligibility Status Inconsistency",
                    )
                elif (
                    not should_be_not_eligible
                    and "not eligible" in elig_raw.lower()
                ):
                    record_issue(
                        "Eligibility Status",
                        "Eligible for next semester / Program Completed Successfully",
                        elig_raw,
                        f"Eligibility Status '{elig_raw}' conflicts with Successful Result Status (0 failed subjects in current semester and all lower semesters cleared).",
                        severity="Critical",
                        category="Eligibility Status Inconsistency",
                    )

        # '$$' notation must agree with the lower-semester clearance status
        if "Result Status" in df_audited.columns:
            has_dd = "$$" in status_raw.replace(" ", "")
            if has_dd != has_lower_sem_fail:
                record_issue(
                    "Result Status",
                    "With '$$' (lower semester not cleared)"
                    if has_lower_sem_fail
                    else "Without '$$' (all lower semesters cleared)",
                    status_raw,
                    "'$$' in Result Status is printed but no lower-semester failure / AD is recorded."
                    if has_dd
                    else "Lower semester not cleared (F/AD in an earlier semester) but '$$' is missing in Result Status.",
                    severity="Warning",
                    category="Lower Semester Clearance Notation",
                )

        if row_error_messages:
            students_with_errors_count += 1
        verification_errors_col.append(" | ".join(row_error_messages))

    df_audited["Verification Errors"] = verification_errors_col

    df_verification_report = pd.DataFrame(
        report_rows,
        columns=[
            "Student SAP No",
            "Roll No",
            "Candidate Name",
            "Field Name",
            "Expected Value",
            "Actual Value",
            "Error Description",
            "Severity",
        ],
    )

    total_students = len(df_audited)
    total_errors = len(report_rows)
    critical_count = sum(1 for r in report_rows if r["Severity"] == "Critical")
    warning_count = sum(1 for r in report_rows if r["Severity"] == "Warning")

    summary_rows = [
        {"Metric": "Total Students Verified", "Count": total_students},
        {
            "Metric": "Total Students Verified Clean (Zero Errors)",
            "Count": total_students - students_with_errors_count,
        },
        {
            "Metric": "Total Students with Errors / Audit Flags",
            "Count": students_with_errors_count,
        },
        {"Metric": "Total Errors & Audit Flags Found", "Count": total_errors},
        {"Metric": "Total Critical Errors", "Count": critical_count},
        {"Metric": "Total Warnings / Special-Case Flags", "Count": warning_count},
    ]
    df_verification_summary = pd.DataFrame(summary_rows)

    error_wise_rows = []
    for cat, total_cat_count in sorted(
        category_counter.items(), key=lambda x: (-x[1], x[0])
    ):
        crit_c = category_severity_counter.get((cat, "Critical"), 0)
        warn_c = category_severity_counter.get((cat, "Warning"), 0)
        error_wise_rows.append(
            {
                "Error Category": cat,
                "Critical": crit_c,
                "Warning": warn_c,
                "Total": total_cat_count,
            }
        )
    if not error_wise_rows:
        error_wise_rows.append(
            {
                "Error Category": "No Errors Detected",
                "Critical": 0,
                "Warning": 0,
                "Total": 0,
            }
        )
    df_error_wise_summary = pd.DataFrame(error_wise_rows)

    return (
        df_audited,
        df_verification_report,
        df_verification_summary,
        df_error_wise_summary,
        error_cells_set,
        dict(error_messages_map),
    )


# ==============================================================================
# EXCEL WORKBOOK FORMATTING & EXPORT
# (2-ROW MERGED GROUP HEADERS, HOVER COMMENTS, & STUDENT_WISE_VIEW ORDERING)
# ==============================================================================

_SUBJECT_PASTEL_PALETTE = [
    "F9C6D0",  # rose
    "FFD8B1",  # peach
    "FFF2B3",  # butter
    "C9EBC0",  # mint
    "B8E6E0",  # aqua
    "BFE0F5",  # sky
    "CFCDF5",  # periwinkle
    "E4C9F2",  # lilac
    "F5C6E8",  # orchid
    "D5E8B0",  # lime
    "FAD4C0",  # apricot
    "C0DCE8",  # powder blue
]


def generate_pastel_palette(n):
    """Return n unique pastel hex colours (curated first, then generated dynamically)."""
    palette = list(_SUBJECT_PASTEL_PALETTE[:n])
    used = set(palette)
    i = 0
    while len(palette) < n:
        hue = (i * 0.618033988749895) % 1.0
        r, g, b = colorsys.hls_to_rgb(hue, 0.86, 0.55)
        hex_color = f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"
        i += 1
        if hex_color not in used:
            used.add(hex_color)
            palette.append(hex_color)
    return palette


def _compute_header_groups(cols):
    """
    Build group spans for Row 1 & Row 2 headers in Result_Data:
    - Subject groups: merged across all columns of each SubN ('Subject 1'..'Subject N')
    - Semester groups: merged across all columns of each Sem_X ('Semester I'..'Semester V')
    - Standalone columns: merged vertically across Row 1 & Row 2
    """
    col_meta = []
    for idx, c in enumerate(cols, start=1):
        c_str = str(c)
        m_sub = re.match(r"^Sub(\d+)\s+(.*)$", c_str)
        m_sem = re.match(r"^Sem_([IVXLCDM\d]+)_(.*)$", c_str)
        if m_sub:
            col_meta.append(
                {
                    "col_idx": idx,
                    "col_name": c_str,
                    "group_type": "subject",
                    "group_key": int(m_sub.group(1)),
                    "group_label": f"Subject {m_sub.group(1)}",
                    "sub_label": m_sub.group(2),
                }
            )
        elif m_sem:
            col_meta.append(
                {
                    "col_idx": idx,
                    "col_name": c_str,
                    "group_type": "semester",
                    "group_key": m_sem.group(1),
                    "group_label": f"Semester {m_sem.group(1)}",
                    "sub_label": m_sem.group(2),
                }
            )
        else:
            col_meta.append(
                {
                    "col_idx": idx,
                    "col_name": c_str,
                    "group_type": "single",
                    "group_key": c_str,
                    "group_label": c_str,
                    "sub_label": c_str,
                }
            )

    spans = []
    for item in col_meta:
        if (
            spans
            and item["group_type"] in ("subject", "semester")
            and spans[-1]["group_type"] == item["group_type"]
            and spans[-1]["group_key"] == item["group_key"]
        ):
            spans[-1]["end_col"] = item["col_idx"]
            spans[-1]["items"].append(item)
        else:
            spans.append(
                {
                    "group_type": item["group_type"],
                    "group_key": item["group_key"],
                    "group_label": item["group_label"],
                    "start_col": item["col_idx"],
                    "end_col": item["col_idx"],
                    "items": [item],
                }
            )
    return col_meta, spans


_EXCEL_NUMERIC_COL_RE = re.compile(
    r"^(Sub\d+ (IA Marks|TH Marks|Total Marks|GP|Credits \(C\)|CG)"
    r"|Sem_[IVXLCDM\d]+_(Credits|CG|SGPA)"
    r"|Total Marks Obt\.|Max Total Marks|Percentage|Total Credit|Total CG"
    r"|CGPA|Cumulative Percentage)$"
)
_EXCEL_DECIMAL_COL_RE = re.compile(r"(SGPA|CGPA|Percentage)$")


def _cell_text(value):
    """String form of an openpyxl cell value ('' for None, keeps numeric 0 as '0')."""
    return "" if value is None else str(value).strip()


def _coerce_numeric_for_excel(df):
    """
    Write purely numeric marks / GP / credits / CG / SGPA / percentage values as real
    Excel numbers (so SUM, sort, filter work) instead of text. IDs, Attempts and any
    special-notation values ('AB', '17 +7', '$', '--', blanks) stay exactly as extracted.
    The audit engine keeps working on the original text DataFrame.
    """
    out = df.copy()
    for c in out.columns:
        c_str = str(c)
        if not _EXCEL_NUMERIC_COL_RE.match(c_str):
            continue
        keep_decimal = bool(_EXCEL_DECIMAL_COL_RE.search(c_str))

        def _conv(v, keep_decimal=keep_decimal):
            if isinstance(v, bool):
                return v
            if isinstance(v, str):
                t = v.strip().replace("%", "")
                if keep_decimal and t.endswith("#"):
                    t = t[:-1]  # '#' (O.229) is shown through a number format instead
                if re.fullmatch(r"\d+", t):
                    return int(t)
                if re.fullmatch(r"\d+\.\d+", t):
                    return float(t)
                return v
            if isinstance(v, float) and v.is_integer() and not keep_decimal:
                return int(v)
            return v

        out[c] = out[c].map(_conv).astype(object)
    return out


def save_to_formatted_excel(
    sheets_dict,
    exam_type="Regular",
    error_cells_set=None,
    error_messages_map=None,
):
    """
    Export formatted Excel workbook:
    - Sheet order: Page_1, Page_2, Result_Data, Student_Wise_View,
      Verification_Report, Verification_Summary, Error_Wise_Summary.
    - Result_Data features 2-row merged Subject & Semester group headers.
    - Discrepancy cells are highlighted with Solid Red Fill + Bold Yellow Font
      and include an Excel Hover Comment detailing Expected vs Actual values.
    """
    if error_cells_set is None:
        error_cells_set = set()
    if error_messages_map is None:
        error_messages_map = {}

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            start_row_idx = 1 if sheet_name == "Result_Data" else 0
            df_to_write = (
                _coerce_numeric_for_excel(df) if sheet_name == "Result_Data" else df
            )
            df_to_write.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=start_row_idx,
            )
            ws = writer.sheets[sheet_name]
            ws.views.sheetView[0].showGridLines = False

            thin_side = Side(style="thin", color="D9D9D9")
            thick_right_side = Side(style="medium", color="1F4E79")
            thin_border = Border(
                left=thin_side,
                right=thin_side,
                top=thin_side,
                bottom=thin_side,
            )

            header_fill_navy = PatternFill(
                start_color="1B365D", end_color="1B365D", fill_type="solid"
            )
            header_fill_subject = PatternFill(
                start_color="1A5276", end_color="1A5276", fill_type="solid"
            )
            header_fill_semester = PatternFill(
                start_color="5B2C6F", end_color="5B2C6F", fill_type="solid"
            )
            header_fill_marks = PatternFill(
                start_color="2E4053", end_color="2E4053", fill_type="solid"
            )
            header_fill_final = PatternFill(
                start_color="145A32", end_color="145A32", fill_type="solid"
            )
            header_fill_audit = PatternFill(
                start_color="7B241C", end_color="7B241C", fill_type="solid"
            )
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_font_subject = Font(
                name="Calibri", size=11, bold=True, color="1F1F1F"
            )

            pwd_row_fill = PatternFill(
                start_color="E8F4F8", end_color="E8F4F8", fill_type="solid"
            )
            carried_forward_fill = PatternFill(
                start_color="EAECEE", end_color="EAECEE", fill_type="solid"
            )
            fresh_mark_fill = PatternFill(
                start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"
            )
            fail_fill = PatternFill(
                start_color="FADBD8", end_color="FADBD8", fill_type="solid"
            )
            fail_font = Font(name="Calibri", size=11, color="9C0006")
            unsuccessful_row_font = Font(name="Calibri", size=11, color="C00000")
            normal_font = Font(name="Calibri", size=11, color="000000")

            verification_error_fill = PatternFill(
                start_color="FF0000", end_color="FF0000", fill_type="solid"
            )
            verification_error_font = Font(
                name="Calibri", size=11, bold=True, color="FFFF00"
            )
            # Special notations (informational, NOT errors)
            special_fills = {
                "grace": PatternFill(start_color="D9C3F0", end_color="D9C3F0", fill_type="solid"),
                "ad": PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
                "lower_unclear": PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
            }
            special_fonts = {
                "grace": Font(name="Calibri", size=11, bold=True, color="4B1E78"),
                "ad": Font(name="Calibri", size=11, bold=True, color="843C0C"),
                "lower_unclear": Font(name="Calibri", size=11, bold=True, color="843C0C"),
            }

            if sheet_name == "Page_1":
                for col_idx in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill_navy
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                for row_idx in range(2, ws.max_row + 1):
                    for c_i in (1, 2):
                        c_cell = ws.cell(row=row_idx, column=c_i)
                        c_cell.border = thin_border
                        c_cell.alignment = Alignment(
                            vertical="center", horizontal="left"
                        )

            elif sheet_name == "Page_2":
                for col_idx in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill_navy
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                for row_idx in range(2, ws.max_row + 1):
                    for col_idx, col_name in enumerate(df.columns, start=1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        if col_name == "Credits" and isinstance(cell.value, (int, float)):
                            cell.number_format = "0.00"  # shown as printed: 2.00 / 4.00
                        if col_name in [
                            "Credits",
                            "Internal Max Marks",
                            "Sem End Max Marks",
                            "Total Max Marks",
                            "Course Code",
                        ]:
                            cell.alignment = Alignment(
                                vertical="center", horizontal="center"
                            )
                        else:
                            cell.alignment = Alignment(
                                vertical="center", horizontal="left"
                            )

            elif sheet_name == "Result_Data":
                cols_list = list(df.columns)
                col_index_map = {c: i + 1 for i, c in enumerate(cols_list)}

                name_col_idx = col_index_map.get("Candidate Full Name", 5)
                freeze_col_letter = get_column_letter(name_col_idx + 1)
                ws.freeze_panes = f"{freeze_col_letter}3"

                final_section_cols = {
                    "Total Credit",
                    "Total CG",
                    "CGPA",
                    "Final Grade",
                    "Cumulative Percentage",
                    "Result Status",
                    "Grade",
                    "Eligibility Status",
                }

                col_meta, header_spans = _compute_header_groups(cols_list)

                subject_end_col_indices = [
                    sp["end_col"]
                    for sp in header_spans
                    if sp["group_type"] == "subject"
                ]
                sem_end_col_indices = [
                    sp["end_col"]
                    for sp in header_spans
                    if sp["group_type"] == "semester"
                ]
                pre_total_credit_col_indices = []
                if "Total Credit" in col_index_map and col_index_map["Total Credit"] > 1:
                    pre_total_credit_col_indices.append(
                        col_index_map["Total Credit"] - 1
                    )

                all_thick_right_col_indices = set(
                    subject_end_col_indices
                    + sem_end_col_indices
                    + pre_total_credit_col_indices
                )

                subject_numbers = [
                    sp["group_key"]
                    for sp in header_spans
                    if sp["group_type"] == "subject"
                ]
                subject_pastel_fills = {
                    num: PatternFill(
                        start_color=hex_c, end_color=hex_c, fill_type="solid"
                    )
                    for num, hex_c in zip(
                        subject_numbers,
                        generate_pastel_palette(len(subject_numbers)),
                    )
                }

                ws.row_dimensions[1].height = 24
                ws.row_dimensions[2].height = 28

                # Build 2-Row Merged Group Headers (Row 1: Group Banners, Row 2: Sub-Headers)
                for sp in header_spans:
                    g_type = sp["group_type"]
                    c1, c2 = sp["start_col"], sp["end_col"]

                    if g_type == "subject":
                        sub_num = sp["group_key"]
                        fill_to_use = subject_pastel_fills.get(
                            sub_num, header_fill_subject
                        )
                        font_to_use = (
                            header_font_subject
                            if sub_num in subject_pastel_fills
                            else header_font
                        )
                        if c1 != c2:
                            ws.merge_cells(
                                start_row=1,
                                start_column=c1,
                                end_row=1,
                                end_column=c2,
                            )
                        ws.cell(row=1, column=c1).value = sp["group_label"]
                        for item in sp["items"]:
                            ci = item["col_idx"]
                            ws.cell(row=2, column=ci).value = item["sub_label"]
                            r_side = (
                                thick_right_side
                                if ci in all_thick_right_col_indices
                                else thin_side
                            )
                            for r_hdr in (1, 2):
                                h_cell = ws.cell(row=r_hdr, column=ci)
                                h_cell.fill = fill_to_use
                                h_cell.font = font_to_use
                                h_cell.border = Border(
                                    left=thin_side,
                                    right=r_side if (r_hdr == 2 or ci == c2) else thin_side,
                                    top=thin_side,
                                    bottom=thin_side,
                                )
                                h_cell.alignment = Alignment(
                                    horizontal="center",
                                    vertical="center",
                                    wrap_text=True,
                                )

                    elif g_type == "semester":
                        if c1 != c2:
                            ws.merge_cells(
                                start_row=1,
                                start_column=c1,
                                end_row=1,
                                end_column=c2,
                            )
                        ws.cell(row=1, column=c1).value = sp["group_label"]
                        for item in sp["items"]:
                            ci = item["col_idx"]
                            ws.cell(row=2, column=ci).value = item["sub_label"]
                            r_side = (
                                thick_right_side
                                if ci in all_thick_right_col_indices
                                else thin_side
                            )
                            for r_hdr in (1, 2):
                                h_cell = ws.cell(row=r_hdr, column=ci)
                                h_cell.fill = header_fill_semester
                                h_cell.font = header_font
                                h_cell.border = Border(
                                    left=thin_side,
                                    right=r_side if (r_hdr == 2 or ci == c2) else thin_side,
                                    top=thin_side,
                                    bottom=thin_side,
                                )
                                h_cell.alignment = Alignment(
                                    horizontal="center",
                                    vertical="center",
                                    wrap_text=True,
                                )

                    else:
                        col_name = sp["group_label"]
                        ws.merge_cells(
                            start_row=1,
                            start_column=c1,
                            end_row=2,
                            end_column=c1,
                        )
                        ws.cell(row=1, column=c1).value = col_name
                        if col_name == "Verification Errors":
                            fill_s = header_fill_audit
                        elif col_name in final_section_cols:
                            fill_s = header_fill_final
                        elif col_name in (
                            "Total Marks Obt.",
                            "Max Total Marks",
                            "Percentage",
                        ):
                            fill_s = header_fill_marks
                        else:
                            fill_s = header_fill_navy

                        r_side = (
                            thick_right_side
                            if c1 in all_thick_right_col_indices
                            else thin_side
                        )
                        for r_hdr in (1, 2):
                            h_cell = ws.cell(row=r_hdr, column=c1)
                            h_cell.fill = fill_s
                            h_cell.font = header_font
                            h_cell.border = Border(
                                left=thin_side,
                                right=r_side,
                                top=thin_side,
                                bottom=thin_side,
                            )
                            h_cell.alignment = Alignment(
                                horizontal="center",
                                vertical="center",
                                wrap_text=True,
                            )

                # Precompute subject blocks dynamically from header_spans
                sub_blocks = []
                for sp in header_spans:
                    if sp["group_type"] == "subject":
                        g_col = next(
                            (
                                it["col_idx"]
                                for it in sp["items"]
                                if it["sub_label"] == "Grade"
                            ),
                            None,
                        )
                        if g_col:
                            sub_blocks.append(
                                (g_col, sp["start_col"], sp["end_col"])
                            )

                is_pwd_col_idx = col_index_map.get("Is_PwD")
                status_col_idx = col_index_map.get("Result Status")
                grade_col_idx = col_index_map.get("Grade")
                final_grade_col_idx = col_index_map.get("Final Grade")
                elig_col_idx = col_index_map.get("Eligibility Status")

                # Format Data Rows (starting at Row 3)
                for row_idx in range(3, ws.max_row + 1):
                    df_row_idx = row_idx - 3
                    status_val = (
                        str(ws.cell(row=row_idx, column=status_col_idx).value or "")
                        if status_col_idx
                        else ""
                    )
                    grade_val = (
                        str(ws.cell(row=row_idx, column=grade_col_idx).value or "")
                        if grade_col_idx
                        else ""
                    )
                    final_grade_val = (
                        str(ws.cell(row=row_idx, column=final_grade_col_idx).value or "")
                        if final_grade_col_idx
                        else ""
                    )
                    eligibility_val = (
                        str(ws.cell(row=row_idx, column=elig_col_idx).value or "")
                        if elig_col_idx
                        else ""
                    )

                    is_unsuccessful_row = (
                        "Unsuccessful" in status_val
                        or "Not Eligible" in eligibility_val
                        or grade_val.strip() == "F"
                        or final_grade_val.strip() == "F"
                    )

                    is_pwd_row = (
                        str(ws.cell(row=row_idx, column=is_pwd_col_idx).value or "").strip()
                        == "Yes"
                        if is_pwd_col_idx
                        else False
                    )

                    failed_cols_in_row = set()
                    carried_cols_in_row = set()
                    fresh_mark_cols_in_row = set()

                    for g_col, start_col, end_col in sub_blocks:
                        g_val = str(
                            ws.cell(row=row_idx, column=g_col).value or ""
                        ).strip()
                        if "F" in g_val:
                            for c in range(start_col, end_col + 1):
                                failed_cols_in_row.add(c)

                        for c in range(start_col, end_col + 1):
                            c_name = cols_list[c - 1]
                            c_val = _cell_text(
                                ws.cell(row=row_idx, column=c).value
                            )

                            if "$" in c_val:
                                carried_cols_in_row.add(c)
                            elif (
                                "Re-examination" in exam_type
                                and (
                                    "IA Marks" in c_name or "TH Marks" in c_name
                                )
                                and c_val
                                and c_val not in ["--", "NV", "AB", "NaN", "nan"]
                            ):
                                fresh_mark_cols_in_row.add(c)

                    for col_idx, col_name in enumerate(cols_list, start=1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        r_side = (
                            thick_right_side
                            if col_idx in all_thick_right_col_indices
                            else thin_side
                        )
                        cell.border = Border(
                            left=thin_side,
                            right=r_side,
                            top=thin_side,
                            bottom=thin_side,
                        )

                        if (
                            col_name in ("Candidate Full Name", "Verification Errors")
                            or "Course Title" in col_name
                        ):
                            cell.alignment = Alignment(
                                vertical="center", horizontal="left"
                            )
                        else:
                            cell.alignment = Alignment(
                                vertical="center", horizontal="center"
                            )

                        if isinstance(cell.value, float) and _EXCEL_DECIMAL_COL_RE.search(
                            str(col_name)
                        ):
                            cell.number_format = "0.00"

                        if col_idx in failed_cols_in_row:
                            cell.font = fail_font
                        elif is_unsuccessful_row:
                            cell.font = unsuccessful_row_font
                        else:
                            cell.font = normal_font

                        if col_idx in failed_cols_in_row:
                            cell.fill = fail_fill
                        elif col_idx in fresh_mark_cols_in_row:
                            cell.fill = fresh_mark_fill
                        elif col_idx in carried_cols_in_row:
                            cell.fill = carried_forward_fill
                        elif is_pwd_row:
                            cell.fill = pwd_row_fill

                        # Special notations ('#', 'AD', '$$') get their own colour and hover note
                        if (df_row_idx, col_name) not in error_cells_set and (
                            col_name == "Result Status"
                            or _SEM_SGPA_COL_RE.match(str(col_name))
                        ):
                            raw_row = df.iloc[df_row_idx]
                            sp_kind = classify_special_notation(col_name, raw_row[col_name])
                            if sp_kind:
                                cell.fill = special_fills[sp_kind]
                                cell.font = special_fonts[sp_kind]
                                cell.comment = Comment(
                                    special_notation_note(sp_kind, raw_row),
                                    "Result Verification Tool",
                                )
                                if sp_kind == "grace" and isinstance(cell.value, (int, float)):
                                    cell.number_format = '0.00"#"'

                        # HIGHEST PRIORITY: Highlight verification errors with solid red fill, bold yellow font, & hover note
                        if (df_row_idx, col_name) in error_cells_set:
                            cell.fill = verification_error_fill
                            cell.font = verification_error_font
                            err_notes = error_messages_map.get(
                                (df_row_idx, col_name)
                            )
                            if err_notes and col_name != "Verification Errors":
                                cell.comment = Comment(
                                    "\n".join(err_notes),
                                    "Verification Audit Engine",
                                )

            elif sheet_name in [
                "Verification_Report",
                "Verification_Summary",
                "Error_Wise_Summary",
            ]:
                ws.freeze_panes = "A2"
                for col_idx in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill_navy
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                    cell.border = thin_border

                crit_fill = PatternFill(
                    start_color="FADBD8", end_color="FADBD8", fill_type="solid"
                )
                crit_font = Font(name="Calibri", size=11, bold=True, color="9C0006")
                warn_fill = PatternFill(
                    start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"
                )
                warn_font = Font(name="Calibri", size=11, bold=True, color="7D6608")

                sev_col_i = (
                    list(df.columns).index("Severity") + 1
                    if "Severity" in df.columns
                    else None
                )

                for row_idx in range(2, ws.max_row + 1):
                    sev_val = (
                        str(ws.cell(row=row_idx, column=sev_col_i).value or "").strip()
                        if sev_col_i
                        else ""
                    )

                    for col_idx, col_name in enumerate(df.columns, start=1):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        cell.border = thin_border
                        if col_name in [
                            "Candidate Name",
                            "Error Description",
                            "Field Name",
                            "Metric",
                            "Error Category",
                        ]:
                            cell.alignment = Alignment(
                                vertical="center", horizontal="left", wrap_text=True
                            )
                        else:
                            cell.alignment = Alignment(
                                vertical="center", horizontal="center"
                            )

                        if sev_val == "Critical":
                            cell.fill = crit_fill
                            cell.font = (
                                crit_font if col_name == "Severity" else normal_font
                            )
                        elif sev_val == "Warning":
                            cell.fill = warn_fill
                            cell.font = (
                                warn_font if col_name == "Severity" else normal_font
                            )
                        else:
                            cell.font = normal_font

            # Auto-fit columns (skipping Row 1 merged group banners on Result_Data)
            min_row_for_width = 2 if sheet_name == "Result_Data" else 1
            for col_idx in range(1, len(df.columns) + 1):
                max_len = 0
                col_letter = get_column_letter(col_idx)
                for r_i in range(min_row_for_width, ws.max_row + 1):
                    val_len = len(str(ws.cell(row=r_i, column=col_idx).value or ""))
                    if val_len > max_len:
                        max_len = val_len
                ws.column_dimensions[col_letter].width = min(
                    max(max_len + 3, 9), 65
                )

            # IMMEDIATELY ADD Student_Wise_View AFTER Result_Data
            if sheet_name == "Result_Data":
                _add_student_wise_view_sheet(
                    writer,
                    df,
                    exam_type=exam_type,
                    error_cells_set=error_cells_set,
                    error_messages_map=error_messages_map,
                )

    output.seek(0)
    return output


# ==============================================================================
# STUDENT-WISE VIEW SHEET (POSITIONED IMMEDIATELY AFTER RESULT_DATA)
# ==============================================================================

_SWV_SHEET_NAME = "Student_Wise_View"


def _swv_value(val, numeric=False):
    """Normalise a Result_Data cell for display (NaN -> '', optional numeric cast)."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(val, bool):
        return val
    if isinstance(val, numbers.Real):
        return int(val) if float(val).is_integer() else float(val)
    text = str(val).strip()
    if numeric:
        if re.fullmatch(r"\d+", text):
            return int(text)
        if re.fullmatch(r"\d+\.\d+", text):
            return float(text)
    return text


def _add_student_wise_view_sheet(
    writer,
    df,
    exam_type="Regular",
    error_cells_set=None,
    error_messages_map=None,
):
    """Create the 'Student_Wise_View' worksheet immediately after 'Result_Data'."""
    if error_cells_set is None:
        error_cells_set = set()
    if error_messages_map is None:
        error_messages_map = {}

    ws = writer.book.create_sheet(_SWV_SHEET_NAME)
    ws.views.sheetView[0].showGridLines = False

    if df is None or df.empty:
        ws["A1"] = "No student records available."
        return ws

    # Column widths are sized from the actual data so that nothing wraps unnecessarily
    # (e.g. 'Unsuccessful$$', 'Not Eligible for next semester', long course titles/names).
    def _max_len(col_names):
        longest = 0
        for cn in col_names:
            if cn in df.columns:
                for val in df[cn].tolist():
                    longest = max(longest, len(str(_swv_value(val))))
        return longest

    title_cols = [c for c in df.columns if re.match(r"^Sub\d+ Course Title$", str(c))]
    name_cols = ["Candidate Full Name", "PRN No"]
    col_a_w = 21
    col_b_w = min(max(_max_len(title_cols + name_cols) + 4, 24), 60)
    col_i_w = min(
        max(_max_len(["Result Status", "Eligibility Status", "Grade"]) + 4, 12), 45
    )
    swv_widths = [col_a_w, col_b_w, 12, 12, 12, 12, 12, 12, col_i_w]
    for col_letter, width in zip("ABCDEFGHI", swv_widths):
        ws.column_dimensions[col_letter].width = width

    thin_side = Side(style="thin", color="D9D9D9")
    border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )
    fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_subject = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")
    fill_semester = PatternFill(start_color="5B2C6F", end_color="5B2C6F", fill_type="solid")
    fill_final = PatternFill(start_color="145A32", end_color="145A32", fill_type="solid")
    fill_label = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
    fill_fail = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    fill_carried = PatternFill(start_color="EAECEE", end_color="EAECEE", fill_type="solid")
    fill_grace = PatternFill(start_color="D9C3F0", end_color="D9C3F0", fill_type="solid")
    fill_ad = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    font_grace = Font(name="Calibri", size=11, bold=True, color="4B1E78")
    font_ad = Font(name="Calibri", size=11, bold=True, color="843C0C")
    fill_fresh = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    fill_error = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    font_banner = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_label = Font(name="Calibri", size=11, bold=True, color="1B365D")
    font_value_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_normal = Font(name="Calibri", size=11, color="000000")
    font_fail = Font(name="Calibri", size=11, color="9C0006")
    font_bad = Font(name="Calibri", size=11, bold=True, color="C00000")
    font_error = Font(name="Calibri", size=11, bold=True, color="FFFF00")

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def put(
        r1,
        c1,
        r2,
        c2,
        value=None,
        font=None,
        fill=None,
        align=None,
        fmt=None,
        comment_text=None,
    ):
        if (r1, c1) != (r2, c2):
            ws.merge_cells(
                start_row=r1, start_column=c1, end_row=r2, end_column=c2
            )
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                cell = ws.cell(row=rr, column=cc)
                cell.border = border
                if font is not None:
                    cell.font = font
                if fill is not None:
                    cell.fill = fill
                if align is not None:
                    cell.alignment = align
        top_left = ws.cell(row=r1, column=c1)
        top_left.value = value
        if fmt and value != "":
            top_left.number_format = fmt
        if comment_text:
            top_left.comment = Comment(comment_text, "Verification Audit Engine")
        return top_left

    cols = list(df.columns)
    sub_nums = sorted(
        {
            int(m.group(1))
            for c in cols
            if (m := re.match(r"^Sub(\d+)\s", str(c)))
        }
    )

    sem_names = []
    for c in cols:
        m = re.match(
            r"^Sem_([IVXLCDM\d]+)_(Credits|CG|SGPA|Attempts|Fail_Info)$",
            str(c),
        )
        if m and m.group(1) not in sem_names:
            sem_names.append(m.group(1))
    sem_names.sort(key=sem_rank)

    def get(row, col, numeric=False):
        if col not in row.index:
            return ""
        return _swv_value(row[col], numeric=numeric)

    # The ID column is labelled exactly as the gazette prints it (ABC ID / APAAR ID)
    swv_id_label = "APAAR ID" if "APAAR ID" in df.columns else "ABC ID"

    info_layout = [
        [
            ("SNo", (1, 1), (2, 2)),
            ("SAP No", (3, 4), (5, 6)),
            ("Roll No", (7, 7), (8, 9)),
        ],
        [
            ("PRN No", (1, 1), (2, 2)),
            (swv_id_label, (3, 4), (5, 6)),
            ("Gender", (7, 7), (8, 9)),
        ],
        [
            ("Candidate Full Name", (1, 1), (2, 2)),
            ("Examination Type", (3, 4), (5, 6)),
            ("Is_PwD", (7, 7), (8, 9)),
        ],
    ]
    sub_headers = [
        "Sub Code",
        "Course Title",
        "IA Marks",
        "TH Marks",
        "Total Marks",
        "Grade",
        "GP",
        "Credits",
        "CG",
    ]
    sem_headers = ["Semester", "Credits", "CG", "SGPA", "Attempts", "Fail_Info"]
    overall_items = [
        ("Grade", False, None),
        ("Total Marks Obt.", True, None),
        ("Max Total Marks", True, None),
        ("Percentage", True, '0.00"%"'),
        ("Result Status", False, None),
        ("Eligibility Status", False, None),
    ]

    total_students = len(df)
    r = 1
    for s_idx, (df_row_idx, row) in enumerate(df.iterrows(), start=1):
        put(
            r,
            1,
            r,
            9,
            f"Student {s_idx} of {total_students}",
            font_banner,
            fill_navy,
            center,
        )
        ws.row_dimensions[r].height = 24
        r += 1
        for info_row in info_layout:
            for label, (lc1, lc2), (vc1, vc2) in info_row:
                put(r, lc1, r, lc2, label, font_label, fill_label, left)
                is_err = (df_row_idx, label) in error_cells_set
                c_notes = "\n".join(
                    error_messages_map.get((df_row_idx, label), [])
                )
                put(
                    r,
                    vc1,
                    r,
                    vc2,
                    str(get(row, label)),
                    font_error if is_err else font_value_bold,
                    fill_error if is_err else None,
                    left,
                    comment_text=c_notes if is_err else None,
                )
            r += 1

        put(
            r,
            1,
            r,
            9,
            "Subject-Wise Performance",
            font_header,
            fill_subject,
            center,
        )
        r += 1
        for c_idx, h in enumerate(sub_headers, start=1):
            put(r, c_idx, r, c_idx, h, font_header, fill_subject, center)
        r += 1

        subject_rows = 0
        for n in sub_nums:
            code = get(row, f"Sub{n} Code")
            if str(code).strip() == "":
                continue
            title = get(row, f"Sub{n} Course Title")
            col_names_for_sub = [
                f"Sub{n} Code",
                f"Sub{n} Course Title",
                f"Sub{n} IA Marks",
                f"Sub{n} TH Marks",
                f"Sub{n} Total Marks",
                f"Sub{n} Grade",
                f"Sub{n} GP",
                f"Sub{n} Credits (C)",
                f"Sub{n} CG",
            ]
            values = [
                code,
                title,
                get(row, f"Sub{n} IA Marks", True),
                get(row, f"Sub{n} TH Marks", True),
                get(row, f"Sub{n} Total Marks", True),
                get(row, f"Sub{n} Grade"),
                get(row, f"Sub{n} GP", True),
                get(row, f"Sub{n} Credits (C)", True),
                get(row, f"Sub{n} CG", True),
            ]
            is_failed = "F" in str(values[5]).strip()
            for c_idx, (v, src_col) in enumerate(
                zip(values, col_names_for_sub), start=1
            ):
                v_str = str(v)
                is_err = (df_row_idx, src_col) in error_cells_set
                c_notes = "\n".join(
                    error_messages_map.get((df_row_idx, src_col), [])
                )
                if is_err:
                    fill_c, font_c = fill_error, font_error
                elif is_failed:
                    fill_c, font_c = fill_fail, font_fail
                elif "$" in v_str:
                    fill_c, font_c = fill_carried, font_normal
                elif (
                    "Re-examination" in exam_type
                    and c_idx in (3, 4)
                    and v_str not in ["", "--", "NV", "AB", "NaN", "nan"]
                ):
                    fill_c, font_c = fill_fresh, font_normal
                else:
                    fill_c, font_c = None, font_normal
                put(
                    r,
                    c_idx,
                    r,
                    c_idx,
                    v,
                    font_c,
                    fill_c,
                    left if c_idx == 2 else center,
                    comment_text=c_notes if is_err else None,
                )
            if len(str(title)) > col_b_w - 3:
                ws.row_dimensions[r].height = 30
            r += 1
            subject_rows += 1

        if subject_rows == 0:
            put(
                r,
                1,
                r,
                9,
                "No subject-wise data available for this student.",
                font_normal,
                None,
                center,
            )
            r += 1

        put(r, 1, r, 6, "Semester Details", font_header, fill_semester, center)
        put(r, 7, r, 9, "Overall Result Details", font_header, fill_final, center)
        r += 1
        for c_idx, h in enumerate(sem_headers, start=1):
            put(r, c_idx, r, c_idx, h, font_header, fill_semester, center)
        put(r, 7, r, 8, "Metric", font_header, fill_final, center)
        put(r, 9, r, 9, "Value", font_header, fill_final, center)
        r += 1

        sem_rows = []
        for s in sem_names:
            s_cols = [
                f"Sem_{s}_Credits",
                f"Sem_{s}_CG",
                f"Sem_{s}_SGPA",
                f"Sem_{s}_Attempts",
                f"Sem_{s}_Fail_Info",
            ]
            vals = [
                get(row, s_cols[0], True),
                get(row, s_cols[1], True),
                get(row, s_cols[2], True),
                get(row, s_cols[3]),
                get(row, s_cols[4]),
            ]
            if any(str(v).strip() != "" for v in vals) or any(
                (df_row_idx, sc) in error_cells_set for sc in s_cols
            ):
                sem_rows.append((s, [f"Semester {s}"] + vals, [None] + s_cols))

        block_rows = max(len(sem_rows), len(overall_items), 1)
        for k in range(block_rows):
            if k < len(sem_rows):
                _, s_vals, s_src_cols = sem_rows[k]
                for c_idx, (v, sc_name) in enumerate(
                    zip(s_vals, s_src_cols), start=1
                ):
                    is_err = (
                        sc_name is not None
                        and (df_row_idx, sc_name) in error_cells_set
                    )
                    c_notes = (
                        "\n".join(
                            error_messages_map.get((df_row_idx, sc_name), [])
                        )
                        if is_err
                        else None
                    )
                    sp_kind = (
                        classify_special_notation(sc_name, row.get(sc_name, ""))
                        if (sc_name is not None and c_idx == 4 and not is_err)
                        else None
                    )
                    cell_val, cell_fmt = v, ("0.00" if c_idx == 4 else None)
                    cell_font = font_error if is_err else font_normal
                    cell_fill = fill_error if is_err else None
                    if sp_kind == "grace":
                        cell_val = _swv_value(str(v).replace("#", ""), numeric=True)
                        cell_fmt = '0.00"#"'
                        cell_font, cell_fill = font_grace, fill_grace
                        c_notes = special_notation_note(sp_kind, row)
                    elif sp_kind == "ad":
                        cell_font, cell_fill = font_ad, fill_ad
                        c_notes = special_notation_note(sp_kind, row)
                    put(
                        r + k,
                        c_idx,
                        r + k,
                        c_idx,
                        cell_val,
                        cell_font,
                        cell_fill,
                        center,
                        fmt=cell_fmt,
                        comment_text=c_notes,
                    )
            elif k == 0:
                put(
                    r,
                    1,
                    r,
                    6,
                    "No semester data available.",
                    font_normal,
                    None,
                    center,
                )

            if k < len(overall_items):
                label, numeric, fmt = overall_items[k]
                v = get(row, label, numeric)
                is_err = (df_row_idx, label) in error_cells_set
                c_notes = (
                    "\n".join(error_messages_map.get((df_row_idx, label), []))
                    if is_err
                    else None
                )
                lower_unclear = (
                    not is_err and classify_special_notation(label, v) == "lower_unclear"
                )
                bad = label in ("Result Status", "Eligibility Status") and (
                    "Unsuccessful" in str(v) or "Not Eligible" in str(v)
                )
                put(r + k, 7, r + k, 8, label, font_label, fill_label, left)
                put(
                    r + k,
                    9,
                    r + k,
                    9,
                    v,
                    font_error if is_err else (font_bad if bad else font_value_bold),
                    fill_error if is_err else (fill_ad if lower_unclear else None),
                    center,
                    fmt=fmt if isinstance(v, (int, float)) else None,
                    comment_text=c_notes
                    or (special_notation_note("lower_unclear", row) if lower_unclear else None),
                )
                if len(str(v)) > col_i_w - 3:
                    ws.row_dimensions[r + k].height = 30
        r += block_rows

        ver_err_text = str(get(row, "Verification Errors")).strip()
        if ver_err_text:
            put(r, 1, r, 2, "Verification Errors", font_error, fill_error, center)
            put(r, 3, r, 9, ver_err_text, font_error, fill_error, left)
            # Height follows the amount of text (merged C:I width) so long error lists are not cut off
            err_chars_per_line = max(sum(swv_widths[2:]) - 4, 20)
            err_lines = max(1, -(-len(ver_err_text) // err_chars_per_line))
            ws.row_dimensions[r].height = max(32, 15 * err_lines + 4)
            r += 1

        r += 2

    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    return ws


# ==============================================================================
# INTERACTIVE STREAMLIT HTML PREVIEW (WITH 2-ROW MERGED HEADERS & HOVER TOOLTIPS)
# ==============================================================================


def render_detailed_table_html(
    df, exam_type="Regular", error_cells_set=None, error_messages_map=None
):
    """Renders interactive HTML preview in Streamlit with 2-row merged Subject/Semester headers."""
    if df.empty:
        st.info("No result records found.")
        return

    if error_cells_set is None:
        error_cells_set = set()
    if error_messages_map is None:
        error_messages_map = {}

    cols = list(df.columns)
    _, header_spans = _compute_header_groups(cols)

    final_section_cols = {
        "Total Credit",
        "Total CG",
        "CGPA",
        "Final Grade",
        "Cumulative Percentage",
        "Result Status",
        "Grade",
        "Eligibility Status",
    }

    subject_end_cols = [
        cols[sp["end_col"] - 1]
        for sp in header_spans
        if sp["group_type"] == "subject"
    ]
    sem_end_cols = [
        cols[sp["end_col"] - 1]
        for sp in header_spans
        if sp["group_type"] == "semester"
    ]

    pre_tc_cols = []
    if "Total Credit" in cols:
        tc_i = cols.index("Total Credit")
        if tc_i > 0:
            pre_tc_cols.append(cols[tc_i - 1])

    all_thick_cols = set(subject_end_cols + sem_end_cols + pre_tc_cols)

    sub_ranges = []
    for sp in header_spans:
        if sp["group_type"] == "subject":
            g_col = next(
                (
                    it["col_name"]
                    for it in sp["items"]
                    if it["sub_label"] == "Grade"
                ),
                None,
            )
            if g_col:
                sub_ranges.append(
                    (g_col, sp["start_col"] - 1, sp["end_col"] - 1)
                )

    html = """
    <style>
    .detailed-table-container {
        width: 100%;
        overflow-x: auto;
        max-height: 580px;
        border: 1px solid #d3d3d3;
        border-radius: 6px;
        margin-top: 10px;
    }
    .detailed-table {
        border-collapse: collapse;
        font-family: Calibri, sans-serif;
        font-size: 13px;
        width: 100%;
        white-space: nowrap;
    }
    .detailed-table th {
        color: white;
        padding: 6px 10px;
        text-align: center;
        border: 1px solid #d9d9d9;
        font-weight: bold;
    }
    .detailed-table thead tr:first-child th {
        position: sticky;
        top: 0;
        z-index: 3;
    }
    .detailed-table thead tr:nth-child(2) th {
        position: sticky;
        top: 31px;
        z-index: 2;
    }
    .th-navy { background-color: #1B365D; }
    .th-subject { background-color: #1A5276; }
    .th-semester { background-color: #5B2C6F; }
    .th-marks { background-color: #2E4053; }
    .th-final { background-color: #145A32; }
    .th-audit { background-color: #7B241C; }

    .detailed-table td {
        padding: 6px 10px;
        text-align: center;
        border: 1px solid #e0e0e0;
        color: #000000;
    }
    .text-left-cell { text-align: left !important; }
    .unsuccessful-row-text { color: #C00000 !important; }
    .thick-right { border-right: 3px solid #1F4E79 !important; }
    .pwd-row { background-color: #E8F4F8 !important; }
    .carried-forward-cell { background-color: #EAECEE !important; }
    .fresh-mark-cell { background-color: #FFF3CD !important; }
    .grace-cell { background-color: #D9C3F0 !important; color: #4B1E78 !important; font-weight: bold; }
    .ad-cell { background-color: #F8CBAD !important; color: #843C0C !important; font-weight: bold; }
    .fail-cell {
        background-color: #FADBD8 !important;
        color: #9C0006 !important;
    }
    .verification-error-cell {
        background-color: #FF0000 !important;
        color: #FFFF00 !important;
        font-weight: bold !important;
        cursor: help;
    }
    </style>
    <div class="detailed-table-container">
    <table class="detailed-table">
    <thead>
    <tr>
    """

    # Row 1 of thead (merged group headers + rowspan=2 for standalone columns)
    for sp in header_spans:
        g_type = sp["group_type"]
        span_len = sp["end_col"] - sp["start_col"] + 1
        end_col_name = cols[sp["end_col"] - 1]
        thick_cls = " thick-right" if end_col_name in all_thick_cols else ""

        if g_type == "subject":
            html += f'<th colspan="{span_len}" class="th-subject{thick_cls}">{sp["group_label"]}</th>'
        elif g_type == "semester":
            html += f'<th colspan="{span_len}" class="th-semester{thick_cls}">{sp["group_label"]}</th>'
        else:
            col = sp["group_label"]
            if col == "Verification Errors":
                cls_name = "th-audit"
            elif col in final_section_cols:
                cls_name = "th-final"
            elif col in ("Total Marks Obt.", "Max Total Marks", "Percentage"):
                cls_name = "th-marks"
            else:
                cls_name = "th-navy"
            html += f'<th rowspan="2" class="{cls_name}{thick_cls}">{col}</th>'

    html += "</tr><tr>"

    # Row 2 of thead (sub-headers for Subject and Semester groups)
    for sp in header_spans:
        if sp["group_type"] in ("subject", "semester"):
            base_cls = (
                "th-subject" if sp["group_type"] == "subject" else "th-semester"
            )
            for it in sp["items"]:
                c_name = it["col_name"]
                thick_cls = " thick-right" if c_name in all_thick_cols else ""
                html += f'<th class="{base_cls}{thick_cls}">{it["sub_label"]}</th>'

    html += "</tr></thead><tbody>"

    for df_row_idx, row in df.iterrows():
        is_pwd = str(row.get("Is_PwD", "")).strip() == "Yes"
        status_val = str(row.get("Result Status", ""))
        grade_val = str(row.get("Grade", ""))
        final_grade_val = str(row.get("Final Grade", ""))
        eligibility_val = str(row.get("Eligibility Status", ""))

        is_unsuccessful_row = (
            "Unsuccessful" in status_val
            or "Not Eligible" in eligibility_val
            or grade_val.strip() == "F"
            or final_grade_val.strip() == "F"
        )

        failed_cols = set()
        carried_cols = set()
        fresh_cols = set()

        for g_col, start_i, end_i in sub_ranges:
            if "F" in str(row.get(g_col, "")).strip():
                for col_name in cols[start_i : end_i + 1]:
                    failed_cols.add(col_name)

            for col_name in cols[start_i : end_i + 1]:
                c_val = str(row.get(col_name, "")).strip()
                if "$" in c_val:
                    carried_cols.add(col_name)
                elif (
                    "Re-examination" in exam_type
                    and ("IA Marks" in col_name or "TH Marks" in col_name)
                    and c_val
                    and c_val not in ["--", "NV", "AB", "NaN", "nan"]
                ):
                    fresh_cols.add(col_name)

        row_classes = []
        if is_pwd:
            row_classes.append("pwd-row")
        if is_unsuccessful_row:
            row_classes.append("unsuccessful-row-text")
        row_cls_str = f' class="{" ".join(row_classes)}"' if row_classes else ""

        html += f"<tr{row_cls_str}>"
        for col in cols:
            val = row[col] if pd.notna(row[col]) else ""
            classes = []
            if col in all_thick_cols:
                classes.append("thick-right")
            if (
                col in ("Candidate Full Name", "Verification Errors")
                or "Course Title" in col
            ):
                classes.append("text-left-cell")

            title_attr = ""
            if (df_row_idx, col) in error_cells_set:
                classes.append("verification-error-cell")
                notes = error_messages_map.get((df_row_idx, col), [])
                if notes:
                    escaped_notes = (
                        "&#10;".join(notes)
                        .replace('"', "&quot;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                    )
                    title_attr = f' title="{escaped_notes}"'
            elif col in failed_cols:
                classes.append("fail-cell")
            elif col in fresh_cols:
                classes.append("fresh-mark-cell")
            elif col in carried_cols:
                classes.append("carried-forward-cell")

            if "verification-error-cell" not in classes:
                sp_kind = classify_special_notation(col, val)
                if sp_kind:
                    classes.append("grace-cell" if sp_kind == "grace" else "ad-cell")
                    note = special_notation_note(sp_kind, row).replace('"', "&quot;").replace("\n", "&#10;")
                    title_attr = f' title="{note}"'

            cls_str = f' class="{" ".join(classes)}"' if classes else ""
            html += f"<td{cls_str}{title_attr}>{val}</td>"
        html += "</tr>"

    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)


# ==============================================================================
# CACHED PIPELINE & MAIN STREAMLIT ENTRY POINTS
# ==============================================================================


@st.cache_data(show_spinner=False)
def process_gazette_pdf_bytes(file_bytes):
    """
    Cached end-to-end extraction, background verification, and Excel generation
    so tab switching and filtering in Streamlit are instantaneous.
    """
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        total_pages = len(pdf.pages)

        p1_text = pdf.pages[0].extract_text() or ""
        df_p1, exam_type = parse_page_1(p1_text)

        p2_page = pdf.pages[1] if total_pages > 1 else pdf.pages[0]
        df_p2 = parse_page_2(p2_page)

        course_map = {}
        if (
            not df_p2.empty
            and "Course Code" in df_p2.columns
            and "Course Title" in df_p2.columns
        ):
            course_map = dict(zip(df_p2["Course Code"], df_p2["Course Title"]))

        df_result_raw = parse_result_pages(
            pdf.pages,
            course_title_map=course_map,
            exam_type=exam_type,
        )

        (
            df_result_audited,
            df_verification_report,
            df_verification_summary,
            df_error_wise_summary,
            error_cells_set,
            error_messages_map,
        ) = verify_and_audit_results(df_result_raw, df_p2)

        # Label the ID column exactly as the result pages print it (ABC ID vs APAAR ID)
        header_probe = ""
        for _pg in pdf.pages[2:4]:
            header_probe += _pg.extract_text() or ""
        if re.search(r"APAAR\s*ID", header_probe, re.IGNORECASE) and (
            "ABC ID" in df_result_audited.columns
        ):
            df_result_audited = df_result_audited.rename(columns={"ABC ID": "APAAR ID"})
            error_cells_set = {
                (i, "APAAR ID" if c == "ABC ID" else c) for i, c in error_cells_set
            }
            error_messages_map = {
                (k[0], "APAAR ID" if k[1] == "ABC ID" else k[1]): v
                for k, v in error_messages_map.items()
            }
            if "Field Name" in df_verification_report.columns:
                df_verification_report = df_verification_report.copy()
                df_verification_report["Field Name"] = df_verification_report[
                    "Field Name"
                ].replace({"ABC ID": "APAAR ID"})

        sheets_data = {
            "Page_1": df_p1,
            "Page_2": df_p2,
            "Result_Data": df_result_audited,
            "Verification_Report": df_verification_report,
            "Verification_Summary": df_verification_summary,
            "Error_Wise_Summary": df_error_wise_summary,
        }

        excel_bytes_io = save_to_formatted_excel(
            sheets_data,
            exam_type=exam_type,
            error_cells_set=error_cells_set,
            error_messages_map=error_messages_map,
        )

        return {
            "total_pages": total_pages,
            "exam_type": exam_type,
            "df_p1": df_p1,
            "df_p2": df_p2,
            "df_result_audited": df_result_audited,
            "df_verification_report": df_verification_report,
            "df_verification_summary": df_verification_summary,
            "df_error_wise_summary": df_error_wise_summary,
            "error_cells_set": error_cells_set,
            "error_messages_map": error_messages_map,
            "excel_bytes": excel_bytes_io.getvalue(),
        }


TOOL_DESCRIPTION_MD = """
**What this tool does**

Upload an SVKM Result Gazette PDF (Regular, ATKT or Re-examination) and the tool will:

1. **Extract** every student's details, subject-wise marks, grades, GP, credits, CG, semester summaries and
   final result exactly as printed in the PDF (values such as `--`, `AB`, `$`, `~`, `17 +7` are preserved as they appear).
2. **Verify** the gazette in the background by independently recalculating totals, grades, GP, credits, CG, SGPA,
   CGPA, percentage, fail counts, result status and eligibility against the Page 2 course scheme.
3. **Export** a formatted Excel workbook with the audit results highlighted.

**How to use it**

1. Upload one or more gazette PDFs (page 1 = details, page 2 = course scheme, remaining pages = results).
2. Review the KPI banner and the preview tabs.
3. Click **Download Audited Excel Report**.

**Excel workbook contents**

| Sheet | Contents |
|---|---|
| `Page_1` | Institution, programme, semester, academic year and examination details |
| `Page_2` | Course codes, titles, credits and maximum marks |
| `Result_Data` | One row per student with merged Subject / Semester headers |
| `Student_Wise_View` | One printable section per student (details, subjects, semester and overall summary) |
| `Verification_Report` | Every discrepancy: student, field, expected value, actual value, description, severity |
| `Verification_Summary` | Students verified, clean vs flagged, total errors, critical vs warning counts |
| `Error_Wise_Summary` | Error counts by category |

**Colour guide**

- 🟥 **Solid red cell with bold yellow text** - value failed verification (hover over it for expected vs actual).
- 🟫 **Light red** - failed subject (grade F).
- ⬜ **Grey** - carried-forward (`$`) value.
- 🟪 **Purple** - SGPA marked `#`: eligible for O.229 grace marks. The unused part of the 10 grace marks is added to the SGPA (0.01 per unused mark, e.g. 10 unused = +0.10). Not an error; the tool checks the SGPA including this bonus.
- 🟧 **Orange** - `AD` (Attendance Defaulter in a lower semester) or `$$` after the Result Status (lower semester not cleared). Not an error; such a candidate is correctly *Not Eligible for next semester* even if successful in the current semester.
- 🟨 **Yellow** - fresh marks in a Re-examination.
- 🟦 **Light blue row** - candidate with disability (`~`).

**Good to know**

- *Critical* flags are calculation or data errors; *Warning* flags are special cases such as **Result in Abeyance**.
- A blank cell means nothing is printed in the PDF; `--` means `--` is printed in the PDF.
- Numeric marks are written to Excel as real numbers, so they can be summed, sorted and filtered.
- The tool reads the PDF only; it never alters the gazette values, it only flags discrepancies.
"""


def render_tool_description():
    """Collapsible user-facing description shown at the top of the page."""
    with st.expander("ℹ️ About this tool - what it does and how to read the output", expanded=False):
        st.markdown(TOOL_DESCRIPTION_MD)


def show():
    """Main view function invoked by app.py router."""
    st.title("🛡️ Intelligent Result Verification & Audit Tool")
    st.caption(
        "Automated Result Gazette Extraction (As-Is PDF Fidelity), Background Mathematical & Rule Verification, and Discrepancy Audit Engine."
    )

    uploaded_files = st.file_uploader(
        "Upload Result Gazette PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        for f_idx, uploaded_file in enumerate(uploaded_files):
            st.subheader(f"📄 Auditing: {uploaded_file.name}")

            with st.spinner(f"Extracting & auditing {uploaded_file.name}..."):
                res = process_gazette_pdf_bytes(uploaded_file.getvalue())

            total_pages = res["total_pages"]
            exam_type = res["exam_type"]
            df_p1 = res["df_p1"]
            df_p2 = res["df_p2"]
            df_result_audited = res["df_result_audited"]
            df_verification_report = res["df_verification_report"]
            df_verification_summary = res["df_verification_summary"]
            df_error_wise_summary = res["df_error_wise_summary"]
            error_cells_set = res["error_cells_set"]
            error_messages_map = res["error_messages_map"]
            excel_bytes = res["excel_bytes"]

            st.info(f"Total Pages Detected: **{total_pages}**")

            # Executive KPI Metrics Banner
            total_verified = len(df_result_audited)
            total_errors = len(df_verification_report)
            students_with_err = (
                df_verification_report["Student SAP No"].nunique()
                if not df_verification_report.empty
                else 0
            )
            crit_errors = (
                int((df_verification_report["Severity"] == "Critical").sum())
                if not df_verification_report.empty
                else 0
            )

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Students Verified", total_verified)
            k2.metric("Total Errors / Flags Found", total_errors)
            k3.metric("Students with Discrepancies", students_with_err)
            k4.metric("Critical Rule Violations", crit_errors)

            # Preview Tabs
            t1, t2, t3, t4, t5 = st.tabs(
                [
                    "📊 Audited Result Data",
                    "🚨 Verification Report",
                    "📈 Verification Summary",
                    "📋 Metadata (Page 1)",
                    "📚 Courses (Page 2)",
                ]
            )

            with t1:
                col_hdr, col_flt = st.columns([3, 1])
                with col_hdr:
                    st.markdown(
                        "##### 📌 Audited Student Results (Erroneous cells highlighted in **Solid Red Fill & Bold Yellow Font** — hover over red cells for details)"
                    )
                with col_flt:
                    only_errors = st.checkbox(
                        "🔍 Show Only Discrepancy Rows",
                        value=False,
                        key=f"flt_err_{f_idx}_{uploaded_file.name}",
                    )

                if only_errors and "Verification Errors" in df_result_audited.columns:
                    df_preview = df_result_audited[
                        df_result_audited["Verification Errors"]
                        .astype(str)
                        .str.strip()
                        != ""
                    ]
                else:
                    df_preview = df_result_audited

                render_detailed_table_html(
                    df_preview,
                    exam_type=exam_type,
                    error_cells_set=error_cells_set,
                    error_messages_map=error_messages_map,
                )
            with t2:
                st.markdown(
                    "##### 🚨 Detailed Verification Report (`Verification_Report`)"
                )
                if df_verification_report.empty:
                    st.success(
                        "All student records passed 100% of verification checks with zero discrepancies!"
                    )
                else:
                    st.dataframe(df_verification_report, use_container_width=True)
            with t3:
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    st.markdown("##### 📌 Overall Verification Summary")
                    st.dataframe(df_verification_summary, use_container_width=True)
                with col_s2:
                    st.markdown("##### 📂 Error-Wise Category Summary")
                    st.dataframe(df_error_wise_summary, use_container_width=True)
            with t4:
                st.dataframe(df_p1, use_container_width=True)
            with t5:
                st.dataframe(df_p2, use_container_width=True)

            export_file_name = (
                f"{uploaded_file.name.rsplit('.', 1)[0]}_Audited_Report.xlsx"
            )
            st.download_button(
                label=f"⬇️ Download Audited Excel Report ({export_file_name})",
                data=excel_bytes,
                file_name=export_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_btn_{f_idx}_{uploaded_file.name}",
            )
            st.divider()

    render_tool_description()


def main():
    """Alias for show() to support direct script execution or alt import."""
    show()


if __name__ == "__main__":
    st.set_page_config(
        page_title="Intelligent Result Verification & Audit Tool",
        page_icon="🛡️",
        layout="wide",
    )
    show()