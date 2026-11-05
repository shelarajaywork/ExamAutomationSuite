import io
import re
from collections import Counter
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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


def to_decimal(val):
    """Safely convert a numeric value or string to Decimal."""
    if val is None:
        return None
    s = str(val).strip().replace("%", "")
    if not s or s in ["--", "AB", "NV", "nan", "NaN", "None"]:
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
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
    if base == "0":
        base = "O"
    return base, {
        "has_carry": has_carry,
        "has_hash": has_hash,
        "has_reval": has_reval,
    }


def parse_mark_component(raw_val):
    """
    Parse an IA, TH, or Total mark cell and detect special cases:
    - AB (Absent - valid entry, treated as 0 numeric marks and F grade)
    - NV (Null & Void)
    - $ (Carried forward)
    - # or + (Ordinance 0.229 / Grace marks e.g. '10 +2', '8 +4', '15 +9' - valid entry)
    - ® (Revaluation)
    - -- (Not Applicable / Blank)
    """
    s = str(raw_val if raw_val is not None else "").strip()
    info = {
        "raw": s,
        "is_empty": s in ["", "--", "nan", "NaN", "None"],
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

    # Check for grace marks syntax: e.g. "10 +2", "8 +4", "15 +9", "14 + 10", "21+3"
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
# HELPER PARSING FUNCTIONS
# ==============================================================================


def clean_text_preserve_symbols(val):
    """Clean whitespace while preserving academic notations ($, ~, /, #, +, ®)."""
    if val is None:
        return ""
    return str(val).strip()


def clean_dashes_and_artifacts(text):
    """Remove continuous dash strings (---...) and redundant whitespace."""
    if not text:
        return ""
    text = re.sub(r"[-_=]{3,}", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_page_1(text):
    """Extract metadata key-values and accurately identify Regular vs Re-examination."""
    metadata = []
    text_clean = re.sub(r"\s+", " ", text)

    inst_match = re.search(
        r"(Shri Vile Parle Kelavani Mandal[^\n\r]*[\s\S]*?(?:AUTONOMOUS|\(Empowered Autonomous\)|\(Autonomous\)))",
        text,
        re.IGNORECASE,
    )
    inst_name = (
        inst_match.group(1).replace("\n", " ")
        if inst_match
        else "SVKM Educational Institution (AUTONOMOUS)"
    )
    metadata.append({"Property": "Institution / Trust", "Value": inst_name})

    prog_match = re.search(
        r"(Bachelor of [^\n\r,]+|Master of [^\n\r,]+|M\.Com[^\n\r,]*|B\.Sc[^\n\r,]*|B\.Com[^\n\r,]*)",
        text,
        re.IGNORECASE,
    )
    metadata.append(
        {
            "Property": "Programme",
            "Value": prog_match.group(1).strip() if prog_match else "N/A",
        }
    )

    sem_match = re.search(r"Semester\s+([IVXLCDM\d]+)", text, re.IGNORECASE)
    metadata.append(
        {
            "Property": "Semester",
            "Value": f"Semester {sem_match.group(1)}" if sem_match else "N/A",
        }
    )

    ncrf_match = re.search(r"NCr?F\s*Level[:\s]*([0-9.]+)", text, re.IGNORECASE)
    metadata.append(
        {
            "Property": "NCrF Level",
            "Value": ncrf_match.group(1) if ncrf_match else "N/A",
        }
    )

    ay_match = re.search(r"A\.Y\.?\s*([0-9]{4}\s*-\s*[0-9]{4})", text, re.IGNORECASE)
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
        f"{exam_type} held in {exam_session_match.group(1)}".strip()
        if exam_session_match
        else exam_type
    )

    metadata.append({"Property": "Examination Type", "Value": exam_type})
    metadata.append({"Property": "Examination Session", "Value": exam_session})

    return pd.DataFrame(metadata), exam_type


def parse_page_2(page):
    """Extract course catalogue and schema from Page 2 (supports both '--' and blank Sem End Max Marks)."""
    text = page.extract_text() or ""
    lines = text.split("\n")
    courses = []

    course_pattern = re.compile(
        r"^([A-Z0-9]{5,12})\s+(.+?)\s+(\d+\.\d{2})\s+(\d+|--)(?:\s+(\d+|--))?\s+(\d+)\b"
    )

    for line in lines:
        m = course_pattern.match(line.strip())
        if m:
            code, title, creds, int_m, sem_m, tot_m = m.groups()
            int_val = int(int_m) if (int_m and int_m != "--") else 0
            sem_val = int(sem_m) if (sem_m and sem_m != "--") else 0
            courses.append(
                {
                    "Course Code": code.strip(),
                    "Course Title": title.strip(),
                    "Credits": float(creds),
                    "Internal Max Marks": int_val,
                    "Sem End Max Marks": sem_val,
                    "Total Max Marks": int(tot_m),
                }
            )

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


def clean_page_lines(raw_text):
    """Filter out non-student header/footer lines, separator dashes, and execution logs."""
    lines = raw_text.split("\n")
    cleaned_lines = []

    for line in lines:
        l_str = line.strip()
        if not l_str:
            continue

        if re.fullmatch(r"[-_=\s]+", l_str):
            continue

        if any(
            h in l_str
            for h in [
                "Shri Vile Parle Kelavani Mandal",
                "SVKM's",
                "Narsee Monjee College",
                "UPG College",
                "MITHIBAI COLLEGE",
                "AFFILIATED TO UNIVERSITY",
                "NAAC RE-ACCREDITED",
                "NAAC Reaccredited",
                "Module Credits",
                "Consolidated Result",
                "COURSE DETAILS",
                "Candidate's Full Name",
                "MRK G GP C",
                "CG ; CG ;",
            ]
        ):
            continue

        if re.search(
            r"^\s*/\s*:\s*Female|Grade\s+Marks\s+GP|Date:\s*\(Execution|Page:\s*\d+/\d+|CONTROLLER OF EXAMINATIONS|PRINCIPAL",
            l_str,
            re.IGNORECASE,
        ):
            continue

        cleaned_lines.append(l_str)

    return cleaned_lines


def parse_header_line(line):
    """Accurately extracts SNo, SAP, Roll, PRN, Name, Gender, ABC ID, Abeyance, and LD/PwD flag."""
    line_clean = clean_dashes_and_artifacts(line)
    is_abeyance = bool(re.search(r"Result\s+in\s+abeyance", line_clean, re.I))
    if is_abeyance:
        line_clean = re.sub(r"Result\s+in\s+abeyance", "", line_clean, flags=re.I).strip()

    parts = line_clean.split()
    if len(parts) < 3 or not parts[0].isdigit() or not parts[1].isdigit():
        return None

    s_no = parts[0]
    sap_no = parts[1]
    roll_no = parts[2]

    rem_parts = parts[3:]
    prn_no = ""
    abc_id = ""

    if rem_parts and (
        rem_parts[0].startswith("MU")
        or (rem_parts[0].isdigit() and len(rem_parts[0]) >= 14)
    ):
        prn_no = rem_parts.pop(0)

    if rem_parts and rem_parts[-1].isdigit() and len(rem_parts[-1]) == 12:
        abc_id = rem_parts.pop()

    full_name = " ".join(rem_parts)
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


def split_pipe_row(raw_line, label):
    """Split row on '|' preserving empty cells for practicals alignment."""
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
    """Parse a single subject's Total-line segment preserving all marks, grades, and $ symbols."""
    seg = (seg or "").strip()
    if not seg:
        return "", "", "", ""

    m_start = re.match(
        r"^([0-9]+(?:\s*\+\s*[0-9]+)?|[0-9\+ABNV®#\$]+|\-\-)\s+(.*)$", seg
    )
    if not m_start:
        tokens = seg.split()
        if len(tokens) == 1:
            return tokens[0], "", "", ""
        return "", "", "", ""

    marks = m_start.group(1)
    rest = m_start.group(2).strip()

    rest = re.sub(r"\$([0-9]+)", r"$ \1", rest)
    rest = re.sub(r"([A-Za-z0-9\+\-]+)\s+\$", r"\1$", rest)
    rest = re.sub(r"\bF00\b", "F 0 0", rest)
    rest = re.sub(r"\bF\s+00\b", "F 0 0", rest)
    rest = re.sub(r"\bPS(\d+)\b", r"P$ \1", rest)

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


def extract_subjects_universal(chunk_lines):
    """Universally extracts Subject Codes, IA Marks, TH Marks, Total details, and Grand Total across all exam formats."""
    subj_codes = []
    ia_scores = []
    th_scores = []
    tot_info = []
    cg_scores = []
    gt_score = ""
    percentage = ""

    for l in chunk_lines:
        gt_match = re.search(r"(\d+/\d+)\s*([\d\.]+%?)?", l)
        if gt_match:
            gt_score = gt_match.group(1)
            percentage = gt_match.group(2) or ""
            break

    ia_with_codes = None
    for l in chunk_lines:
        if l.strip().startswith("IA") and re.search(r"[A-Z]{3,8}\d{2,4}", l):
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
                m_gt = re.search(r"(\d+/\d+)\s*([\d\.]+%?)?", seg)
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
        for l in chunk_lines:
            codes = re.findall(r"\b[A-Z]{2,8}\d{2,4}[A-Z0-9]?\b", l)
            if len(codes) >= 3:
                subj_codes = codes
                break
        for l in chunk_lines:
            if l.startswith("IA"):
                ia_scores = split_pipe_row(l, "IA")

    total_line_idx = None
    for i, l in enumerate(chunk_lines):
        if l.startswith("TH"):
            th_scores = split_pipe_row(l, "TH")
        elif l.startswith("Total"):
            tot_info = split_pipe_row(l, "Total")
            total_line_idx = i

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


def extract_semester_and_final_metrics(rem_str):
    """Dynamically parses semester progression metrics, individual grade value, and clean status."""
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
            and sem_raw[:-1]
            in ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]
        ):
            sem_num = sem_raw[:-1]
        else:
            sem_num = sem_raw

        body = raw_body[len(m_num.group(1)) :]

        c_m = re.search(r"\bC\s*[-:\.]\s*(\d+|--)", body)
        if not c_m:
            c_lead = re.match(r"^[-:\.]\s*(\d+)", body)
            c_val = c_lead.group(1) if c_lead else ""
        else:
            c_val = c_m.group(1) if c_m.group(1) != "--" else ""

        cg_m = re.search(r"\bCG\s*[-:\.]\s*(\d+|--)", body)
        sgpa_m = re.search(r"\bSGPA\s*[-:\.]\s*([0-9\.]+|AD|--)", body)
        att_m = re.search(r"Attempt[/s]*[-:\.]\s*(\d+)", body, re.I)
        fail_m = re.search(r"\bF\s*\(\s*(\d+)\s*\)", body, re.I)

        sem_key = f"Sem_{sem_num}"
        sem_data[f"{sem_key}_Credits"] = c_val
        sem_data[f"{sem_key}_CG"] = (
            cg_m.group(1) if (cg_m and cg_m.group(1) != "--") else ""
        )
        sem_data[f"{sem_key}_SGPA"] = (
            sgpa_m.group(1) if (sgpa_m and sgpa_m.group(1) != "--") else ""
        )
        sem_data[f"{sem_key}_Attempts"] = att_m.group(1) if att_m else ""
        sem_data[f"{sem_key}_Fail_Info"] = (
            f"F({fail_m.group(1)})" if fail_m else ""
        )

    tot_cred = re.search(r"Total\s+Credit\s*:\s*(\d+|--)?", rem_clean, re.I)
    tot_cg = re.search(r"Total\s+CG\s*:\s*(\d+|--)?", rem_clean, re.I)
    cgpa = re.search(r"\bCGPA\s*:\s*([0-9\.]+|--)?", rem_clean, re.I)

    final_grade_m = re.search(
        r"Final\s+Grade\s*:\s*([OABCDPF0][\+\-!]?|--)", rem_clean, re.I
    )
    final_grade_val = (
        final_grade_m.group(1).strip()
        if (final_grade_m and final_grade_m.group(1) != "--")
        else ""
    )
    if final_grade_val == "0":
        final_grade_val = "O"

    sem_grade_m = re.search(
        r"Grade\s*[-:\s]\s*([OABCDPF0][\+\-!]?|--)(?:\s|$|\]|;|,)", rem_clean, re.I
    )
    extracted_grade = (
        sem_grade_m.group(1).strip()
        if (sem_grade_m and sem_grade_m.group(1) != "--")
        else ""
    )
    if extracted_grade == "0":
        extracted_grade = "O"

    cum_perc = re.search(
        r"Cum\.?\s*Perc\.?\s*:\s*([0-9\.]+%?|--)?", rem_clean, re.I
    )

    remark_m = re.search(r"Remark\s*:\s*([^\];]+)", rem_clean, re.I)
    raw_remark = remark_m.group(1).strip() if remark_m else ""
    status_m = re.search(
        r"\b(Successful|Unsuccessful)([\$]*)", raw_remark, re.I
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
        "Total Credit": (
            tot_cred.group(1)
            if (tot_cred and tot_cred.group(1) and tot_cred.group(1) != "--")
            else ""
        ),
        "Total CG": (
            tot_cg.group(1)
            if (tot_cg and tot_cg.group(1) and tot_cg.group(1) != "--")
            else ""
        ),
        "CGPA": (
            cgpa.group(1)
            if (cgpa and cgpa.group(1) and cgpa.group(1) != "--")
            else ""
        ),
        "Final Grade": final_grade_val,
        "Cumulative Percentage": (
            cum_perc.group(1)
            if (cum_perc and cum_perc.group(1) and cum_perc.group(1) != "--")
            else ""
        ),
        "Result Status": result_status,
        "Grade": extracted_grade,
        "Eligibility Status": eligibility_status,
    }


def parse_result_pages(
    pdf_pages, course_title_map=None, exam_type="Regular"
):
    """Extract all student records and generate a structured Result_Data dataset."""
    if course_title_map is None:
        course_title_map = {}

    all_lines = []

    for p_idx, page in enumerate(pdf_pages):
        p_no = p_idx + 1
        p_text = page.extract_text() or ""

        if "Consolidated Result" not in p_text and not re.search(
            r"^\d{1,4}\s+\d{9,14}", p_text, re.M
        ):
            if len(pdf_pages) > 2 and p_no < 3:
                continue

        cleaned = clean_page_lines(p_text)
        all_lines.extend(cleaned)

    student_chunks = []
    current_chunk = []

    for line in all_lines:
        if re.match(r"^\d{1,4}\s+\d{9,14}\s+[A-Z0-9]+", line):
            if current_chunk:
                student_chunks.append(current_chunk)
                current_chunk = []
        if current_chunk or re.match(r"^\d{1,4}\s+\d{9,14}\s+[A-Z0-9]+", line):
            current_chunk.append(line)

    if current_chunk:
        student_chunks.append(current_chunk)

    detailed_records = []
    all_discovered_sem_keys = set()

    for chunk in student_chunks:
        if not chunk:
            continue

        header_info = parse_header_line(chunk[0])
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

        (
            subj_codes,
            ia_scores,
            th_scores,
            tot_info,
            cg_scores,
            gt_score,
            percentage,
        ) = extract_subjects_universal(chunk)

        rem_str = " ".join(chunk)
        eval_metrics = extract_semester_and_final_metrics(rem_str)
        all_discovered_sem_keys.update(eval_metrics["semesters"].keys())

        marks_obt, max_marks = "", ""
        if gt_score and "/" in gt_score:
            parts = gt_score.split("/")
            marks_obt = parts[0].strip()
            max_marks = parts[1].strip()

        raw_percentage_str = percentage.replace("%", "").strip() if percentage else ""

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

            cg_entry = cg_scores[idx] if idx < len(cg_scores) else ""
            cg_clean = cg_entry.replace("$", "").replace("#", "").strip()
            if cg_clean.isdigit():
                cg_val = int(cg_clean)
            elif (
                gp.replace("$", "").replace("#", "").isdigit()
                and creds.replace("$", "").replace("#", "").isdigit()
            ):
                cg_val = int(gp.replace("$", "").replace("#", "")) * int(
                    creds.replace("$", "").replace("#", "")
                )
            else:
                cg_val = ""

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

    roman_order = [
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

    def sem_sort_key(col_name):
        m = re.search(r"Sem_([IVXLCDM\d]+)_", col_name)
        if m:
            num = m.group(1).upper()
            return (
                roman_order.index(num)
                if num in roman_order
                else (int(num) if num.isdigit() else 99)
            )
        return 99

    sem_suffix_order = ["Credits", "SGPA", "Attempts", "CG", "Fail_Info"]

    def sem_col_sort_key(col_name):
        m = re.match(r"^Sem_([IVXLCDM\d]+)_(.*)$", col_name)
        sem_name = m.group(1) if m else ""
        suffix = m.group(2) if m else ""
        suffix_rank = (
            sem_suffix_order.index(suffix)
            if suffix in sem_suffix_order
            else len(sem_suffix_order)
        )
        return (sem_sort_key(col_name), sem_name, suffix_rank, col_name)

    sorted_sem_cols = sorted(all_discovered_sem_keys, key=sem_col_sort_key)
    df_detailed = pd.DataFrame(detailed_records)

    if not df_detailed.empty:
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
            c
            for c in df_detailed.columns
            if re.match(r"^Sub\d+", str(c))
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
        df_detailed = df_detailed.reindex(columns=all_ordered)

        non_empty_cols = [
            c
            for c in df_detailed.columns
            if not df_detailed[c]
            .fillna("")
            .astype(str)
            .str.strip()
            .eq("")
            .all()
        ]
        df_detailed = df_detailed[non_empty_cols]

    return df_detailed


# ==============================================================================
# INTELLIGENT RESULT VERIFICATION AND AUDIT ENGINE
# ==============================================================================


def verify_and_audit_results(df_result, df_p2):
    """
    Comprehensive Result Verification and Audit Engine.
    Validates every calculation, grade, credit, SGPA, CGPA, result status,
    summary field, and special case in Result_Data.

    Rules Incorporated:
    1. Lower semester Fail_Info F(n) with n > 0 validly makes Eligibility Status
       'Not Eligible for next semester' even when the current semester is passed.
    2. 'AB' in IA Marks or TH Marks is a valid entry (evaluated as 0 marks and F grade).
    3. Grace marks with '+' sign in TH Marks (or IA Marks) are valid entries and
       included in Subject Total Marks and component passing checks.
    4. Component passing requires >= 40% of Internal Max Marks (when > 0) and
       >= 40% of Sem End Max Marks (when > 0).

    Returns:
    - df_audited: Result_Data with the 'Verification Errors' column appended
    - df_verification_report: Detailed row-level error report for 'Verification_Report'
    - df_verification_summary: High-level KPI summary table
    - df_error_wise_summary: Error category breakdown table
    - error_cells_set: Set of (row_idx, col_name) tuples to highlight in red fill + bold yellow font
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
        return df_result, empty_report, empty_summary, empty_cat, set()

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

    # Discover subject numbers (Sub1..SubN)
    cols = list(df_audited.columns)
    sub_nums = sorted(
        {
            int(m.group(1))
            for c in cols
            if (m := re.match(r"^Sub(\d+)\s", str(c)))
        }
    )

    # Discover semester prefixes in ascending order
    roman_order = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]
    sem_prefixes = []
    for c in cols:
        m = re.match(r"^Sem_([IVXLCDM\d]+)_", str(c))
        if m and m.group(1) not in sem_prefixes:
            sem_prefixes.append(m.group(1))

    def _sem_rank(s):
        return (
            roman_order.index(s.upper())
            if s.upper() in roman_order
            else (int(s) if s.isdigit() else 99)
        )

    sem_prefixes.sort(key=_sem_rank)
    highest_sem = sem_prefixes[-1] if sem_prefixes else None
    lower_sems = sem_prefixes[:-1] if len(sem_prefixes) > 1 else []

    # Check if cumulative columns exist in df_audited
    has_cumulative_cols = any(
        c in df_audited.columns
        for c in ["Total Credit", "Total CG", "CGPA", "Final Grade"]
    )

    report_rows = []
    error_cells_set = set()
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
            error_cells_set.add((row_idx, "Verification Errors"))
            row_error_messages.append(
                f"[{severity}] {field_name}: {description} (Expected: {expected_val}, Actual: {actual_val})"
            )
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

        # Check ABC ID and accidental concatenation of 13-digit ID into Candidate Full Name
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

        # Check Result in Abeyance special case
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
        # 2. SUBJECT-LEVEL INDEPENDENT VALIDATIONS (Sub1..SubN)
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

            # Validate IA Marks ('AB' and '+' grace marks are valid entries)
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
                # Valid 'AB' entry in IA Marks
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

            # Validate TH Marks ('AB' and '+' grace marks are valid entries)
            if sem_max == Decimal("0"):
                # Internal-only course (e.g., UBNSS174, UBYOG176, Field Project)
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
                    # Valid 'AB' entry in TH Marks
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

            # Compute Expected Subject Total Marks = Subject IA Marks + Subject TH Marks (including valid '+' grace marks)
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

            # Validate Subject Grade (accounting for AB/NV and 40% component-level minimum passing criteria)
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

            # Validate Subject Credits (C) against Page 2 Scheme (0 is valid when Grade is F)
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

            # Validate Subject CG = Subject GP * Subject Credit (0 if Grade = F)
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
                    f"Missing Subject CG for {sub_code}.",
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
        # 3. SEMESTER-LEVEL VALIDATIONS (HIGHEST SEMESTER & LOWER SEMESTERS)
        # ------------------------------------------------------------------
        # Check if any lower semester has Fail_Info = F(n) where n > 0
        lower_sem_fail_counts = {}
        for s in lower_sems:
            f_col = f"Sem_{s}_Fail_Info"
            if f_col in df_audited.columns:
                f_val = str(row.get(f_col, "")).strip()
                m_f = re.search(r"F\s*\(\s*(\d+)\s*\)", f_val, re.I)
                if m_f and int(m_f.group(1)) > 0:
                    lower_sem_fail_counts[s] = int(m_f.group(1))
        has_lower_sem_fail = len(lower_sem_fail_counts) > 0

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

            # Highest Semester Credits = sum of all subject credits (blank '--' is valid when failed_subject_count > 0)
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

            # Highest Semester CG = sum of all subject CG values (blank '--' is valid when failed_subject_count > 0)
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

            # Highest Semester SGPA = Highest Semester CG / Highest Semester Credits (blank '--' is valid when failed_subject_count > 0)
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
                    elif expected_sgpa is not None and abs(
                        actual_sem_sgpa - expected_sgpa
                    ) > Decimal("0.01"):
                        record_issue(
                            col_sem_sgpa,
                            f"{expected_sgpa:.2f}",
                            str(row.get(col_sem_sgpa, "")),
                            f"{col_sem_sgpa} mismatch: {int(sum_sub_cg)} / {int(sum_sub_scheme_credits)} = {expected_sgpa:.2f}, but found {actual_sem_sgpa:.2f}.",
                            severity="Critical",
                            category="Semester SGPA Mismatch",
                        )

            # Semester Fail_Info must equal F(n) where n is count of subjects with Grade F
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

            # Semester Grade derived from Semester SGPA (blank '--' or 'F' is valid if failed_subject_count > 0)
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
        # 4. OVERALL RESULT CALCULATIONS (GRAND TOTAL, PERCENTAGE, STATUS)
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
                # Percentage is intentionally omitted in Gazette for Unsuccessful students
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

        # Cumulative validations (when Total Credit / Total CG / CGPA / Final Grade exist)
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

        # Result Status validation (current semester status)
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

        # Eligibility Status consistency check (incorporating lower semester F(n) with n > 0)
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
                        else "lower semester Fail_Info "
                        + ", ".join(
                            f"Sem {k}: F({v})"
                            for k, v in lower_sem_fail_counts.items()
                        )
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
                        f"Eligibility Status '{elig_raw}' conflicts with Successful Result Status (0 failed subjects in current and lower semesters).",
                        severity="Critical",
                        category="Eligibility Status Inconsistency",
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
    )


# ==============================================================================
# EXCEL WORKBOOK FORMATTING & EXPORT (WITH ERROR HIGHLIGHTING & AUDIT SHEETS)
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
    import colorsys

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


def save_to_formatted_excel(
    sheets_dict, exam_type="Regular", error_cells_set=None
):
    """Export formatted Excel workbook with verification highlighting (solid red fill + bold yellow font) and audit sheets."""
    if error_cells_set is None:
        error_cells_set = set()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
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

            # Mandatory Verification Discrepancy Style: Solid Red Fill + Bold Yellow Font
            verification_error_fill = PatternFill(
                start_color="FF0000", end_color="FF0000", fill_type="solid"
            )
            verification_error_font = Font(
                name="Calibri", size=11, bold=True, color="FFFF00"
            )

            if sheet_name == "Page_1":
                for col_idx in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = header_fill_navy
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row=row_idx, column=1).border = thin_border
                    ws.cell(row=row_idx, column=1).alignment = Alignment(
                        vertical="center", horizontal="left"
                    )
                    ws.cell(row=row_idx, column=2).border = thin_border
                    ws.cell(row=row_idx, column=2).alignment = Alignment(
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
                name_col_idx = (
                    list(df.columns).index("Candidate Full Name") + 1
                    if "Candidate Full Name" in df.columns
                    else 5
                )
                freeze_col_letter = get_column_letter(name_col_idx + 1)
                ws.freeze_panes = f"{freeze_col_letter}2"

                final_section_cols = [
                    "Total Credit",
                    "Total CG",
                    "CGPA",
                    "Final Grade",
                    "Cumulative Percentage",
                    "Result Status",
                    "Grade",
                    "Eligibility Status",
                ]

                subject_end_col_indices = [
                    idx + 1
                    for idx, c in enumerate(df.columns)
                    if str(c).endswith(" CG")
                ]

                sem_end_col_indices = []
                current_sem = None
                last_idx = None
                for idx, c in enumerate(df.columns, start=1):
                    m = re.match(r"^Sem_([IVXLCDM\d]+)_", str(c))
                    if m:
                        s_name = m.group(1)
                        if current_sem and s_name != current_sem:
                            sem_end_col_indices.append(last_idx)
                        current_sem = s_name
                        last_idx = idx
                if last_idx and last_idx not in sem_end_col_indices:
                    sem_end_col_indices.append(last_idx)

                pre_total_credit_col_indices = []
                if "Total Credit" in df.columns:
                    tc_idx = list(df.columns).index("Total Credit")
                    if tc_idx > 0:
                        pre_total_credit_col_indices.append(tc_idx)

                all_thick_right_col_indices = set(
                    subject_end_col_indices
                    + sem_end_col_indices
                    + pre_total_credit_col_indices
                )

                subject_numbers = []
                for c in df.columns:
                    m_sub = re.match(r"^Sub(\d+)\s", str(c))
                    if m_sub and int(m_sub.group(1)) not in subject_numbers:
                        subject_numbers.append(int(m_sub.group(1)))
                subject_pastel_fills = {
                    num: PatternFill(
                        start_color=hex_c, end_color=hex_c, fill_type="solid"
                    )
                    for num, hex_c in zip(
                        subject_numbers,
                        generate_pastel_palette(len(subject_numbers)),
                    )
                }
                header_font_subject = Font(
                    name="Calibri", size=11, bold=True, color="1F1F1F"
                )

                for col_idx, col_name in enumerate(df.columns, start=1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.font = header_font

                    if col_name == "Verification Errors":
                        cell.fill = header_fill_audit
                    elif str(col_name).startswith("Sem_"):
                        cell.fill = header_fill_semester
                    elif col_name in final_section_cols:
                        cell.fill = header_fill_final
                    elif col_name in [
                        "Total Marks Obt.",
                        "Max Total Marks",
                        "Percentage",
                    ]:
                        cell.fill = header_fill_marks
                    elif str(col_name).startswith("Sub"):
                        m_sub = re.match(r"^Sub(\d+)\s", str(col_name))
                        if m_sub and int(m_sub.group(1)) in subject_pastel_fills:
                            cell.fill = subject_pastel_fills[
                                int(m_sub.group(1))
                            ]
                            cell.font = header_font_subject
                        else:
                            cell.fill = header_fill_subject
                    else:
                        cell.fill = header_fill_navy

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
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )

                grade_col_indices = [
                    idx + 1
                    for idx, c in enumerate(df.columns)
                    if str(c).endswith(" Grade")
                    and not str(c).startswith("Final")
                    and not str(c).startswith("Overall")
                    and str(c) != "Grade"
                ]

                sub_blocks = []
                for g_col in grade_col_indices:
                    start_col = g_col - 5
                    end_col = g_col + 3
                    sub_blocks.append((g_col, start_col, end_col))

                is_pwd_col_idx = (
                    list(df.columns).index("Is_PwD") + 1
                    if "Is_PwD" in df.columns
                    else None
                )

                for row_idx in range(2, ws.max_row + 1):
                    df_row_idx = row_idx - 2
                    status_val = str(
                        ws.cell(
                            row=row_idx,
                            column=list(df.columns).index("Result Status") + 1,
                        ).value
                        or ""
                        if "Result Status" in df.columns
                        else ""
                    )
                    grade_val = str(
                        ws.cell(
                            row=row_idx,
                            column=list(df.columns).index("Grade") + 1,
                        ).value
                        or ""
                        if "Grade" in df.columns
                        else ""
                    )
                    final_grade_val = str(
                        ws.cell(
                            row=row_idx,
                            column=list(df.columns).index("Final Grade") + 1,
                        ).value
                        or ""
                        if "Final Grade" in df.columns
                        else ""
                    )
                    eligibility_val = str(
                        ws.cell(
                            row=row_idx,
                            column=list(df.columns).index("Eligibility Status")
                            + 1,
                        ).value
                        or ""
                        if "Eligibility Status" in df.columns
                        else ""
                    )

                    is_unsuccessful_row = (
                        "Unsuccessful" in status_val
                        or "Not Eligible" in eligibility_val
                        or grade_val.strip() == "F"
                        or final_grade_val.strip() == "F"
                    )

                    is_pwd_row = False
                    if is_pwd_col_idx:
                        is_pwd_row = (
                            str(
                                ws.cell(
                                    row=row_idx, column=is_pwd_col_idx
                                ).value
                                or ""
                            ).strip()
                            == "Yes"
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
                            c_name = df.columns[c - 1]
                            c_val = str(
                                ws.cell(row=row_idx, column=c).value or ""
                            ).strip()

                            if "$" in c_val:
                                carried_cols_in_row.add(c)
                            elif (
                                exam_type == "Re-examination"
                                and (
                                    "IA Marks" in c_name or "TH Marks" in c_name
                                )
                                and c_val
                                and c_val not in ["--", "NV", "AB", "NaN", "nan"]
                            ):
                                fresh_mark_cols_in_row.add(c)

                    for col_idx in range(1, len(df.columns) + 1):
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

                        col_name = df.columns[col_idx - 1]
                        if (
                            col_name in ["Candidate Full Name", "Verification Errors"]
                            or "Course Title" in col_name
                        ):
                            cell.alignment = Alignment(
                                vertical="center", horizontal="left"
                            )
                        else:
                            cell.alignment = Alignment(
                                vertical="center", horizontal="center"
                            )

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

                        # HIGHEST PRIORITY: Highlight verification errors with solid red fill & bold yellow font
                        if (df_row_idx, col_name) in error_cells_set:
                            cell.fill = verification_error_fill
                            cell.font = verification_error_font

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

                for row_idx in range(2, ws.max_row + 1):
                    sev_val = ""
                    if "Severity" in df.columns:
                        sev_col_i = list(df.columns).index("Severity") + 1
                        sev_val = str(
                            ws.cell(row=row_idx, column=sev_col_i).value or ""
                        ).strip()

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
                            cell.font = crit_font if col_name == "Severity" else normal_font
                        elif sev_val == "Warning":
                            cell.fill = warn_fill
                            cell.font = warn_font if col_name == "Severity" else normal_font
                        else:
                            cell.font = normal_font

            # Auto-fit columns
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_len = len(str(cell.value or ""))
                    if val_len > max_len:
                        max_len = val_len
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 8), 65)

        # Build Student_Wise_View with error cell highlighting propagated
        _add_student_wise_view_sheet(
            writer,
            sheets_dict.get("Result_Data"),
            exam_type=exam_type,
            error_cells_set=error_cells_set,
        )

    output.seek(0)
    return output


# ==============================================================================
# STUDENT-WISE VIEW SHEET (WITH VERIFICATION HIGHLIGHTING)
# ==============================================================================

_SWV_SHEET_NAME = "Student_Wise_View"


def _swv_value(val, numeric=False):
    """Normalise a Result_Data cell for display (NaN -> '', optional numeric cast)."""
    import numbers

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
    writer, df, exam_type="Regular", error_cells_set=None
):
    """Append the 'Student_Wise_View' worksheet with verification error highlighting."""
    from openpyxl.worksheet.properties import PageSetupProperties

    if error_cells_set is None:
        error_cells_set = set()

    ws = writer.book.create_sheet(_SWV_SHEET_NAME)
    ws.views.sheetView[0].showGridLines = False

    for col_letter, width in zip(
        "ABCDEFGHI", [20, 50, 12, 12, 12, 12, 12, 12, 12]
    ):
        ws.column_dimensions[col_letter].width = width

    if df is None or df.empty:
        ws["A1"] = "No student records available."
        return ws

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

    def put(r1, c1, r2, c2, value=None, font=None, fill=None, align=None, fmt=None):
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
        return top_left

    cols = list(df.columns)
    sub_nums = []
    for c in cols:
        m = re.match(r"^Sub(\d+)\s", str(c))
        if m and int(m.group(1)) not in sub_nums:
            sub_nums.append(int(m.group(1)))
    sub_nums.sort()

    sem_names = []
    for c in cols:
        m = re.match(
            r"^Sem_([IVXLCDM\d]+)_(Credits|SGPA|Attempts|Fail_Info)$", str(c)
        )
        if m and m.group(1) not in sem_names:
            sem_names.append(m.group(1))

    def get(row, col, numeric=False):
        if col not in row.index:
            return ""
        return _swv_value(row[col], numeric=numeric)

    info_layout = [
        [("SNo", (1, 1), (2, 2)), ("SAP No", (3, 4), (5, 6)), ("Roll No", (7, 7), (8, 9))],
        [("PRN No", (1, 1), (2, 2)), ("ABC ID", (3, 4), (5, 6)), ("Gender", (7, 7), (8, 9))],
        [
            ("Candidate Full Name", (1, 1), (2, 2)),
            ("Examination Type", (3, 4), (5, 6)),
            ("Is_PwD", (7, 7), (8, 9)),
        ],
    ]
    sub_headers = [
        "Sub Code", "Course Title", "IA Marks", "TH Marks",
        "Total Marks", "Grade", "GP", "Credits", "CG",
    ]
    sem_headers = ["Semester", "Credits", "SGPA", "Attempts", "Fail_Info"]
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
        put(r, 1, r, 9, f"Student {s_idx} of {total_students}", font_banner, fill_navy, center)
        ws.row_dimensions[r].height = 24
        r += 1
        for info_row in info_layout:
            for label, (lc1, lc2), (vc1, vc2) in info_row:
                put(r, lc1, r, lc2, label, font_label, fill_label, left)
                is_err = (df_row_idx, label) in error_cells_set
                put(
                    r,
                    vc1,
                    r,
                    vc2,
                    str(get(row, label)),
                    font_error if is_err else font_value_bold,
                    fill_error if is_err else None,
                    left,
                )
            r += 1

        put(r, 1, r, 9, "Subject-Wise Performance", font_header, fill_subject, center)
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
            for c_idx, (v, src_col) in enumerate(zip(values, col_names_for_sub), start=1):
                v_str = str(v)
                if (df_row_idx, src_col) in error_cells_set:
                    fill_c, font_c = fill_error, font_error
                elif is_failed:
                    fill_c, font_c = fill_fail, font_fail
                elif "$" in v_str:
                    fill_c, font_c = fill_carried, font_normal
                elif (
                    exam_type == "Re-examination"
                    and c_idx in (3, 4)
                    and v_str not in ["", "--", "NV", "AB", "NaN", "nan"]
                ):
                    fill_c, font_c = fill_fresh, font_normal
                else:
                    fill_c, font_c = None, font_normal
                put(r, c_idx, r, c_idx, v, font_c, fill_c, left if c_idx == 2 else center)
            if len(str(title)) > 52:
                ws.row_dimensions[r].height = 30
            r += 1
            subject_rows += 1

        if subject_rows == 0:
            put(r, 1, r, 9, "No subject-wise data available for this student.", font_normal, None, center)
            r += 1

        put(r, 1, r, 5, "Semester Details", font_header, fill_semester, center)
        put(r, 6, r, 9, "Overall Result Details", font_header, fill_final, center)
        r += 1
        for c_idx, h in enumerate(sem_headers, start=1):
            put(r, c_idx, r, c_idx, h, font_header, fill_semester, center)
        put(r, 6, r, 7, "Metric", font_header, fill_final, center)
        put(r, 8, r, 9, "Value", font_header, fill_final, center)
        r += 1

        sem_rows = []
        for s in sem_names:
            s_cols = [
                f"Sem_{s}_Credits",
                f"Sem_{s}_SGPA",
                f"Sem_{s}_Attempts",
                f"Sem_{s}_Fail_Info",
            ]
            vals = [
                get(row, s_cols[0], True),
                get(row, s_cols[1], True),
                get(row, s_cols[2]),
                get(row, s_cols[3]),
            ]
            if any(str(v).strip() != "" for v in vals) or any(
                (df_row_idx, sc) in error_cells_set for sc in s_cols
            ):
                sem_rows.append((s, [f"Semester {s}"] + vals, [None] + s_cols))

        block_rows = max(len(sem_rows), len(overall_items), 1)
        for k in range(block_rows):
            if k < len(sem_rows):
                _, s_vals, s_src_cols = sem_rows[k]
                for c_idx, (v, sc_name) in enumerate(zip(s_vals, s_src_cols), start=1):
                    is_err = sc_name is not None and (df_row_idx, sc_name) in error_cells_set
                    put(
                        r + k,
                        c_idx,
                        r + k,
                        c_idx,
                        v,
                        font_error if is_err else font_normal,
                        fill_error if is_err else None,
                        center,
                        fmt="0.00" if c_idx == 3 else None,
                    )
            elif k == 0:
                put(r, 1, r, 5, "No semester data available.", font_normal, None, center)

            if k < len(overall_items):
                label, numeric, fmt = overall_items[k]
                v = get(row, label, numeric)
                is_err = (df_row_idx, label) in error_cells_set
                bad = label in ("Result Status", "Eligibility Status") and (
                    "Unsuccessful" in str(v) or "Not Eligible" in str(v)
                )
                put(r + k, 6, r + k, 7, label, font_label, fill_label, left)
                put(
                    r + k,
                    8,
                    r + k,
                    9,
                    v,
                    font_error if is_err else (font_bad if bad else font_value_bold),
                    fill_error if is_err else None,
                    center,
                    fmt=fmt if isinstance(v, (int, float)) else None,
                )
                if label == "Eligibility Status" and len(str(v)) > 22:
                    ws.row_dimensions[r + k].height = 30
        r += block_rows

        # Display Verification Errors banner if any exist for this student
        ver_err_text = str(get(row, "Verification Errors")).strip()
        if ver_err_text:
            put(r, 1, r, 2, "Verification Errors", font_error, fill_error, center)
            put(r, 3, r, 9, ver_err_text, font_error, fill_error, left)
            ws.row_dimensions[r].height = 32
            r += 1

        r += 2

    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    return ws


# ==============================================================================
# INTERACTIVE STREAMLIT HTML PREVIEW (WITH RED FILL + BOLD YELLOW FONT ERRORS)
# ==============================================================================


def render_detailed_table_html(df, exam_type="Regular", error_cells_set=None):
    """Renders interactive HTML preview in Streamlit with visual cues and red/yellow error cells."""
    if df.empty:
        st.info("No result records found.")
        return

    if error_cells_set is None:
        error_cells_set = set()

    cols = list(df.columns)
    final_section_cols = [
        "Total Credit",
        "Total CG",
        "CGPA",
        "Final Grade",
        "Cumulative Percentage",
        "Result Status",
        "Grade",
        "Eligibility Status",
    ]

    subject_end_cols = [c for c in cols if str(c).endswith(" CG")]
    sem_end_cols = []
    current_sem = None
    last_col = None
    for c in cols:
        m = re.match(r"^Sem_([IVXLCDM\d]+)_", str(c))
        if m:
            s_name = m.group(1)
            if current_sem and s_name != current_sem:
                sem_end_cols.append(last_col)
            current_sem = s_name
            last_col = c
    if last_col and last_col not in sem_end_cols:
        sem_end_cols.append(last_col)

    pre_tc_cols = []
    if "Total Credit" in cols:
        tc_i = cols.index("Total Credit")
        if tc_i > 0:
            pre_tc_cols.append(cols[tc_i - 1])

    all_thick_cols = set(subject_end_cols + sem_end_cols + pre_tc_cols)

    grade_cols = [
        c
        for c in cols
        if str(c).endswith(" Grade")
        and not str(c).startswith("Final")
        and not str(c).startswith("Overall")
        and str(c) != "Grade"
    ]

    sub_ranges = []
    for g in grade_cols:
        g_idx = cols.index(g)
        s_code_col = cols[g_idx - 5]
        e_cg_col = cols[g_idx + 3]
        sub_ranges.append((g, s_code_col, e_cg_col, g_idx - 5, g_idx + 3))

    html = """
    <style>
    .detailed-table-container {
        width: 100%;
        overflow-x: auto;
        max-height: 550px;
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
        padding: 8px 10px;
        text-align: center;
        position: sticky;
        top: 0;
        z-index: 2;
        border: 1px solid #d9d9d9;
        font-weight: bold;
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

    .thick-right {
        border-right: 3px solid #1F4E79 !important;
    }
    .pwd-row {
        background-color: #E8F4F8 !important;
    }
    .carried-forward-cell {
        background-color: #EAECEE !important;
    }
    .fresh-mark-cell {
        background-color: #FFF3CD !important;
    }
    .fail-cell {
        background-color: #FADBD8 !important;
        color: #9C0006 !important;
    }
    .verification-error-cell {
        background-color: #FF0000 !important;
        color: #FFFF00 !important;
        font-weight: bold !important;
    }
    </style>
    <div class="detailed-table-container">
    <table class="detailed-table">
    <thead>
    <tr>
    """
    for col in cols:
        th_cls = ["thick-right"] if col in all_thick_cols else []
        if col == "Verification Errors":
            th_cls.append("th-audit")
        elif str(col).startswith("Sem_"):
            th_cls.append("th-semester")
        elif col in final_section_cols:
            th_cls.append("th-final")
        elif col in ["Total Marks Obt.", "Max Total Marks", "Percentage"]:
            th_cls.append("th-marks")
        elif str(col).startswith("Sub"):
            th_cls.append("th-subject")
        else:
            th_cls.append("th-navy")

        html += f'<th class="{" ".join(th_cls)}">{col}</th>'
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

        for g_col, _, _, start_i, end_i in sub_ranges:
            if "F" in str(row.get(g_col, "")).strip():
                for col_name in cols[start_i : end_i + 1]:
                    failed_cols.add(col_name)

            for col_name in cols[start_i : end_i + 1]:
                c_val = str(row.get(col_name, "")).strip()
                if "$" in c_val:
                    carried_cols.add(col_name)
                elif (
                    exam_type == "Re-examination"
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
        row_cls_str = (
            f' class="{" ".join(row_classes)}"' if row_classes else ""
        )

        html += f"<tr{row_cls_str}>"
        for col in cols:
            val = row[col] if pd.notna(row[col]) else ""
            classes = []
            if col in all_thick_cols:
                classes.append("thick-right")
            if (
                col in ["Candidate Full Name", "Verification Errors"]
                or "Course Title" in col
            ):
                classes.append("text-left-cell")

            if (df_row_idx, col) in error_cells_set:
                classes.append("verification-error-cell")
            elif col in failed_cols:
                classes.append("fail-cell")
            elif col in fresh_cols:
                classes.append("fresh-mark-cell")
            elif col in carried_cols:
                classes.append("carried-forward-cell")

            cls_str = f' class="{" ".join(classes)}"' if classes else ""
            html += f"<td{cls_str}>{val}</td>"
        html += "</tr>"

    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)


# ==============================================================================
# MAIN ENTRY POINTS
# ==============================================================================


def show():
    """Main view function invoked by app.py router."""
    st.title("🛡️ Intelligent Result Verification & Audit Tool")
    st.caption(
        "Automated Result Gazette Extraction, Mathematical & Rule Verification, and Discrepancy Audit Engine."
    )

    uploaded_files = st.file_uploader(
        "Upload Result Gazette PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.subheader(f"📄 Auditing: {uploaded_file.name}")

            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                st.info(f"Total Pages Detected: **{total_pages}**")

                # 1. Page 1 Metadata & Exam Type Detection
                p1_text = pdf.pages[0].extract_text() or ""
                df_p1, exam_type = parse_page_1(p1_text)

                # 2. Page 2 Course Scheme
                p2_page = pdf.pages[1] if total_pages > 1 else pdf.pages[0]
                df_p2 = parse_page_2(p2_page)

                course_map = {}
                if (
                    not df_p2.empty
                    and "Course Code" in df_p2.columns
                    and "Course Title" in df_p2.columns
                ):
                    course_map = dict(
                        zip(df_p2["Course Code"], df_p2["Course Title"])
                    )

                # 3. Dynamic Results Parsing
                df_result_raw = parse_result_pages(
                    pdf.pages,
                    course_title_map=course_map,
                    exam_type=exam_type,
                )

                # 4. Run Intelligent Result Verification and Audit Engine
                (
                    df_result_audited,
                    df_verification_report,
                    df_verification_summary,
                    df_error_wise_summary,
                    error_cells_set,
                ) = verify_and_audit_results(df_result_raw, df_p2)

                # 5. Executive KPI Metrics Banner
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

                # 6. Preview Tabs
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
                    st.markdown(
                        "##### 📌 Audited Student Results (Erroneous cells highlighted in **Solid Red Fill & Bold Yellow Font**)"
                    )
                    render_detailed_table_html(
                        df_result_audited,
                        exam_type=exam_type,
                        error_cells_set=error_cells_set,
                    )
                with t2:
                    st.markdown("##### 🚨 Detailed Verification Report (`Verification_Report`)")
                    if df_verification_report.empty:
                        st.success("All student records passed 100% of verification checks with zero discrepancies!")
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

                # 7. Save & Download Audited Excel Workbook
                sheets_data = {
                    "Page_1": df_p1,
                    "Page_2": df_p2,
                    "Result_Data": df_result_audited,
                    "Verification_Report": df_verification_report,
                    "Verification_Summary": df_verification_summary,
                    "Error_Wise_Summary": df_error_wise_summary,
                }

                excel_bytes = save_to_formatted_excel(
                    sheets_data,
                    exam_type=exam_type,
                    error_cells_set=error_cells_set,
                )

                export_file_name = (
                    f"{uploaded_file.name.rsplit('.', 1)[0]}_Audited_Report.xlsx"
                )
                st.download_button(
                    label=f"⬇️ Download Audited Excel Report ({export_file_name})",
                    data=excel_bytes,
                    file_name=export_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.divider()


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