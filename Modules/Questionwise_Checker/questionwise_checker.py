"""
==================================================
QUESTIONWISE CHECKER (MULTI-FILE & GMR INTEGRATED)
==================================================

PURPOSE
--------------------------------------------------
This Streamlit tool helps exam staff:

1. Open password-protected or unprotected Excel mark sheets exported from
   the examination system (supports multiple uploads simultaneously).
2. Optionally upload a GMR Excel file to match records by PRN/Student Number 
   and Subject Code/Modulecode to fetch 'Internal Marks' (Internal Total Obtained),
   'Total Max Marks' (from Composite AGR type), 'Semester Total Max Marks' (from 
   Semester Total AGR type), calculate 'Total Marks Obtained', and derive 'Grade' 
   dynamically based on percentage criteria.
3. Clean the raw export (remove blank/footer rows, drop unused columns).
4. Sort records so each student's rows appear in a fixed order:
   Examiner -> Moderator -> Reval 1 -> Reval 2.
5. Calculate 'Change %' for Moderator rows:
   Change % = (Moderator TotalObtained - Examiner TotalObtained) / Semester Total Max Marks * 100.
6. Check paper eligibility for 100% moderation if >10% mark change occurs 
   in >50% of moderated answer books.
7. Display 5 summary cards (Total Students, Positive Changes, Negative Changes, 
   No Changes, and 10% Change Moderation Metrics) and three preview tabs.
8. Export processed output Excel files:
   - Original Cleaned File: Reval rows removed, Moderator filtered, legacy format.
   - Formatted File: All UserTypes, GMR columns, Change %, center alignment, F2 freeze, 
     thick student borders, green/red text, and automatic Excel pop-up message if eligible.
   - Consolidated Multi-Sheet Excel option for multiple files with pop-up messages per eligible sheet.
==================================================
"""

from datetime import datetime
import io
import os
import re
import tempfile
import zipfile

import msoffcrypto
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st


# ==================================================
# CONFIGURATION
# ==================================================

PASSWORDS = [
    "Exam@108",
    "UPG@123",
    "Exam@105"
]

FIXED_COLUMNS = [
    "SrNo",
    "ExamAssingment_ID",
    "Subject Name-Subject code",
    "OnScreenID",
    "SubjectName",
    "Semester",
    "RollNo",
    "PRNNumber",
    "Name",
    "CampusName",
    "UserType",
    "MarkAttendance",
    "Email",
    "Semester Total Max Marks",
    "Internal Marks",
    "Total Max Marks",
    "Total Marks Obtained",
    "Grade",
    "Change %",
]

# Columns to drop from the original standard clean downloaded Excel files
COLUMNS_TO_DROP_STANDARD = [
    "Semester",
    "ExamAssingment_ID",
    "OnScreenID",
    "RollNo",
    "Name",
    "UserType",
    "CampusName",
    "MarkAttendance",
    "Email",
    "Grade",
    "Internal Marks",
    "Total Max Marks",
    "Semester Total Max Marks",
    "Total Marks Obtained",
    "Change %",
]

# Columns to drop from the formatted downloaded Excel files
COLUMNS_TO_DROP_FORMATTED = [
    "Semester",
    "ExamAssingment_ID",
    "OnScreenID",
    "RollNo",
    "Name",
    "CampusName",
    "MarkAttendance",
]

PREVIEW_HIDDEN_COLUMNS = [
    "OnScreenID",
    "RollNo",
    "Name",
    "CampusName",
    "Subject Name-Subject code",
    "SubjectName",
    "Semester",
]

USER_TYPE_ORDER = {
    "Examiner": 1,
    "Moderator": 2,
    "Reval 1": 3,
    "Reval 2": 4,
}

ROW_COLOURS = {
    "Moderator": "background-color:#032E15;",
    "Reval 1": "background-color:#162456;color:white;",
    "Reval 2": "background-color:#290245;color:white;",
}

INCREASE_STYLE = "color:#9ACD32;font-weight:bold;"  # mark went up / grade improved
DECREASE_STYLE = "color:#FF4500;font-weight:bold;"  # mark went down / grade degraded

GRADE_RANKS = {
    "O": 8,
    "A+": 7,
    "A": 6,
    "B+": 5,
    "B": 4,
    "C": 3,
    "P": 2,
    "F": 1
}


# ==================================================
# STEP 1: DECRYPT + LOAD UPLOADED FILES
# ==================================================
def decrypt_and_load(uploaded_file):
    """
    Attempts to load unprotected Excel files directly first.
    If protected, iterates through PASSWORDS list to decrypt.
    """
    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, dtype=object, keep_default_na=False)
        return df, "Unprotected (Direct)"
    except Exception:
        pass

    for password in PASSWORDS:
        decrypted_path = None
        try:
            uploaded_file.seek(0)

            office_file = msoffcrypto.OfficeFile(uploaded_file)
            office_file.load_key(password=password)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                office_file.decrypt(tmp)
                decrypted_path = tmp.name

            df = pd.read_excel(decrypted_path, dtype=object, keep_default_na=False)
            return df, password

        except Exception:
            continue

        finally:
            if decrypted_path and os.path.exists(decrypted_path):
                os.remove(decrypted_path)

    return None, None


def load_gmr_file(gmr_file):
    """Loads GMR file directly or decrypts if password protected."""
    if gmr_file is None:
        return None
    df, _ = decrypt_and_load(gmr_file)
    return df


# ==================================================
# STEP 2: CLEAN, SORT, & GMR LOOKUP MATCHING
# ==================================================
def clean_and_sort(df):
    """Clean up raw Questionwise export and sort rows."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    if "PRNNumber" in df.columns:
        df = df[df["PRNNumber"].astype(str).str.strip() != ""]

    df = df.dropna(how="all")

    if "SrNo" in df.columns:
        df = df[~df["SrNo"].astype(str).str.contains("Count:", case=False, na=False)]
    else:
        first_col = df.columns[0]
        df = df[~df[first_col].astype(str).str.contains("Count:", case=False, na=False)]

    if "PRNNumber" in df.columns:
        df["PRNNumber"] = pd.to_numeric(df["PRNNumber"], errors="coerce")
        df = df[df["PRNNumber"].notna()]

    if "PRNNumber" in df.columns and "UserType" in df.columns:
        df["SortOrder"] = df["UserType"].map(USER_TYPE_ORDER).fillna(99)
        df = df.sort_values(by=["PRNNumber", "SortOrder"]).drop(columns=["SortOrder"])

    df = df.reset_index(drop=True)
    return df


def calculate_grade(total_obtained, total_max):
    """Calculates Grade based on percentage of Total Obtained vs Total Max Marks."""
    try:
        obtained = float(total_obtained)
        max_m = float(total_max)

        if max_m <= 0:
            return ""

        pct = (obtained / max_m) * 100.0

        if pct >= 90:
            return "O"
        elif pct >= 80:
            return "A+"
        elif pct >= 70:
            return "A"
        elif pct >= 60:
            return "B+"
        elif pct >= 55:
            return "B"
        elif pct >= 50:
            return "C"
        elif pct >= 40:
            return "P"
        else:
            return "F"
    except (ValueError, TypeError):
        return ""


def apply_gmr_lookup(df, gmr_df):
    """
    Performs lookup matching between Questionwise and GMR data:
    - Internal Marks: from rows where AGR type contains "Internal Total"
    - Total Max Marks: MAX Marks from rows where AGR type contains "Composite "
    - Semester Total Max Marks: MAX Marks from rows where AGR type contains "Semester Total"
    - Total Marks Obtained: TotalObtainedScore + Internal Marks
    - Grade: Calculated based on percentage criteria
    """
    df = df.copy()
    
    df["Internal Marks"] = ""
    df["Total Max Marks"] = ""
    df["Semester Total Max Marks"] = ""
    df["Total Marks Obtained"] = ""
    df["Grade"] = ""

    if gmr_df is None or gmr_df.empty:
        return df

    gmr_clean = gmr_df.copy()
    gmr_clean.columns = gmr_clean.columns.astype(str).str.strip()

    required_gmr_cols = ["Student Number", "Modulecode", "AGR type", "Grades"]
    if not all(col in gmr_clean.columns for col in required_gmr_cols):
        return df

    # Prepare Lookup Keys
    df_key = (
        df["PRNNumber"].astype(str).str.strip() + 
        df["Subject Name-Subject code"].astype(str).str.strip()
    )

    gmr_clean["LookupKey"] = (
        gmr_clean["Student Number"].astype(str).str.strip() + 
        gmr_clean["Modulecode"].astype(str).str.strip()
    )

    # 1. Fetch Internal Marks
    internal_df = gmr_clean[
        gmr_clean["AGR type"].astype(str).str.contains("Internal Total", case=False, na=False)
    ]
    internal_obtained_map = dict(zip(internal_df["LookupKey"], internal_df["Grades"]))
    df["Internal Marks"] = df_key.map(internal_obtained_map).fillna("")

    # 2. Fetch Total Max Marks from Composite AGR type
    composite_max_map = {}
    if "MAX Marks" in gmr_clean.columns:
        composite_df = gmr_clean[
            gmr_clean["AGR type"].astype(str).str.contains("Composite ", case=False, na=False)
        ]
        composite_max_map = dict(zip(composite_df["LookupKey"], composite_df["MAX Marks"]))
    df["Total Max Marks"] = df_key.map(composite_max_map).fillna("")

    # 3. Fetch Semester Total Max Marks from Semester Total AGR type
    sem_total_max_map = {}
    if "MAX Marks" in gmr_clean.columns:
        sem_total_df = gmr_clean[
            gmr_clean["AGR type"].astype(str).str.contains("Semester Total", case=False, na=False)
        ]
        sem_total_max_map = dict(zip(sem_total_df["LookupKey"], sem_total_df["MAX Marks"]))
    df["Semester Total Max Marks"] = df_key.map(sem_total_max_map).fillna("")

    # 4. Calculate Total Marks Obtained
    def compute_total_obtained(row):
        tot_score = row.get("TotalObtainedScore", "")
        int_marks = row.get("Internal Marks", "")

        try:
            val_tot = float(str(tot_score).strip())
        except (ValueError, TypeError):
            val_tot = 0.0

        try:
            val_int = float(str(int_marks).strip())
        except (ValueError, TypeError):
            val_int = 0.0

        if str(tot_score).strip() == "" and str(int_marks).strip() == "":
            return ""

        return val_tot + val_int

    df["Total Marks Obtained"] = df.apply(compute_total_obtained, axis=1)

    # 5. Calculate Grade
    df["Grade"] = df.apply(
        lambda r: calculate_grade(r["Total Marks Obtained"], r["Total Max Marks"]), axis=1
    )

    return df


def calculate_paper_max_marks(df):
    """Extracts sum of question max marks from question column headers."""
    question_columns = [col for col in df.columns if col not in FIXED_COLUMNS]
    total_q_max = 0.0
    for col in question_columns:
        match = re.search(r"\(([\d.]+)\)", col)
        if match:
            try:
                total_q_max += float(match.group(1))
            except ValueError:
                pass
    return total_q_max


def calculate_moderator_change_pct(df):
    """
    Calculates Change % for Moderator rows using Semester Total Max Marks:
    Change % = (TotalObtainedScore (Moderator) - TotalObtainedScore (Examiner)) / Semester Total Max Marks * 100
    """
    df = df.copy()
    df["Change %"] = pd.Series([""] * len(df), index=df.index, dtype="object")

    if "PRNNumber" not in df.columns or "UserType" not in df.columns or "TotalObtainedScore" not in df.columns:
        return df

    paper_q_max = calculate_paper_max_marks(df)

    for prn, group in df.groupby("PRNNumber", dropna=False):
        examiner_rows = group[group["UserType"].astype(str).str.strip() == "Examiner"]
        moderator_rows = group[group["UserType"].astype(str).str.strip() == "Moderator"]

        if examiner_rows.empty or moderator_rows.empty:
            continue

        examiner_row = examiner_rows.iloc[0]
        moderator_row = moderator_rows.iloc[0]

        try:
            ex_score = float(examiner_row["TotalObtainedScore"])
            mod_score = float(moderator_row["TotalObtainedScore"])
            
            # 1. Try Semester Total Max Marks from GMR lookup
            sem_max_marks = 0.0
            sem_max_val = str(moderator_row.get("Semester Total Max Marks", "")).strip()
            if sem_max_val != "":
                try:
                    sem_max_marks = float(sem_max_val)
                except ValueError:
                    sem_max_marks = 0.0

            # 2. Fallback to Total Max Marks if GMR Semester Total is not available
            if sem_max_marks <= 0:
                tot_max_val = str(moderator_row.get("Total Max Marks", "")).strip()
                if tot_max_val != "":
                    try:
                        sem_max_marks = float(tot_max_val)
                    except ValueError:
                        sem_max_marks = 0.0

            # 3. Fallback to sum of question max marks from column headers
            if sem_max_marks <= 0:
                sem_max_marks = paper_q_max

            if sem_max_marks > 0:
                change_pct = ((mod_score - ex_score) / sem_max_marks) * 100.0
                df.loc[moderator_rows.index[0], "Change %"] = round(change_pct, 2)
        except (ValueError, TypeError):
            continue

    return df


def reorder_gmr_columns(df):
    """
    Reorders GMR and calculated columns right after TotalObtainedScore in the exact sequence:
    Semester Total Max Marks, Internal Marks, Total Max Marks, Total Marks Obtained, Grade, Change %, Email
    """
    df = df.copy()
    target_tail = [
        "Semester Total Max Marks",
        "Internal Marks",
        "Total Max Marks",
        "Total Marks Obtained",
        "Grade",
        "Change %",
        "Email",
    ]
    if "TotalObtainedScore" in df.columns:
        cols = df.columns.tolist()
        tot_idx = cols.index("TotalObtainedScore")
        head = cols[:tot_idx + 1]
        tail = [c for c in target_tail if c in cols]
        remaining = [c for c in cols if c not in head and c not in tail]
        df = df[head + tail + remaining]
    return df


# ==================================================
# STEP 2B: PROCESS DATA FOR OUTPUT EXCEL DOWNLOAD
# ==================================================
def process_data_for_download(df, include_all_usertypes=False):
    """Applies download-specific transformations to DataFrame."""
    download_df = df.copy()

    if not include_all_usertypes:
        # Remove Reval 1 and Reval 2 Rows
        if "UserType" in download_df.columns:
            reval_mask = download_df["UserType"].astype(str).str.strip().str.lower().isin(["reval 1", "reval 2"])
            download_df = download_df[~reval_mask]

        # Filter Moderator Rows
        if "PRNNumber" in download_df.columns and "UserType" in download_df.columns:
            moderator_prns = set(
                download_df[download_df["UserType"].astype(str).str.strip().str.lower() == "moderator"]["PRNNumber"]
            )

            def filter_rows(row):
                prn = row["PRNNumber"]
                u_type = str(row["UserType"]).strip().lower()
                if prn in moderator_prns:
                    return u_type == "moderator"
                return True

            download_df = download_df[download_df.apply(filter_rows, axis=1)]

    # Replace "AB" in TotalObtainedScore
    if "TotalObtainedScore" in download_df.columns:
        def adjust_score(row):
            score = str(row["TotalObtainedScore"]).strip()
            attendance = str(row.get("MarkAttendance", "")).strip().upper()
            
            if score.upper() == "AB":
                if attendance == "UFM":
                    return "UFM"
                return "ABSENT"
            return score

        download_df["TotalObtainedScore"] = download_df.apply(adjust_score, axis=1)

    # Drop Unwanted Columns
    cols_to_remove = COLUMNS_TO_DROP_FORMATTED if include_all_usertypes else COLUMNS_TO_DROP_STANDARD
    download_df = download_df.drop(columns=cols_to_remove, errors="ignore")

    # Rename Columns
    if "Subject Name-Subject code" in download_df.columns:
        download_df = download_df.rename(columns={"Subject Name-Subject code": "Subject code"})

    if "PRNNumber" in download_df.columns:
        download_df = download_df.rename(columns={"PRNNumber": "SAP ID"})

    download_df = download_df.reset_index(drop=True)

    # Make SrNo Continuous (1, 2, 3, ...)
    if "SrNo" in download_df.columns:
        download_df["SrNo"] = download_df.index + 1

    return download_df


# ==================================================
# STEP 3: DISPLAY FORMATTING HELPERS
# ==================================================
def format_value(value):
    """Formats values cleanly for preview display."""
    text = str(value).strip()

    if text == "" or text.upper() == "NA" or text.upper() == "NAN":
        return "" if text == "" else "NA"

    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
        return f"{number:g}"
    except (TypeError, ValueError):
        return text


def highlight_rows(data, question_columns):
    """Applies CSS styles for Streamlit preview table safely."""
    styles = pd.DataFrame("", index=data.index, columns=data.columns)

    for idx in data.index:
        if "UserType" in data.columns:
            user_type = str(data.loc[idx, "UserType"]).strip()
            if user_type in ROW_COLOURS:
                styles.loc[idx, :] += ROW_COLOURS[user_type]

    if "PRNNumber" in data.columns and "UserType" in data.columns:
        for _, group in data.groupby("PRNNumber", dropna=False):
            examiner_rows = group[group["UserType"] == "Examiner"]
            if examiner_rows.empty:
                continue
            examiner_row = examiner_rows.iloc[0]

            for idx in group.index:
                if str(data.loc[idx, "UserType"]) == "Examiner":
                    continue

                # Question-wise marks highlighting
                for col in question_columns:
                    if col not in data.columns:
                        continue

                    examiner_value = str(examiner_row[col]).strip()
                    current_value = str(data.loc[idx, col]).strip()

                    if examiner_value == "" or examiner_value.upper() == "NA":
                        continue

                    if current_value == "" or current_value.upper() == "NA":
                        styles.loc[idx, col] += DECREASE_STYLE
                        continue

                    try:
                        ex_mark = float(examiner_value)
                        cur_mark = float(current_value)
                    except (TypeError, ValueError):
                        continue

                    if cur_mark > ex_mark:
                        styles.loc[idx, col] += INCREASE_STYLE
                    elif cur_mark < ex_mark:
                        styles.loc[idx, col] += DECREASE_STYLE

                # Grade Conditional Formatting
                if "Grade" in data.columns:
                    ex_grade = str(examiner_row.get("Grade", "")).strip().upper()
                    cur_grade = str(data.loc[idx, "Grade"]).strip().upper()

                    if ex_grade in GRADE_RANKS and cur_grade in GRADE_RANKS:
                        ex_rank = GRADE_RANKS[ex_grade]
                        cur_rank = GRADE_RANKS[cur_grade]

                        if cur_rank > ex_rank:
                            styles.loc[idx, "Grade"] += INCREASE_STYLE
                        elif cur_rank < ex_rank:
                            styles.loc[idx, "Grade"] += DECREASE_STYLE

                # Total Marks Obtained Conditional Formatting
                if "Total Marks Obtained" in data.columns:
                    ex_tot_str = str(examiner_row.get("Total Marks Obtained", "")).strip()
                    cur_tot_str = str(data.loc[idx, "Total Marks Obtained"]).strip()
                    if ex_tot_str != "" and cur_tot_str != "":
                        try:
                            ex_tot = float(ex_tot_str)
                            cur_tot = float(cur_tot_str)
                            if cur_tot > ex_tot:
                                styles.loc[idx, "Total Marks Obtained"] += INCREASE_STYLE
                            elif cur_tot < ex_tot:
                                styles.loc[idx, "Total Marks Obtained"] += DECREASE_STYLE
                        except (ValueError, TypeError):
                            pass

                # Change % Conditional Formatting
                if "Change %" in data.columns:
                    chg_str = str(data.loc[idx, "Change %"]).strip()
                    if chg_str != "":
                        try:
                            chg_val = float(chg_str)
                            if chg_val > 0:
                                styles.loc[idx, "Change %"] += INCREASE_STYLE
                            elif chg_val < 0:
                                styles.loc[idx, "Change %"] += DECREASE_STYLE
                        except (ValueError, TypeError):
                            pass

    if "TotalObtainedScore" in data.columns:
        styles.loc[:, "TotalObtainedScore"] += "font-weight:bold;"

    return styles


def style_preview(df, question_columns, hide_columns=PREVIEW_HIDDEN_COLUMNS):
    """Builds styled interactive preview table for web display."""
    display_df = df.drop(columns=hide_columns, errors="ignore").copy()
    display_df.index = range(1, len(display_df) + 1)
    return (
        display_df.style
        .format(format_value)
        .apply(lambda data: highlight_rows(data, question_columns), axis=None)
    )


# ==================================================
# STEP 4: SUMMARY STATISTICS & METADATA
# ==================================================
def count_reviewers(df, user_type):
    """Number of distinct students that have a row of the given UserType."""
    if "UserType" not in df.columns or "PRNNumber" not in df.columns:
        return 0
    return (
        df[df["UserType"].astype(str).str.strip().eq(user_type)]["PRNNumber"]
        .nunique()
    )


def score_change_summary(df, review_type):
    """Compare review_type TotalObtainedScore against Examiner's TotalObtainedScore."""
    increased = decreased = unchanged = 0

    if "TotalObtainedScore" not in df.columns or "UserType" not in df.columns:
        return increased, decreased, unchanged

    for _, group in df.groupby("PRNNumber", dropna=False):
        examiner_rows = group[group["UserType"] == "Examiner"]
        review_rows = group[group["UserType"] == review_type]

        if examiner_rows.empty or review_rows.empty:
            continue

        try:
            examiner_score = float(examiner_rows.iloc[0]["TotalObtainedScore"])
            review_score = float(review_rows.iloc[0]["TotalObtainedScore"])
        except (TypeError, ValueError):
            continue

        if review_score > examiner_score:
            increased += 1
        elif review_score < examiner_score:
            decreased += 1
        else:
            unchanged += 1

    return increased, decreased, unchanged


def compute_moderation_10pct_metrics(df):
    """
    Computes:
    - Moderator Count
    - Count of Moderator records with |Change %| > 10%
    - Change % (>10%) ratio & percentage
    - Eligibility for 100% Moderation (>50% threshold)
    """
    if "Change %" not in df.columns or "UserType" not in df.columns:
        return 0, 0, 0.0, False

    mod_rows = df[df["UserType"].astype(str).str.strip() == "Moderator"]
    mod_count = len(mod_rows)

    if mod_count == 0:
        return 0, 0, 0.0, False

    count_gt_10pct = 0
    for val in mod_rows["Change %"]:
        try:
            if abs(float(val)) > 10.0:
                count_gt_10pct += 1
        except (ValueError, TypeError):
            continue

    ratio_gt_10pct = (count_gt_10pct / mod_count) if mod_count > 0 else 0.0
    is_eligible_100pct_moderation = ratio_gt_10pct > 0.50

    return mod_count, count_gt_10pct, ratio_gt_10pct, is_eligible_100pct_moderation


def build_paper_info(df):
    """Builds header info line for display."""
    if df.empty:
        return ""

    semester = str(df.iloc[0].get("Semester", "")).strip()
    subject_code = str(df.iloc[0].get("Subject Name-Subject code", "")).strip()
    subject_name = str(df.iloc[0].get("SubjectName", "")).strip()

    return f"{semester}  -  {subject_code}  -  {subject_name}"


def generate_output_filename(df, idx=1, prefix=""):
    """Generates dynamic Excel filename."""
    if df.empty:
        return f"{prefix}Questionwise_Output_{idx}_{datetime.now().strftime('%d-%m-%y')}.xlsx"

    subject_name = str(df.iloc[0].get("SubjectName", "Subject")).strip()
    semester_raw = str(df.iloc[0].get("Semester", "Semester")).strip()

    semester_short = re.sub(r"(?i)\bsemester\b", "Sem", semester_raw).strip()

    unique_prns = df["PRNNumber"].nunique() if "PRNNumber" in df.columns else len(df)
    date_str = datetime.now().strftime("%d-%m-%y")

    filename = f"{prefix}{subject_name}_{semester_short}_{unique_prns}_{date_str}.xlsx"
    return filename


# ==================================================
# STEP 5: EXCEL GENERATION & STYLING
# ==================================================
def sanitize_sheet_name(subject_identifier):
    """Sanitizes SubjectName / Subject Code into a valid Excel sheet name (max 31 chars)."""
    if not subject_identifier:
        return "Questionwise Data"

    name = str(subject_identifier)
    name = re.sub(r"\[[^\]]*\]", "", name)
    name = re.sub(r"[\\/?*\[\]:]", "", name)
    name = re.sub(r"\s+", " ", name).strip()

    if not name:
        name = "Questionwise Data"

    return name[:31]


def get_sheet_name_for_df(df, default_idx=1):
    """Retrieves sheet name based on 'Subject Name-Subject code' or 'SubjectName'."""
    if not df.empty:
        if "SubjectName" in df.columns and str(df.iloc[0]["SubjectName"]).strip():
            return sanitize_sheet_name(df.iloc[0]["SubjectName"])
        if "Subject Name-Subject code" in df.columns and str(df.iloc[0]["Subject Name-Subject code"]).strip():
            return sanitize_sheet_name(df.iloc[0]["Subject Name-Subject code"])
    return f"Sheet_{default_idx}"


def add_sheet_popup_validation(worksheet, is_eligible):
    """
    Applies Data Validation Input Message to cell A1.
    When a user clicks on/selects cell A1 in Excel, a dynamic floating notification pop-up appears.
    """
    if not is_eligible:
        return

    dv = DataValidation(
        type="custom",
        formula1="TRUE",
        allow_blank=True,
        showInputMessage=True,
        showErrorMessage=False,
        promptTitle="⚠️ 100% Moderation Alert",
        prompt="The 10% change in marks is observed in more than 50% of the moderated answer books. The paper is eligible for 100% Moderation."
    )
    worksheet.add_data_validation(dv)
    dv.add(worksheet["A1"])


def to_excel_bytes(df):
    """ORIGINAL Standard clean output without GMR columns."""
    download_df = process_data_for_download(df, include_all_usertypes=False)
    buffer = io.BytesIO()

    sheet_name = "Questionwise Data"
    if not df.empty and "SubjectName" in df.columns:
        subj = str(df.iloc[0]["SubjectName"]).strip()
        if subj:
            sheet_name = sanitize_sheet_name(subj)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        download_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        worksheet.views.sheetView[0].showGridLines = False
        worksheet.freeze_panes = "A2"
        worksheet.row_dimensions[1].height = 30

        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000")
        )

        center_wrap_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        center_align = Alignment(horizontal="center", vertical="center")
        grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        cols = download_df.columns.tolist()
        center_col_indices = []

        for col_idx, col_name in enumerate(cols, start=1):
            if col_name in ["SrNo", "Subject code"]:
                center_col_indices.append(col_idx)

        if "SAP ID" in cols and "TotalObtainedScore" in cols:
            sap_idx = cols.index("SAP ID") + 1
            tot_idx = cols.index("TotalObtainedScore") + 1
            for idx in range(sap_idx, tot_idx + 1):
                center_col_indices.append(idx)

        center_col_indices = set(center_col_indices)
        total_score_col_idx = cols.index("TotalObtainedScore") + 1 if "TotalObtainedScore" in cols else None

        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border = thin_border
                
                if cell.row == 1:
                    cell.font = Font(bold=True)
                    cell.fill = grey_fill
                    cell.alignment = center_wrap_align
                elif total_score_col_idx and cell.column == total_score_col_idx:
                    cell.font = Font(bold=True)

                if cell.row > 1 and cell.column in center_col_indices:
                    cell.alignment = center_align

        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)

    return buffer.getvalue()


def to_excel_bytes_formatted(df):
    """
    Formatted output Excel containing ALL UserTypes, GMR columns, Change %,
    center alignment, fixed width 15 for GMR/Change %/Email columns, row height 30 for header with wrap text,
    darker green text (#006100), light green fill (#E2EFDA) for Moderator rows, thick bottom border per student, 
    F2 freeze pane, conditional formatting, and interactive Data Validation pop-up if eligible.
    """
    download_df = process_data_for_download(df, include_all_usertypes=True)
    buffer = io.BytesIO()

    sheet_name = get_sheet_name_for_df(df)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        download_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        worksheet.views.sheetView[0].showGridLines = False
        worksheet.freeze_panes = "F2"
        worksheet.row_dimensions[1].height = 30

        thin_side = Side(style="thin", color="000000")
        medium_side = Side(style="medium", color="000000")
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        center_wrap_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        center_align = Alignment(horizontal="center", vertical="center")
        
        grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        mod_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        # Darker green font for text legibility
        dark_green_font = Font(color="006100", bold=True)
        red_font = Font(color="FF4500", bold=True)

        cols = download_df.columns.tolist()
        center_col_indices = []

        fixed_15_cols = [
            "Semester Total Max Marks", "Internal Marks", 
            "Total Max Marks", "Total Marks Obtained", "Grade", "Change %"
        ]

        align_center_col_names = [
            "SrNo", "Subject code", "Grade", "Internal Marks", 
            "Total Max Marks", "Semester Total Max Marks", "Total Marks Obtained", "Change %"
        ]

        for col_idx, col_name in enumerate(cols, start=1):
            if col_name in align_center_col_names:
                center_col_indices.append(col_idx)

        if "SAP ID" in cols and "TotalObtainedScore" in cols:
            sap_idx = cols.index("SAP ID") + 1
            tot_idx = cols.index("TotalObtainedScore") + 1
            for idx in range(sap_idx, tot_idx + 1):
                center_col_indices.append(idx)

        center_col_indices = set(center_col_indices)
        total_score_col_idx = cols.index("TotalObtainedScore") + 1 if "TotalObtainedScore" in cols else None
        user_type_col_idx = cols.index("UserType") if "UserType" in cols else None

        # Basic styling & Moderator Light Green Fill
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            row_idx = row[0].row - 2  # Convert to 0-based dataframe index
            is_moderator_row = False
            if row_idx >= 0 and user_type_col_idx is not None and row_idx < len(download_df):
                u_type = str(download_df.iloc[row_idx].get("UserType", "")).strip().lower()
                is_moderator_row = (u_type == "moderator")

            for cell in row:
                cell.border = thin_border
                
                if cell.row == 1:
                    cell.font = Font(bold=True)
                    cell.fill = grey_fill
                    cell.alignment = center_wrap_align
                else:
                    if is_moderator_row:
                        cell.fill = mod_fill
                    if total_score_col_idx and cell.column == total_score_col_idx:
                        cell.font = Font(bold=True)

                if cell.row > 1 and cell.column in center_col_indices:
                    cell.alignment = center_align

        # Conditional Formatting on Exported Rows
        if "UserType" in download_df.columns and "SAP ID" in download_df.columns:
            question_cols = [c for c in download_df.columns if c not in FIXED_COLUMNS and c not in COLUMNS_TO_DROP_FORMATTED]
            
            for sap_id, group in download_df.groupby("SAP ID", dropna=False):
                examiner_rows = group[group["UserType"] == "Examiner"]
                if examiner_rows.empty:
                    continue
                examiner_row = examiner_rows.iloc[0]

                for row_idx in group.index:
                    if str(download_df.loc[row_idx, "UserType"]).strip() == "Examiner":
                        continue

                    excel_row_num = row_idx + 2

                    # Marks Formatting
                    for q_col in question_cols:
                        if q_col not in download_df.columns:
                            continue
                        col_idx = cols.index(q_col) + 1
                        ex_val_str = str(examiner_row[q_col]).strip()
                        cur_val_str = str(download_df.loc[row_idx, q_col]).strip()

                        if ex_val_str == "" or ex_val_str.upper() == "NA":
                            continue

                        cell = worksheet.cell(row=excel_row_num, column=col_idx)

                        if cur_val_str == "" or cur_val_str.upper() == "NA":
                            cell.font = red_font
                            continue

                        try:
                            ex_val = float(ex_val_str)
                            cur_val = float(cur_val_str)
                            if cur_val > ex_val:
                                cell.font = dark_green_font
                            elif cur_val < ex_val:
                                cell.font = red_font
                        except (ValueError, TypeError):
                            pass

                    # Grade Formatting
                    if "Grade" in download_df.columns:
                        grade_col_idx = cols.index("Grade") + 1
                        ex_grade = str(examiner_row.get("Grade", "")).strip().upper()
                        cur_grade = str(download_df.loc[row_idx, "Grade"]).strip().upper()

                        if ex_grade in GRADE_RANKS and cur_grade in GRADE_RANKS:
                            cell = worksheet.cell(row=excel_row_num, column=grade_col_idx)
                            if GRADE_RANKS[cur_grade] > GRADE_RANKS[ex_grade]:
                                cell.font = dark_green_font
                            elif GRADE_RANKS[cur_grade] < GRADE_RANKS[ex_grade]:
                                cell.font = red_font

                    # Total Marks Obtained Formatting
                    if "Total Marks Obtained" in download_df.columns:
                        tot_obt_col_idx = cols.index("Total Marks Obtained") + 1
                        ex_tot_str = str(examiner_row.get("Total Marks Obtained", "")).strip()
                        cur_tot_str = str(download_df.loc[row_idx, "Total Marks Obtained"]).strip()

                        if ex_tot_str != "" and cur_tot_str != "":
                            cell = worksheet.cell(row=excel_row_num, column=tot_obt_col_idx)
                            try:
                                ex_tot = float(ex_tot_str)
                                cur_tot = float(cur_tot_str)
                                if cur_tot > ex_tot:
                                    cell.font = dark_green_font
                                elif cur_tot < ex_tot:
                                    cell.font = red_font
                            except (ValueError, TypeError):
                                pass

                    # Change % Formatting
                    if "Change %" in download_df.columns:
                        chg_col_idx = cols.index("Change %") + 1
                        chg_str = str(download_df.loc[row_idx, "Change %"]).strip()

                        if chg_str != "":
                            cell = worksheet.cell(row=excel_row_num, column=chg_col_idx)
                            try:
                                chg_val = float(chg_str)
                                if chg_val > 0:
                                    cell.font = dark_green_font
                                elif chg_val < 0:
                                    cell.font = red_font
                            except (ValueError, TypeError):
                                pass

                # Thick bottom border per student group
                last_row_idx = group.index[-1] + 2
                for c_idx in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=last_row_idx, column=c_idx)
                    cell.border = Border(left=thin_side, right=thin_side, top=cell.border.top, bottom=medium_side)

        # Apply interactive pop-up Data Validation message if paper is eligible for 100% Moderation
        _, _, _, is_eligible = compute_moderation_10pct_metrics(df)
        add_sheet_popup_validation(worksheet, is_eligible)

        # Auto-size columns / Apply fixed width 15
        for col in worksheet.columns:
            col_name = str(col[0].value).strip()
            col_letter = get_column_letter(col[0].column)
            if col_name in fixed_15_cols:
                worksheet.column_dimensions[col_letter].width = 15
            else:
                max_len = 0
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)

    return buffer.getvalue()


def create_consolidated_multi_sheet_excel(dfs_list):
    """
    Clubs all processed dataframes into a single Excel workbook 
    with separate sheets named by Subject Name-Subject code containing ALL UserTypes,
    Change %, center alignment, F2 freeze pane, thick student borders, and interactive pop-up message if eligible.
    """
    buffer = io.BytesIO()
    used_sheet_names = set()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for idx, (df, orig_name) in enumerate(dfs_list, start=1):
            download_df = process_data_for_download(df, include_all_usertypes=True)
            
            base_sheet_name = get_sheet_name_for_df(df, idx)
            sheet_name = base_sheet_name
            
            counter = 1
            while sheet_name in used_sheet_names:
                sheet_name = f"{base_sheet_name[:28]}_{counter}"
                counter += 1
            used_sheet_names.add(sheet_name)

            download_df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]

            worksheet.views.sheetView[0].showGridLines = False
            worksheet.freeze_panes = "F2"
            worksheet.row_dimensions[1].height = 30

            thin_side = Side(style="thin", color="000000")
            medium_side = Side(style="medium", color="000000")
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

            center_wrap_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            center_align = Alignment(horizontal="center", vertical="center")
            
            grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            mod_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

            dark_green_font = Font(color="006100", bold=True)
            red_font = Font(color="FF4500", bold=True)

            cols = download_df.columns.tolist()
            center_col_indices = []

            fixed_15_cols = [
                "Semester Total Max Marks", "Internal Marks", 
                "Total Max Marks", "Total Marks Obtained", "Grade", "Change %"
            ]

            align_center_col_names = [
                "SrNo", "Subject code", "Grade", "Internal Marks", 
                "Total Max Marks", "Semester Total Max Marks", "Total Marks Obtained", "Change %"
            ]

            for col_idx, col_name in enumerate(cols, start=1):
                if col_name in align_center_col_names:
                    center_col_indices.append(col_idx)

            if "SAP ID" in cols and "TotalObtainedScore" in cols:
                sap_idx = cols.index("SAP ID") + 1
                tot_idx = cols.index("TotalObtainedScore") + 1
                for i in range(sap_idx, tot_idx + 1):
                    center_col_indices.append(i)

            center_col_indices = set(center_col_indices)
            total_score_col_idx = cols.index("TotalObtainedScore") + 1 if "TotalObtainedScore" in cols else None
            user_type_col_idx = cols.index("UserType") if "UserType" in cols else None

            # Basic styling & Moderator Light Green Fill
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                row_idx = row[0].row - 2
                is_moderator_row = False
                if row_idx >= 0 and user_type_col_idx is not None and row_idx < len(download_df):
                    u_type = str(download_df.iloc[row_idx].get("UserType", "")).strip().lower()
                    is_moderator_row = (u_type == "moderator")

                for cell in row:
                    cell.border = thin_border
                    
                    if cell.row == 1:
                        cell.font = Font(bold=True)
                        cell.fill = grey_fill
                        cell.alignment = center_wrap_align
                    else:
                        if is_moderator_row:
                            cell.fill = mod_fill
                        if total_score_col_idx and cell.column == total_score_col_idx:
                            cell.font = Font(bold=True)

                    if cell.row > 1 and cell.column in center_col_indices:
                        cell.alignment = center_align

            # Conditional Formatting & Student Group Borders
            if "UserType" in download_df.columns and "SAP ID" in download_df.columns:
                question_cols = [c for c in download_df.columns if c not in FIXED_COLUMNS and c not in COLUMNS_TO_DROP_FORMATTED]
                
                for sap_id, group in download_df.groupby("SAP ID", dropna=False):
                    examiner_rows = group[group["UserType"] == "Examiner"]
                    if examiner_rows.empty:
                        continue
                    examiner_row = examiner_rows.iloc[0]

                    for row_idx in group.index:
                        if str(download_df.loc[row_idx, "UserType"]).strip() == "Examiner":
                            continue

                        excel_row_num = row_idx + 2

                        for q_col in question_cols:
                            if q_col not in download_df.columns:
                                continue
                            col_idx = cols.index(q_col) + 1
                            ex_val_str = str(examiner_row[q_col]).strip()
                            cur_val_str = str(download_df.loc[row_idx, q_col]).strip()

                            if ex_val_str == "" or ex_val_str.upper() == "NA":
                                continue

                            cell = worksheet.cell(row=excel_row_num, column=col_idx)

                            if cur_val_str == "" or cur_val_str.upper() == "NA":
                                cell.font = red_font
                                continue

                            try:
                                ex_val = float(ex_val_str)
                                cur_val = float(cur_val_str)
                                if cur_val > ex_val:
                                    cell.font = dark_green_font
                                elif cur_val < ex_val:
                                    cell.font = red_font
                            except (ValueError, TypeError):
                                pass

                        if "Grade" in download_df.columns:
                            grade_col_idx = cols.index("Grade") + 1
                            ex_grade = str(examiner_row.get("Grade", "")).strip().upper()
                            cur_grade = str(download_df.loc[row_idx, "Grade"]).strip().upper()

                            if ex_grade in GRADE_RANKS and cur_grade in GRADE_RANKS:
                                cell = worksheet.cell(row=excel_row_num, column=grade_col_idx)
                                if GRADE_RANKS[cur_grade] > GRADE_RANKS[ex_grade]:
                                    cell.font = dark_green_font
                                elif GRADE_RANKS[cur_grade] < GRADE_RANKS[ex_grade]:
                                    cell.font = red_font

                        if "Total Marks Obtained" in download_df.columns:
                            tot_obt_col_idx = cols.index("Total Marks Obtained") + 1
                            ex_tot_str = str(examiner_row.get("Total Marks Obtained", "")).strip()
                            cur_tot_str = str(download_df.loc[row_idx, "Total Marks Obtained"]).strip()

                            if ex_tot_str != "" and cur_tot_str != "":
                                cell = worksheet.cell(row=excel_row_num, column=tot_obt_col_idx)
                                try:
                                    ex_tot = float(ex_tot_str)
                                    cur_tot = float(cur_tot_str)
                                    if cur_tot > ex_tot:
                                        cell.font = dark_green_font
                                    elif cur_tot < ex_tot:
                                        cell.font = red_font
                                except (ValueError, TypeError):
                                    pass

                        if "Change %" in download_df.columns:
                            chg_col_idx = cols.index("Change %") + 1
                            chg_str = str(download_df.loc[row_idx, "Change %"]).strip()

                            if chg_str != "":
                                cell = worksheet.cell(row=excel_row_num, column=chg_col_idx)
                                try:
                                    chg_val = float(chg_str)
                                    if chg_val > 0:
                                        cell.font = dark_green_font
                                    elif chg_val < 0:
                                        cell.font = red_font
                                except (ValueError, TypeError):
                                    pass

                    # Thick bottom border per student group
                    last_row_idx = group.index[-1] + 2
                    for c_idx in range(1, worksheet.max_column + 1):
                        cell = worksheet.cell(row=last_row_idx, column=c_idx)
                        cell.border = Border(left=thin_side, right=thin_side, top=cell.border.top, bottom=medium_side)

            # Apply interactive pop-up Data Validation message if paper is eligible for 100% Moderation
            _, _, _, is_eligible = compute_moderation_10pct_metrics(df)
            add_sheet_popup_validation(worksheet, is_eligible)

            for col in worksheet.columns:
                col_name = str(col[0].value).strip()
                col_letter = get_column_letter(col[0].column)
                if col_name in fixed_15_cols:
                    worksheet.column_dimensions[col_letter].width = 15
                else:
                    max_len = 0
                    for cell in col:
                        if cell.value is not None:
                            max_len = max(max_len, len(str(cell.value)))
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)

    return buffer.getvalue()


def create_zip_file(excel_files):
    """Creates a ZIP archive containing multiple Excel file bytes."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for filename, file_bytes in excel_files:
            zip_file.writestr(filename, file_bytes)
    return zip_buffer.getvalue()


# ==================================================
# MAIN ENTRY POINT
# ==================================================
def show():
    """Render the Questionwise Checker page in Streamlit."""
    st.title("📝 Questionwise Checker")
    st.caption("Upload password-protected or unprotected Excel files.")

    # Side-by-side Upload Controls
    col_up1, col_up2 = st.columns(2)

    with col_up1:
        uploaded_files = st.file_uploader(
            "Upload Questionwise Excel/Excels",
            type=["xlsx"],
            accept_multiple_files=True
        )

    with col_up2:
        gmr_file = st.file_uploader(
            "Upload GMR Excel",
            type=["xlsx", "xls"]
        )

    if not uploaded_files:
        return

    # Load GMR File if present
    gmr_df = None
    if gmr_file is not None:
        with st.spinner("Loading GMR Excel..."):
            gmr_df = load_gmr_file(gmr_file)
            if gmr_df is not None:
                st.caption("✅ GMR File uploaded and loaded successfully.")
            else:
                st.warning("⚠️ Could not load GMR Excel file.")

    processed_outputs_clean = []
    dfs_for_multi_sheet = []

    # Process each uploaded Questionwise file independently
    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        st.markdown("---")
        st.subheader(f"📁 File {idx}: {uploaded_file.name}")

        try:
            with st.spinner(f"Decrypting and processing {uploaded_file.name}..."):
                df, used_password = decrypt_and_load(uploaded_file)

            if df is None:
                st.error(f"Unable to open {uploaded_file.name}. Password not recognised.")
                continue

            st.caption(f"Opened successfully using: {used_password}")

            # Clean and Sort
            df = clean_and_sort(df)

            if df.empty:
                st.warning(f"No valid student rows found in {uploaded_file.name}.")
                continue

            if "PRNNumber" not in df.columns or "UserType" not in df.columns:
                st.error(f"Missing required 'PRNNumber' or 'UserType' columns in {uploaded_file.name}.")
                continue

            # Apply GMR Lookup Match & Change % Calculation
            df = apply_gmr_lookup(df, gmr_df)
            df = calculate_moderator_change_pct(df)
            df = reorder_gmr_columns(df)
            
            dfs_for_multi_sheet.append((df, uploaded_file.name))

            question_columns = [col for col in df.columns if col not in FIXED_COLUMNS]

            # Summary statistics
            total_students = df["PRNNumber"].nunique()
            moderation_count = count_reviewers(df, "Moderator")
            reval1_count = count_reviewers(df, "Reval 1")
            reval2_count = count_reviewers(df, "Reval 2")

            moderator_positive, moderator_negative, moderator_no_change = score_change_summary(df, "Moderator")
            reval1_positive, reval1_negative, reval1_no_change = score_change_summary(df, "Reval 1")
            reval2_positive, reval2_negative, reval2_no_change = score_change_summary(df, "Reval 2")

            mod_cnt, cnt_gt_10, ratio_gt_10, is_eligible = compute_moderation_10pct_metrics(df)

            # 5 Summary Cards Grid
            col1, col2, col3, col4, col5 = st.columns(5)

            with col1:
                st.info(
                    f"Total Students: {total_students}\n\n"
                    f"Moderation Count: {moderation_count}\n\n"
                    f"Reval 1 Count: {reval1_count}\n\n"
                    f"Reval 2 Count: {reval2_count}"
                )

            with col2:
                st.success(
                    f"Positive Change in Marks:\n\n"
                    f"Moderation: {moderator_positive}\n\n"
                    f"Reval 1: {reval1_positive}\n\n"
                    f"Reval 2: {reval2_positive}"
                )

            with col3:
                st.error(
                    f"Negative Change in Marks:\n\n"
                    f"Moderation: {moderator_negative}\n\n"
                    f"Reval 1: {reval1_negative}\n\n"
                    f"Reval 2: {reval2_negative}"
                )

            with col4:
                st.warning(
                    f"No Change in Marks:\n\n"
                    f"Moderation: {moderator_no_change}\n\n"
                    f"Reval 1: {reval1_no_change}\n\n"
                    f"Reval 2: {reval2_no_change}"
                )

            with col5:
                # 5th Card: Change % Metrics
                pct_str = f"{ratio_gt_10 * 100:.1f}%"
                color_style = "color:#FF4500;" if is_eligible else "color:#9ACD32;"
                
                st.markdown(
                    f"""
                    <div style="background-color:#1E293B;padding:12px;border-radius:8px;color:white;border:1px solid #334155;">
                        <span style="font-size:15px;font-weight:600;">Eligible for 100% Moderation?</span><br/>
                        <span style="font-size:15px;">Moderators |Change %| > 10%: <b>{cnt_gt_10} / {mod_cnt}</b></span><br/>
                        <span style="font-size:15px;">Change % (>10%):</span><br/>
                        <span style="font-size:45px;font-weight:bold;{color_style}">{pct_str}</span>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            if is_eligible:
                st.error("Above 10% change in marks is observed in more than 50% of the moderated answer books. This module is eligible for 100% Moderation.")

            st.subheader(build_paper_info(df))

            # Filter preview subsets
            no_attendance_hidden_columns = PREVIEW_HIDDEN_COLUMNS + ["MarkAttendance"]

            review_prns = df[
                df["UserType"].astype(str).str.strip().isin(["Moderator", "Reval 1", "Reval 2"])
            ]["PRNNumber"].unique()
            review_df = df[df["PRNNumber"].isin(review_prns)].copy()

            changed_prns = []
            if "TotalObtainedScore" in df.columns and "UserType" in df.columns:
                for prn, group in df.groupby("PRNNumber", dropna=False, sort=False):
                    examiner_rows = group[group["UserType"] == "Examiner"]
                    reviewer_rows = group[group["UserType"].isin(["Moderator", "Reval 1", "Reval 2"])]
                    if not examiner_rows.empty and not reviewer_rows.empty:
                        try:
                            ex_score = float(examiner_rows.iloc[0]["TotalObtainedScore"])
                            for _, rev_row in reviewer_rows.iterrows():
                                rev_score = float(rev_row["TotalObtainedScore"])
                                if rev_score != ex_score:
                                    changed_prns.append(prn)
                                    break
                        except (TypeError, ValueError):
                            continue

            changes_df = df[df["PRNNumber"].isin(changed_prns)].copy() if changed_prns else pd.DataFrame()

            # Display Tabs
            tab1, tab2, tab3 = st.tabs([
                "📋 All Data Preview", 
                "🔍 Moderation / Revaluation Cases Only", 
                "⚡ Changes Only"
            ])

            with tab1:
                st.subheader("All Data Preview")
                st.write(style_preview(df, question_columns))

            with tab2:
                st.subheader("Moderation / Revaluation Cases Only")
                if not review_df.empty:
                    st.write(style_preview(review_df, question_columns, hide_columns=no_attendance_hidden_columns))
                else:
                    st.info("No moderation or revaluation cases found in this file.")

            with tab3:
                st.subheader("Changes Only")
                if not changes_df.empty:
                    st.write(style_preview(changes_df, question_columns, hide_columns=no_attendance_hidden_columns))
                else:
                    st.info("No score changes detected between examiner and reviewers.")

            # Generate individual Excel Outputs
            out_bytes_clean = to_excel_bytes(df)
            out_bytes_fmt = to_excel_bytes_formatted(df)
            out_filename = generate_output_filename(df, idx)
            out_filename_fmt = generate_output_filename(df, idx, prefix="Formatted_")
            
            # Save original clean outputs for original bulk download ZIP
            processed_outputs_clean.append((out_filename, out_bytes_clean))

            # Side-by-Side Individual File Download Buttons
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label=f"📥 Download Cleaned File: {out_filename}",
                    data=out_bytes_clean,
                    file_name=out_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_clean_{idx}"
                )
            with dl_col2:
                st.download_button(
                    label=f"🎨 Download Formatted File: {out_filename_fmt}",
                    data=out_bytes_fmt,
                    file_name=out_filename_fmt,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_fmt_{idx}"
                )

        except Exception as error:
            st.error(f"Error processing file {uploaded_file.name}: {error}")

    # Bulk Download Options
    if processed_outputs_clean:
        st.markdown("---")
        st.header("📦 Bulk Download Options")
        
        bulk_col1, bulk_col2 = st.columns(2)

        # 1. Original ZIP File Download (Clean Files)
        with bulk_col1:
            zip_filename = f"Questionwise_Processed_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.zip"
            zip_bytes = create_zip_file(processed_outputs_clean)

            st.download_button(
                label="📦 Download All Output Excel Files (ZIP)",
                data=zip_bytes,
                file_name=zip_filename,
                mime="application/zip",
                key="dl_zip_all"
            )

        # 2. Consolidated Multi-Sheet Excel Download
        with bulk_col2:
            if len(dfs_for_multi_sheet) > 1:
                multi_sheet_bytes = create_consolidated_multi_sheet_excel(dfs_for_multi_sheet)
                multi_sheet_filename = f"Consolidated_Questionwise_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.xlsx"

                st.download_button(
                    label="📊 Download Consolidated Multi-Sheet Excel",
                    data=multi_sheet_bytes,
                    file_name=multi_sheet_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_multi_sheet_all"
                )