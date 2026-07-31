import io
import re
from datetime import datetime
from functools import lru_cache

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st


# ==========================================
# OPTIONAL DEPENDENCY: deep_translator
# ==========================================
# The plugin/module loader that scans for show()/main() has to *import*
# this file first. A top-level `from deep_translator import GoogleTranslator`
# means: if that package isn't installed in the hosting environment, the
# import blows up before show() is ever discovered -- which is exactly the
# "No module named 'deep_translator'" error you hit. To fix that properly,
# add this line to your requirements.txt:
#
#     deep_translator
#
# But even without that fix, the app below now degrades gracefully: the
# translator is created lazily (only the first time it's actually needed),
# any failure is caught, and NAME_MARAT is simply left blank instead of the
# whole module refusing to load.
_translator = None
_translator_unavailable = False


def _get_translator():
    """Lazily create and cache the GoogleTranslator instance.

    Returns None (and remembers the failure) if deep_translator is not
    installed or fails to initialize, instead of raising.
    """
    global _translator, _translator_unavailable
    if _translator is not None:
        return _translator
    if _translator_unavailable:
        return None
    try:
        from deep_translator import GoogleTranslator
        _translator = GoogleTranslator(source="en", target="mr")
    except Exception:
        _translator_unavailable = True
        _translator = None
    return _translator


# ==========================================
# 1. REQUIRED OUTPUT COLUMNS SPECIFICATION
# ==========================================
TARGET_COLUMNS = [
    "LOTNO", "CONVID", "FACULTY", "PRNERN", "PROGTYPE", "APPL_NO", "SEAT_NO",
    "COLL_NO", "COLL_NAME", "COLL_NAMEM", "STUDLASTNAME", "STUDFIRSTNAME",
    "STUDMIDDDLENAME", "STUDMOTHERNAME", "NAME", "NAME_MARAT", "SEX", "ABBR",
    "CLASS", "MCLASS", "SUB1", "SUB1_NAME", "SUB1_NAMEM", "SUB2", "SUB2_NAME",
    "SUB2_NAMEM", "DEGNM", "MDEGNM", "SUBDEGNM", "MSUBDEGNM", "MONTH", "MMONTH"
]

MKCL_SOURCE_MAPPING = {
    "Student Number": "SEAT_NO",
    "Student Name": "NAME",
    "Gender": "SEX",
    "Final Grade": "CLASS",
    "Highest month of passing": "MONTH"
}


# ==========================================
# 2. ROBUST DATA CLEANING & MATCHING HELPERS
# ==========================================
def sanitize_key(val) -> str:
    """
    Solves the Excel XLOOKUP mismatch by coercing integers, floats,
    text strings, scientific notation, and non-breaking spaces into 
    a single clean, uniform string key.
    """
    if pd.isna(val) or val is None:
        return ""
    
    clean_str = str(val).replace("\xa0", " ").strip()
    
    if "." in clean_str:
        try:
            float_val = float(clean_str)
            if float_val.is_integer():
                return str(int(float_val))
        except ValueError:
            pass
        clean_str = clean_str.split(".")[0].strip()

    return clean_str.upper()


@lru_cache(maxsize=4096)
def _translate_cached(clean_name: str) -> str:
    """Cached lookup so repeated names across rows only hit the API once."""
    translator = _get_translator()
    if translator is None:
        return ""
    try:
        result = translator.translate(clean_name)
        return result if result else ""
    except Exception:
        return ""


def transliterate_name_to_marathi(name_str: str) -> str:
    """
    Transliterates English student name to Marathi Devanagari script using deep_translator.
    Returns "" (without raising) if deep_translator is unavailable or the
    translation call fails for any reason, e.g. no network access.
    """
    if pd.isna(name_str) or not str(name_str).strip():
        return ""
    clean_name = str(name_str).strip()
    return _translate_cached(clean_name)


def derive_faculty(degnm: str) -> str:
    """
    Derive Faculty based on keywords/patterns in DEGNM (case-insensitive):
    - Commerce: COMMERCE, Management, Finance, Economics, B.Com, BCom
    - Science: SCIENCE, BSc, B.Sc, B.Sc.
    - Arts: ARTS, B.A.
    """
    if pd.isna(degnm) or not str(degnm).strip():
        return ""
    
    degnm_str = str(degnm).strip()
    
    commerce_pattern = r'(?i)\b(commerce|accounting|management|finance|financial|marketing|business|economics|b\.?com|m\.?com)\b'
    if re.search(commerce_pattern, degnm_str):
        return "Commerce"
        
    science_pattern = r'(?i)\b(science|artificial|intelligent|data|b\.?sc\.?|m\.?sc\.?)\b'
    if re.search(science_pattern, degnm_str):
        return "Science"
        
    arts_pattern = r'(?i)\b(arts|entertainment|media|film|b\.?a\.?|m\.?a\.?)\b'
    if re.search(arts_pattern, degnm_str):
        return "Arts"
        
    return ""


def derive_subdegnm(degnm: str) -> str:
    """
    Derive SUBDEGNM based on degree level:
    - Master/Masters/begins with 'M' -> 'Two Year Degree Course'
    - Bachelor/Bachelors/begins with 'B' -> 'Three Year Degree Course'
    """
    if pd.isna(degnm) or not str(degnm).strip():
        return ""
    
    degnm_str = str(degnm).strip()
    if re.search(r'\bMASTERS?\b', degnm_str, re.IGNORECASE) or degnm_str.upper().startswith('M'):
        return "Two Year Degree Course"
    if re.search(r'\bBACHELORS?\b', degnm_str, re.IGNORECASE) or degnm_str.upper().startswith('B'):
        return "Three Year Degree Course"
        
    return ""


def map_gender(val) -> str:
    """Convert Gender values to SEX: Male = 1 and Female = 2."""
    if pd.isna(val):
        return ""
    
    val_str = str(val).strip().lower()
    if val_str in ["male", "m", "1"]:
        return "1"
    elif val_str in ["female", "f", "2"]:
        return "2"
    return str(val)


def format_class_grade(val) -> str:
    """Format CLASS grade values (e.g., "'A' Grade", "'B+' Grade")."""
    if pd.isna(val) or not str(val).strip():
        return ""
    
    val_str = str(val).strip()
    clean_val = re.sub(r"(?i)\bgrade\b", "", val_str).strip(" '\"")
    if clean_val:
        return f"'{clean_val}' Grade"
    return ""


# ==========================================
# 3. DYNAMIC MASTER LOOKUP MAP CREATION
# ==========================================
def find_priority_column(columns, keyword_groups):
    """Finds the best-matching column from a list of priority keyword groups."""
    cols_lower = {c: c.lower() for c in columns}
    for keywords in keyword_groups:
        for c in columns:
            cl = cols_lower[c]
            if any(k in cl for k in keywords):
                return c
    return None


def build_master_lookup(master_files):
    """
    Reads all uploaded Student Master Data files across all sheets.
    Builds Student Number -> Program Name dictionary.
    """
    student_lookup = {}
    master_records_count = 0

    for file in master_files:
        excel_obj = pd.ExcelFile(file)
        for sheet_name in excel_obj.sheet_names:
            df_master = pd.read_excel(excel_obj, sheet_name=sheet_name, dtype=str)
            master_records_count += len(df_master)

            df_master.columns = [str(c).replace("\xa0", " ").strip() for c in df_master.columns]

            stu_col = find_priority_column(
                df_master.columns,
                [["student number"], ["student_number"]]
            )
            if not stu_col and len(df_master.columns) > 0:
                stu_col = df_master.columns[0]

            prog_col = find_priority_column(
                df_master.columns,
                [
                    ["program name"],
                    ["program_name"],
                    ["programme name"],
                    ["degree name"],
                ]
            )
            if not prog_col and len(df_master.columns) >= 204:
                prog_col = df_master.columns[203]

            if stu_col and prog_col:
                for _, row in df_master.iterrows():
                    key = sanitize_key(row[stu_col])
                    prog_val = row[prog_col]
                    if key and pd.notna(prog_val) and str(prog_val).strip():
                        student_lookup[key] = str(prog_val).strip()

    return student_lookup, master_records_count


# ==========================================
# 4. CORE ETL & MERGE PROCESSING
# ==========================================
def process_data(mkcl_files, student_lookup, college_choice):
    """
    Merges MKCL files across sheets, performs Python lookup for DEGNM,
    applies dynamic row filtering (Convocation Number blank, CGPA and all existing 
    SEM_GPA columns present), eliminates duplicate rows, and sorts by SEAT_NO.
    """
    mkcl_dfs = []
    total_mkcl_raw_rows = 0

    for file in mkcl_files:
        excel_obj = pd.ExcelFile(file)
        for sheet_name in excel_obj.sheet_names:
            df = pd.read_excel(excel_obj, sheet_name=sheet_name, dtype=str)
            total_mkcl_raw_rows += len(df)

            df.columns = [str(c).replace("\xa0", " ").strip() for c in df.columns]
            
            # 1. Convocation Number filter: must be blank/null/empty
            conv_col = next((c for c in df.columns if c.lower() == "convocation number"), None)
            if conv_col:
                df = df[df[conv_col].isna() | (df[conv_col].astype(str).str.strip() == "")]

            # 2. CGPA filter: must contain a value
            cgpa_col = next((c for c in df.columns if c.lower() == "cgpa"), None)
            if cgpa_col:
                df = df[df[cgpa_col].notna() & (df[cgpa_col].astype(str).str.strip() != "")]
            else:
                df = df.iloc[0:0]

            # 3. Dynamic Semester GPA filter: All present SEM X_GPA columns must contain a value
            target_sem_gpas = ["SEM 1_GPA", "SEM 2_GPA", "SEM 3_GPA", "SEM 4_GPA", "SEM 5_GPA", "SEM 6_GPA"]
            for sem_gpa_name in target_sem_gpas:
                sem_col = next((c for c in df.columns if c.lower() == sem_gpa_name.lower()), None)
                if sem_col:
                    df = df[df[sem_col].notna() & (df[sem_col].astype(str).str.strip() != "")]

            if not df.empty:
                mkcl_dfs.append(df)

    if not mkcl_dfs:
        empty_df = pd.DataFrame(columns=TARGET_COLUMNS)
        return empty_df, pd.DataFrame(), total_mkcl_raw_rows, 0, 0, 0

    merged_mkcl = pd.concat(mkcl_dfs, ignore_index=True)
    total_filtered_rows = len(merged_mkcl)

    matched_count = 0
    unmatched_count = 0
    deg_names = []
    unmatched_rows = []

    stu_num_col = next((c for c in merged_mkcl.columns if "student number" in c.lower()), None)
    if not stu_num_col and len(merged_mkcl.columns) > 0:
        stu_num_col = merged_mkcl.columns[0]

    for idx, row in merged_mkcl.iterrows():
        stu_id = sanitize_key(row[stu_num_col]) if stu_num_col else ""
        prog_name = student_lookup.get(stu_id, "")

        if prog_name:
            matched_count += 1
            deg_names.append(prog_name)
        else:
            unmatched_count += 1
            deg_names.append("")
            unmatched_rows.append(row.to_dict())

    merged_mkcl["DEGNM"] = deg_names

    out_df = pd.DataFrame(columns=TARGET_COLUMNS)

    for src_col, target_col in MKCL_SOURCE_MAPPING.items():
        matched_src = next((c for c in merged_mkcl.columns if c.lower() == src_col.lower()), None)
        if matched_src:
            out_df[target_col] = merged_mkcl[matched_src].apply(sanitize_key) if target_col == "SEAT_NO" else merged_mkcl[matched_src].astype(str).str.strip()

    prn_col = next((c for c in merged_mkcl.columns if "prn" in c.lower()), None)
    if prn_col:
        out_df["PRNERN"] = merged_mkcl[prn_col].apply(sanitize_key)

    out_df["DEGNM"] = merged_mkcl["DEGNM"].astype(str).str.strip()
    out_df["PROGTYPE"] = "DEGREE"

    if college_choice == "Narsee Monjee College of Commerce & Economics (Autonomous)":
        out_df["COLL_NO"] = "205"
        out_df["COLL_NAME"] = "Narsee Monjee College of Commerce & Economics (Autonomous)"
    else:
        out_df["COLL_NO"] = "598"
        out_df["COLL_NAME"] = "UPG College of Arts, Science & Commerce (AUTONOMOUS)"

    out_df["SEX"] = out_df["SEX"].apply(map_gender)
    out_df["CLASS"] = out_df["CLASS"].apply(format_class_grade)
    out_df["FACULTY"] = out_df["DEGNM"].apply(derive_faculty)
    out_df["SUBDEGNM"] = out_df["DEGNM"].apply(derive_subdegnm)

    # Transliterate NAME column to Marathi for NAME_MARAT
    out_df["NAME_MARAT"] = out_df["NAME"].apply(transliterate_name_to_marathi)

    out_df = out_df.fillna("").astype(str).replace("nan", "")
    out_df = out_df[TARGET_COLUMNS]

    # Deduplicate rows in the output dataset
    out_df = out_df.drop_duplicates().reset_index(drop=True)

    # Sort output dataframe by SEAT_NO
    out_df = out_df.sort_values(by="SEAT_NO", ascending=True).reset_index(drop=True)

    unmatched_df = pd.DataFrame(unmatched_rows)
    if not unmatched_df.empty:
        unmatched_df = unmatched_df.drop_duplicates().reset_index(drop=True)

    return out_df, unmatched_df, total_mkcl_raw_rows, total_filtered_rows, matched_count, unmatched_count


# ==========================================
# 5. OPENPYXL EXCEL GENERATION & STYLING
# ==========================================
def generate_formatted_excel(df: pd.DataFrame) -> bytes:
    """Generates formatted Excel workbook with default row height 20."""
    wb = openpyxl.Workbook()
    
    # SHEET 1: Convocation Data
    ws = wb.active
    ws.title = "Convocation Data"
    ws.views.sheetView[0].showGridLines = False

    font_regular = Font(name="Times New Roman", size=11, bold=False)
    font_header = Font(name="Times New Roman", size=11, bold=True)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=False)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    thin_border_side = Side(style="thin", color="D3D3D3")
    cell_border = Border(
        left=thin_border_side, right=thin_border_side,
        top=thin_border_side, bottom=thin_border_side
    )

    ws.append(list(df.columns))
    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = cell_border

    for row_values in df.itertuples(index=False):
        ws.append(list(row_values))

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = font_regular
            cell.border = cell_border
            if cell.column_letter in ["A", "B", "E", "G", "H", "Q", "AE"]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Set row height to 20 for Sheet 1
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 20

    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # SHEET 2: Statistics
    ws_stats = wb.create_sheet(title="Statistics")
    ws_stats.views.sheetView[0].showGridLines = False

    stats_headers = ["PROGRAM NAME", "MALE (1)", "FEMALE (2)", "TOTAL COUNT"]
    ws_stats.append(stats_headers)

    for col_idx in range(1, len(stats_headers) + 1):
        cell = ws_stats.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = cell_border

    if not df.empty and "DEGNM" in df.columns and "SEX" in df.columns:
        program_groups = df.groupby("DEGNM")
        
        grand_male = 0
        grand_female = 0
        grand_total = 0

        for prog_name, group in program_groups:
            display_prog = prog_name if prog_name.strip() else "[UNMATCHED PROGRAM]"
            male_cnt = (group["SEX"] == "1").sum()
            female_cnt = (group["SEX"] == "2").sum()
            total_cnt = len(group)

            grand_male += male_cnt
            grand_female += female_cnt
            grand_total += total_cnt

            row_data = [display_prog, male_cnt, female_cnt, total_cnt]
            ws_stats.append(row_data)

        total_row_idx = ws_stats.max_row + 1
        ws_stats.append(["TOTAL", grand_male, grand_female, grand_total])

        for row in ws_stats.iter_rows(min_row=2, max_row=ws_stats.max_row, min_col=1, max_col=4):
            is_total_row = (row[0].row == total_row_idx)
            for idx, cell in enumerate(row):
                cell.font = Font(name="Times New Roman", size=11, bold=is_total_row)
                cell.border = cell_border
                if idx == 0:
                    cell.alignment = left_align
                else:
                    cell.alignment = center_align

    # Set row height to 20 for Sheet 2
    for r in range(1, ws_stats.max_row + 1):
        ws_stats.row_dimensions[r].height = 20

    for col in ws_stats.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws_stats.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer.getvalue()


# ==========================================
# 6. MODULE ENTRY POINT FOR STREAMLIT ROUTING
# ==========================================
def show():
    """Main UI layout function called by app.py routing."""
    st.title("👩🏻‍🎓 Convocation Data Generator 👨🏻‍🎓")
    st.markdown(
        "Upload **MKCL Reports** and **Student Master Data** files. The app automatically handles "
        "Excel text/number mismatches to fetch degree/program names accurately, standardizes columns, "
        "and generates formatted convocation reports."
    )

    if _get_translator() is None:
        st.warning(
            "⚠️ The `deep_translator` package isn't installed, so the **NAME_MARAT** "
            "(Marathi name) column will be left blank. Add `deep_translator` to your "
            "requirements.txt and redeploy to enable automatic transliteration.",
            icon="⚠️",
        )

    # 3-Column horizontal single-line layout for controls & uploaders
    col1, col2, col3 = st.columns(3)

    with col1:
        st.text("🏫 Select College")
        college_option = st.selectbox(
            "Select College",
            options=[
                "Narsee Monjee College of Commerce & Economics (Autonomous)",
                "UPG College of Arts, Science & Commerce (AUTONOMOUS)"
            ],
            label_visibility="collapsed"
        )

    with col2:
        mkcl_files = st.file_uploader(
            "Upload MKCL Report Files (.xlsx / .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="mkcl_uploader"
        )

    with col3:
        master_files = st.file_uploader(
            "Upload Student Master Data Files (.xlsx / .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="master_uploader"
        )

    st.markdown("---")

    process_ready = bool(mkcl_files and master_files)
    process_btn = st.button("🚀 Process & Generate Convocation File", type="primary", disabled=not process_ready)

    if not process_ready:
        st.info("💡 Upload at least one MKCL Report file AND one Student Master Data file to enable processing.")

    if process_btn and process_ready:
        try:
            with st.spinner("Building Student Master Lookup Map, Processing MKCL Reports & Transliterating Names to Marathi..."):
                student_lookup, master_total_records = build_master_lookup(master_files)

                final_df, unmatched_df, mkcl_raw_count, filtered_count, matched_count, unmatched_count = process_data(
                    mkcl_files, student_lookup, college_option
                )

            st.subheader("📊 Processing Summary Statistics")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Master Records Loaded", f"{master_total_records:,}")
            m2.metric("MKCL Raw Records", f"{mkcl_raw_count:,}")
            m3.metric("Filtered Records (Conv/CGPA/GPAs)", f"{filtered_count:,}")
            m4.metric("Program Matches Found", f"{matched_count:,}")
            m5.metric("Unmatched Records", f"{unmatched_count:,}")

            if unmatched_count > 0:
                st.warning(f"⚠️ **Exception Report:** {unmatched_count} record(s) could not find a matching Program Name in Master Data.")
                with st.expander("🔍 View Unmatched Exceptions"):
                    st.dataframe(unmatched_df, use_container_width=True)

            if final_df.empty:
                st.error("❌ No valid output records generated. Please verify your filter criteria or uploaded files.")
            else:
                st.subheader("Final Processed Data Preview")
                st.dataframe(final_df, use_container_width=True)

                excel_data = generate_formatted_excel(final_df)

                college_tag = "NMC" if "Narsee Monjee" in college_option else "UPG"
                current_date_str = datetime.now().strftime("%d-%m-%Y")
                dynamic_filename = f"Convocation_Data_{college_tag}_{current_date_str}.xlsx"

                st.download_button(
                    label="📥 Download Standardized Convocation Excel File",
                    data=excel_data,
                    file_name=dynamic_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ An error occurred during processing: {str(e)}")
            st.exception(e)


# Alias main to show for flexibility
main = show

if __name__ == "__main__":
    show()