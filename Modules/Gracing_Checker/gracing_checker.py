"""
==================================================
GRACING CHECKER MODULE
==================================================

PURPOSE
--------------------------------------------------
This Streamlit tool processes examination gracing sheets:

1. Loads uploaded Excel files containing pre/post gracing records.
2. Restructures student records into a consolidated tabular layout with exact column order:
   - Module Code, Module Desc, Credit, Internal Actual (Max), Internal Actual,
     Sem Actual (Max), Sem Actual, Sem Grace, Composite Score (Max), 
     Composite Score, Overall Grade, Completion Status after Gracing, Remarks & Scale.
   - Merged student header row across all columns:
     "Student Number | Student Name | Additional ID | Gender"
   - Red text highlight on Composite Score if component marks don't sum up.
   - Summary row showing Total Composite Score, Total Composite Score (Max), 
     and Percentage in Overall Grade.
   - Center alignment for columns Credit through Overall Grade.
   - Highlights 'Completed Unsuccessfuly' statuses in RED.
3. Identifies all instances where grace marks were awarded.
4. Computes high-level statistics & summary reports.
5. Displays tab-based previews with 1-based indexing:
   - 📊 Processed Data (First Tab)
   - 📋 Graced Cases Details
   - 📚 Subject Summary
   - 👤 Student Summary
6. Generates a formatted multi-sheet Excel report for download.
==================================================
"""

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
def parse_scale(scale_val):
    """Parses maximum mark integer from scale strings (e.g., 'R050' -> '50', 'S020' -> '20')."""
    s = str(scale_val).strip()
    digits = re.findall(r'\d+', s)
    if digits:
        return str(int(digits[0]))
    return s if s.upper() not in ['NAN', 'NONE', 'NA'] else ''


# ==================================================
# STEP 1: PROCESSED DATA GENERATION
# ==================================================
def create_processed_data(df):
    """
    Restructures raw examination records into a clean student-wise dataset:
    - Merged Student Info Row across all columns: "Student Number | Name | Add ID | Gender"
    - Strict Column Order:
      1. Module Code
      2. Module Desc
      3. Credit
      4. Internal Actual (Max)
      5. Internal Actual
      6. Sem Actual (Max)
      7. Sem Actual
      8. Sem Grace
      9. Composite Score (Max)
      10. Composite Score
      11. Overall Grade
      12. Completion Status after Gracing
      13. Remarks & Scale
    """
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    student_ids = df['Student Number'].dropna().unique()
    all_processed_rows = []

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
        all_processed_rows.append({
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
            'Comp_Mismatch': False
        })

        modules = module_df['Module Code'].dropna().unique()
        calc_total_comp = 0.0
        calc_total_max_comp = 0.0

        for m_code in modules:
            m_sub = module_df[module_df['Module Code'] == m_code]
            m_desc = m_sub.iloc[0]['Module Desc']

            # Non-zero credit from Overall Grade or first available row
            credit_row = m_sub[m_sub['Appraisal Type Code'].astype(str).str.strip() == 'Overall Grade']
            credit = credit_row.iloc[0]['Credit'] if not credit_row.empty else m_sub.iloc[0]['Credit']

            # Extraction helper for appraisal type values and scale max marks
            def get_val_and_scale(code):
                sub = m_sub[m_sub['Appraisal Type Code'].astype(str).str.strip() == code]
                if not sub.empty:
                    val = sub['Marks after Gracing'].iloc[0]
                    scale = sub['Scale'].iloc[0] if 'Scale' in sub.columns else ''
                    return val, parse_scale(scale)
                return '', ''

            internal_actual, internal_max = get_val_and_scale('Internal Actual')
            sem_actual, sem_max = get_val_and_scale('Sem Actual')
            sem_grace, _ = get_val_and_scale('Sem Grace')
            composite_score, composite_max = get_val_and_scale('Composite Score')
            overall_grade, _ = get_val_and_scale('Overall Grade')

            sub_overall = m_sub[m_sub['Appraisal Type Code'].astype(str).str.strip() == 'Overall Grade']
            completion_status = sub_overall['Completion Status after Gracing'].iloc[0] if not sub_overall.empty else ''

            # Merge Appraisal Remarks, General Remarks, Scale Type
            app_rem = str(sub_overall['Appraisal Remarks'].iloc[0]).strip() if not sub_overall.empty else ''
            gen_rem = str(sub_overall['General Remarks'].iloc[0]).strip() if not sub_overall.empty else ''
            scale_t = str(sub_overall['Scale Type'].iloc[0]).strip() if not sub_overall.empty else ''

            remarks_list = [r for r in [app_rem, gen_rem, scale_t] if r and r.upper() not in ['NAN', 'NONE', 'NA']]
            remarks_combined = " | ".join(remarks_list)

            # Verification: Internal Actual + Sem Actual + Sem Grace == Composite Score
            is_mismatch = False
            try:
                i_val = float(internal_actual) if str(internal_actual).strip() not in ['', 'nan', 'NA'] else 0.0
                s_val = float(sem_actual) if str(sem_actual).strip() not in ['', 'nan', 'NA'] else 0.0
                g_val = float(sem_grace) if str(sem_grace).strip() not in ['', 'nan', 'NA'] else 0.0
                c_val = float(composite_score) if str(composite_score).strip() not in ['', 'nan', 'NA'] else 0.0

                calc_total_comp += c_val

                sum_components = i_val + s_val + g_val
                if str(composite_score).strip() not in ['', 'nan', 'NA']:
                    if abs(sum_components - c_val) >= 0.01:
                        is_mismatch = True
            except (TypeError, ValueError):
                pass

            # Calculate total max composite marks
            try:
                cm_val = float(composite_max) if str(composite_max).strip() not in ['', 'nan', 'NA'] else 0.0
                calc_total_max_comp += cm_val
            except (TypeError, ValueError):
                pass

            all_processed_rows.append({
                'Module Code': m_code,
                'Module Desc': m_desc,
                'Credit': credit,
                'Internal Actual (Max)': internal_max,
                'Internal Actual': internal_actual,
                'Sem Actual (Max)': sem_max,
                'Sem Actual': sem_actual,
                'Sem Grace': sem_grace,
                'Composite Score (Max)': composite_max,
                'Composite Score': composite_score,
                'Overall Grade': overall_grade,
                'Completion Status after Gracing': completion_status,
                'Remarks & Scale': remarks_combined,
                'Row_Type': 'DATA',
                'Comp_Mismatch': is_mismatch
            })

        # 2. Total & Percentage Summary Row
        tot_val_raw = total_row.iloc[0]['Marks after Gracing'] if not total_row.empty else (total_row.iloc[0]['Marks before Gracing'] if not total_row.empty else '')
        perc_val_raw = perc_row.iloc[0]['Marks after Gracing'] if not perc_row.empty else (perc_row.iloc[0]['Marks before Gracing'] if not perc_row.empty else '')

        total_mismatch = False
        try:
            tot_num = float(tot_val_raw)
            if abs(calc_total_comp - tot_num) >= 0.01:
                total_mismatch = True
            tot_display = f"{tot_num:g}"
        except (TypeError, ValueError):
            tot_display = str(tot_val_raw)

        try:
            perc_num = float(perc_val_raw)
            perc_display = f"{perc_num:g}%"
        except (TypeError, ValueError):
            perc_display = f"{perc_val_raw}%" if str(perc_val_raw).strip() != '' else ''

        tot_max_display = f"{int(calc_total_max_comp)}" if calc_total_max_comp.is_integer() else f"{calc_total_max_comp:g}"

        all_processed_rows.append({
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
            'Comp_Mismatch': total_mismatch
        })

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
    - Red text for mismatched Composite Scores
    - Red font for 'Completed Unsuccessfuly'
    """
    display_df = proc_df.drop(columns=['Row_Type', 'Comp_Mismatch'], errors='ignore').copy()
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
            row_type = proc_df.iloc[orig_idx]['Row_Type'] if 'Row_Type' in proc_df.columns else ''
            is_mismatch = proc_df.iloc[orig_idx]['Comp_Mismatch'] if 'Comp_Mismatch' in proc_df.columns else False
            status_val = str(data.loc[idx, 'Completion Status after Gracing']).strip()

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

            # 2. Total Summary Row (Soft Yellow Background + Bold text)
            if row_type == 'SUMMARY':
                styles.loc[idx, :] += "background-color: #FFF2CC; font-weight: bold; color: #000000;"
                if is_mismatch and 'Composite Score' in data.columns:
                    styles.loc[idx, 'Composite Score'] += "color: red; font-weight: bold;"
                continue

            # 3. Red text highlight if Composite Score mismatched
            if is_mismatch and 'Composite Score' in data.columns:
                styles.loc[idx, 'Composite Score'] += "color: red; font-weight: bold;"

            # 4. Red font for 'Completed Unsuccessfuly'
            if status_val.lower() == 'completed unsuccessfuly' and 'Completion Status after Gracing' in data.columns:
                styles.loc[idx, 'Completion Status after Gracing'] += "color: red; font-weight: bold;"

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


# ==================================================
# STEP 4: EXCEL REPORT GENERATION
# ==================================================
def to_excel_bytes(proc_df, graced_df, subject_summary, student_summary):
    """Generates styled multi-sheet Excel output file."""
    buffer = io.BytesIO()
    clean_proc = proc_df.drop(columns=['Row_Type', 'Comp_Mismatch'], errors='ignore')

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        clean_proc.to_excel(writer, index=False, sheet_name="Processed Data")
        if not graced_df.empty:
            graced_df.to_excel(writer, index=False, sheet_name="Graced Cases Details")
            subject_summary.to_excel(writer, index=False, sheet_name="Subject Summary")
            student_summary.to_excel(writer, index=False, sheet_name="Student Summary")

        workbook = writer.book

        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000")
        )
        center_align = Alignment(horizontal="center", vertical="center")
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        student_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        summary_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        red_font = Font(color="FF0000", bold=True)
        bold_blue_font = Font(color="1F4E78", bold=True)
        bold_font = Font(bold=True)

        sheets = ["Processed Data"] + (["Graced Cases Details", "Subject Summary", "Student Summary"] if not graced_df.empty else [])

        for sheet_name in sheets:
            worksheet = workbook[sheet_name]

            # Disable gridlines & freeze header
            worksheet.views.sheetView[0].showGridLines = False
            worksheet.freeze_panes = "A2"

            if sheet_name == "Processed Data":
                cols = clean_proc.columns.tolist()
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
                comp_score_idx = cols.index('Composite Score') + 1 if 'Composite Score' in cols else None
                status_col_idx = cols.index('Completion Status after Gracing') + 1 if 'Completion Status after Gracing' in cols else None

                for row_idx in range(1, len(clean_proc) + 2):
                    if row_idx == 1:
                        for col_idx in range(1, len(cols) + 1):
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.border = thin_border
                            cell.font = bold_font
                            cell.fill = header_fill
                            cell.alignment = center_align
                        continue

                    proc_idx = row_idx - 2
                    row_type_val = proc_df.iloc[proc_idx]['Row_Type'] if (0 <= proc_idx < len(proc_df) and 'Row_Type' in proc_df.columns) else ''
                    is_mismatch = proc_df.iloc[proc_idx]['Comp_Mismatch'] if (0 <= proc_idx < len(proc_df) and 'Comp_Mismatch' in proc_df.columns) else False

                    is_student_header = (row_type_val == 'HEADER')
                    is_summary_row = (row_type_val == 'SUMMARY')

                    status_val = str(worksheet.cell(row=row_idx, column=status_col_idx).value).strip() if status_col_idx else ''

                    if is_student_header:
                        # Merge student header across all data columns
                        worksheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=len(cols))
                        for col_idx in range(1, len(cols) + 1):
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.border = thin_border
                            cell.font = bold_blue_font
                            cell.fill = student_header_fill
                    else:
                        for col_idx in range(1, len(cols) + 1):
                            cell = worksheet.cell(row=row_idx, column=col_idx)
                            cell.border = thin_border

                            if is_summary_row:
                                cell.font = bold_font
                                cell.fill = summary_fill
                            elif col_idx in max_cols_indices:
                                cell.font = bold_font

                            if col_idx in center_cols_indices:
                                cell.alignment = center_align

                            # Red font for mismatched composite score
                            if is_mismatch and col_idx == comp_score_idx:
                                cell.font = red_font

                            # Red font for unsuccessful completion status
                            if status_col_idx and col_idx == status_col_idx and status_val.lower() == 'completed unsuccessfuly':
                                cell.font = red_font
            else:
                for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                    for cell in row:
                        cell.border = thin_border
                        if cell.row == 1:
                            cell.font = bold_font
                            cell.fill = header_fill
                            cell.alignment = center_align

            # Auto-size columns
            for col in worksheet.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        max_len = max(max_len, len(val_str))
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 14)

    return buffer.getvalue()


# ==================================================
# MAIN ENTRY POINT
# ==================================================
def show():
    """Renders the Gracing Checker tool page in Streamlit."""
    st.title("🎁 Gracing Checker")
    st.caption("Upload an examination result Excel file to analyze awarded grace marks.")

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

        # Generate Processed Data & Gracing Analysis
        proc_df = create_processed_data(df)
        graced_df = process_gracing_data(df)
        subject_summary = get_subject_summary(graced_df)
        student_summary = get_student_summary(graced_df)

        # Statistics
        total_students = df['Student Number'].nunique()
        graced_students_count = graced_df['Student Number'].nunique() if not graced_df.empty else 0
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

        st.markdown("---")

        # Tab-based Preview Windows (Processed Data tab is first)
        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Processed Data", 
            "📋 Graced Cases Details", 
            "📚 Subject Summary", 
            "👤 Student Summary"
        ])

        with tab1:
            st.subheader("Processed Data Preview")
            if not proc_df.empty:
                st.write(style_processed_preview(proc_df))
            else:
                st.info("No processed data generated.")

        with tab2:
            st.subheader("Graced Cases Details")
            if not graced_df.empty:
                st.write(style_preview(graced_df))
            else:
                st.info("No grace marks were awarded in this file.")

        with tab3:
            st.subheader("Subject-Wise Gracing Summary")
            if not subject_summary.empty:
                st.write(style_preview(subject_summary))
            else:
                st.info("No subject-wise gracing data available.")

        with tab4:
            st.subheader("Student-Wise Gracing Summary")
            if not student_summary.empty:
                st.write(style_preview(student_summary))
            else:
                st.info("No student-wise gracing data available.")

        # Export Section
        st.markdown("---")
        st.subheader("📥 Export Gracing Report")

        excel_bytes = to_excel_bytes(proc_df, graced_df, subject_summary, student_summary)
        date_str = datetime.now().strftime("%d-%m-%y")
        output_filename = f"Gracing_Analysis_Report_{date_str}.xlsx"

        st.download_button(
            label=f"📥 Download Complete Gracing Report ({output_filename})",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as error:
        st.error(f"Something went wrong while processing the file: {error}")