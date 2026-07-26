"""
==================================================
GRACING CHECKER MODULE
==================================================

PURPOSE
--------------------------------------------------
This Streamlit tool processes examination gracing sheets:

1. Loads uploaded Excel files containing pre/post gracing records.
2. Identifies all instances where grace marks were awarded to students.
3. Computes summary statistics:
   - Total Students evaluated vs. Students Graced
   - Total Gracing Instances & Total Grace Marks Awarded
   - Subject-wise gracing distribution & student-wise summaries
4. Displays interactive preview tabs with 1-based indexing:
   - Graced Cases Details
   - Subject Summary
   - Student Summary
   - Single Student Details (with Previous / Next navigation)
   - All Raw Data
5. Generates a formatted multi-sheet Excel report for download.
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
# STEP 1: DATA EXTRACTION & PROCESSING
# ==================================================
def process_gracing_data(df):
    """
    Extracts detailed gracing records from the raw examination dataset.
    
    Identifies rows where 'Sem Grace' or 'Grace on Total marks' > 0 and 
    extracts before/after marks, composite scores, overall grades, and status.
    """
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()

    # Filter out empty rows and student summary rows (e.g. 'Total', 'Percentage')
    non_summary = df[
        df['Module Desc'].notna() & 
        (~df['Student Name'].astype(str).str.upper().isin(['TOTAL', 'PERCENTAGE']))
    ].copy()

    # Identify rows corresponding to grace marks
    grace_rows = non_summary[
        non_summary['Appraisal Type Code'].astype(str).str.strip().isin(['Sem Grace', 'Grace on Total marks'])
    ].copy()

    # Convert grace marks to numeric
    grace_rows['Grace_Marks_Val'] = pd.to_numeric(grace_rows['Marks after Gracing'], errors='coerce').fillna(0)

    # Filter strictly rows where grace marks were awarded (> 0)
    awarded_grace = grace_rows[grace_rows['Grace_Marks_Val'] > 0].copy()

    records = []
    for _, g_row in awarded_grace.iterrows():
        s_num = g_row['Student Number']
        m_code = g_row['Module Code']

        # Get all evaluation component rows for this student and module
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

    graced_df = pd.DataFrame(records)
    return graced_df


def get_subject_summary(graced_df):
    """Generates subject-wise aggregation of grace marks awarded."""
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
    """Generates student-wise aggregation listing graced subjects and marks."""
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
# STEP 2: DISPLAY FORMATTING HELPERS
# ==================================================
def format_value(value):
    """Formats cell values cleanly for Streamlit interactive preview."""
    text = str(value).strip()
    if text == "" or text.upper() == "NA" or text.upper() == "NAN":
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
        return f"{number:g}"
    except (TypeError, ValueError):
        return text


def style_preview(df):
    """Prepares dataframe with 1-based indexing for preview display."""
    display_df = df.copy()
    display_df.index = range(1, len(display_df) + 1)
    return display_df.style.format(format_value)


# ==================================================
# STEP 3: EXCEL REPORT GENERATION
# ==================================================
def to_excel_bytes(graced_df, subject_summary, student_summary):
    """
    Converts dataframes to a styled multi-sheet Excel file:
    - Sheets: Graced Cases Details, Subject Summary, Student Summary
    - Centered bold header with grey background
    - Frozen top row and hidden default gridlines
    - Thin cell borders and auto-sized column widths
    """
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
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
        grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        for sheet_name in ["Graced Cases Details", "Subject Summary", "Student Summary"]:
            worksheet = workbook[sheet_name]

            # Disable default gridlines & freeze top row
            worksheet.views.sheetView[0].showGridLines = False
            worksheet.freeze_panes = "A2"

            # Apply cell formatting
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.border = thin_border
                    
                    if cell.row == 1:
                        cell.font = Font(bold=True)
                        cell.fill = grey_fill
                        cell.alignment = center_align

            # Auto-size columns based on maximum content length
            for col in worksheet.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    return buffer.getvalue()


# ==================================================
# MAIN ENTRY POINT
# ==================================================
def show():
    """Renders the Gracing Checker tool page in Streamlit."""
    st.title("✅ Gracing Checker")
    st.caption("Upload an examination result Excel file to analyze awarded grace marks.")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file is None:
        return

    try:
        # Load raw file
        with st.spinner("Loading and analyzing gracing data..."):
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

        # Process gracing data
        graced_df = process_gracing_data(df)
        subject_summary = get_subject_summary(graced_df)
        student_summary = get_student_summary(graced_df)

        # High-level statistics
        total_students = df['Student Number'].nunique()
        graced_students_count = graced_df['Student Number'].nunique() if not graced_df.empty else 0
        total_graced_cases = len(graced_df)
        total_grace_marks = graced_df['Grace Marks'].sum() if not graced_df.empty else 0

        # Update global session state tracking
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

        # Filter out summary rows for single student details
        non_summary = df[
            df['Module Desc'].notna() & 
            (~df['Student Name'].astype(str).str.upper().isin(['TOTAL', 'PERCENTAGE']))
        ].copy()

        unique_students = non_summary['Student Number'].dropna().unique().tolist()

        # Initialize student index in session state if not present
        if "gracing_student_idx" not in st.session_state or st.session_state.gracing_student_idx >= len(unique_students):
            st.session_state.gracing_student_idx = 0

        # Interactive Preview Tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📋 Graced Cases Details", 
            "📚 Subject Summary", 
            "👤 Student Summary", 
            "🔍 Single Student Details",
            "📄 All Raw Data"
        ])

        with tab1:
            st.subheader("Graced Cases Details")
            if not graced_df.empty:
                st.write(style_preview(graced_df))
            else:
                st.info("No grace marks were awarded in this file.")

        with tab2:
            st.subheader("Subject-Wise Gracing Summary")
            if not subject_summary.empty:
                st.write(style_preview(subject_summary))
            else:
                st.info("No subject-wise gracing data available.")

        with tab3:
            st.subheader("Student-Wise Gracing Summary")
            if not student_summary.empty:
                st.write(style_preview(student_summary))
            else:
                st.info("No student-wise gracing data available.")

        with tab4:
            st.subheader("Single Student Details")
            
            if unique_students:
                # Navigation Controls (Previous / Dropdown / Next)
                nav_col1, nav_col2, nav_col3 = st.columns([1, 3, 1])

                with nav_col1:
                    if st.button("◀ Previous", key="prev_student_btn", use_container_width=True):
                        if st.session_state.gracing_student_idx > 0:
                            st.session_state.gracing_student_idx -= 1
                            st.rerun()

                with nav_col3:
                    if st.button("Next ▶", key="next_student_btn", use_container_width=True):
                        if st.session_state.gracing_student_idx < len(unique_students) - 1:
                            st.session_state.gracing_student_idx += 1
                            st.rerun()

                with nav_col2:
                    selected_student = st.selectbox(
                        "Jump to Student Number:",
                        options=unique_students,
                        index=st.session_state.gracing_student_idx,
                        key="student_select_box"
                    )
                    # Sync dropdown selection index
                    st.session_state.gracing_student_idx = unique_students.index(selected_student)

                current_student_id = unique_students[st.session_state.gracing_student_idx]

                # Filter student specific rows
                student_df = non_summary[non_summary['Student Number'] == current_student_id]

                if not student_df.empty:
                    st.markdown("---")
                    
                    # Student Metadata Header Cards
                    first_row = student_df.iloc[0]
                    std_num = first_row['Student Number']
                    std_name = first_row.get('Student Name', '')
                    add_id = first_row.get('Additional ID', '')
                    gender = first_row.get('Gender', '')

                    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
                    with m_col1:
                        st.info(f"**Student Number:**\n\n{std_num}")
                    with m_col2:
                        st.success(f"**Student Name:**\n\n{std_name}")
                    with m_col3:
                        st.warning(f"**Additional ID:**\n\n{add_id}")
                    with m_col4:
                        st.error(f"**Gender:**\n\n{gender}")

                    st.markdown("---")

                    # Extract Module Code, Module Desc, and Credit
                    if 'Appraisal Type Code' in student_df.columns:
                        modules_df = student_df[
                            student_df['Appraisal Type Code'].astype(str).str.strip() == 'Overall Grade'
                        ][['Module Code', 'Module Desc', 'Credit']].drop_duplicates()
                    else:
                        modules_df = pd.DataFrame()

                    if modules_df.empty:
                        modules_df = student_df[['Module Code', 'Module Desc', 'Credit']].drop_duplicates()

                    st.write(style_preview(modules_df))
            else:
                st.info("No student records found in the uploaded file.")

        with tab5:
            st.subheader("All Raw Data Preview")
            st.dataframe(df, use_container_width=True)

        # Export Section
        st.markdown("---")
        st.subheader("📥 Export Gracing Report")

        if not graced_df.empty:
            excel_bytes = to_excel_bytes(graced_df, subject_summary, student_summary)
            date_str = datetime.now().strftime("%d-%m-%y")
            output_filename = f"Gracing_Analysis_Report_{date_str}.xlsx"

            st.download_button(
                label=f"📥 Download Complete Gracing Report ({output_filename})",
                data=excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No gracing data available to export.")

    except Exception as error:
        st.error(f"Something went wrong while processing the file: {error}")