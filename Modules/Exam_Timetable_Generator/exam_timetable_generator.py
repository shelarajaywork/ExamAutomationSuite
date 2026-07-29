import datetime
import re
import io
import networkx as nx
import pandas as pd
import streamlit as st

# ===================== HELPER FUNCTIONS =====================

def canonical_subject(s: str) -> str:
    """Normalises module descriptions for scheduling by removing common language/level suffixes."""
    if not isinstance(s, str):
        s = str(s) if pd.notna(s) else ""
    t = s.strip().upper()
    
    # Remove specific language variations
    t = t.replace(" IN ENGLISH", "").replace(" ENGLISH", "")
    
    # Trim trailing Roman numerals (I, II, III)
    while t.endswith(" I") or t.endswith(" II") or t.endswith(" III"):
        t = t.rsplit(" ", 1)[0]
        
    # Standardise multiple spaces
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def short_year(y) -> str:
    """Formats academic year string to short format (e.g. '2024-2025' -> '24-25')."""
    y_str = str(y).strip()
    if "-" in y_str:
        parts = y_str.split("-")
        if len(parts) == 2 and len(parts[0]) >= 2 and len(parts[1]) >= 2:
            return f"{parts[0][-2:]}-{parts[1][-2:]}"
    return y_str


def parse_duration_to_seconds(dur_val) -> int:
    """Parses duration string (HH:MM or HH:MM:SS) or timedelta object into seconds."""
    if pd.isna(dur_val) or str(dur_val).strip() == "":
        return 0
    
    dur_str = str(dur_val).strip()
    parts = dur_str.split(":")
    try:
        if len(parts) >= 2:
            hrs = int(parts[0])
            mins = int(parts[1])
            return hrs * 3600 + mins * 60
        else:
            return int(float(dur_str)) * 3600
    except ValueError:
        return 0


def format_duration(seconds: int) -> str:
    """Formats total seconds into 'HH:MM:SS' string format."""
    hrs = seconds // 3600
    mins = (seconds % 3600) // 60
    return f"{hrs:02d}:{mins:02d}:00"


def format_time_range(start_time_str: str, duration_seconds: int) -> str:
    """Calculates end time and returns a formatted range '10:00 AM - 12:00 PM'."""
    start_dt = datetime.datetime.strptime(start_time_str, "%I:%M %p")
    end_dt = start_dt + datetime.timedelta(seconds=duration_seconds)
    return f"{start_dt.strftime('%I:%M %p')} - {end_dt.strftime('%I:%M %p')}"


def map_colors_to_dates(color_dict: dict, start_date: datetime.date) -> dict:
    """Maps color/slot integers (1, 2, 3...) to actual date objects, automatically skipping Sundays."""
    max_slot = max(color_dict.values()) if color_dict else 0
    day_date_map = {}
    current_date = start_date
    slot = 1

    while slot <= max_slot:
        if current_date.weekday() != 6:  # 6 corresponds to Sunday in Python
            day_date_map[slot] = current_date
            slot += 1
        current_date += datetime.timedelta(days=1)
        
    return day_date_map


def generate_time_options():
    """Generates time options from 07:00 AM to 03:00 PM with 30-minute intervals."""
    times = []
    current = datetime.datetime.strptime("07:00 AM", "%I:%M %p")
    end = datetime.datetime.strptime("03:00 PM", "%I:%M %p")
    while current <= end:
        times.append(current.strftime("%I:%M %p"))
        current += datetime.timedelta(minutes=30)
    return times


# ===================== CORE SCHEDULING ALGORITHM =====================

def generate_timetable(df: pd.DataFrame, is_reexam: bool, start_date: datetime.date, start_time_str: str):
    # Standardise column headers to lowercase without spaces
    col_map = {str(col).strip().lower().replace(" ", ""): col for col in df.columns}

    # Map required column names
    c_sub = col_map.get("moduledescription", col_map.get("subjectname"))
    c_stud = col_map.get("studentnumber")
    c_prog = col_map.get("programname", col_map.get("programmename"))
    c_college = col_map.get("schoolname", col_map.get("collegename"))
    c_ay = col_map.get("currentacademicyear", col_map.get("academicyear"))
    c_sem = col_map.get("currentsession", col_map.get("semester"))
    c_dur = col_map.get("examduration", col_map.get("duration"))
    
    # Additional required columns
    c_prog_abbr = col_map.get("programabbreviation")
    c_mod_abbr = col_map.get("moduleabbreviation")
    c_credit = col_map.get("credit")
    c_student_na = col_map.get("studentna")

    # Validate required columns
    missing = []
    if not c_sub: missing.append("Module Description")
    if not c_stud: missing.append("Student Number")
    if not c_prog: missing.append("Program Name")
    if not c_college: missing.append("School Name")
    if not c_ay: missing.append("Current Academic Year")
    if not c_sem: missing.append("Current Session")
    if not c_dur: missing.append("Exam Duration")

    if missing:
        st.error(f"Missing required columns in uploaded Excel file(s): {', '.join(missing)}")
        return None

    # Filter strictly for 'WRIT' rows from 'Student na' column if present
    if c_student_na and c_student_na in df.columns:
        df = df[df[c_student_na].astype(str).str.strip().str.upper() == "WRIT"].copy()

    # Pre-process dataframe records
    df = df.dropna(subset=[c_sub, c_stud]).copy()
    
    subj_data = {}         
    subj_sem = {}          
    subj_dur = {}          
    student_to_nodes = {}  

    re_row_students = {}   
    re_row_ays = {}        

    for _, row in df.iterrows():
        subject = str(row[c_sub]).strip()
        stud = str(row[c_stud]).strip()
        if not subject or not stud:
            continue

        canon_sub = canonical_subject(subject)
        sem = str(row[c_sem]).strip()
        ay = short_year(row[c_ay])
        dur_txt = str(row[c_dur]).strip()
        dur_sec = parse_duration_to_seconds(dur_txt)
        dur_fmt = format_duration(dur_sec)
        college = str(row[c_college]).strip()
        prog = str(row[c_prog]).strip()

        prog_abbr = str(row[c_prog_abbr]).strip() if c_prog_abbr and pd.notna(row[c_prog_abbr]) else ""
        mod_abbr = str(row[c_mod_abbr]).strip() if c_mod_abbr and pd.notna(row[c_mod_abbr]) else ""
        credit_val = row[c_credit] if c_credit and pd.notna(row[c_credit]) else ""

        node_key = f"{canon_sub}|{sem}|{dur_txt}"

        if node_key not in subj_data:
            subj_data[node_key] = {}
            subj_sem[node_key] = sem
            subj_dur[node_key] = dur_fmt

        # Structuring data by exam mode
        if not is_reexam:
            row_key = f"{prog_abbr}|{prog}|{ay}|{mod_abbr}|{subject}|{credit_val}"
            if college not in subj_data[node_key]:
                subj_data[node_key][college] = {}
            if row_key not in subj_data[node_key][college]:
                subj_data[node_key][college][row_key] = set()
            subj_data[node_key][college][row_key].add(stud)
        else:
            row_key = f"{college}|{prog_abbr}|{prog}|{sem}|{mod_abbr}|{subject}|{credit_val}|{dur_txt}"
            if row_key not in re_row_students:
                re_row_students[row_key] = set()
                re_row_ays[row_key] = set()
            re_row_students[row_key].add(stud)
            re_row_ays[row_key].add(ay)

        if stud not in student_to_nodes:
            student_to_nodes[stud] = set()
        student_to_nodes[stud].add(node_key)

    # Build Graph for DSATUR Clash Resolution
    G = nx.Graph()
    for node_key in subj_data.keys():
        if not is_reexam:
            total_students = sum(
                len(studs)
                for col in subj_data[node_key].values()
                for studs in col.values()
            )
        else:
            total_students = sum(
                len(re_row_students[rk])
                for rk in re_row_students
                if node_key in rk
            )
        G.add_node(node_key, weight=total_students)

    for stud, nodes in student_to_nodes.items():
        node_list = list(nodes)
        for i in range(len(node_list)):
            for j in range(i + 1, len(node_list)):
                G.add_edge(node_list[i], node_list[j])

    # Perform DSATUR Coloring
    color_assignment = nx.coloring.greedy_color(G, strategy="DSATUR")
    color_assignment = {node: c + 1 for node, c in color_assignment.items()}

    # Map slot integers to actual date objects
    day_date_map = map_colors_to_dates(color_assignment, start_date)

    # Build Final Output Dataframe
    output_rows = []

    if is_reexam:
        for row_key, stud_set in re_row_students.items():
            parts = row_key.split("|")
            coll_val, p_abbr, prog_val, sem_val, m_abbr, sub_val, cred_val, dur_val = (
                parts[0], parts[1], parts[2], parts[3], parts[4], parts[5], parts[6], parts[7]
            )
            
            node_key = f"{canonical_subject(sub_val)}|{sem_val}|{dur_val}"
            assigned_slot = color_assignment.get(node_key, 1)
            dur_sec = parse_duration_to_seconds(dur_val)

            output_rows.append({
                "School Name": coll_val,
                "Program Abbreviation": p_abbr,
                "Program Name": prog_val,
                "Current Session": sem_val,
                "Module Abbreviation": m_abbr,
                "Module Description": sub_val,
                "Current Academic Year": ", ".join(sorted(re_row_ays[row_key])),
                "Credit": cred_val,
                "Students": len(stud_set),
                "Exam Duration": format_duration(dur_sec),
                "Exam Date": day_date_map.get(assigned_slot, ""),
                "Time": format_time_range(start_time_str, dur_sec),
                "Day": f"Day {assigned_slot}"
            })
    else:
        for node_key, colleges in subj_data.items():
            assigned_slot = color_assignment.get(node_key, 1)
            dur_sec = parse_duration_to_seconds(subj_dur[node_key])

            for college, row_keys in colleges.items():
                for row_key, stud_set in row_keys.items():
                    p = row_key.split("|")
                    p_abbr, prog_val, ay_val, m_abbr, sub_val, cred_val = (
                        p[0], p[1], p[2], p[3], p[4], p[5]
                    )
                    output_rows.append({
                        "School Name": college,
                        "Program Abbreviation": p_abbr,
                        "Program Name": prog_val,
                        "Current Session": subj_sem[node_key],
                        "Module Abbreviation": m_abbr,
                        "Module Description": sub_val,
                        "Current Academic Year": ay_val,
                        "Credit": cred_val,
                        "Students": len(stud_set),
                        "Exam Duration": subj_dur[node_key],
                        "Exam Date": day_date_map.get(assigned_slot, ""),
                        "Time": format_time_range(start_time_str, dur_sec),
                        "Day": f"Day {assigned_slot}"
                    })

    # Strict Output Column Order
    target_columns = [
        "School Name", "Program Abbreviation", "Program Name", "Current Session",
        "Module Abbreviation", "Module Description", "Current Academic Year",
        "Credit", "Students", "Exam Duration", "Exam Date", "Time", "Day"
    ]
    
    out_df = pd.DataFrame(output_rows)
    for col in target_columns:
        if col not in out_df.columns:
            out_df[col] = ""
            
    return out_df[target_columns]


# ===================== STREAMLIT USER INTERFACE =====================

def main():
    st.set_page_config(page_title="Exam Timetable Generator", layout="wide", page_icon="📅")
    
    st.title("📅 Exam Timetable Generator")
    st.markdown("Upload your exam registration data to generate clash-free timetables automatically.")
    st.divider()

    # --- CONFIGURATION PANEL IN MAIN WINDOW ---
    with st.container():
        col1, col2, col3 = st.columns(3)

        with col1:
            exam_type = st.radio(
                "Select Examination Type:",
                ("Regular Examination", "Re-Examination")
            )
            is_reexam = (exam_type == "Re-Examination")

        with col2:
            exam_start_date = st.date_input(
                "Exam Start Date:",
                value=datetime.date.today(),
                format="DD-MM-YYYY",
                help="Select the starting date for the examination schedule."
            )

        with col3:
            time_options = generate_time_options()
            default_index = time_options.index("10:00 AM") if "10:00 AM" in time_options else 0
            
            exam_start_time_str = st.selectbox(
                "Exam Start Time:",
                options=time_options,
                index=default_index,
                help="Select standard start time between 07:00 AM and 03:00 PM (30 min intervals)."
            )

    st.divider()

    # --- CONDITIONAL UPLOAD SECTIONS ---
    if not is_reexam:
        st.info("Upload All Subject Student Data from ZACAD_REPORT.")
        uploaded_files = st.file_uploader(
            "Upload All Subjects Students Data files [ZACAD_REPORT]. (.xlsx, .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="regular_uploader"
        )
    else:
        st.subheader("Re-Examination Data Upload")
        uploaded_files = st.file_uploader(
            "Upload Re-Examination Applications Lists Excel Files [ZREEXAM_REPORT] (.xlsx, .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="reexam_uploader"
        )

    if uploaded_files:
        try:
            # Combine all uploaded files into a single DataFrame
            df_list = [pd.read_excel(file) for file in uploaded_files]
            df_input = pd.concat(df_list, ignore_index=True)
            
            st.success(f"Successfully loaded {len(uploaded_files)} file(s) with a total of {len(df_input)} rows.")

            with st.expander("Preview Combined Data"):
                st.dataframe(df_input.head(10))

            if st.button("🚀 Generate Timetable", type="primary"):
                with st.spinner("Building conflict graph and calculating schedules..."):
                    result_df = generate_timetable(
                        df=df_input,
                        is_reexam=is_reexam,
                        start_date=exam_start_date,
                        start_time_str=exam_start_time_str
                    )

                if result_df is not None and not result_df.empty:
                    st.success("Timetable generated successfully!")
                    
                    st.subheader(f"Generated Schedule ({exam_type})")
                    st.dataframe(result_df, use_container_width=True)

                    # Export to Excel with custom formatting
                    output_bytes = io.BytesIO()
                    with pd.ExcelWriter(output_bytes, engine="openpyxl") as writer:
                        sheet_name = "Re-Examination" if is_reexam else "Regular Examination"
                        result_df.to_excel(writer, index=False, sheet_name=sheet_name)
                        
                        workbook = writer.book
                        worksheet = writer.sheets[sheet_name]
                        
                        # Freeze top header row
                        worksheet.freeze_panes = "A2"
                        
                        # Style headers: Bold font + Light Grey Fill (#D3D3D3)
                        from openpyxl.styles import Font, PatternFill
                        
                        header_font = Font(bold=True)
                        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
                        
                        for col_num in range(1, len(result_df.columns) + 1):
                            cell = worksheet.cell(row=1, column=col_num)
                            cell.font = header_font
                            cell.fill = header_fill

                        # Apply Short Date format (DD-MM-YYYY) to 'Exam Date' column
                        date_col_idx = result_df.columns.get_loc("Exam Date") + 1
                        for row_num in range(2, len(result_df) + 2):
                            cell = worksheet.cell(row=row_num, column=date_col_idx)
                            if cell.value:
                                cell.number_format = "DD-MM-YYYY"

                    # Dynamic Output File Name
                    curr_date_str = datetime.date.today().strftime("%d-%m-%Y")
                    if not is_reexam:
                        file_name = f"Regular Examination Timetable {curr_date_str}.xlsx"
                    else:
                        file_name = f"Re-examination Timetable {curr_date_str}.xlsx"
                    
                    st.download_button(
                        label="📥 Download Timetable Excel",
                        data=output_bytes.getvalue(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error processing files: {e}")

if __name__ == "__main__":
    main()