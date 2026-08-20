import os
import re
import glob
import ssl
import shutil
import tempfile
import gdown
import datetime
import urllib3
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Disable SSL verification for internal network proxies/firewalls
ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration
GDRIVE_PARENT_FOLDER_ID = "1W4VVj16v-JAJDJIiUCxx2EAXDETfVqEq"
LOCAL_DOWNLOAD_DIR = os.path.join(tempfile.gettempdir(), "osm_download_files")


def get_college_abbr(college_full: str) -> str:
    """Intelligently map full college name to standardized short code."""
    c = str(college_full).upper()
    if "NARSEE MONJEE" in c or "NMCCE" in c or "NMC" in c:
        return "NMC"
    elif "USHA PRAVIN GANDHI" in c or "UPG" in c:
        return "UPG"
    elif "MITHIBAI" in c or "MBC" in c:
        return "MBC"
    elif "D.J. SANGHVI" in c or "DJSCE" in c or "DJ" in c:
        return "DJSCE"
    elif "BHAGUBHAI" in c or "BNCP" in c:
        return "BNCP"
    elif "JITENDRA CHAUHAN" in c or "JCL" in c:
        return "JCL"
    elif "PRAVIN GANDHI" in c or "PGCL" in c:
        return "PGCL"
    else:
        cleaned = re.sub(r"SVKM['’]S|COLLEGE|OF|COMMERCE|ECONOMICS|ARTS|SCIENCE", "", c)
        words = re.findall(r'[A-Z0-9]+', cleaned)
        if words:
            abbr = "".join([w[0] for w in words])
            return abbr[:5] if len(abbr) > 1 else college_full[:10]
        return college_full[:10]


def extract_college_name(filename: str) -> tuple:
    """Extract full college name and short abbreviation from filename."""
    base = os.path.basename(filename)
    match = re.split(r'Evaluation Dashboard Details', base, flags=re.IGNORECASE)
    if match and match[0].strip():
        college_raw = match[0].strip()
        college_raw = re.sub(r'[\-_()]+$', '', college_raw).strip()
    else:
        college_raw = "SVKM Institute"
    
    college_abbr = get_college_abbr(college_raw)
    return college_raw, college_abbr


@st.cache_data(ttl=120)  # Auto-refreshes every 2 minutes
def load_and_clean_osm_data(parent_folder_id: str):
    """
    Cleans local download cache, downloads fresh files from Google Drive,
    and cleans them per business rules.
    """
    if os.path.exists(LOCAL_DOWNLOAD_DIR):
        try:
            shutil.rmtree(LOCAL_DOWNLOAD_DIR, ignore_errors=True)
        except Exception:
            pass
            
    os.makedirs(LOCAL_DOWNLOAD_DIR, exist_ok=True)
    folder_url = f"https://drive.google.com/drive/folders/{parent_folder_id}"
    
    try:
        gdown.download_folder(
            url=folder_url, 
            output=LOCAL_DOWNLOAD_DIR, 
            quiet=True,
            verify=False
        )
    except Exception as e:
        st.error(f"Error downloading data from Google Drive: {e}")
        return None, 0

    excel_files = glob.glob(os.path.join(LOCAL_DOWNLOAD_DIR, "**", "*.xlsx"), recursive=True) + \
                  glob.glob(os.path.join(LOCAL_DOWNLOAD_DIR, "**", "*.xls"), recursive=True)

    if not excel_files:
        return None, 0

    all_dfs = []
    file_count = 0

    for file in excel_files:
        file_path_norm = os.path.normpath(file).lower()
        if "osm dashboard" in file_path_norm or "osm" in file_path_norm or "evaluation" in file_path_norm:
            try:
                raw_df = pd.read_excel(file)
                
                # 1. Drop completely blank rows
                raw_df = raw_df.dropna(how='all')
                
                # 2. Exclude rows containing "Count:" in any column
                mask_count = raw_df.astype(str).apply(
                    lambda row: row.str.contains(r'Count\s*:', case=False, na=False).any(), 
                    axis=1
                )
                clean_df = raw_df[~mask_count].copy()
                
                # 3. Rename "Semester/Trimester" to "Semester"
                if 'Semester/Trimester' in clean_df.columns:
                    clean_df = clean_df.rename(columns={'Semester/Trimester': 'Semester'})
                
                # 4. Ignore CampusName, StreamName, and UploadCount
                drop_cols = [c for c in ['CampusName', 'StreamName', 'UploadCount'] if c in clean_df.columns]
                clean_df = clean_df.drop(columns=drop_cols)
                
                # 5. Extract College information
                full_col, short_col = extract_college_name(file)
                clean_df['College_Full'] = full_col
                clean_df['College'] = short_col
                clean_df['Source_File'] = os.path.basename(file)
                
                # 6. Derive Exam Type
                clean_df['Exam_Type'] = clean_df['CategoryName'].astype(str).apply(
                    lambda x: "Re-examination" if "[re-exam]" in x.lower() or "re-exam" in x.lower() else "Regular Examination"
                )
                
                all_dfs.append(clean_df)
                file_count += 1
            except Exception:
                pass

    if not all_dfs:
        return None, 0

    merged_df = pd.concat(all_dfs, ignore_index=True)

    # Standardize missing strings
    merged_df['RoleName'] = merged_df['RoleName'].fillna('Unassigned')
    merged_df['CourseName'] = merged_df['CourseName'].fillna('Unknown Programme')
    merged_df['Semester'] = merged_df['Semester'].fillna('General')
    merged_df['CategoryName'] = merged_df['CategoryName'].fillna('Unspecified Module')
    merged_df['ExaminerName'] = merged_df['ExaminerName'].fillna('Unassigned Examiner')

    # Convert numeric fields
    num_cols = ['AbsentCount', 'PresentCount', 'CheckCount', 'InprogressCount', 'RejectCount', 'UnCheckCount']
    for c in num_cols:
        if c in merged_df.columns:
            merged_df[c] = pd.to_numeric(merged_df[c], errors='coerce').fillna(0).astype(int)
        else:
            merged_df[c] = 0

    # Parse and validate dates strictly as dd/mm/yyyy
    merged_df['ExamDate_Parsed'] = pd.to_datetime(merged_df['ExamDate'], format='%d/%m/%Y', errors='coerce')
    merged_df['AssignedDate_Parsed'] = pd.to_datetime(merged_df['AssignedDateTime'], format='%d/%m/%Y', errors='coerce')
    merged_df['EvaluationLastDate_Parsed'] = pd.to_datetime(merged_df['EvaluationLastDate'], format='%d/%m/%Y', errors='coerce')

    # Re-format raw date strings to uniform dd/mm/yyyy
    merged_df['ExamDate_Display'] = merged_df['ExamDate_Parsed'].dt.strftime('%d/%m/%Y').fillna('-')
    merged_df['AssignedDate_Display'] = merged_df['AssignedDate_Parsed'].dt.strftime('%d/%m/%Y').fillna('-')
    merged_df['EvaluationLastDate_Display'] = merged_df['EvaluationLastDate_Parsed'].dt.strftime('%d/%m/%Y').fillna('-')

    # Evaluation window calculation
    merged_df['Eval_Window_Days'] = (merged_df['EvaluationLastDate_Parsed'] - merged_df['AssignedDate_Parsed']).dt.days
    merged_df['Eval_Window_Days'] = merged_df['Eval_Window_Days'].apply(lambda x: x if x >= 0 else np.nan)

    merged_df['Is_Deadline_Extended'] = merged_df['OldEvaluationLastDate'].apply(
        lambda x: "Yes" if str(x).strip() not in ['-', 'nan', 'None', ''] else "No"
    )

    return merged_df, file_count


def highlight_pending_evaluations(row):
    """
    Applies custom row background styling based on EvaluationLastDate for pending items:
    - Red: EvaluationLastDate has passed
    - Soft Yellow: EvaluationLastDate is today
    - Soft Blue: EvaluationLastDate is tomorrow
    """
    today = datetime.date.today()
    tomorrow = today + datetime.timedelta(days=1)
    
    last_date = row.get('EvaluationLastDate_Parsed', None)
    if pd.isna(last_date) or last_date is None:
        return [''] * len(row)
    
    last_date_val = last_date.date() if isinstance(last_date, pd.Timestamp) else last_date
    
    if last_date_val < today:
        return ['background-color: rgba(239, 68, 68, 0.22); color: inherit; font-weight: 500;'] * len(row)
    elif last_date_val == today:
        return ['background-color: rgba(234, 179, 8, 0.22); color: inherit; font-weight: 500;'] * len(row)
    elif last_date_val == tomorrow:
        return ['background-color: rgba(59, 130, 246, 0.22); color: inherit; font-weight: 500;'] * len(row)
    
    return [''] * len(row)


def render_kpi_card(title, value, subtext="", card_mode="normal"):
    """Render an equal-sized, theme-adaptive KPI Card."""
    if card_mode == "red":
        border_color = "#ef4444"
        bg_color = "rgba(239, 68, 68, 0.12)"
        val_color = "#ef4444"
    elif card_mode == "green":
        border_color = "#10b981"
        bg_color = "rgba(16, 185, 129, 0.12)"
        val_color = "#10b981"
    else:
        border_color = "rgba(128, 128, 128, 0.22)"
        bg_color = "var(--secondary-background-color, rgba(128, 128, 128, 0.05))"
        val_color = "var(--text-color, #1e293b)"

    html_code = f"""
    <div style="
        background-color: {bg_color};
        border: 1.5px solid {border_color};
        border-radius: 12px;
        padding: 12px 14px;
        height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-shadow: 0 2px 5px rgba(0,0,0,0.04);
        box-sizing: border-box;
    ">
        <div style="font-size: 0.78rem; font-weight: 600; color: var(--text-color, #64748b); opacity: 0.85; line-height: 1.2;">
            {title}
        </div>
        <div style="font-size: 1.45rem; font-weight: 800; color: {val_color}; line-height: 1.1; margin: 2px 0;">
            {value}
        </div>
        <div style="font-size: 0.72rem; font-weight: 500; color: {val_color}; opacity: 0.95; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
            {subtext}
        </div>
    </div>
    """
    st.markdown(html_code, unsafe_allow_html=True)


def show():
    """Main render function for OSM Examination Analytics Dashboard."""
    
    # Theme-aware CSS
    st.markdown("""
    <style>
        .portal-header {
            font-size: 1.95rem;
            font-weight: 800;
            color: var(--text-color, #1e3a8a);
            margin-bottom: 0.1rem;
            line-height: 1.2;
        }
        .portal-sub {
            font-size: 0.88rem;
            color: var(--text-color, #64748b);
            opacity: 0.8;
            margin-bottom: 0.8rem;
        }
        .filter-box {
            background-color: var(--secondary-background-color, rgba(128, 128, 128, 0.05));
            border: 1px solid rgba(128, 128, 128, 0.2);
            border-radius: 12px;
            padding: 14px 18px 10px 18px;
            margin-bottom: 18px;
        }
        div[data-testid="stTabs"] > div:first-child {
            position: sticky;
            top: 0;
            background-color: var(--background-color, #ffffff);
            z-index: 99;
            padding-top: 6px;
            padding-bottom: 6px;
            border-bottom: 2px solid rgba(128, 128, 128, 0.2);
        }
    </style>
    """, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # COMPACT HEADER: TITLE & ICON-ONLY BUTTONS ON SAME LINE
    # -------------------------------------------------------------------------
    h_col1, h_col2, h_col3 = st.columns([8.6, 0.7, 0.7])
    
    with h_col1:
        st.markdown('<div class="portal-header">🎓 Online Screen Marking (OSM) Analytics</div>', unsafe_allow_html=True)
        st.markdown('<div class="portal-sub">Executive Insights, Workload Analytics, SLA Deadlines & Rejection Management</div>', unsafe_allow_html=True)
    with h_col2:
        if st.button("🔄", help="Sync Fresh Data from Google Drive", key="btn_top_sync_fresh", use_container_width=True):
            st.cache_data.clear()
            if os.path.exists(LOCAL_DOWNLOAD_DIR):
                shutil.rmtree(LOCAL_DOWNLOAD_DIR, ignore_errors=True)
            st.rerun()
    with h_col3:
        if st.button("🏠", help="Return to Main Dashboard", key="btn_top_home_dash", use_container_width=True):
            st.session_state.selected_tool = "🏠 Dashboard"
            st.rerun()

    # Load dataset
    with st.spinner("Syncing and processing evaluation data from Google Drive..."):
        df_raw, file_count = load_and_clean_osm_data(GDRIVE_PARENT_FOLDER_ID)

    if df_raw is None or len(df_raw) == 0:
        st.warning("⚠️ No valid evaluation records found in the Google Drive repository. Click 🔄 above to refresh.")
        return

    # -------------------------------------------------------------------------
    # RESPONSIVE CASCADING / INTERDEPENDENT FILTER ENGINE
    # -------------------------------------------------------------------------
    st.markdown('<div class="filter-box">', unsafe_allow_html=True)
    f_head_col1, f_head_col2 = st.columns([8.5, 1.5])
    with f_head_col1:
        st.markdown("##### 🔍 Responsive Cascading Filters")
    with f_head_col2:
        if st.button("Clear All Filters", use_container_width=True, key="btn_clear_osm_filters"):
            for k in list(st.session_state.keys()):
                if k.startswith("osm_filter_"):
                    del st.session_state[k]
            st.rerun()

    # Current filter state values
    cur_college = st.session_state.get("osm_filter_college", "All")
    cur_ay = st.session_state.get("osm_filter_ay", "All")
    cur_exam_type = st.session_state.get("osm_filter_exam_type", "All")
    cur_sem = st.session_state.get("osm_filter_sem", "All")
    cur_role = st.session_state.get("osm_filter_role", "All")
    cur_prog = st.session_state.get("osm_filter_prog", "All")
    cur_mod = st.session_state.get("osm_filter_module", "All")
    cur_start_d = st.session_state.get("osm_filter_start_date", None)
    cur_end_d = st.session_state.get("osm_filter_end_date", None)

    active_filters = {
        'College': cur_college,
        'AYear': cur_ay,
        'Exam_Type': cur_exam_type,
        'Semester': cur_sem,
        'RoleName': cur_role,
        'CourseName': cur_prog,
        'CategoryName': cur_mod,
        'date_range': (cur_start_d, cur_end_d) if (cur_start_d and cur_end_d) else None
    }

    # Cross-filtering helper function
    def get_cascaded_subset(exclude_field=None):
        temp = df_raw.copy()
        for field, val in active_filters.items():
            if field == exclude_field or val == "All" or val is None:
                continue
            if field == 'date_range':
                s_d, e_d = val
                if s_d and e_d:
                    temp = temp[
                        (temp['ExamDate_Parsed'].dt.date >= s_d) &
                        (temp['ExamDate_Parsed'].dt.date <= e_d)
                    ]
            elif field == 'AYear':
                temp = temp[temp['AYear'].astype(str).str.startswith(str(val))]
            elif field in temp.columns:
                temp = temp[temp[field].astype(str) == str(val)]
        return temp

    # Dynamically compute options based on other active selections
    df_for_col = get_cascaded_subset('College')
    college_opts = ["All"] + sorted(df_for_col['College'].dropna().unique().tolist())
    if cur_college not in college_opts:
        st.session_state["osm_filter_college"] = "All"

    df_for_ay = get_cascaded_subset('AYear')
    ay_vals = [str(int(x)) for x in df_for_ay['AYear'].dropna().unique() if not np.isnan(x)]
    ay_opts = ["All"] + sorted(ay_vals)
    if cur_ay not in ay_opts:
        st.session_state["osm_filter_ay"] = "All"

    df_for_et = get_cascaded_subset('Exam_Type')
    exam_type_opts = ["All"] + sorted(df_for_et['Exam_Type'].dropna().unique().tolist())
    if cur_exam_type not in exam_type_opts:
        st.session_state["osm_filter_exam_type"] = "All"

    df_for_sem = get_cascaded_subset('Semester')
    sem_opts = ["All"] + sorted(df_for_sem['Semester'].dropna().astype(str).unique().tolist())
    if cur_sem not in sem_opts:
        st.session_state["osm_filter_sem"] = "All"

    df_for_role = get_cascaded_subset('RoleName')
    role_opts = ["All"] + sorted(df_for_role['RoleName'].dropna().astype(str).unique().tolist())
    if cur_role not in role_opts:
        st.session_state["osm_filter_role"] = "All"

    df_for_prog = get_cascaded_subset('CourseName')
    prog_opts = ["All"] + sorted(df_for_prog['CourseName'].dropna().astype(str).unique().tolist())
    if cur_prog not in prog_opts:
        st.session_state["osm_filter_prog"] = "All"

    df_for_mod = get_cascaded_subset('CategoryName')
    mod_opts = ["All"] + sorted(df_for_mod['CategoryName'].dropna().astype(str).unique().tolist())
    if cur_mod not in mod_opts:
        st.session_state["osm_filter_module"] = "All"

    # Filter Row 1
    r1_1, r1_2, r1_3, r1_4 = st.columns(4)
    with r1_1:
        idx_col = college_opts.index(st.session_state.get("osm_filter_college", "All"))
        sel_college = st.selectbox("College", college_opts, index=idx_col, key="osm_filter_college")

    with r1_2:
        idx_ay = ay_opts.index(st.session_state.get("osm_filter_ay", "All"))
        sel_ay = st.selectbox("Academic Year (AYear)", ay_opts, index=idx_ay, key="osm_filter_ay")

    with r1_3:
        idx_et = exam_type_opts.index(st.session_state.get("osm_filter_exam_type", "All"))
        sel_exam_type = st.selectbox("Exam Type", exam_type_opts, index=idx_et, key="osm_filter_exam_type")

    with r1_4:
        idx_sem = sem_opts.index(st.session_state.get("osm_filter_sem", "All"))
        sel_sem = st.selectbox("Semester", sem_opts, index=idx_sem, key="osm_filter_sem")

    # Filter Row 2
    r2_1, r2_2, r2_3, r2_4 = st.columns(4)
    with r2_1:
        idx_role = role_opts.index(st.session_state.get("osm_filter_role", "All"))
        sel_role = st.selectbox("Evaluated By (Role)", role_opts, index=idx_role, key="osm_filter_role")

    with r2_2:
        idx_prog = prog_opts.index(st.session_state.get("osm_filter_prog", "All"))
        sel_programme = st.selectbox("Programme", prog_opts, index=idx_prog, key="osm_filter_prog")

    with r2_3:
        idx_mod = mod_opts.index(st.session_state.get("osm_filter_module", "All"))
        sel_module = st.selectbox("Module (Subject)", mod_opts, index=idx_mod, key="osm_filter_module")

    with r2_4:
        min_exam_d = df_raw['ExamDate_Parsed'].min()
        max_exam_d = df_raw['ExamDate_Parsed'].max()
        start_d_default = min_exam_d.date() if pd.notnull(min_exam_d) else datetime.date.today()
        end_d_default = max_exam_d.date() if pd.notnull(max_exam_d) else datetime.date.today()

        d_sub1, d_sub2 = st.columns(2)
        with d_sub1:
            sel_start_date = st.date_input("Start Exam Date", value=start_d_default, format="DD/MM/YYYY", key="osm_filter_start_date")
        with d_sub2:
            sel_end_date = st.date_input("End Exam Date", value=end_d_default, format="DD/MM/YYYY", key="osm_filter_end_date")

    st.markdown('</div>', unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # COMPUTE FINAL FILTERED DATASET
    # -------------------------------------------------------------------------
    df_filtered = df_raw.copy()

    if sel_college != "All":
        df_filtered = df_filtered[df_filtered['College'] == sel_college]
    if sel_ay != "All":
        df_filtered = df_filtered[df_filtered['AYear'].astype(str).str.startswith(sel_ay)]
    if sel_exam_type != "All":
        df_filtered = df_filtered[df_filtered['Exam_Type'] == sel_exam_type]
    if sel_sem != "All":
        df_filtered = df_filtered[df_filtered['Semester'] == sel_sem]
    if sel_role != "All":
        df_filtered = df_filtered[df_filtered['RoleName'] == sel_role]
    if sel_programme != "All":
        df_filtered = df_filtered[df_filtered['CourseName'] == sel_programme]
    if sel_module != "All":
        df_filtered = df_filtered[df_filtered['CategoryName'] == sel_module]
    if sel_start_date and sel_end_date:
        df_filtered = df_filtered[
            (df_filtered['ExamDate_Parsed'].dt.date >= sel_start_date) &
            (df_filtered['ExamDate_Parsed'].dt.date <= sel_end_date)
        ]

    if len(df_filtered) == 0:
        st.warning("⚠️ No records match the selected filter combination. Click **Clear All Filters** above to reset.")
        return

    # -------------------------------------------------------------------------
    # REAL-TIME DYNAMIC KPI CALCULATIONS (BOUND TO df_filtered)
    # -------------------------------------------------------------------------
    tot_present = int(df_filtered['PresentCount'].sum())
    tot_absent = int(df_filtered['AbsentCount'].sum())
    tot_checked = int(df_filtered['CheckCount'].sum())
    tot_inprogress = int(df_filtered['InprogressCount'].sum())
    tot_unchecked = int(df_filtered['UnCheckCount'].sum())
    tot_rejected = int(df_filtered['RejectCount'].sum())
    tot_answerbooks = tot_present

    eval_progress_pct = (tot_checked / tot_answerbooks * 100) if tot_answerbooks > 0 else 0.0
    attendance_pct = (tot_present / (tot_present + tot_absent) * 100) if (tot_present + tot_absent) > 0 else 0.0

    today = datetime.date.today()
    overdue_records = df_filtered[
        (df_filtered['UnCheckCount'] > 0) & 
        (df_filtered['EvaluationLastDate_Parsed'].dt.date < today)
    ]
    tot_overdue_books = int(overdue_records['UnCheckCount'].sum())

    # -------------------------------------------------------------------------
    # EQUAL-SIZED 7-COLUMN KPI CARD SCORECARD
    # -------------------------------------------------------------------------
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    
    with k1:
        render_kpi_card("Answerbooks Total", f"{tot_answerbooks:,}", f"{len(df_filtered)} Batches", "normal")
    with k2:
        render_kpi_card("Evaluated (Checked)", f"{tot_checked:,}", f"📈 {eval_progress_pct:.1f}% Completed", "normal")
    with k3:
        render_kpi_card("In-Progress", f"{tot_inprogress:,}", "⏳ Under Evaluation", "normal")
    with k4:
        render_kpi_card("Pending (Unchecked)", f"{tot_unchecked:,}", "⚠️ Awaiting Check", "normal")
    with k5:
        if tot_rejected > 0:
            render_kpi_card("Rejected Answerbooks", f"{tot_rejected:,}", "🚨 Action Required", "red")
        else:
            render_kpi_card("Rejected Answerbooks", "0", "✅ 0 Rejections", "green")
    with k6:
        render_kpi_card("Overdue Papers", f"{tot_overdue_books:,}", "⏰ Passed Deadline" if tot_overdue_books > 0 else "✅ SLA On Track", "normal")
    with k7:
        render_kpi_card("Student Attendance", f"{attendance_pct:.1f}%", f"{tot_absent:,} Absentees", "normal")

    st.markdown("<hr style='margin-top: 0.8rem; margin-bottom: 1.2rem; border-color: rgba(128,128,128,0.2);'>", unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # NAVIGATION TABS
    # -------------------------------------------------------------------------
    tab_summary, tab_prog, tab_module, tab_college, tab_workload, tab_tat, tab_preview = st.tabs([
        "📊 Executive Overview",
        "🎓 Programme Analytics",
        "📚 Module Analytics",
        "🏛️ College Comparisons",
        "👨‍🏫 Evaluator Workload",
        "⏳ Turnaround & Deadlines",
        "📋 Interactive Previews"
    ])

    def apply_chart_theme(fig, height=360):
        fig.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            height=height,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        fig.update_xaxes(gridcolor='rgba(128,128,128,0.15)', zerolinecolor='rgba(128,128,128,0.15)')
        fig.update_yaxes(gridcolor='rgba(128,128,128,0.15)', zerolinecolor='rgba(128,128,128,0.15)')
        return fig

    # =========================================================================
    # TAB 1: EXECUTIVE OVERVIEW
    # =========================================================================
    with tab_summary:
        c1, c2 = st.columns([6, 4])
        
        with c1:
            st.markdown("##### 📈 Exam Type & Evaluation Status")
            exam_summary = df_filtered.groupby(['College', 'Exam_Type'])[['CheckCount', 'UnCheckCount']].sum().reset_index()
            fig_et = px.bar(
                exam_summary,
                x='College',
                y=['CheckCount', 'UnCheckCount'],
                color_discrete_map={'CheckCount': '#10b981', 'UnCheckCount': '#ef4444'},
                barmode='stack',
                title="Checked vs. Pending Answerbooks by College & Exam Type",
                facet_col='Exam_Type'
            )
            fig_et.update_layout(legend_title_text='Status')
            st.plotly_chart(apply_chart_theme(fig_et, 380), use_container_width=True)

        with c2:
            st.markdown("##### 🍩 Evaluation Status Split")
            status_df = pd.DataFrame({
                'Status': ['Checked', 'In-Progress', 'Pending (Unchecked)', 'Rejected'],
                'Count': [tot_checked, tot_inprogress, tot_unchecked, tot_rejected]
            })
            status_df = status_df[status_df['Count'] > 0]

            fig_donut = px.pie(
                status_df,
                names='Status',
                values='Count',
                hole=0.52,
                color='Status',
                color_discrete_map={
                    'Checked': '#10b981', 
                    'In-Progress': '#f59e0b', 
                    'Pending (Unchecked)': '#ef4444',
                    'Rejected': '#dc2626'
                }
            )
            fig_donut.update_traces(textposition='inside', textinfo='percent+label')
            fig_donut.update_layout(showlegend=False)
            st.plotly_chart(apply_chart_theme(fig_donut, 380), use_container_width=True)

        r2_1, r2_2 = st.columns(2)
        with r2_1:
            st.markdown("##### 👥 Evaluator Role Share")
            role_df = df_filtered.groupby('RoleName')[['PresentCount', 'CheckCount']].sum().reset_index()
            fig_role = px.pie(
                role_df,
                names='RoleName',
                values='PresentCount',
                color_discrete_sequence=px.colors.qualitative.Safe,
                hole=0.4
            )
            st.plotly_chart(apply_chart_theme(fig_role, 320), use_container_width=True)

        with r2_2:
            st.markdown("##### 📅 Examination Timeline Volume")
            date_trend = df_filtered.groupby('ExamDate_Parsed')['PresentCount'].sum().reset_index().dropna()
            fig_trend = px.area(
                date_trend,
                x='ExamDate_Parsed',
                y='PresentCount',
                labels={'ExamDate_Parsed': 'Exam Date (dd/mm/yyyy)', 'PresentCount': 'Answerbooks Scheduled'},
                color_discrete_sequence=['#3b82f6']
            )
            fig_trend.update_xaxes(tickformat="%d/%m/%Y")
            st.plotly_chart(apply_chart_theme(fig_trend, 320), use_container_width=True)

    # =========================================================================
    # TAB 2: PROGRAMME ANALYTICS
    # =========================================================================
    with tab_prog:
        st.markdown("##### 🎓 Programme Evaluation Analytics")
        prog_df = df_filtered.groupby(['College', 'CourseName', 'Semester'])[
            ['PresentCount', 'AbsentCount', 'CheckCount', 'InprogressCount', 'RejectCount', 'UnCheckCount']
        ].sum().reset_index()
        prog_df['Completion_%'] = (prog_df['CheckCount'] / prog_df['PresentCount'] * 100).fillna(0).round(1)

        fig_prog_bar = px.bar(
            prog_df.sort_values(by='PresentCount', ascending=False).head(15),
            x='PresentCount',
            y='CourseName',
            color='College',
            orientation='h',
            title="Top 15 Programmes by Volume",
            labels={'PresentCount': 'Total Answerbooks', 'CourseName': 'Programme'}
        )
        fig_prog_bar.update_layout(yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(apply_chart_theme(fig_prog_bar, 420), use_container_width=True)

        st.markdown("###### 📋 Programme Evaluation Summary")
        st.dataframe(prog_df.sort_values(by='PresentCount', ascending=False), use_container_width=True)

    # =========================================================================
    # TAB 3: MODULE ANALYTICS
    # =========================================================================
    with tab_module:
        st.markdown("##### 📚 Module / Subject Evaluation Tracking")
        mod_df = df_filtered.groupby(['CategoryName', 'CourseName', 'Semester', 'Exam_Type'])[
            ['PresentCount', 'CheckCount', 'RejectCount', 'UnCheckCount']
        ].sum().reset_index()
        mod_df['Completion_%'] = (mod_df['CheckCount'] / mod_df['PresentCount'] * 100).fillna(0).round(1)

        m1, m2 = st.columns([7, 3])
        with m1:
            fig_mod = px.bar(
                mod_df.sort_values(by='PresentCount', ascending=False).head(15),
                x='PresentCount',
                y='CategoryName',
                color='Completion_%',
                color_continuous_scale='Greens',
                orientation='h',
                title="Top 15 Modules by Volume",
                labels={'PresentCount': 'Answerbooks', 'CategoryName': 'Module'}
            )
            fig_mod.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(apply_chart_theme(fig_mod, 420), use_container_width=True)

        with m2:
            st.markdown("###### 📊 Module Completion Scorecard")
            tot_mods = len(mod_df)
            completed_mods = len(mod_df[mod_df['Completion_%'] == 100.0])
            pending_mods = tot_mods - completed_mods
            
            st.metric("Total Modules", tot_mods)
            st.metric("100% Completed", completed_mods, delta=f"{(completed_mods/tot_mods*100):.1f}%" if tot_mods > 0 else "0%")
            st.metric("Pending Evaluation", pending_mods, delta=f"-{pending_mods}" if pending_mods > 0 else "0", delta_color="inverse")

        st.markdown("###### 📋 Complete Module Tracking Table")
        st.dataframe(mod_df.sort_values(by='PresentCount', ascending=False), use_container_width=True)

    # =========================================================================
    # TAB 4: COLLEGE COMPARISONS
    # =========================================================================
    with tab_college:
        st.markdown("##### 🏛️ College Performance Benchmarks")
        col_comp = df_filtered.groupby(['College', 'College_Full'])[
            ['PresentCount', 'AbsentCount', 'CheckCount', 'RejectCount', 'UnCheckCount']
        ].sum().reset_index()
        col_comp['Attendance_%'] = (col_comp['PresentCount'] / (col_comp['PresentCount'] + col_comp['AbsentCount']) * 100).fillna(0).round(1)
        col_comp['Evaluation_%'] = (col_comp['CheckCount'] / col_comp['PresentCount'] * 100).fillna(0).round(1)

        cc1, cc2 = st.columns(2)
        with cc1:
            fig_col_vol = px.bar(
                col_comp,
                x='College',
                y='PresentCount',
                color='College',
                text='PresentCount',
                title="Answerbooks Evaluated by College",
                color_discrete_sequence=px.colors.qualitative.Bold
            )
            fig_col_vol.update_layout(showlegend=False)
            st.plotly_chart(apply_chart_theme(fig_col_vol, 360), use_container_width=True)

        with cc2:
            fig_col_pct = px.bar(
                col_comp,
                x='College',
                y='Evaluation_%',
                color='College',
                text='Evaluation_%',
                title="Evaluation Progress (%) by College",
                color_discrete_sequence=px.colors.qualitative.Vivid
            )
            fig_col_pct.update_layout(showlegend=False, yaxis_range=[0, 105])
            st.plotly_chart(apply_chart_theme(fig_col_pct, 360), use_container_width=True)

        st.markdown("###### 📋 College Performance Roster")
        st.dataframe(col_comp, use_container_width=True)

    # =========================================================================
    # TAB 5: EVALUATOR WORKLOAD
    # =========================================================================
    with tab_workload:
        st.markdown("##### 👨‍🏫 Evaluator Workload & Performance")
        ex_summary = df_filtered.groupby(['ExaminerName', 'College', 'RoleName', 'Mobile'])[
            ['PresentCount', 'CheckCount', 'InprogressCount', 'RejectCount', 'UnCheckCount']
        ].sum().reset_index()
        ex_summary['Completion_%'] = (ex_summary['CheckCount'] / ex_summary['PresentCount'] * 100).fillna(0).round(1)

        ew1, ew2 = st.columns([6, 4])
        with ew1:
            top_ex = ex_summary.sort_values(by='CheckCount', ascending=False).head(12)
            fig_top_ex = px.bar(
                top_ex,
                x='CheckCount',
                y='ExaminerName',
                color='College',
                orientation='h',
                title="Top 12 Evaluators by Completed Papers",
                labels={'CheckCount': 'Answerbooks Checked', 'ExaminerName': 'Evaluator'}
            )
            fig_top_ex.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(apply_chart_theme(fig_top_ex, 400), use_container_width=True)

        with ew2:
            fig_box = px.box(
                ex_summary,
                x='RoleName',
                y='PresentCount',
                color='RoleName',
                points="all",
                title="Assignment Spread per Evaluator",
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig_box.update_layout(showlegend=False)
            st.plotly_chart(apply_chart_theme(fig_box, 400), use_container_width=True)

        st.markdown("###### 📋 Evaluator Directory & Progress")
        st.dataframe(ex_summary.sort_values(by='PresentCount', ascending=False), use_container_width=True)

    # =========================================================================
    # TAB 6: TURNAROUND & DEADLINES
    # =========================================================================
    with tab_tat:
        st.markdown("##### ⏳ Turnaround Time (TAT) & Deadline Analytics")
        
        t1, t2 = st.columns([6, 4])
        with t1:
            tat_dist = df_filtered['Eval_Window_Days'].dropna().value_counts().reset_index()
            tat_dist.columns = ['Allocated_Days', 'Batch_Count']
            tat_dist = tat_dist.sort_values(by='Allocated_Days')

            fig_tat = px.bar(
                tat_dist,
                x='Allocated_Days',
                y='Batch_Count',
                title="Allocated Evaluation Window (Days from Assignment to Deadline)",
                labels={'Allocated_Days': 'Days Allocated', 'Batch_Count': 'Batches'},
                color_discrete_sequence=['#6366f1']
            )
            st.plotly_chart(apply_chart_theme(fig_tat, 360), use_container_width=True)

        with t2:
            ext_counts = df_filtered['Is_Deadline_Extended'].value_counts().reset_index()
            ext_counts.columns = ['Deadline_Extended', 'Count']
            fig_ext = px.pie(
                ext_counts,
                names='Deadline_Extended',
                values='Count',
                title="Batches with Extended Deadlines",
                color='Deadline_Extended',
                color_discrete_map={'Yes': '#ef4444', 'No': '#10b981'},
                hole=0.45
            )
            st.plotly_chart(apply_chart_theme(fig_ext, 360), use_container_width=True)

        st.markdown("###### 📆 Deadline Tracking Schedule (dd/mm/yyyy)")
        tat_sched = df_filtered.groupby(['ExamDate_Display', 'AssignedDate_Display', 'EvaluationLastDate_Display', 'Is_Deadline_Extended'])[
            ['PresentCount', 'CheckCount', 'RejectCount', 'UnCheckCount']
        ].sum().reset_index()
        tat_sched.columns = ['Exam Date', 'Assigned Date', 'Evaluation Deadline', 'Deadline Extended', 'Present', 'Checked', 'Rejected', 'Unchecked']
        st.dataframe(tat_sched, use_container_width=True)

    # =========================================================================
    # TAB 7: INTERACTIVE PREVIEWS & DRILLDOWNS
    # =========================================================================
    with tab_preview:
        st.markdown("### 📋 Interactive Evaluation Previews")

        # 1. PENDING EVALUATIONS PREVIEW
        df_pending = df_filtered[df_filtered['UnCheckCount'] > 0].copy()
        
        if len(df_pending) > 0:
            st.markdown("#### ⚠️ Pending Evaluations Preview (`UnCheckCount > 0`)")
            st.caption("🔴 **Red**: Overdue (Deadline Passed) | 🟡 **Yellow**: Due Today | 🔵 **Blue**: Due Tomorrow")

            cols_preview = [
                'College', 'CourseName', 'Semester', 'CategoryName', 'RoleName', 
                'ExaminerName', 'Mobile', 'ExamDate_Display', 'EvaluationLastDate_Display', 
                'PresentCount', 'CheckCount', 'InprogressCount', 'UnCheckCount', 'EvaluationLastDate_Parsed'
            ]
            display_pending = df_pending[[c for c in cols_preview if c in df_pending.columns]].sort_values(
                by='EvaluationLastDate_Parsed', ascending=True
            )
            
            styled_pending = display_pending.style.apply(highlight_pending_evaluations, axis=1)
            
            st.dataframe(
                styled_pending, 
                column_config={"EvaluationLastDate_Parsed": None},
                use_container_width=True
            )

            csv_pending = display_pending.drop(columns=['EvaluationLastDate_Parsed'], errors='ignore').to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Pending Evaluations (CSV)",
                data=csv_pending,
                file_name="osm_pending_evaluations.csv",
                mime="text/csv"
            )
            st.markdown("<hr style='margin: 1.5rem 0; border-color: rgba(128,128,128,0.2);'>", unsafe_allow_html=True)
        else:
            st.success("🎉 All answerbooks under the current selection are 100% evaluated (`UnCheckCount = 0`).")
            st.markdown("<hr style='margin: 1.5rem 0; border-color: rgba(128,128,128,0.2);'>", unsafe_allow_html=True)

        # 2. REJECTED ANSWERBOOKS PREVIEW
        df_rejected = df_filtered[df_filtered['RejectCount'] > 0].copy()
        
        if len(df_rejected) > 0:
            st.markdown("#### 🚫 Rejected Answerbooks Preview (`RejectCount > 0`)")
            st.caption("Batches where answerbooks were flagged or rejected during evaluation/moderation.")

            cols_rej = [
                'College', 'CourseName', 'Semester', 'CategoryName', 'RoleName', 
                'ExaminerName', 'ExaminerEmail', 'Mobile', 'ExamDate_Display', 
                'PresentCount', 'RejectCount', 'CheckCount'
            ]
            display_rejected = df_rejected[[c for c in cols_rej if c in df_rejected.columns]].sort_values(
                by='RejectCount', ascending=False
            )
            
            st.dataframe(display_rejected, use_container_width=True)

            csv_rej = display_rejected.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Rejected Records (CSV)",
                data=csv_rej,
                file_name="osm_rejected_answerbooks.csv",
                mime="text/csv"
            )
            st.markdown("<hr style='margin: 1.5rem 0; border-color: rgba(128,128,128,0.2);'>", unsafe_allow_html=True)
        else:
            st.info("✅ No rejected answerbooks found (`RejectCount = 0`) under current selection.")
            st.markdown("<hr style='margin: 1.5rem 0; border-color: rgba(128,128,128,0.2);'>", unsafe_allow_html=True)

        # 3. COMPLETED EVALUATIONS PREVIEW
        st.markdown("#### ✅ Completed Evaluations Preview (`UnCheckCount == 0`)")
        df_completed = df_filtered[df_filtered['UnCheckCount'] == 0].copy()
        
        search_query = st.text_input("🔍 Search Completed Batches:", "")
        
        cols_comp = [
            'College', 'CourseName', 'Semester', 'CategoryName', 'RoleName', 
            'ExaminerName', 'ExamDate_Display', 'EvaluationLastDate_Display', 
            'PresentCount', 'CheckCount', 'Exam_Type'
        ]
        display_comp = df_completed[[c for c in cols_comp if c in df_completed.columns]]

        if search_query:
            mask = display_comp.astype(str).apply(
                lambda row: row.str.contains(search_query, case=False)
            ).any(axis=1)
            display_comp = display_comp[mask]

        st.dataframe(display_comp, use_container_width=True)

        csv_completed = display_comp.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Completed Evaluations (CSV)",
            data=csv_completed,
            file_name="osm_completed_evaluations.csv",
            mime="text/csv"
        )