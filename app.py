"""
==================================================
EXAM TOOLS - MAIN APP
==================================================
Centralized Examination Automation & Analytics Suite
"""
import os
import sys

# Ensure workspace root is always in sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import streamlit as st

try:
    from Modules.Auth.auth import require_auth, render_logout_button
except Exception:
    from Auth.auth import require_auth, render_logout_button

try:
    from Modules.UI.icons import get_svg_icon, get_tool_icon
    from Modules.UI.theme import apply_theme_css, render_compact_theme_toggle, get_current_theme
except Exception:
    from UI.icons import get_svg_icon, get_tool_icon
    from UI.theme import apply_theme_css, render_compact_theme_toggle, get_current_theme

# --------------------
# Module Imports
# --------------------
try:
    from Modules.CED_Examination_Reports.ced_examination_reports import show as ced_examination_reports
except Exception:
    try:
        from Modules.OSM_Dashboard.osm_dashboard import show as ced_examination_reports
    except Exception as import_error:
        ced_examination_reports = None
        _ced_import_error_message = str(import_error)
    else:
        _ced_import_error_message = None
else:
    _ced_import_error_message = None

try:
    from Modules.Questionwise_Checker.questionwise_checker import show as questionwise_checker
except Exception as import_error:
    questionwise_checker = None
    _qc_import_error_message = str(import_error)
else:
    _qc_import_error_message = None

try:
    from Modules.Letters_to_Ext_Exmnrs.letters_to_ext_exmnrs import show as letters_to_ext_examiners
except Exception as import_error:
    try:
        from Modules.Letters_to_Ext_Exmnrs.letters_to_ext_exmnrs import main as letters_to_ext_examiners
    except Exception as import_error_2:
        letters_to_ext_examiners = None
        _lte_import_error_message = str(import_error_2)
    else:
        _lte_import_error_message = None
else:
    _lte_import_error_message = None

try:
    from Modules.Rank_List_Generator.rank_list_generator import show as rank_list_generator
except Exception as import_error:
    rank_list_generator = None
    _rlg_import_error_message = str(import_error)
else:
    _rlg_import_error_message = None

try:
    from Modules.Result_Gazette_to_Excel.result_gazette_to_excel import show as result_gazette_to_excel
except Exception:
    try:
        from Modules.Result_Gazette_to_Excel.result_gazette_to_excel import main as result_gazette_to_excel
    except Exception as import_error:
        result_gazette_to_excel = None
        _rgte_import_error_message = str(import_error)
    else:
        _rgte_import_error_message = None
else:
    _rgte_import_error_message = None

try:
    from Modules.Grade_Sheet_to_Excel.grade_sheet_to_excel import show as grade_sheet_to_excel
except Exception:
    try:
        from Modules.Grade_Sheet_to_Excel.grade_sheet_to_excel import main as grade_sheet_to_excel
    except Exception as import_error:
        grade_sheet_to_excel = None
        _gste_import_error_message = str(import_error)
    else:
        _gste_import_error_message = None
else:
    _gste_import_error_message = None

try:
    from Modules.Gracing_Checker.gracing_checker import show as gracing_checker
except Exception as import_error:
    gracing_checker = None
    _gc_import_error_message = str(import_error)
else:
    _gc_import_error_message = None

try:
    from Modules.Exam_Timetable_Generator.exam_timetable_generator import main as exam_timetable_generator
except Exception as import_error:
    try:
        from Modules.Exam_Timetable_Generator.exam_timetable_generator import show as exam_timetable_generator
    except Exception as import_error_2:
        exam_timetable_generator = None
        _etg_import_error_message = str(import_error_2)
    else:
        _etg_import_error_message = None
else:
    _etg_import_error_message = None

try:
    from Modules.Convocation_Data_Generator.convocation_data_generator import show as convocation_data_generator
except Exception as import_error:
    try:
        from Modules.Convocation_Data_Generator.convocation_data_generator import main as convocation_data_generator
    except Exception as import_error_2:
        convocation_data_generator = None
        _cdg_import_error_message = str(import_error_2)
    else:
        _cdg_import_error_message = None
else:
    _cdg_import_error_message = None

try:
    from Modules.Data_Analysis.data_analysis import show as data_analysis
except Exception as import_error:
    data_analysis = None
    _da_import_error_message = str(import_error)
else:
    _da_import_error_message = None

try:
    from Modules.Merge_Files.merge_files import show as merge_files
except Exception as import_error:
    merge_files = None
    _mf_import_error_message = str(import_error)
else:
    _mf_import_error_message = None


# --------------------
# Page Configuration
# --------------------
st.set_page_config(
    page_title="SVKM Exam Tools",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply dynamic WCAG AAA accessible Tailwind Slate theme CSS
apply_theme_css()

# --------------------
# Authentication Gate
# --------------------
require_auth()

# --------------------
# Shared Session State & Menu
# --------------------
tools_list = [
    "Dashboard",
    "OSM Dashboard",
    "Exam Timetable Generator",
    "Gracing Checker",
    "Questionwise Checker",
    "Letters to Ext. Examiners",
    "Rank List Generator",
    "Result Gazette PDF to Excel",
    "Grade Sheet to Excel",
    "Convocation Data Generator",
    "Duplicate Checker",
    "Result Analysis",
    "Data Analysis",
    "Data Cleaner",
    "Merge Files",
    "SharePoint Import",
    "Settings",
]

def normalize_tool_name(name: str) -> str:
    """Normalizes tool name to match clean list, preserving compatibility with legacy emoji links."""
    if not name:
        return "Dashboard"
    clean = name
    for emoji in ["🏠", "📊", "📅", "✅", "📝", "✉️", "🏆", "📑", "🎓", "🔍", "📈", "🧹", "🔗", "☁️", "⚙️"]:
        clean = clean.replace(emoji, "").strip()
    return clean if clean in tools_list else "Dashboard"

if "selected_tool" not in st.session_state:
    st.session_state.selected_tool = "Dashboard"
else:
    st.session_state.selected_tool = normalize_tool_name(st.session_state.selected_tool)

def update_menu():
    st.session_state.selected_tool = st.session_state.sidebar_radio

# --------------------
# Sidebar
# --------------------
brand_icon = get_svg_icon("gradesheet", size=22, color="var(--primary-btn-bg)")
st.sidebar.markdown(
    f"""
    <div style="display: flex; align-items: center; gap: 8px; padding: 4px 2px 2px 2px;">
        <div class="ui-icon-box" style="margin: 0; width: 34px; height: 34px;">
            {brand_icon}
        </div>
        <div>
            <div style="font-size: 1.05rem; font-weight: 700; color: var(--text-primary); line-height: 1.15;">Exam Suite</div>
            <div style="font-size: 0.70rem; color: var(--text-muted); font-weight: 600;">SVKM CED</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.markdown("<hr style='margin: 0.5rem 0 0.75rem 0; border: none; border-top: 1px solid var(--border-color);'/>", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Select Tool",
    tools_list,
    index=tools_list.index(st.session_state.selected_tool),
    key="sidebar_radio",
    on_change=update_menu,
    label_visibility="collapsed"
)

render_logout_button()

# --------------------
# App Router
# --------------------

# 1. Main Dashboard
if st.session_state.selected_tool == "Dashboard":
    building_icon = get_svg_icon("building", size=14, color="currentColor")
    
    st.markdown(
        f"""
        <div style="margin-bottom: 1.5rem;">
            <div class="ui-badge" style="margin-bottom: 0.6rem;">
                {building_icon} SVKM's Centralized Examination Division
            </div>
            <h1 style="font-size: 2.1rem; font-weight: 700; margin-bottom: 0.35rem;">
                Examination Automation & Analytics Suite
            </h1>
            <p style="color: var(--text-secondary); font-size: 0.95rem; margin: 0;">
                Select an examination tool below to process records, analyze results, and generate official documentation.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("<hr style='margin: 0 0 1.5rem 0; border: none; border-top: 1px solid var(--border-color);'/>", unsafe_allow_html=True)

    tools_info = [
        {
            "name": "OSM Dashboard",
            "icon": "osm",
            "desc": "Track Online Screen Marking evaluation status, examiner workload, turnaround times, and pending papers.",
        },
        {
            "name": "Exam Timetable Generator",
            "icon": "calendar",
            "desc": "Generate clash-free exam timetables for Regular and Re-Examinations using DSATUR graph coloring.",
        },
        {
            "name": "Gracing Checker",
            "icon": "gracing",
            "desc": "Check and validate grace marks, ordinance rules, and eligibility criteria for candidates.",
        },
        {
            "name": "Questionwise Checker",
            "icon": "questionwise",
            "desc": "Verify question-wise marks, compare Examiner vs. Moderator scores, and export cleaned mark sheets.",
        },
        {
            "name": "Letters to Ext. Examiners",
            "icon": "mail",
            "desc": "Select examiners from evaluation dashboards and generate draft appointment/thanking emails with attached PDFs.",
        },
        {
            "name": "Rank List Generator",
            "icon": "rank",
            "desc": "Process student CGPAs across programs to extract top 5% rank holders and export formatted Word documents.",
        },
        {
            "name": "Result Gazette PDF to Excel",
            "icon": "gazette",
            "desc": "Extract metadata, course schemes, student result summaries, and subject-level data from PDF gazettes to Excel.",
        },
        {
            "name": "Grade Sheet to Excel",
            "icon": "gradesheet",
            "desc": "Extract student demographics, multi-semester progression, CA/ESE marks, and grade summaries from Grade Sheet PDFs to Excel.",
        },
        {
            "name": "Convocation Data Generator",
            "icon": "convocation",
            "desc": "Process MKCL reports and Master Data to generate standardized convocation Excel reports.",
        },
        {
            "name": "Duplicate Checker",
            "icon": "duplicate",
            "desc": "Detect and remove duplicate student entries across examination sheets.",
        },
        {
            "name": "Result Analysis",
            "icon": "analysis",
            "desc": "Analyze pass/fail statistics, subject performance, and score distributions.",
        },
        {
            "name": "Data Analysis",
            "icon": "data_analysis",
            "desc": "Upload Student Master Data, Gracing, GMR, or Moderation reports and view KPI & chart dashboards.",
        },
        {
            "name": "Data Cleaner",
            "icon": "cleaner",
            "desc": "Clean up stray characters, format standard headers, and strip invalid rows.",
        },
        {
            "name": "Merge Files",
            "icon": "merge",
            "desc": "Combine multiple mark sheets or result exports into a single master sheet.",
        },
        {
            "name": "SharePoint Import",
            "icon": "sharepoint",
            "desc": "Import and sync mark sheets directly from institutional SharePoint drives.",
        },
    ]

    for i in range(0, len(tools_info), 2):
        col1, col2 = st.columns(2)
        with col1:
            tool = tools_info[i]
            tool_icon_svg = get_svg_icon(tool["icon"], size=20, color="var(--primary-btn-bg)")
            with st.container(border=True):
                st.markdown(
                    f"""
                    <div style="display: flex; align-items: center; margin-bottom: 0.5rem;">
                        <div class="ui-icon-box">
                            {tool_icon_svg}
                        </div>
                        <h3 style="margin: 0; font-size: 1.15rem; font-weight: 600;">{tool["name"]}</h3>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                st.write(tool["desc"])
                if st.button(f"Open {tool['name']}", key=f"btn_{i}", use_container_width=True):
                    st.session_state.selected_tool = tool["name"]
                    st.rerun()

        if i + 1 < len(tools_info):
            with col2:
                tool = tools_info[i + 1]
                tool_icon_svg = get_svg_icon(tool["icon"], size=20, color="var(--primary-btn-bg)")
                with st.container(border=True):
                    st.markdown(
                        f"""
                        <div style="display: flex; align-items: center; margin-bottom: 0.5rem;">
                            <div class="ui-icon-box">
                                {tool_icon_svg}
                            </div>
                            <h3 style="margin: 0; font-size: 1.15rem; font-weight: 600;">{tool["name"]}</h3>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    st.write(tool["desc"])
                    if st.button(f"Open {tool['name']}", key=f"btn_{i+1}", use_container_width=True):
                        st.session_state.selected_tool = tool["name"]
                        st.rerun()

# 2. OSM Dashboard
elif st.session_state.selected_tool == "OSM Dashboard":
    if ced_examination_reports is not None:
        ced_examination_reports()
    else:
        st.title("OSM Dashboard")
        st.error(
            "This tool could not be loaded due to an import error:\n\n"
            f"`{_ced_import_error_message}`\n\n"
            "Ensure that `Modules/CED_Examination_Reports/ced_examination_reports.py` exists "
            "and defines a `show()` entry point function."
        )

# 3. Exam Timetable Generator
elif st.session_state.selected_tool == "Exam Timetable Generator":
    if exam_timetable_generator is not None:
        exam_timetable_generator()
    else:
        st.title("Exam Timetable Generator")
        st.error(f"`{_etg_import_error_message}`")

# 4. Gracing Checker
elif st.session_state.selected_tool == "Gracing Checker":
    if gracing_checker is not None:
        gracing_checker()
    else:
        st.title("Gracing Checker")
        st.info("This tool is under development.")

# 5. Questionwise Checker
elif st.session_state.selected_tool == "Questionwise Checker":
    if questionwise_checker is not None:
        questionwise_checker()
    else:
        st.title("Questionwise Checker")
        st.error(f"`{_qc_import_error_message}`")

# 6. Letters to Ext. Examiners
elif st.session_state.selected_tool == "Letters to Ext. Examiners":
    if letters_to_ext_examiners is not None:
        letters_to_ext_examiners()
    else:
        st.title("Letters to Ext. Examiners")
        st.error(f"`{_lte_import_error_message}`")

# 7. Rank List Generator
elif st.session_state.selected_tool == "Rank List Generator":
    if rank_list_generator is not None:
        rank_list_generator()
    else:
        st.title("Rank List Generator")
        st.error(f"`{_rlg_import_error_message}`")

# 8. Result Gazette PDF to Excel
elif st.session_state.selected_tool == "Result Gazette PDF to Excel":
    if result_gazette_to_excel is not None:
        result_gazette_to_excel()
    else:
        st.title("Result Gazette PDF to Excel")
        st.error(
            "This tool could not be loaded due to an import error:\n\n"
            f"`{_rgte_import_error_message}`\n\n"
            "Ensure that `Modules/Result_Gazette_to_Excel/result_gazette_to_excel.py` exists "
            "and defines a `show()` or `main()` entry point function."
        )

# 9. Grade Sheet to Excel
elif st.session_state.selected_tool == "Grade Sheet to Excel":
    if grade_sheet_to_excel is not None:
        grade_sheet_to_excel()
    else:
        st.title("Grade Sheet to Excel")
        st.error(
            "This tool could not be loaded due to an import error:\n\n"
            f"`{_gste_import_error_message}`\n\n"
            "Ensure that `Modules/Grade_Sheet_to_Excel/grade_sheet_to_excel.py` exists "
            "and defines a `show()` or `main()` entry point function."
        )

# 10. Convocation Data Generator
elif st.session_state.selected_tool == "Convocation Data Generator":
    if convocation_data_generator is not None:
        convocation_data_generator()
    else:
        st.title("Convocation Data Generator")
        st.error(f"`{_cdg_import_error_message}`")

# 11. Duplicate Checker
elif st.session_state.selected_tool == "Duplicate Checker":
    st.title("Duplicate Checker")
    st.info("Tool under development")

# 12. Result Analysis
elif st.session_state.selected_tool == "Result Analysis":
    st.title("Result Analysis")
    st.info("Tool under development")

# 13. Data Analysis
elif st.session_state.selected_tool == "Data Analysis":
    if data_analysis is not None:
        data_analysis()
    else:
        st.title("Data Analysis")
        st.error(f"`{_da_import_error_message}`")

# 14. Data Cleaner
elif st.session_state.selected_tool == "Data Cleaner":
    st.title("Data Cleaner")
    st.info("Tool under development")

# 15. Merge Files
elif st.session_state.selected_tool == "Merge Files":
    if merge_files is not None:
        merge_files()
    else:
        st.title("Merge Files")
        st.error(f"`{_mf_import_error_message}`")

# 16. SharePoint Import
elif st.session_state.selected_tool == "SharePoint Import":
    st.title("SharePoint Import")
    st.info("Tool under development")

# 17. Settings
elif st.session_state.selected_tool == "Settings":
    st.title("Settings")
    st.info("Application settings will appear here")