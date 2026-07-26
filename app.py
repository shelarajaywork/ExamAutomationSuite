"""
==================================================
EXAM TOOLS - MAIN APP
==================================================

PURPOSE
--------------------------------------------------
This is the entry point of the Streamlit app. It handles:
  1. Page setup (title, icon, layout).
  2. The sidebar menu & Dashboard tool cards.
  3. Routing to the right "page" based on user selection.

The actual logic for each tool lives in its own file inside the
`Modules/` subfolders (e.g. `Modules/Questionwise_Checker/questionwise_checker.py`,
`Modules/Gracing_Checker/gracing_checker.py`).
==================================================
"""

import streamlit as st

# Import Questionwise Checker module tool
try:
    from Modules.Questionwise_Checker.questionwise_checker import show as questionwise_checker
except Exception as import_error:
    questionwise_checker = None
    _qc_import_error_message = str(import_error)
else:
    _qc_import_error_message = None

# Import Gracing Checker module tool
try:
    from Modules.Gracing_Checker.gracing_checker import show as gracing_checker
except Exception as import_error:
    gracing_checker = None
    _gc_import_error_message = str(import_error)
else:
    _gc_import_error_message = None


# --------------------
# Page Configuration
# --------------------
st.set_page_config(
    page_title="Exam Tools",
    page_icon="🎓",
    layout="wide",
)

# --------------------
# Shared Session State
# --------------------
if "students_processed" not in st.session_state:
    st.session_state.students_processed = 0

if "files_uploaded" not in st.session_state:
    st.session_state.files_uploaded = 0

if "selected_tool" not in st.session_state:
    st.session_state.selected_tool = "🏠 Dashboard"

# Synchronize menu selection with session state
def update_menu():
    st.session_state.selected_tool = st.session_state.sidebar_radio

# --------------------
# Sidebar
# --------------------
st.sidebar.title("🎓 Exam Tools")

tools_list = [
    "🏠 Dashboard",
    "✅ Gracing Checker",
    "📝 Questionwise Checker",
    "📂 DigiLocker Tracker",
    "🔍 Duplicate Checker",
    "📊 Result Analysis",
    "🧹 Data Cleaner",
    "🔗 Merge Files",
    "☁️ SharePoint Import",
    "⚙️ Settings",
]

menu = st.sidebar.radio(
    "Select Tool",
    tools_list,
    index=tools_list.index(st.session_state.selected_tool),
    key="sidebar_radio",
    on_change=update_menu,
)

# --------------------
# Dashboard Page
# --------------------
if st.session_state.selected_tool == "🏠 Dashboard":

    st.title("🎓 Examination Automation Suite")
    st.caption("Select a tool below to begin processing examination records.")
    st.markdown("---")

    # Defined tool cards with icons and descriptions
    tools_info = [
        {
            "name": "✅ Gracing Checker",
            "desc": "Check and validate grace marks, ordinance rules, and eligibility criteria for candidates.",
        },
        {
            "name": "📝 Questionwise Checker",
            "desc": "Verify question-wise marks, compare Examiner vs. Moderator scores, and export cleaned mark sheets.",
        },
        {
            "name": "📂 DigiLocker Tracker",
            "desc": "Track, validate, and identify DigiLocker upload errors and discrepancies.",
        },
        {
            "name": "🔍 Duplicate Checker",
            "desc": "Detect and remove duplicate student entries across examination sheets.",
        },
        {
            "name": "📊 Result Analysis",
            "desc": "Analyze pass/fail statistics, subject performance, and score distributions.",
        },
        {
            "name": "🧹 Data Cleaner",
            "desc": "Clean up stray characters, format standard headers, and strip invalid rows.",
        },
        {
            "name": "🔗 Merge Files",
            "desc": "Combine multiple mark sheets or result exports into a single master sheet.",
        },
        {
            "name": "☁️ SharePoint Import",
            "desc": "Import and sync mark sheets directly from institutional SharePoint drives.",
        },
    ]

    # Grid layout for Tool Cards (2 cards per row)
    for i in range(0, len(tools_info), 2):
        col1, col2 = st.columns(2)
        
        with col1:
            tool = tools_info[i]
            with st.container(border=True):
                st.subheader(tool["name"])
                st.write(tool["desc"])
                btn_label = tool["name"].split(" ", 1)[1] if " " in tool["name"] else tool["name"]
                if st.button(f"Open {btn_label}", key=f"btn_{i}"):
                    st.session_state.selected_tool = tool["name"]
                    st.rerun()

        if i + 1 < len(tools_info):
            with col2:
                tool = tools_info[i + 1]
                with st.container(border=True):
                    st.subheader(tool["name"])
                    st.write(tool["desc"])
                    btn_label = tool["name"].split(" ", 1)[1] if " " in tool["name"] else tool["name"]
                    if st.button(f"Open {btn_label}", key=f"btn_{i+1}"):
                        st.session_state.selected_tool = tool["name"]
                        st.rerun()

# --------------------
# Gracing Checker
# --------------------
elif st.session_state.selected_tool == "✅ Gracing Checker":

    if gracing_checker is not None:
        gracing_checker()
    else:
        st.title("✅ Gracing Checker")
        if _gc_import_error_message:
            st.info("This tool file (`Modules/Gracing_Checker/gracing_checker.py`) is currently blank or under development.")
        else:
            st.error(
                "This tool could not be loaded due to an import error:\n\n"
                f"`{_gc_import_error_message}`\n\n"
                "Check that `Modules/Gracing_Checker/gracing_checker.py` defines a `show()` function."
            )

# --------------------
# Questionwise Checker
# --------------------
elif st.session_state.selected_tool == "📝 Questionwise Checker":

    if questionwise_checker is not None:
        questionwise_checker()
    else:
        st.title("📝 Questionwise Checker")
        st.error(
            "This tool could not be loaded due to an import error:\n\n"
            f"`{_qc_import_error_message}`\n\n"
            "Check that `Modules/Questionwise_Checker/questionwise_checker.py` exists and that "
            "all its dependencies (see requirements.txt) are installed."
        )

# --------------------
# DigiLocker Tracker
# --------------------
elif st.session_state.selected_tool == "📂 DigiLocker Tracker":

    st.title("📂 DigiLocker Tracker")

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV File",
        type=["xlsx", "csv"],
    )

    if uploaded_file:
        st.success("File Uploaded Successfully")

# --------------------
# Duplicate Checker
# --------------------
elif st.session_state.selected_tool == "🔍 Duplicate Checker":

    st.title("🔍 Duplicate Checker")
    st.info("Tool under development")

# --------------------
# Result Analysis
# --------------------
elif st.session_state.selected_tool == "📊 Result Analysis":

    st.title("📊 Result Analysis")
    st.info("Tool under development")

# --------------------
# Data Cleaner
# --------------------
elif st.session_state.selected_tool == "🧹 Data Cleaner":

    st.title("🧹 Data Cleaner")
    st.info("Tool under development")

# --------------------
# Merge Files
# --------------------
elif st.session_state.selected_tool == "🔗 Merge Files":

    st.title("🔗 Merge Files")
    st.info("Tool under development")

# --------------------
# SharePoint Import
# --------------------
elif st.session_state.selected_tool == "☁️ SharePoint Import":

    st.title("☁️ SharePoint Import")
    st.info("Tool under development")

# --------------------
# Settings
# --------------------
elif st.session_state.selected_tool == "⚙️ Settings":

    st.title("⚙️ Settings")
    st.info("Application settings will appear here")