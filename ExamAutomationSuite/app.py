"""
==================================================
EXAM TOOLS - MAIN APP
==================================================

PURPOSE
--------------------------------------------------
This is the entry point of the Streamlit app. It only handles:
  1. Page setup (title, icon, layout).
  2. The sidebar menu.
  3. Routing to the right "page" based on what's selected in the menu.

The actual logic for each tool lives in its own file inside the
`Modules/` folder (e.g. `Modules/questionwise_checker.py`). Keeping
each tool in its own file/module means this file stays short and easy
to scan, and each tool can be worked on independently.

Run this app with:
    streamlit run app.py
==================================================
"""

import streamlit as st

# The Questionwise Checker tool lives in its own module. Import errors
# here (e.g. a typo in the filename, or a missing dependency) would
# otherwise crash the entire app before it even starts, so we catch
# that and show a clear message instead.
try:
    from Modules.questionwise_checker import show as questionwise_checker
except Exception as import_error:
    questionwise_checker = None
    _import_error_message = str(import_error)
else:
    _import_error_message = None


# --------------------
# Page Configuration
# --------------------
# `set_page_config` must be the first Streamlit command in the script.
st.set_page_config(
    page_title="Exam Tools",
    page_icon="🎓",
    layout="wide",
)

# --------------------
# Shared session state
# --------------------
# st.session_state persists values across reruns *for the current
# user's session* (Streamlit reruns this whole script top-to-bottom
# every time the user interacts with a widget). We use it here so the
# Dashboard page can show real numbers, fed in by whichever tool the
# user actually used - see Modules/questionwise_checker.py, which sets
# `students_processed` and `files_uploaded` after processing a file.
if "students_processed" not in st.session_state:
    st.session_state.students_processed = 0

if "files_uploaded" not in st.session_state:
    st.session_state.files_uploaded = 0

# --------------------
# Sidebar
# --------------------
st.sidebar.title("🎓 Exam Tools")

menu = st.sidebar.radio(
    "Select Tool",
    [
        "🏠 Dashboard",
        "📝 Questionwise Checker",
        "📂 DigiLocker Tracker",
        "🔍 Duplicate Checker",
        "✅ ABC Validation",
        "📊 Result Analysis",
        "🧹 Data Cleaner",
        "🔗 Merge Files",
        "☁️ SharePoint Import",
        "⚙️ Settings",
    ],
)

# --------------------
# Dashboard
# --------------------
if menu == "🏠 Dashboard":

    st.title("🎓 Examination Automation Suite")
    st.markdown("---")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        # Real value, updated by Questionwise Checker after it
        # processes a file (falls back to 0 before any file is run).
        st.metric("Students Processed", st.session_state.students_processed)

    with col2:
        # Placeholder until Duplicate Checker is built - wire this up
        # to st.session_state the same way once that tool is live.
        st.metric("Duplicate Records", "0")

    with col3:
        st.metric("Files Uploaded", st.session_state.files_uploaded)

    with col4:
        # Placeholder until tools report errors back via
        # st.session_state (e.g. a running "errors_found" counter).
        st.metric("Errors Found", "0")

    st.markdown("---")
    st.subheader("Welcome")

    st.info(
        """
        This application will help manage:

        • DigiLocker Error Tracking

        • Duplicate Record Validation

        • ABC ID Validation

        • Result Analysis

        • Data Cleaning

        • SharePoint Integration
        """
    )

# --------------------
# DigiLocker Tracker
# --------------------
elif menu == "📂 DigiLocker Tracker":

    st.title("📂 DigiLocker Tracker")

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV File",
        type=["xlsx", "csv"],
    )

    if uploaded_file:
        st.success("File Uploaded Successfully")

# --------------------
# Questionwise Checker
# --------------------
elif menu == "📝 Questionwise Checker":

    if questionwise_checker is not None:
        questionwise_checker()
    else:
        # This shows up instead of a hard crash if Modules/questionwise_checker.py
        # is missing, has a syntax error, or a missing dependency (e.g. msoffcrypto
        # not installed) - the sidebar/other tools still work normally.
        st.title("📝 Questionwise Checker")
        st.error(
            "This tool could not be loaded due to an import error:\n\n"
            f"`{_import_error_message}`\n\n"
            "Check that `Modules/questionwise_checker.py` exists and that "
            "all its dependencies (see requirements.txt) are installed."
        )

# --------------------
# Duplicate Checker
# --------------------
elif menu == "🔍 Duplicate Checker":

    st.title("🔍 Duplicate Checker")
    st.info("Tool under development")

# --------------------
# ABC Validation
# --------------------
elif menu == "✅ ABC Validation":

    st.title("✅ ABC Validation")
    st.info("Tool under development")

# --------------------
# Result Analysis
# --------------------
elif menu == "📊 Result Analysis":

    st.title("📊 Result Analysis")
    st.info("Tool under development")

# --------------------
# Data Cleaner
# --------------------
elif menu == "🧹 Data Cleaner":

    st.title("🧹 Data Cleaner")
    st.info("Tool under development")

# --------------------
# Merge Files
# --------------------
elif menu == "🔗 Merge Files":

    st.title("🔗 Merge Files")
    st.info("Tool under development")

# --------------------
# SharePoint Import
# --------------------
elif menu == "☁️ SharePoint Import":

    st.title("☁️ SharePoint Import")
    st.info("Tool under development")

# --------------------
# Settings
# --------------------
elif menu == "⚙️ Settings":

    st.title("⚙️ Settings")
    st.info("Application settings will appear here")