"""
==================================================
AUTH MODULE - SVKM EXAM AUTOMATION SUITE
==================================================
Handles password-based authentication, session state management,
login UI rendering, and user logout for the examination suite.
"""

import hmac
import streamlit as st

DEFAULT_PASSWORD = "Exam@111"


def get_auth_password() -> str:
    """Retrieves authentication password from st.secrets if configured, else returns default."""
    if "auth" in st.secrets and "password" in st.secrets["auth"]:
        return str(st.secrets["auth"]["password"])
    return DEFAULT_PASSWORD


def verify_password(entered_password: str, expected_password: str) -> bool:
    """Verifies that the entered password matches using constant-time comparison."""
    if not entered_password:
        return False
    return hmac.compare_digest(entered_password.strip(), expected_password.strip())


def logout():
    """Logs the user out by clearing authentication state and rerunning."""
    st.session_state["authenticated"] = False
    st.rerun()


def render_login_page():
    """Renders the centralized examination suite password-only login page."""
    expected_password = get_auth_password()

    # CSS to completely hide the sidebar and center the login page
    login_css = """
    <style>
    /* Hide sidebar and toggle controls during login */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    [data-testid="stSidebarCollapsedControl"] {
        display: none !important;
    }
    .badge-ced {
        display: inline-block;
        background-color: #dbeafe;
        color: #1e40af;
        font-size: 0.85rem;
        font-weight: 600;
        padding: 0.35rem 0.9rem;
        border-radius: 9999px;
        margin-bottom: 1rem;
    }
    </style>
    """
    st.markdown(login_css, unsafe_allow_html=True)

    # Centered layout using 3 columns
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(
            """
            <div style="text-align: center; margin-top: 3rem; margin-bottom: 1.5rem;">
                <div class="badge-ced">🏛️ SVKM's Centralized Examination Division</div>
                <h1 style="font-size: 2rem; font-weight: 700; color: #1e3a8a; margin-bottom: 0.2rem;">
                    🎓 Exam Automation Suite
                </h1>
                <p style="color: #64748b; font-size: 0.95rem;">
                    Enter your access password to unlock examination tools and analytics
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.container(border=True):
            with st.form("login_form", clear_on_submit=False):
                entered_password = st.text_input(
                    "🔑 Access Password",
                    type="password",
                    placeholder="Enter access password",
                    help="Enter the examination suite access password"
                )

                st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                submit_button = st.form_submit_button("🚀 Sign In", use_container_width=True, type="primary")

                if submit_button:
                    if not entered_password:
                        st.error("⚠️ Please enter the password.")
                    else:
                        if verify_password(entered_password, expected_password):
                            st.session_state["authenticated"] = True
                            st.rerun()
                        else:
                            st.error("❌ Invalid password. Please try again.")

        st.caption(
            "<div style='text-align: center; color: #94a3b8; font-size: 0.8rem; margin-top: 1.5rem;'>"
            "🔒 Protected System • SVKM CED Internal Access Only"
            "</div>",
            unsafe_allow_html=True
        )


def render_logout_button():
    """Renders a clean logout button at the bottom of the sidebar."""
    if not st.session_state.get("authenticated", False):
        return

    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Log Out", use_container_width=True, key="sidebar_logout_btn"):
        logout()


def require_auth():
    """
    Main authentication gatekeeper.
    Call this at the beginning of app.py.
    If not authenticated, renders login page and halts execution.
    """
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if not st.session_state["authenticated"]:
        render_login_page()
        st.stop()
