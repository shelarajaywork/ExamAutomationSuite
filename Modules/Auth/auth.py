"""
==================================================
AUTH MODULE - SVKM EXAM AUTOMATION SUITE
==================================================
Handles password-based authentication, session state management,
clean distraction-free login UI, and user logout for the examination suite.
"""

import hmac
import streamlit as st

try:
    from Modules.UI.icons import get_svg_icon
    from Modules.UI.theme import apply_theme_css, get_current_theme
except Exception:
    try:
        from ..UI.icons import get_svg_icon
        from ..UI.theme import apply_theme_css, get_current_theme
    except Exception:
        from UI.icons import get_svg_icon
        from UI.theme import apply_theme_css, get_current_theme

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
    """
    Renders an ultra-compact, perfectly centered, distraction-free login card
    with high-contrast inputs and pure Light Mode styling.
    """
    expected_password = get_auth_password()

    login_css = """
    <style>
    /* Completely hide sidebar and header on login */
    [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], header[data-testid="stHeader"] {
        display: none !important;
    }

    /* Complete suppression of all input instructions, form hints, and submit text */
    [data-testid="stInputInstructions"],
    div[data-testid="stInputInstructions"],
    span[data-testid="stInputInstructions"],
    div[data-testid="stTextInput"] [data-testid="stInputInstructions"],
    div[data-testid="stTextInput"] small,
    div[data-baseweb="input"] ~ div,
    .st-emotion-cache-1vt4y43,
    .st-emotion-cache-1aege1h,
    .st-emotion-cache-1gw59om {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        height: 0 !important;
        width: 0 !important;
        font-size: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        position: absolute !important;
        pointer-events: none !important;
        clip: rect(0, 0, 0, 0) !important;
    }

    /* Clean, distraction-free background canvas */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
        background-color: #F1F5F9 !important;
        background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 50%, #F1F5F9 100%) !important;
        min-height: 100vh !important;
    }

    /* Centered Header Elements */
    .login-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 5px;
        background-color: #EFF6FF;
        color: #1D4ED8;
        border: 1px solid #DBEAFE;
        font-size: 0.74rem;
        font-weight: 600;
        padding: 0.2rem 0.65rem;
        border-radius: 9999px;
        letter-spacing: 0.2px;
        margin: 0 auto 0.6rem auto;
    }

    .login-title {
        font-size: 1.45rem !important;
        font-weight: 700 !important;
        color: #0F172A !important;
        text-align: center !important;
        margin: 0 auto 0.25rem auto !important;
        letter-spacing: -0.025em !important;
        line-height: 1.25 !important;
    }

    .login-desc {
        font-size: 0.82rem !important;
        color: #475569 !important;
        text-align: center !important;
        margin: 0 auto 1.15rem auto !important;
        line-height: 1.45 !important;
        font-weight: 500 !important;
    }

    /* Elevated White Login Container */
    [data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stContainer"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.06), 0 8px 10px -6px rgba(0, 0, 0, 0.02) !important;
        padding: 0.5rem !important;
    }

    /* Password Input - Strict High Contrast Dark Text on Pure White */
    div[data-testid="stTextInput"] {
        margin-bottom: 0.5rem !important;
    }
    div[data-testid="stTextInput"] label {
        text-align: center !important;
        display: block !important;
        width: 100% !important;
        font-size: 0.84rem !important;
        font-weight: 600 !important;
        margin-bottom: 0.35rem !important;
        color: #0F172A !important;
    }
    div[data-baseweb="input"], div[data-baseweb="base-input"] {
        min-height: 38px !important;
        height: 38px !important;
        border-radius: 8px !important;
        background-color: #FFFFFF !important;
        border: 1.5px solid #CBD5E1 !important;
    }
    div[data-baseweb="input"]:focus-within {
        border-color: #2563EB !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.20) !important;
    }
    div[data-baseweb="input"] input, input[type="password"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 0.92rem !important;
        font-weight: 600 !important;
        padding: 0.45rem 0.75rem !important;
        text-align: center !important;
    }
    div[data-baseweb="input"] input::placeholder {
        color: #94A3B8 !important;
        -webkit-text-fill-color: #94A3B8 !important;
        font-weight: 400 !important;
    }

    /* High-Contrast Crisp White Text on Sign In Button */
    button[data-testid="baseButton-primary"],
    button[data-testid="stBaseButton-primary"],
    .stButton > button[kind="primary"],
    .stButton > button {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        min-height: 38px !important;
        height: 38px !important;
        font-size: 0.90rem !important;
        font-weight: 700 !important;
        box-shadow: 0 2px 4px 0 rgba(37, 99, 235, 0.2) !important;
        transition: all 0.15s ease-in-out !important;
        cursor: pointer !important;
    }
    button[data-testid="baseButton-primary"]:hover,
    .stButton > button:hover {
        background-color: #1D4ED8 !important;
        box-shadow: 0 4px 10px 0 rgba(37, 99, 235, 0.35) !important;
    }
    button[data-testid="baseButton-primary"] *,
    button[data-testid="baseButton-primary"] p,
    .stButton > button * {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 0.90rem !important;
    }
    </style>
    """
    st.markdown(login_css, unsafe_allow_html=True)

    # Center column structure to guarantee compact width and perfect horizontal centering
    col_l, col_center, col_r = st.columns([1.3, 1.0, 1.3])
    with col_center:
        st.markdown("<div style='height: 8vh;'></div>", unsafe_allow_html=True)
        building_icon = get_svg_icon("building", size=13, color="currentColor")
        
        st.markdown(
            f"""
            <div style="text-align: center; width: 100%; margin-bottom: 0.5rem;">
                <div class="login-badge">
                    {building_icon} SVKM CED
                </div>
                <div class="login-title">
                    Exam Automation Suite
                </div>
                <div class="login-desc">
                    Enter your password to unlock tools and records
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.container(border=True):
            entered_password = st.text_input(
                "Enter Password",
                type="password",
                placeholder="Enter Password",
                key="login_password_field",
                label_visibility="visible",
            )

            st.markdown("<div style='margin-top: 8px;'></div>", unsafe_allow_html=True)
            submit_button = st.button("Sign In", use_container_width=True, type="primary", key="login_submit_btn")

            # Validate when user clicks Sign In or submits input
            if submit_button:
                if not entered_password:
                    st.error("Please enter the password.")
                else:
                    if verify_password(entered_password, expected_password):
                        st.session_state["authenticated"] = True
                        st.rerun()
                    else:
                        st.error("Invalid password. Please try again.")
            elif entered_password and entered_password.strip() == expected_password.strip():
                st.session_state["authenticated"] = True
                st.rerun()

        shield_icon = get_svg_icon("shield", size=12, color="currentColor")
        st.markdown(
            f"""
            <div style="text-align: center; color: #64748B; font-size: 0.74rem; margin-top: 1rem; font-weight: 500; display: flex; align-items: center; justify-content: center; gap: 4px;">
                {shield_icon} Protected System • SVKM CED Internal Access Only
            </div>
            """,
            unsafe_allow_html=True
        )


def render_logout_button():
    """Renders a clean, compact logout button in the sidebar with vector icon."""
    if not st.session_state.get("authenticated", False):
        return

    st.sidebar.markdown("<hr style='margin: 0.5rem 0 0.5rem 0; border: none; border-top: 1px solid var(--border-color);'/>", unsafe_allow_html=True)
    if st.sidebar.button("Log Out", use_container_width=True, key="sidebar_logout_btn"):
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
