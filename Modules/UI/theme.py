"""
==================================================
THEME & DESIGN SYSTEM (BLUE & WHITE LIGHT MODE)
==================================================
Clean, professional royal blue and white enterprise aesthetic
inspired by modern UI/UX design standards.
Strictly locked to high-contrast Light Mode.
"""

import streamlit as st

try:
    from Modules.UI.icons import get_svg_icon
except Exception:
    try:
        from .icons import get_svg_icon
    except Exception:
        try:
            from UI.icons import get_svg_icon
        except Exception:
            get_svg_icon = lambda name, size=20, color="currentColor": ""


def get_current_theme() -> str:
    """Strictly returns 'light' mode."""
    return "light"


def set_theme(theme_name: str):
    """No-op: App is permanently in Light mode."""
    pass


def toggle_theme():
    """No-op: App is permanently in Light mode."""
    pass


def render_compact_theme_toggle(key_suffix: str = "main", location: str = None):
    """No-op: Theme toggle button removed as per user request."""
    pass


def render_theme_switcher(location: str = "sidebar"):
    """No-op: Theme toggle button removed as per user request."""
    pass


def apply_theme_css():
    """
    Applies pure, high-contrast Light Mode design system across all pages.
    Features elevated white cards, pill badges, vibrant royal blue primary accents,
    high-contrast typography, and smooth micro-interactions.
    """
    # High-contrast Light Mode Palette (Tailwind Slate / Royal Blue)
    bg_main = "#F8FAFC"              # Slate-50 Clean Canvas
    bg_card = "#FFFFFF"              # Pure White Elevated Card
    bg_card_hover = "#FFFFFF"
    bg_sidebar = "#FFFFFF"           # Clean White Sidebar
    border_color = "#E2E8F0"         # Slate-200 Soft 1px Border
    border_hover = "#CBD5E1"         # Slate-300 Hover Border
    text_primary = "#0F172A"         # Slate-900 Deep High Contrast (16.5:1 ratio)
    text_secondary = "#334155"       # Slate-700 Subtitle (8.5:1 ratio)
    text_muted = "#64748B"           # Slate-500 Captions
    placeholder_color = "#94A3B8"    # Slate-400 Placeholder
    input_bg = "#FFFFFF"             # Pure White Input
    input_border = "#CBD5E1"         # Slate-300 Input Border
    input_text = "#0F172A"           # Slate-900 Input Text
    input_focus = "#2563EB"          # Royal Blue Focus Ring
    primary_btn_bg = "#2563EB"       # Royal Blue-600 Action Button
    primary_btn_hover = "#1D4ED8"    # Royal Blue-700 Hover
    primary_btn_active = "#1E40AF"   # Royal Blue-800 Active
    primary_btn_text = "#FFFFFF"     # Pure White
    badge_bg = "#EFF6FF"             # Soft Blue-50 Badge
    badge_text = "#1D4ED8"           # Cobalt-700
    badge_border = "#DBEAFE"         # Blue-100
    hover_bg = "#F1F5F9"             # Slate-100 Hover
    card_shadow = "0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.02)"
    icon_bg = "#EFF6FF"
    icon_color = "#2563EB"

    theme_css = f"""
    <style>
    /* ================= DESIGN SYSTEM VARIABLES ================= */
    :root {{
        --app-bg: {bg_main};
        --card-bg: {bg_card};
        --card-hover-bg: {bg_card_hover};
        --sidebar-bg: {bg_sidebar};
        --border-color: {border_color};
        --border-hover: {border_hover};
        --text-primary: {text_primary};
        --text-secondary: {text_secondary};
        --text-muted: {text_muted};
        --placeholder-color: {placeholder_color};
        --input-bg: {input_bg};
        --input-border: {input_border};
        --input-text: {input_text};
        --input-focus: {input_focus};
        --primary-btn-bg: {primary_btn_bg};
        --primary-btn-hover: {primary_btn_hover};
        --primary-btn-active: {primary_btn_active};
        --primary-btn-text: {primary_btn_text};
        --badge-bg: {badge_bg};
        --badge-text: {badge_text};
        --badge-border: {badge_border};
        --hover-bg: {hover_bg};
        --card-shadow: {card_shadow};
        --icon-bg: {icon_bg};
        --icon-color: {icon_color};
    }}

    /* Global Canvas & Font Stack */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], [data-testid="stHeader"] {{
        background-color: var(--app-bg) !important;
        color: var(--text-primary) !important;
        font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
    }}

    /* Hide Streamlit input instructions globally */
    [data-testid="stInputInstructions"], .st-emotion-cache-1vt4y43 {{
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }}

    /* ================= SIDEBAR COMPACT DENSITY ================= */
    section[data-testid="stSidebar"], [data-testid="stSidebarContent"], [data-testid="stSidebarUserContent"] {{
        background-color: var(--sidebar-bg) !important;
        border-right: 1px solid var(--border-color) !important;
    }}
    section[data-testid="stSidebar"] > div:first-child, [data-testid="stSidebarContent"] > div:first-child {{
        padding: 0.5rem 0.45rem 0.5rem 0.45rem !important;
    }}
    section[data-testid="stSidebar"] * {{
        color: var(--text-primary) !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {{
        color: var(--text-secondary) !important;
        font-weight: 500;
    }}

    /* Tightened Radio Navigation Density */
    div[data-testid="stRadio"] div[role="radiogroup"] {{
        gap: 1px !important;
    }}
    div[data-testid="stRadio"] div[role="radiogroup"] > label, div[data-testid="stRadio"] label {{
        padding: 2px 7px !important;
        margin: 0 !important;
        border-radius: 6px !important;
        line-height: 1.2 !important;
        min-height: 26px !important;
        display: flex !important;
        align-items: center !important;
        transition: all 0.12s ease !important;
    }}
    div[data-testid="stRadio"] div[role="radiogroup"] > label p, div[data-testid="stRadio"] label p {{
        font-size: 0.81rem !important;
        font-weight: 500 !important;
        line-height: 1.2 !important;
        margin: 0 !important;
    }}
    div[data-testid="stRadio"] div[role="radiogroup"] > label:hover, div[data-testid="stRadio"] label:hover {{
        background-color: var(--icon-bg) !important;
        color: var(--primary-btn-bg) !important;
    }}
    div[data-testid="stRadio"] div[role="radiogroup"] > label:hover p, div[data-testid="stRadio"] label:hover p {{
        color: var(--primary-btn-bg) !important;
    }}

    /* Compact Sidebar Logout Button */
    button[key="sidebar_logout_btn"], .stButton > button[key="sidebar_logout_btn"], button[data-testid="baseButton-secondary"][key="sidebar_logout_btn"] {{
        padding: 0.35rem 0.75rem !important;
        font-size: 0.82rem !important;
        min-height: 32px !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        background-color: var(--card-bg) !important;
        border: 1px solid var(--border-color) !important;
        color: var(--text-secondary) !important;
        margin-top: 0.25rem !important;
        transition: all 0.15s ease !important;
    }}
    button[key="sidebar_logout_btn"]:hover, .stButton > button[key="sidebar_logout_btn"]:hover {{
        border-color: #EF4444 !important;
        color: #EF4444 !important;
        background-color: rgba(239, 68, 68, 0.08) !important;
    }}

    /* Typography Hierarchy */
    h1, h2, h3, h4, h5, h6, [data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2, [data-testid="stMarkdownContainer"] h3 {{
        color: var(--text-primary) !important;
        font-weight: 700 !important;
        letter-spacing: -0.025em !important;
    }}
    p, span, div, [data-testid="stMarkdownContainer"] p {{
        color: var(--text-primary);
    }}
    label, [data-testid="stWidgetLabel"] label, [data-testid="stWidgetLabel"] p {{
        color: var(--text-primary) !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
    }}
    .stCaption, [data-testid="stCaptionContainer"] {{
        color: var(--text-muted) !important;
        font-size: 0.85rem !important;
        font-weight: 400 !important;
    }}

    /* Elevated Cards / Containers */
    [data-testid="stVerticalBlockBorderWrapper"], [data-testid="stVerticalBlockBorderWrapper"] > div, div[data-testid="stContainer"] {{
        background-color: var(--card-bg) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 12px !important;
        box-shadow: var(--card-shadow) !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    [data-testid="stVerticalBlockBorderWrapper"]:hover, div[data-testid="stContainer"]:hover {{
        border-color: var(--border-hover) !important;
        box-shadow: 0 14px 28px -5px rgba(0, 0, 0, 0.08), 0 8px 10px -6px rgba(0, 0, 0, 0.03) !important;
    }}

    /* Inputs & Selects - Crisp Dark Text on Pure White */
    div[data-baseweb="input"], div[data-baseweb="base-input"] {{
        background-color: #FFFFFF !important;
        border: 1.5px solid var(--input-border) !important;
        border-radius: 8px !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }}
    div[data-baseweb="input"]:focus-within, div[data-baseweb="base-input"]:focus-within {{
        border-color: var(--input-focus) !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.20) !important;
    }}
    input[type="text"], input[type="password"], input[type="number"], textarea, div[data-baseweb="input"] input {{
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 0.90rem !important;
        font-weight: 600 !important;
        padding: 0.45rem 0.75rem !important;
    }}
    input::placeholder, textarea::placeholder {{
        color: var(--placeholder-color) !important;
        -webkit-text-fill-color: var(--placeholder-color) !important;
        opacity: 1 !important;
        font-weight: 400 !important;
    }}
    div[data-baseweb="select"] > div {{
        background-color: #FFFFFF !important;
        border: 1.5px solid var(--input-border) !important;
        border-radius: 8px !important;
        color: #0F172A !important;
    }}

    /* Primary Action Buttons (Royal Blue) */
    button[data-testid="baseButton-primary"], button[data-testid="stBaseButton-primary"], button[kind="primary"], .stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] > button {{
        background-color: var(--primary-btn-bg) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 0.5rem 1.25rem !important;
        font-size: 0.90rem !important;
        box-shadow: 0 2px 4px 0 rgba(37, 99, 235, 0.2) !important;
        transition: all 0.15s ease-in-out !important;
    }}
    button[data-testid="baseButton-primary"]:hover, button[data-testid="stBaseButton-primary"]:hover, button[kind="primary"]:hover, .stButton > button[kind="primary"]:hover, div[data-testid="stFormSubmitButton"] > button:hover {{
        background-color: var(--primary-btn-hover) !important;
        box-shadow: 0 4px 10px 0 rgba(37, 99, 235, 0.35) !important;
        transform: translateY(-1px) !important;
    }}
    button[data-testid="baseButton-primary"]:active, button[data-testid="stBaseButton-primary"]:active {{
        background-color: var(--primary-btn-active) !important;
        transform: translateY(0) !important;
    }}
    button[data-testid="baseButton-primary"] *, 
    button[data-testid="stBaseButton-primary"] *, 
    button[kind="primary"] *, 
    .stButton > button[kind="primary"] *,
    div[data-testid="stFormSubmitButton"] > button *,
    div[data-testid="stFormSubmitButton"] p,
    button[data-testid="baseButton-primary"] p {{
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }}

    /* Secondary Buttons */
    button[data-testid="baseButton-secondary"], button[data-testid="stBaseButton-secondary"], button[kind="secondary"], .stButton > button[kind="secondary"], .stButton > button {{
        background-color: var(--card-bg) !important;
        color: var(--text-primary) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 0.90rem !important;
        transition: all 0.15s ease-in-out !important;
    }}
    button[data-testid="baseButton-secondary"]:hover, button[data-testid="stBaseButton-secondary"]:hover, button[kind="secondary"]:hover, .stButton > button:hover {{
        border-color: var(--border-hover) !important;
        background-color: var(--hover-bg) !important;
    }}

    /* Pill Badges & Icon Boxes */
    .ui-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background-color: var(--badge-bg);
        color: var(--badge-text);
        border: 1px solid var(--badge-border);
        font-size: 0.78rem;
        font-weight: 600;
        padding: 0.25rem 0.80rem;
        border-radius: 9999px;
        letter-spacing: 0.2px;
    }}
    .ui-icon-box {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 38px;
        height: 38px;
        border-radius: 8px;
        background-color: var(--icon-bg);
        color: var(--icon-color);
        margin-right: 12px;
        flex-shrink: 0;
    }}

    /* Data Tables & DataFrames */
    [data-testid="stDataFrame"] {{
        background-color: var(--card-bg) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 10px !important;
    }}

    /* Streamlit Chrome Cleanup */
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    .viewerBadge_container__1S-5D, .viewerBadge_link__1S-5D, [data-testid="stStatusWidget"] {{
        display: none !important;
    }}
    [data-testid="stSidebarCollapsedControl"] {{
        display: block !important;
        visibility: visible !important;
        z-index: 100000 !important;
        color: var(--text-primary) !important;
    }}
    </style>
    """
    st.markdown(theme_css, unsafe_allow_html=True)


__all__ = [
    "render_compact_theme_toggle",
    "render_theme_switcher",
    "apply_theme_css",
    "get_current_theme",
    "set_theme",
    "toggle_theme",
]
