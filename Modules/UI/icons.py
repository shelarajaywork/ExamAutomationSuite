"""
==================================================
ICONS MODULE - SVKM EXAM AUTOMATION SUITE
==================================================
Provides crisp, accessible vector SVG icons (Lucide/Heroicons standard)
replacing informal emojis across the application.
"""

# Dictionary of raw SVG paths / definitions (Lucide standard 24x24 viewBox)
SVG_PATHS = {
    "dashboard": """<path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>""",
    "osm": """<rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/>""",
    "calendar": """<rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/><path d="m9 16 2 2 4-4"/>""",
    "gracing": """<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10"/><path d="m9 12 2 2 4-4"/>""",
    "questionwise": """<path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><line x1="16" x2="8" y1="13" y2="13"/><line x1="16" x2="8" y1="17" y2="17"/><line x1="10" x2="8" y1="9" y2="9"/>""",
    "mail": """<rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>""",
    "rank": """<path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6"/><path d="M18 9h1.5a2.5 2.5 0 0 0 0-5H18"/><path d="M4 22h16"/><path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22"/><path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22"/><path d="M18 2H6v7a6 6 0 0 0 12 0V2Z"/>""",
    "gazette": """<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>""",
    "gradesheet": """<path d="M21.42 10.922a1 1 0 0 0-.019-.838L12.83 2.18a2 2 0 0 0-1.66 0L2.6 10.084a1 1 0 0 0 0 1.832l8.57 7.904a2 2 0 0 0 1.66 0l8.57-7.904a1 1 0 0 0 .02-.994Z"/><path d="M22 10v6"/><path d="M6 12.5V16a6 3 0 0 0 12 0v-3.5"/>""",
    "convocation": """<path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1-2.5-2.5Z"/><path d="M6 6h10"/><path d="M6 10h10"/><path d="M6 14h6"/><path d="m14 18 2 2 4-4"/>""",
    "duplicate": """<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>""",
    "analysis": """<path d="M3 3v18h18"/><path d="m19 9-5 5-4-4-3 3"/>""",
    "data_analysis": """<line x1="18" x2="18" y1="20" y2="10"/><line x1="12" x2="12" y1="20" y2="4"/><line x1="6" x2="6" y1="20" y2="14"/>""",
    "cleaner": """<path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"/>""",
    "merge": """<circle cx="18" cy="18" r="3"/><circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M6 9v6"/><path d="M9 6h6a3 3 0 0 1 3 3v6"/>""",
    "sharepoint": """<path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/>""",
    "settings": """<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>""",
    "lock": """<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>""",
    "key": """<path d="m21 2-2 2m-1.5 1.5L14 9l-1.5-1.5-3 3L11 12l-1.5 1.5-3-3-4 4a5.5 5.5 0 1 0 7.78 7.78l10.72-10.72A2.5 2.5 0 0 0 21 2Z"/>""",
    "logout": """<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" x2="9" y1="12" y2="12"/>""",
    "sun": """<circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/>""",
    "moon": """<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>""",
    "check": """<polyline points="20 6 9 17 4 12"/>""",
    "alert": """<circle cx="12" cy="12" r="10"/><line x1="12" x2="12" y1="8" y2="12"/><line x1="12" x2="12.01" y1="16" y2="16"/>""",
    "shield": """<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>""",
    "arrow_right": """<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>""",
    "building": """<rect width="16" height="20" x="4" y="2" rx="2" ry="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01"/><path d="M16 6h.01"/><path d="M8 10h.01"/><path d="M16 10h.01"/><path d="M8 14h.01"/><path d="M16 14h.01"/>""",
}

# Mapping of tool names to icon keys
TOOL_ICON_MAP = {
    "Dashboard": "dashboard",
    "OSM Dashboard": "osm",
    "Exam Timetable Generator": "calendar",
    "Gracing Checker": "gracing",
    "Questionwise Checker": "questionwise",
    "Letters to Ext. Examiners": "mail",
    "Rank List Generator": "rank",
    "Result Gazette PDF to Excel": "gazette",
    "Grade Sheet to Excel": "gradesheet",
    "Convocation Data Generator": "convocation",
    "Duplicate Checker": "duplicate",
    "Result Analysis": "analysis",
    "Data Analysis": "data_analysis",
    "Data Cleaner": "cleaner",
    "Merge Files": "merge",
    "SharePoint Import": "sharepoint",
    "Settings": "settings",
}


def get_svg_icon(name: str, size: int = 20, color: str = "currentColor", stroke_width: float = 2.0) -> str:
    """Returns an inline SVG string for the specified icon key."""
    path_content = SVG_PATHS.get(name, SVG_PATHS.get("dashboard"))
    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="{stroke_width}" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle; display: inline-block;">{path_content}</svg>"""


def get_tool_icon(tool_name: str, size: int = 20, color: str = "currentColor") -> str:
    """Returns SVG icon for a given tool name."""
    clean_name = tool_name
    for emoji in ["🏠", "📊", "📅", "✅", "📝", "✉️", "🏆", "📑", "🎓", "🔍", "📈", "🧹", "🔗", "☁️", "⚙️"]:
        clean_name = clean_name.replace(emoji, "").strip()
    icon_key = TOOL_ICON_MAP.get(clean_name, "dashboard")
    return get_svg_icon(icon_key, size=size, color=color)
