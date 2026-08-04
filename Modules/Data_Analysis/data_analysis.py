"""
==================================================
DATA ANALYSIS TOOL
==================================================

PURPOSE
--------------------------------------------------
Lets the user upload one or more Excel files belonging to different
report categories (Student Master Data, Gracing Report, GMR Report,
Moderation Report, ...) and, for the categories that are supported,
renders a KPI + chart dashboard.

Only "Student Master Data" is fully wired up right now (built against
the NMC Student Master Data export). The other categories are stubbed
out so files can already be uploaded; their dashboards will be added
once sample files are available.
==================================================
"""

import io

import pandas as pd
import plotly.express as px
import streamlit as st

# --------------------------------------------------
# Report categories available in this tool
# --------------------------------------------------
REPORT_TYPES = [
    "🎓 Student Master Data",
    "📋 Gracing Report",
    "📄 GMR Report",
    "🧮 Moderation Report",
]

PLOTLY_TEMPLATE = "plotly_white"
COLOR_SEQUENCE = px.colors.qualitative.Set2


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def _find_col(df: pd.DataFrame, *candidates: str):
    """Case/space-insensitive column lookup so minor header variations
    (extra spaces, different casing, punctuation) don't break the tool."""
    normalized = {
        "".join(ch for ch in str(c).lower() if ch.isalnum()): c for c in df.columns
    }
    for cand in candidates:
        key = "".join(ch for ch in cand.lower() if ch.isalnum())
        if key in normalized:
            return normalized[key]
    return None


@st.cache_data(show_spinner=False)
def _read_excel_files(file_bytes_list, file_names):
    """Read & concatenate one or more uploaded Excel files into a single
    DataFrame. Cached on the raw bytes so re-renders don't re-parse."""
    frames = []
    for content, name in zip(file_bytes_list, file_names):
        df = pd.read_excel(io.BytesIO(content))
        df["__source_file__"] = name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def _metric_row(metrics):
    """metrics: list of (label, value) tuples -> rendered as st.metric cards."""
    cols = st.columns(len(metrics))
    for col, (label, value) in zip(cols, metrics):
        col.metric(label, value)


def _bar(df_counts, x, y, title, horizontal=False):
    if horizontal:
        fig = px.bar(
            df_counts, x=y, y=x, orientation="h", title=title,
            color=x, color_discrete_sequence=COLOR_SEQUENCE, text=y,
        )
    else:
        fig = px.bar(
            df_counts, x=x, y=y, title=title,
            color=x, color_discrete_sequence=COLOR_SEQUENCE, text=y,
        )
    fig.update_layout(template=PLOTLY_TEMPLATE, showlegend=False, title_x=0.02)
    fig.update_traces(textposition="outside")
    return fig


def _pie(df_counts, names, values, title):
    fig = px.pie(
        df_counts, names=names, values=values, title=title, hole=0.45,
        color_discrete_sequence=COLOR_SEQUENCE,
    )
    fig.update_layout(template=PLOTLY_TEMPLATE, title_x=0.02)
    fig.update_traces(textinfo="percent+label")
    return fig


# --------------------------------------------------
# Student Master Data dashboard
# --------------------------------------------------
def _student_master_data_dashboard(df: pd.DataFrame):
    # Resolve columns defensively (headers can shift slightly between exports)
    col_student = _find_col(df, "Student Number", "Student ID", "PRN")
    col_program = _find_col(df, "Program Name")
    col_gender = _find_col(df, "Gender")
    col_div = _find_col(df, "Div")
    col_status = _find_col(df, "Final Status")
    col_religion = _find_col(df, "Religion")
    col_caste = _find_col(df, "Social Class", "Caste")
    col_org = _find_col(df, "Org. Name", "Org Name")
    col_domicile = _find_col(df, "Domicile")
    col_nationality = _find_col(df, "Nationality")
    col_disability = _find_col(df, "Disability")
    col_sess = _find_col(df, "Academic Sess", "Academic Session", "Semester", "Sem")

    total_rows = len(df)

    # --------------------------
    # De-duplicate students
    # --------------------------
    # The master data export repeats a student once per semester/session
    # they were registered in during the academic year, so a raw row count
    # overstates headcount. Build a "one row per student" view for every
    # demographic KPI/chart, keeping each student's most recent
    # semester/session row (so status fields like Div / Final Status reflect
    # their latest state rather than an arbitrary earlier one).
    if col_student:
        if col_sess:
            df_sorted = df.sort_values(by=[col_student, col_sess], ascending=[True, False])
        else:
            df_sorted = df
        df_students = df_sorted.drop_duplicates(subset=col_student, keep="first").copy()
        total_students = df_students[col_student].nunique()
        dedup_note = total_rows != total_students
    else:
        # No reliable student identifier found — fall back to raw rows,
        # but flag it so the numbers aren't silently treated as unique.
        df_students = df
        total_students = total_rows
        dedup_note = False

    st.success(f"Loaded **{total_students:,}** unique student(s) from the uploaded file(s).")
    if dedup_note:
        st.caption(
            f"ℹ️ The file contains {total_rows:,} rows in total — students appear more than once "
            f"because they're repeated across semesters/sessions. All figures below count each "
            f"student once, using their most recent semester record."
        )

    # --------------------------
    # KPI cards
    # --------------------------
    st.markdown("### 📌 Key Metrics")

    total_programs = df_students[col_program].nunique() if col_program else "—"
    total_divisions = df_students[col_div].nunique() if col_div else "—"

    attending_pct = "—"
    if col_status:
        vc = df_students[col_status].astype(str).str.strip().value_counts()
        attending = vc.get("Attending", 0)
        attending_pct = f"{(attending / total_students * 100):.1f}%" if total_students else "—"

    _metric_row([
        ("Unique Students", f"{total_students:,}"),
        ("Programs Offered", total_programs),
        ("Divisions", total_divisions),
        ("Attending %", attending_pct),
    ])

    if col_gender:
        gender_vc = df_students[col_gender].astype(str).str.strip().replace({"nan": None}).dropna().value_counts()
        male = gender_vc.get("Male", 0)
        female = gender_vc.get("Female", 0)
        other = int(gender_vc.sum() - male - female)
        _metric_row([
            ("Male Students", f"{male:,}"),
            ("Female Students", f"{female:,}"),
            ("Other / Unspecified", f"{other:,}"),
            ("Institute", df_students[col_org].dropna().iloc[0] if col_org and df_students[col_org].notna().any() else "—"),
        ])

    st.markdown("---")

    # --------------------------
    # Semester-wise headcount (uses the raw, non-deduplicated rows on purpose,
    # since that's exactly what tells us how many students were enrolled in
    # each semester/session)
    # --------------------------
    if col_sess and col_student:
        st.markdown("### 📆 Semester-wise Student Count")

        sorted_sess = sorted(df[col_sess].dropna().unique(), key=lambda v: (str(type(v)), v))
        sess_label = {code: f"Semester {i + 1}" for i, code in enumerate(sorted_sess)}
        df_sess = df.copy()
        df_sess["__semester_label__"] = df_sess[col_sess].map(sess_label)

        sem_counts = (
            df_sess.groupby("__semester_label__")[col_student].nunique()
            .reindex([sess_label[c] for c in sorted_sess])
            .rename_axis("Semester").reset_index(name="Students")
        )

        sc1, sc2 = st.columns([1.3, 1])
        with sc1:
            fig = px.bar(
                sem_counts, x="Semester", y="Students", title="Unique Students per Semester",
                color="Semester", color_discrete_sequence=COLOR_SEQUENCE, text="Students",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE, showlegend=False, title_x=0.02)
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

        with sc2:
            st.dataframe(sem_counts, use_container_width=True, hide_index=True)

        # Semester x Program breakdown (unique students, not row counts)
        if col_program:
            st.markdown("#### Semester-wise Program Strength")
            sem_prog = (
                df_sess.groupby(["__semester_label__", col_program])[col_student].nunique()
                .reset_index(name="Students")
                .rename(columns={"__semester_label__": "Semester", col_program: "Program"})
            )
            fig = px.bar(
                sem_prog, x="Semester", y="Students", color="Program", barmode="stack",
                title="Unique Students per Semester, by Program",
                category_orders={"Semester": [sess_label[c] for c in sorted_sess]},
                color_discrete_sequence=COLOR_SEQUENCE,
            )
            fig.update_layout(template=PLOTLY_TEMPLATE, title_x=0.02, legend_title="Program")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
    elif col_sess and not col_student:
        st.info(
            "A semester/session column was found, but no student identifier column "
            "(Student Number / Student ID / PRN) — can't compute unique semester-wise counts."
        )

    # --------------------------
    # Charts row 1: Program strength + Gender split
    # --------------------------
    c1, c2 = st.columns([1.4, 1])

    with c1:
        if col_program:
            counts = (
                df_students[col_program].astype(str).str.strip().value_counts()
                .rename_axis("Program").reset_index(name="Students")
                .sort_values("Students", ascending=True)
            )
            st.plotly_chart(
                _bar(counts, "Program", "Students", "Unique Students per Program", horizontal=True),
                use_container_width=True,
            )
        else:
            st.info("No 'Program Name' column found in the uploaded data.")

    with c2:
        if col_gender:
            counts = (
                df_students[col_gender].astype(str).str.strip().value_counts()
                .rename_axis("Gender").reset_index(name="Students")
            )
            st.plotly_chart(_pie(counts, "Gender", "Students", "Gender Distribution"), use_container_width=True)
        else:
            st.info("No 'Gender' column found in the uploaded data.")

    # --------------------------
    # Charts row 2: Division strength + Final Status
    # --------------------------
    c3, c4 = st.columns(2)

    with c3:
        if col_div:
            counts = (
                df_students[col_div].astype(str).str.strip().value_counts()
                .rename_axis("Division").reset_index(name="Students")
                .sort_values("Division")
            )
            st.plotly_chart(_bar(counts, "Division", "Students", "Unique Students per Division"),
                             use_container_width=True)
        else:
            st.info("No 'Div' column found in the uploaded data.")

    with c4:
        if col_status:
            counts = (
                df_students[col_status].astype(str).str.strip().value_counts()
                .rename_axis("Status").reset_index(name="Students")
            )
            st.plotly_chart(_pie(counts, "Status", "Students", "Final Status (Attending vs Not Attending)"),
                             use_container_width=True)
        else:
            st.info("No 'Final Status' column found in the uploaded data.")

    # --------------------------
    # Charts row 3: Religion + Social Class / Caste
    # --------------------------
    c5, c6 = st.columns(2)

    with c5:
        if col_religion:
            counts = (
                df_students[col_religion].astype(str).str.strip().replace({"": None, "nan": None}).dropna()
                .value_counts().rename_axis("Religion").reset_index(name="Students")
            )
            st.plotly_chart(_pie(counts, "Religion", "Students", "Religion Distribution"), use_container_width=True)
        else:
            st.info("No 'Religion' column found in the uploaded data.")

    with c6:
        if col_caste:
            counts = (
                df_students[col_caste].astype(str).str.strip().replace({"": None, "nan": None}).dropna()
                .value_counts().rename_axis("Category").reset_index(name="Students")
                .sort_values("Students", ascending=True)
            )
            st.plotly_chart(
                _bar(counts, "Category", "Students", "Category-wise Distribution", horizontal=True),
                use_container_width=True,
            )
        else:
            st.info("No 'Social Class' / 'Caste' column found in the uploaded data.")

    # --------------------------
    # Program x Gender breakdown (stacked)
    # --------------------------
    if col_program and col_gender:
        st.markdown("---")
        st.markdown("### 👥 Program-wise Gender Breakdown")
        pivot = (
            df_students.groupby([col_program, col_gender]).size()
            .reset_index(name="Students")
        )
        fig = px.bar(
            pivot, x=col_program, y="Students", color=col_gender, barmode="stack",
            title="Program-wise Gender Breakdown (unique students)", color_discrete_sequence=COLOR_SEQUENCE,
        )
        fig.update_layout(template=PLOTLY_TEMPLATE, title_x=0.02, xaxis_title="Program", legend_title="Gender")
        fig.update_xaxes(tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)

    # --------------------------
    # Extra info: Domicile / Nationality / Disability (only if present & varied)
    # --------------------------
    extras = []
    for label, col in [("Domicile", col_domicile), ("Nationality", col_nationality), ("Disability", col_disability)]:
        if col and df_students[col].astype(str).str.strip().replace({"": None, "nan": None}).dropna().nunique() > 1:
            extras.append((label, col))

    if extras:
        st.markdown("---")
        st.markdown("### 🌐 Other Demographics")
        cols = st.columns(len(extras))
        for c, (label, col) in zip(cols, extras):
            with c:
                counts = (
                    df_students[col].astype(str).str.strip().replace({"": None, "nan": None}).dropna()
                    .value_counts().rename_axis(label).reset_index(name="Students")
                )
                st.plotly_chart(_pie(counts, label, "Students", label), use_container_width=True)

    # --------------------------
    # Raw data + download
    # --------------------------
    with st.expander(f"🔎 View underlying data ({total_rows:,} raw rows, {total_students:,} unique students)"):
        view_choice = st.radio(
            "Show:", ["Unique students (deduplicated)", "All raw rows"],
            horizontal=True, key="_smd_view_choice",
        )
        st.dataframe(df_students if view_choice.startswith("Unique") else df, use_container_width=True)


# --------------------------------------------------
# Placeholder dashboard for report types not yet built
# --------------------------------------------------
def _coming_soon_dashboard(df: pd.DataFrame, report_label: str):
    st.success(f"**{len(df):,}** row(s) loaded from the uploaded file(s) for **{report_label}**.")
    st.info(
        f"📊 The KPI/chart dashboard for **{report_label}** hasn't been built yet — "
        "share a sample export of this report and it'll be added here, following the same "
        "layout as the Student Master Data dashboard."
    )
    with st.expander("🔎 View uploaded data"):
        st.dataframe(df, use_container_width=True)


# --------------------------------------------------
# Main entry point (called from app.py)
# --------------------------------------------------
def show():
    st.title("📊 Data Analysis")
    st.caption("Upload examination / student data exports to view KPIs and visual analysis.")
    st.markdown("---")

    tabs = st.tabs(REPORT_TYPES)

    for tab, report_type in zip(tabs, REPORT_TYPES):
        with tab:
            uploaded_files = st.file_uploader(
                f"Upload {report_type.split(' ', 1)[1]} file(s) (.xlsx / .xls)",
                type=["xlsx", "xls"],
                accept_multiple_files=True,
                key=f"uploader_{report_type}",
            )

            if not uploaded_files:
                st.info(f"Upload one or more {report_type.split(' ', 1)[1]} files to see the analysis.")
                continue

            try:
                file_bytes = [f.getvalue() for f in uploaded_files]
                file_names = [f.name for f in uploaded_files]
                df = _read_excel_files(file_bytes, file_names)
            except Exception as e:
                st.error(f"Could not read the uploaded file(s): {e}")
                continue

            if df.empty:
                st.warning("No data found in the uploaded file(s).")
                continue

            st.session_state.files_uploaded += len(uploaded_files)

            if report_type == "🎓 Student Master Data":
                _student_master_data_dashboard(df)
            else:
                _coming_soon_dashboard(df, report_type.split(" ", 1)[1])