"""
==================================================
MERGE FILES TOOL (ENHANCED)
==================================================

PURPOSE
--------------------------------------------------
Lets the user upload any number of Excel files whose columns are
"similarly named" (some files may carry extra columns) and merges
them into a single master Excel file.

KEY BEHAVIOURS & ENHANCEMENTS
--------------------------------------------------
1. Only the FIRST sheet of every uploaded file is read.
2. Completely blank rows are automatically removed prior to processing.
3. Column headers are cleaned before matching (spaces trimmed/collapsed).
4. "Source File" column is styled in italic format across previews and Excel outputs.
5. Merged preview displays sample rows from every uploaded file for instant verification.
6. Dynamic "Remove Duplicates" allows users to select custom column combinations.
7. Per-file column report uses 1-based indexing (starts at 1).
8. Compact and organized Merge Options UI section.
==================================================
"""

import io
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def clean_column_name(name: str) -> str:
    """Trim leading/trailing spaces and collapse internal runs of
    whitespace to a single space, while keeping the words themselves
    (and the single space between them) intact."""
    name = str(name)
    name = re.sub(r"\s+", " ", name)  # collapse any run of whitespace to one space
    return name.strip()


def similar(a: str, b: str) -> float:
    """Simple similarity ratio between two strings (0-1)."""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def find_possible_duplicate_columns(all_columns, threshold: float = 0.82):
    """Given the de-duplicated list of final column names across all
    files, flag pairs that are suspiciously similar (likely typos)
    but not identical — these are NOT merged automatically, just
    surfaced as a warning."""
    flagged = []
    cols = list(all_columns)
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            a, b = cols[i], cols[j]
            if a == b:
                continue
            if similar(a, b) >= threshold:
                flagged.append((a, b))
    return flagged


def style_preview(df: pd.DataFrame):
    """Formats values cleanly for preview display and applies italic style to Source File column."""
    def apply_italic_source(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        if "Source File" in data.columns:
            styles["Source File"] = "font-style: italic;"
        return styles

    display_df = df.copy()
    display_df.index = range(1, len(display_df) + 1)
    return display_df.style.apply(apply_italic_source, axis=None)


def autosize_and_style(writer, sheet_name, df):
    """Bold header row, freeze it, italicize Source File column values,
    and auto-size columns for a polished output file."""
    worksheet = writer.sheets[sheet_name]
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    italic_font = Font(italic=True)

    source_col_idx = None

    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill

        if col_name == "Source File":
            source_col_idx = col_idx

        # Auto-size logic
        max_len = len(str(col_name))
        for value in df[col_name].head(200).tolist():
            if pd.isna(value):
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 60)

    # Apply Italic font to Source File data cells in Excel output
    if source_col_idx:
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=source_col_idx)
            cell.font = italic_font

    worksheet.freeze_panes = "A2"


# --------------------------------------------------
# Main entry point
# --------------------------------------------------
def show():
    st.title("🔗 Merge Files")
    st.caption(
        "Upload any number of Excel files with similarly named columns. "
        "Only the first sheet of each file is used. Column headers are "
        "automatically trimmed of stray spaces before matching."
    )
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Upload Excel files to merge",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Upload two or more Excel files to get started.")
        return

    if len(uploaded_files) == 1:
        st.warning("Only one file was uploaded — upload at least two files to merge.")

    # --------------------------------------------------
    # Redesigned Compact Merge Options UI
    # --------------------------------------------------
    with st.expander("⚙️ Merge options", expanded=True):
        opt_col1, opt_col2 = st.columns(2)
        
        with opt_col1:
            case_insensitive = st.checkbox(
                "Treat columns as the same even if casing differs (e.g. 'Roll No' = 'roll no')",
                value=True,
            )
            add_source_col = st.checkbox(
                "Add a 'Source File' column showing original file name (Italic)",
                value=True,
            )

        with opt_col2:
            merge_mode = st.radio(
                "Merge Strategy:",
                options=[
                    "Union — keep all columns",
                    "Intersection — keep common columns only",
                ],
                index=0,
                horizontal=True,
            )

    # --------------------------------------------------
    # Read each file's first sheet + clean columns & remove blank rows
    # --------------------------------------------------
    frames = []
    per_file_report = []
    read_errors = []

    for f in uploaded_files:
        try:
            df = pd.read_excel(f, sheet_name=0)  # first sheet only
        except Exception as e:
            read_errors.append((f.name, str(e)))
            continue

        # Automatically remove completely blank rows
        df = df.dropna(how="all")

        if df.empty:
            st.warning(f"File **'{f.name}'** contains no valid data rows and was skipped.")
            continue

        original_cols = list(df.columns)
        df.columns = [clean_column_name(c) for c in df.columns]

        if case_insensitive:
            canonical_map = {}
            new_cols = []
            for c in df.columns:
                key = c.lower()
                if key not in canonical_map:
                    canonical_map[key] = c
                new_cols.append(canonical_map[key])
            df.columns = new_cols

        if add_source_col:
            df["Source File"] = f.name

        frames.append(df)
        per_file_report.append(
            {
                "File": f.name,
                "Rows": len(df),
                "Columns Found": len(original_cols),
                "Column Names (cleaned)": ", ".join(
                    [c for c in df.columns if c != "Source File"]
                ),
            }
        )

    if read_errors:
        st.error("Some files could not be read and were skipped:")
        for name, err in read_errors:
            st.write(f"- **{name}**: {err}")

    if not frames:
        st.stop()

    # --------------------------------------------------
    # Column report before merging (1-based index)
    # --------------------------------------------------
    st.subheader("📋 Per-file column report")
    report_df = pd.DataFrame(per_file_report)
    report_df.index = range(1, len(report_df) + 1)  # 1-based indexing for report
    st.dataframe(report_df, use_container_width=True)

    all_columns_seen = []
    for df in frames:
        for c in df.columns:
            if c not in all_columns_seen:
                all_columns_seen.append(c)

    possible_dupes = find_possible_duplicate_columns(
        [c for c in all_columns_seen if c != "Source File"]
    )
    if possible_dupes:
        st.warning("⚠️ Possible near-duplicate column names detected (these were "
                   "kept as separate columns — check for typos):")
        for a, b in possible_dupes:
            st.write(f"- **'{a}'** vs **'{b}'**")

    # --------------------------------------------------
    # Merge — union or intersection
    # --------------------------------------------------
    use_intersection = merge_mode.startswith("Intersection")

    if use_intersection:
        common_cols = set(frames[0].columns)
        for df in frames[1:]:
            common_cols &= set(df.columns)
        common_cols = [c for c in all_columns_seen if c in common_cols]

        if not common_cols or (add_source_col and common_cols == ["Source File"]):
            st.error(
                "No columns are common to all uploaded files, so an intersection "
                "merge would produce an empty result. Switch to Union mode, or "
                "check your files' column names above."
            )
            st.stop()

        dropped_per_file = {
            f.name: [c for c in df.columns if c not in common_cols and c != "Source File"]
            for f, df in zip(uploaded_files, frames)
        }
        any_dropped = any(dropped_per_file.values())
        if any_dropped:
            with st.expander("ℹ️ Columns dropped by intersection mode (click to view)"):
                for fname, cols in dropped_per_file.items():
                    if cols:
                        st.write(f"- **{fname}**: {', '.join(cols)}")

        frames = [df[common_cols] for df in frames]

    merged_df = pd.concat(frames, ignore_index=True, sort=False)

    # --------------------------------------------------
    # Dynamic Remove Duplicates Option
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("🧹 Remove Duplicates Options")
    
    dedup_enable = st.checkbox("Enable Duplicate Removal", value=False)
    
    if dedup_enable:
        dedup_col1, dedup_col2 = st.columns([3, 1])
        with dedup_col1:
            all_available_cols = list(merged_df.columns)
            selected_dedup_cols = st.multiselect(
                "Select column(s) to identify duplicate records (leave empty to check full row):",
                options=all_available_cols,
                default=[],
                help="Select one or more columns (e.g. PRNNumber / Roll No). Duplicate records will be removed based on these column combinations."
            )
        with dedup_col2:
            keep_option = st.selectbox(
                "Keep record:",
                options=["First", "Last"],
                index=0
            )

        before_count = len(merged_df)
        subset_cols = selected_dedup_cols if selected_dedup_cols else None
        keep_val = keep_option.lower()

        merged_df = merged_df.drop_duplicates(subset=subset_cols, keep=keep_val)
        removed_count = before_count - len(merged_df)
        
        if removed_count > 0:
            st.success(f"Successfully removed **{removed_count}** duplicate record(s).")
        else:
            st.info("No duplicate records were found based on the selected criteria.")

    # --------------------------------------------------
    # Merged Preview (Sampling Rows from EACH uploaded file)
    # --------------------------------------------------
    st.markdown("---")
    st.subheader("✅ Merged preview")
    st.write(
        f"**{len(merged_df)} total rows** across **{len(merged_df.columns)} columns**, "
        f"combined from **{len(frames)} file(s)**."
    )

    # Multi-file representative preview sample (up to 10 rows per source file)
    if "Source File" in merged_df.columns:
        preview_samples = []
        for src_file, grp in merged_df.groupby("Source File", sort=False):
            preview_samples.append(grp.head(10))
        preview_df = pd.concat(preview_samples, ignore_index=True)
        st.caption("🔍 Preview showing up to the first 10 rows from each uploaded file:")
    else:
        preview_df = merged_df.head(50)

    st.write(style_preview(preview_df))

    # --------------------------------------------------
    # Build downloadable Excel
    # --------------------------------------------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="Merged")
        autosize_and_style(writer, "Merged", merged_df)
    output.seek(0)

    st.markdown("---")
    st.download_button(
        label="⬇️ Download merged Excel file",
        data=output,
        file_name="merged_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # --------------------------------------------------
    # Keep dashboard KPIs meaningful
    # --------------------------------------------------
    st.session_state.files_uploaded += len(frames)
    st.session_state.students_processed += len(merged_df)