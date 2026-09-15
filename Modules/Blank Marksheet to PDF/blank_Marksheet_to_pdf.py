"""
Excel Marksheet -> PDF Converter
=================================
Streamlit app that:
  1. Accepts one or more uploaded Excel files (.xls, .xlsx, or the
     SpreadsheetML/"XML Spreadsheet 2003" .XLS files exported by many
     college ERPs).
  2. Swaps the header text in every matching cell: "IN WORDS" becomes
     "IN FIG." and "IN FIG." becomes "IN WORDS" (a true swap - each
     cell is evaluated against its own original text).
  3. Sets column E width = 11 and column F width = 30 (Excel character
     width units).
  4. Sets the print/page setup to A4, "narrow" margins (0.5" on every
     side) plus a 1 cm gutter added to the left/binding margin.
  5. Removes the sheet-name header (e.g. "Sheet1") and the
     auto page-number footer (e.g. "Page 1") that would otherwise
     print on every page.
  6. Extends the "Examiner's Sign / Date / Page X of Y" footer cell so
     it stays fully visible even though column E is now narrower.
  7. Middle-aligns (vertically centers) every cell's content before
     export.
  8. Optionally verifies each marksheet against an uploaded master
     student-subject data file: for the subject that marksheet
     belongs to, checks that every student the master file says
     should be there is present (by SAP/Student Number), and flags
     any student in the marksheet who isn't in the master's list for
     that subject. If any student is missing, that file's PDF is
     blocked (not generated) until the marksheet is fixed.
  9. Converts each modified workbook to a PDF using the **Microsoft
     Excel application itself** (via COM automation), so no
     LibreOffice install is needed - the PDF is produced by Excel's
     own print engine, using the exact copy of Excel already on the
     PC.
  10. Names each output PDF after its source file (same base name,
     .pdf extension) and offers them for download individually or as
     a single ZIP.

Requirements
------------
This app automates the real Microsoft Excel application to do the
Excel -> PDF rendering, so it requires:
  - Microsoft Excel installed on the machine running this Streamlit
    app (this only works on **Windows**; Excel does not expose this
    COM automation interface on macOS).
  - The `pywin32` Python package, which lets Python drive Excel:

        pip install pywin32

Python packages (requirements.txt):
    streamlit
    openpyxl
    pywin32          (Windows only)

Run with:
    streamlit run streamlit_app.py

Notes
-----
- Excel is launched once, hidden, and reused for every file in a
  batch (much faster than relaunching it per file), then closed again
  at the end - including if an error occurs partway through.
- Each file is opened, saved, and closed with alerts/prompts
  suppressed, so no dialog boxes should ever appear on screen or
  block the conversion.
- Because this drives a real, licensed copy of Excel, only run one
  conversion batch at a time per machine (Excel automation is not
  designed for many simultaneous/parallel users on the same PC).
"""

import io
import re
import sys
import tempfile
import zipfile
from copy import copy
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins

IS_WINDOWS = sys.platform.startswith("win")

# pywin32 only exists (and only works) on Windows, so import it lazily/
# defensively - the rest of the app can still be imported and its
# helper functions unit-tested on any OS.
try:
    import pythoncom
    import win32com.client as win32

    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# Cell text to find/replace (exact, whitespace-trimmed match).
FIND_TEXT = "IN WORDS"
REPLACE_TEXT = "IN FIG."

# Column widths requested (Excel "characters" width units).
COLUMN_WIDTHS = {
    "E": 11,
    "F": 30,
}

# Page setup: A4, narrow margins (Word's "Narrow" preset = 0.5in all
# round) plus a 1 cm gutter added to the left/binding edge.
CM_TO_INCH = 1 / 2.54
NARROW_MARGIN_IN = 0.5
GUTTER_CM = 1.0
GUTTER_IN = GUTTER_CM * CM_TO_INCH

# Excel constants (from the Excel object model - hard-coded here so we
# don't depend on the generated constants module being pre-built).
XL_OPEN_XML_WORKBOOK = 51  # xlOpenXMLWorkbook -> .xlsx
XL_TYPE_PDF = 0  # xlTypePDF, for ExportAsFixedFormat

# Master file column headers we look for (case-insensitive, exact match
# preferred, falling back to substring match if the export renames
# things slightly).
MASTER_MODULE_COL = ("Module Abbreviation",)
MASTER_MODULE_DESC_COL = ("Module Description",)
MASTER_SAP_COL = ("Student Number",)
MASTER_ROLL_COL = ("Student Roll Number",)
MASTER_NAME_COL = ("Student name", "Student Name")

# Fixed column positions (1-based) in the marksheet template itself:
# A = SR No, B = SAP No, C = ROLL No, D = Student name.
MARKSHEET_SAP_COL = 2
MARKSHEET_ROLL_COL = 3
MARKSHEET_NAME_COL = 4

# A SAP/Student Number is treated as a plausible student id if it's a
# run of 6-15 digits (rules out header labels like "SAP"/"No" and
# stray blank cells).
ID_PATTERN = re.compile(r"^\d{6,15}$")

# The marksheet's title row contains the subject code in parentheses,
# e.g. "( UIDLL222 ) DLLE 3" -> subject code "UIDLL222".
SUBJECT_CODE_PATTERN = re.compile(r"\(\s*([A-Za-z0-9]+)\s*\)")


# --------------------------------------------------------------------------
# Master-file verification
# --------------------------------------------------------------------------

def normalize_id(value) -> str:
    """Turn a SAP/roll number cell value (which may come back as a str,
    int, or float depending on how it was stored) into a clean string
    for comparison, e.g. 53013250003.0 -> '53013250003'."""
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value).strip()
    if isinstance(value, int):
        return str(value)
    return str(value).replace("\xa0", " ").strip()


def normalize_code(value) -> str:
    """Normalize a subject/module code for comparison."""
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip().upper()


def _find_column(header_row, *name_variants):
    """Find the 0-based index of the column whose header matches one of
    name_variants (case-insensitive), preferring an exact match and
    falling back to a substring match."""
    normalized = [
        (i, str(v).strip().lower()) for i, v in enumerate(header_row) if v is not None
    ]
    variants_lower = [v.lower() for v in name_variants]
    for i, text in normalized:
        if text in variants_lower:
            return i
    for i, text in normalized:
        if any(v in text for v in variants_lower):
            return i
    return None


def load_master_data(file_bytes: bytes) -> dict:
    """Parse an uploaded master student-subject file into:
        { subject_code: {"description": str, "students": [ {sap, roll, name}, ... ]} }
    Scans every worksheet in the file; a sheet is used if its header
    row contains at least a module-abbreviation column and a
    student-number column.
    """
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    data = {}
    try:
        for ws in wb.worksheets:
            rows_iter = ws.iter_rows(values_only=True)
            try:
                header = next(rows_iter)
            except StopIteration:
                continue

            module_idx = _find_column(header, *MASTER_MODULE_COL)
            sap_idx = _find_column(header, *MASTER_SAP_COL)
            if module_idx is None or sap_idx is None:
                continue  # doesn't look like the master data table

            desc_idx = _find_column(header, *MASTER_MODULE_DESC_COL)
            roll_idx = _find_column(header, *MASTER_ROLL_COL)
            name_idx = _find_column(header, *MASTER_NAME_COL)

            for row in rows_iter:
                if row is None or module_idx >= len(row) or sap_idx >= len(row):
                    continue
                code = normalize_code(row[module_idx])
                sap = normalize_id(row[sap_idx])
                if not code or not sap:
                    continue
                entry = data.setdefault(code, {"description": "", "students": []})
                if not entry["description"] and desc_idx is not None and desc_idx < len(row):
                    desc_val = row[desc_idx]
                    if desc_val:
                        entry["description"] = str(desc_val).replace("\xa0", " ").strip()
                entry["students"].append(
                    {
                        "sap": sap,
                        "roll": normalize_id(row[roll_idx]) if roll_idx is not None and roll_idx < len(row) else "",
                        "name": (
                            str(row[name_idx]).strip()
                            if name_idx is not None and name_idx < len(row) and row[name_idx]
                            else ""
                        ),
                    }
                )
    finally:
        wb.close()
    return data


def extract_subject_code(xlsx_path: Path):
    """Read the subject code (e.g. 'UIDLL222') out of the marksheet's
    own title row, by looking for the first '( CODE )' pattern in the
    first few rows of the sheet."""
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 10:
                break
            for value in row:
                if isinstance(value, str):
                    m = SUBJECT_CODE_PATTERN.search(value)
                    if m:
                        return m.group(1).strip().upper()
    finally:
        wb.close()
    return None


def extract_marksheet_students(xlsx_path: Path) -> list:
    """Read the list of students actually present in the marksheet, by
    scanning column B (SAP No) for plausible student-id values."""
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    students = []
    try:
        ws = wb.worksheets[0]
        for row in ws.iter_rows(values_only=True):
            if len(row) < MARKSHEET_NAME_COL:
                continue
            sap = normalize_id(row[MARKSHEET_SAP_COL - 1])
            if not ID_PATTERN.match(sap):
                continue
            roll = normalize_id(row[MARKSHEET_ROLL_COL - 1]) if len(row) >= MARKSHEET_ROLL_COL else ""
            name_val = row[MARKSHEET_NAME_COL - 1] if len(row) >= MARKSHEET_NAME_COL else None
            name = str(name_val).strip() if name_val else ""
            students.append({"sap": sap, "roll": roll, "name": name})
    finally:
        wb.close()
    return students


def verify_against_master(xlsx_path: Path, master_data: dict) -> dict:
    """Compare the students actually present in this marksheet against
    the master file's expected list for the same subject.

    Returns a dict with:
        subject_code, subject_description, verified (bool - False if
        the subject code couldn't be found in the master file at all),
        master_count, marksheet_count, missing (list of master
        students absent from the marksheet), extra (list of marksheet
        students absent from the master's list for that subject).
    """
    code = extract_subject_code(xlsx_path)
    marksheet_students = extract_marksheet_students(xlsx_path)
    marksheet_sap_set = {s["sap"] for s in marksheet_students}

    result = {
        "subject_code": code,
        "subject_description": "",
        "verified": False,
        "master_count": 0,
        "marksheet_count": len(marksheet_students),
        "missing": [],
        "extra": [],
    }

    if code is None or code not in master_data:
        return result  # no baseline to check against

    subject = master_data[code]
    master_students = subject["students"]
    master_sap_set = {m["sap"] for m in master_students}

    result["verified"] = True
    result["subject_description"] = subject["description"]
    result["master_count"] = len(master_students)
    result["missing"] = [m for m in master_students if m["sap"] not in marksheet_sap_set]
    result["extra"] = [s for s in marksheet_students if s["sap"] not in master_sap_set]
    return result


# --------------------------------------------------------------------------
# Excel COM automation helpers
# --------------------------------------------------------------------------

class ExcelSession:
    """Wraps one hidden Microsoft Excel COM instance, reused across all
    files in a batch. Always quits Excel on exit, even if an error
    occurs, so no orphaned EXCEL.EXE process is left running."""

    def __enter__(self):
        if not IS_WINDOWS:
            raise RuntimeError(
                "Converting via Microsoft Excel requires Windows (Excel's "
                "COM automation interface is not available on macOS/Linux)."
            )
        if not PYWIN32_AVAILABLE:
            raise RuntimeError(
                "The 'pywin32' package is required to automate Excel. "
                "Install it with: pip install pywin32"
            )

        pythoncom.CoInitialize()
        self.app = win32.gencache.EnsureDispatch("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.ScreenUpdating = False
        self.app.AskToUpdateLinks = False
        self.app.EnableEvents = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.app.Quit()
        finally:
            del self.app
            pythoncom.CoUninitialize()
        return False

    def convert_to_xlsx(self, src_path: Path, out_dir: Path) -> Path:
        """Open src_path (.xls, .xlsx, or SpreadsheetML "XML Spreadsheet
        2003" saved with an .xls extension) in Excel and save it as a
        normal .xlsx file in out_dir."""
        out_path = out_dir / f"{src_path.stem}.xlsx"
        wb = self.app.Workbooks.Open(str(src_path.resolve()))
        try:
            wb.SaveAs(str(out_path.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
        finally:
            wb.Close(SaveChanges=False)
        return out_path

    def convert_to_pdf(self, xlsx_path: Path, out_dir: Path) -> Path:
        """Open xlsx_path in Excel and export it to PDF using Excel's own
        print engine (so the page setup we wrote with openpyxl - paper
        size, margins, fit-to-width, headers/footers - is honoured
        exactly as Excel itself would print it)."""
        out_path = out_dir / f"{xlsx_path.stem}.pdf"
        wb = self.app.Workbooks.Open(str(xlsx_path.resolve()))
        try:
            wb.ExportAsFixedFormat(XL_TYPE_PDF, str(out_path.resolve()))
        finally:
            wb.Close(SaveChanges=False)
        return out_path


def excel_is_available() -> tuple[bool, str]:
    """Best-effort check that Excel automation will work, without
    leaving a stray Excel process running. Returns (ok, message)."""
    if not IS_WINDOWS:
        return False, "This machine is not Windows - Excel automation is unavailable here."
    if not PYWIN32_AVAILABLE:
        return False, "The 'pywin32' package is not installed. Run: pip install pywin32"
    try:
        with ExcelSession():
            pass
        return True, "Microsoft Excel is available and was launched successfully."
    except Exception as exc:  # noqa: BLE001
        return False, f"Could not start Microsoft Excel: {exc}"


# --------------------------------------------------------------------------
# Workbook editing (unchanged - still done with openpyxl on the .xlsx
# file, regardless of which application renders the final PDF)
# --------------------------------------------------------------------------

def modify_workbook(xlsx_path: Path) -> int:
    """Apply the text swap, column widths, page setup, header/footer
    removal, and footer-visibility fix to the given .xlsx file in
    place. Returns the number of cells swapped."""
    wb = load_workbook(xlsx_path)
    swapped = 0

    for ws in wb.worksheets:
        # 1. Swap "IN WORDS" <-> "IN FIG." wherever either appears
        # verbatim. Each cell is evaluated independently against its
        # own original value, so this is a true swap, not a two-step
        # replace that would collapse both into the same text.
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                text = cell.value.strip()
                if text == FIND_TEXT:
                    cell.value = REPLACE_TEXT
                    swapped += 1
                elif text == REPLACE_TEXT:
                    cell.value = FIND_TEXT
                    swapped += 1

        # 2. Column widths.
        for col_letter, width in COLUMN_WIDTHS.items():
            ws.column_dimensions[col_letter].width = width

        # 3. Page setup: A4, narrow margins + 1cm gutter on the left.
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "portrait"
        # Fit all columns onto one page width; let rows flow onto as
        # many pages tall as needed (existing page breaks in the sheet
        # are preserved).
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.page_margins = PageMargins(
            left=NARROW_MARGIN_IN + GUTTER_IN,
            right=NARROW_MARGIN_IN,
            top=NARROW_MARGIN_IN,
            bottom=NARROW_MARGIN_IN,
            header=0.3,
            footer=0.3,
        )

        # 4. Remove the sheet-name header (e.g. "Sheet1") and the
        # auto page-number footer (e.g. "Page 1") that print by
        # default. Clear every header/footer variant.
        ws.oddHeader.center.text = None
        ws.oddHeader.left.text = None
        ws.oddHeader.right.text = None
        ws.evenHeader.center.text = None
        ws.evenHeader.left.text = None
        ws.evenHeader.right.text = None
        ws.firstHeader.center.text = None
        ws.firstHeader.left.text = None
        ws.firstHeader.right.text = None

        ws.oddFooter.center.text = None
        ws.oddFooter.left.text = None
        ws.oddFooter.right.text = None
        ws.evenFooter.center.text = None
        ws.evenFooter.left.text = None
        ws.evenFooter.right.text = None
        ws.firstFooter.center.text = None
        ws.firstFooter.left.text = None
        ws.firstFooter.right.text = None

        # 6. Middle-align all data: set vertical alignment to "center"
        # on every used cell, while leaving each cell's existing
        # horizontal alignment, wrap-text, and other alignment
        # settings untouched.
        for row in ws.iter_rows():
            for cell in row:
                alignment = copy(cell.alignment)
                alignment.vertical = "center"
                cell.alignment = alignment

        # 7. Make sure the "Examiner's Sign / Date / Page X of Y"
        # footer line is fully visible. That text lives in a merged
        # cell that was originally sized to fit alongside the old,
        # wider "IN WORDS"/"IN FIG." columns. Now that column E has
        # been narrowed, extend any such merged cell out to the last
        # column so the full text has room to display.
        last_col = ws.max_column
        for merged_range in list(ws.merged_cells.ranges):
            top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            if (
                isinstance(top_left.value, str)
                and "Page :" in top_left.value
                and merged_range.max_col < last_col
            ):
                value = top_left.value
                alignment = copy(top_left.alignment)
                ws.unmerge_cells(str(merged_range))
                ws.merge_cells(
                    start_row=merged_range.min_row,
                    start_column=merged_range.min_col,
                    end_row=merged_range.max_row,
                    end_column=last_col,
                )
                new_top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                new_top_left.value = value
                new_top_left.alignment = alignment

    wb.save(xlsx_path)
    return swapped


def process_one_file(uploaded_file, work_dir: Path, excel: ExcelSession, master_data: dict = None) -> dict:
    """Run the full pipeline for a single uploaded file, using the
    given (already-open) ExcelSession for both conversion steps.

    If master_data is provided, the file is verified against it first:
    if any student the master file lists for that subject is missing
    from the marksheet, conversion is skipped and the file is marked
    "blocked" with the verification details attached.

    Returns a dict with keys: name, pdf_bytes, replaced_count, error,
    blocked, verification.
    """
    original_name = uploaded_file.name
    stem = Path(original_name).stem
    result = {
        "name": original_name,
        "pdf_bytes": None,
        "replaced_count": 0,
        "error": None,
        "blocked": False,
        "verification": None,
    }

    try:
        # Save the upload to disk with its original extension so Excel
        # can open it correctly.
        src_path = work_dir / original_name
        src_path.write_bytes(uploaded_file.getvalue())

        # Step 1: normalise to .xlsx (handles legacy .xls, SpreadsheetML
        # "XML Spreadsheet 2003" .XLS exports, and native .xlsx alike).
        xlsx_dir = work_dir / "xlsx"
        xlsx_dir.mkdir(exist_ok=True)
        xlsx_path = excel.convert_to_xlsx(src_path, xlsx_dir)

        # Step 1b: verify against the master file, if one was provided.
        if master_data is not None:
            verification = verify_against_master(xlsx_path, master_data)
            result["verification"] = verification
            if verification["missing"]:
                result["blocked"] = True
                return result

        # Step 2: apply the requested edits.
        replaced = modify_workbook(xlsx_path)
        result["replaced_count"] = replaced

        # Step 3: render to PDF via Excel's own print engine.
        pdf_dir = work_dir / "pdf"
        pdf_dir.mkdir(exist_ok=True)
        pdf_path = excel.convert_to_pdf(xlsx_path, pdf_dir)

        # Rename to match the original file's base name exactly.
        final_pdf_path = pdf_dir / f"{stem}.pdf"
        if pdf_path != final_pdf_path:
            pdf_path.replace(final_pdf_path)

        result["pdf_bytes"] = final_pdf_path.read_bytes()

    except Exception as exc:  # noqa: BLE001 - surface any error to the UI
        result["error"] = str(exc)

    return result


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

@st.cache_data(show_spinner="Reading master file...")
def load_master_data_cached(file_bytes: bytes) -> dict:
    return load_master_data(file_bytes)


def show():
    st.title("📄 Blank Marksheet to PDF")
    st.write(
        "Upload one or more marksheet Excel files. For each file, this tool will:\n"
        f"1. Swap **\"{FIND_TEXT}\"** \u2194 **\"{REPLACE_TEXT}\"** wherever either appears\n"
        f"2. Set column **E** width to **{COLUMN_WIDTHS['E']}** and column **F** "
        f"width to **{COLUMN_WIDTHS['F']}**\n"
        f"3. Set the page to **A4**, narrow margins, with a **{GUTTER_CM:.0f} cm** "
        "gutter\n"
        "4. Remove the sheet-name header (e.g. \"Sheet1\") and page-number "
        "footer from each page\n"
        "5. Keep the Examiner's Sign / Date / Page X of Y footer line fully visible\n"
        "6. Middle-align (vertically center) all cell content\n"
        "7. Export a PDF with the **same file name** as the original"
    )

    with st.expander("⚙️ Setup requirements (read once)"):
        st.markdown(
            "This app renders the PDFs using **Microsoft Excel itself** (via "
            "COM automation), so it requires:\n\n"
            "- **Windows**, with **Microsoft Excel installed**.\n"
            "- The `pywin32` package: `pip install pywin32`\n\n"
            "Excel is launched hidden in the background, used to convert and "
            "print each file, then closed automatically."
        )
        ok, message = excel_is_available()
        if ok:
            st.success(message)
        else:
            st.error(message)

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📄 1. Marksheet Excel Files")
        uploaded_files = st.file_uploader(
            "Choose Excel file(s)",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
            help="Accepts .xls, .xlsx, and SpreadsheetML/'XML Spreadsheet 2003' "
            "files saved with an .xls/.XLS extension.",
            key="bm_uploaded_files",
        )

    with col2:
        st.markdown("### ✅ 2. Master Student List (Optional)")
        st.write(
            "Upload master file containing student subject enrollment. Each marksheet "
            "will be checked against it (by subject code in title, e.g. `UIDLL222`) to "
            "ensure no enrolled student is missing before PDF creation."
        )
        master_file = st.file_uploader(
            "Master student-subject data file (optional)",
            type=["xls", "xlsx"],
            accept_multiple_files=False,
            key="bm_master_file",
        )

    master_data = None
    if master_file is not None:
        master_data = load_master_data_cached(master_file.getvalue())
        if master_data:
            total_records = sum(len(v["students"]) for v in master_data.values())
            st.success(
                f"Master file loaded: **{len(master_data)}** subject(s), "
                f"**{total_records}** student-subject record(s)."
            )
        else:
            st.error(
                "Could not find the expected columns (Module Abbreviation, "
                "Student Number) in the master file - verification will be "
                "skipped."
            )
            master_data = None

    if uploaded_files:
        st.write(f"**{len(uploaded_files)} file(s) selected.**")
        run = st.button("Convert to PDF", type="primary", key="bm_convert_btn")

        if run:
            results = []
            progress = st.progress(0.0, text="Starting Excel...")

            try:
                with tempfile.TemporaryDirectory() as tmp:
                    work_dir = Path(tmp)
                    with ExcelSession() as excel:
                        for i, uploaded_file in enumerate(uploaded_files):
                            progress.progress(
                                i / len(uploaded_files),
                                text=f"Processing {uploaded_file.name} ...",
                            )
                            results.append(
                                process_one_file(uploaded_file, work_dir, excel, master_data)
                            )
                    progress.progress(1.0, text="Done.")
                st.session_state["results"] = results
            except RuntimeError as exc:
                progress.empty()
                st.error(f"Could not run the conversion: {exc}")

    if "results" in st.session_state:
        results = st.session_state["results"]
        st.subheader("Results")

        successes = [
            r for r in results if r.get("error") is None and r.get("pdf_bytes") is not None
        ]

        for r in results:
            verification = r.get("verification")

            if r.get("blocked", False):
                v = verification or {}
                subject_label = v.get("subject_code") or "unknown subject"
                if v.get("subject_description"):
                    subject_label += f" ({v['subject_description']})"
                st.error(
                    f"🚫 **{r.get('name', 'File')}** — blocked: missing "
                    f"**{len(v.get('missing', []))}** student(s) for "
                    f"**{subject_label}** (master has {v.get('master_count', 0)}, "
                    f"marksheet has {v.get('marksheet_count', 0)})"
                )
                missing_rows = [
                    {"SAP No": m["sap"], "Roll No": m["roll"], "Name": m["name"]}
                    for m in v.get("missing", [])
                ]
                if missing_rows:
                    st.table(missing_rows)
                extra_rows = [
                    {"SAP No": e["sap"], "Roll No": e["roll"], "Name": e["name"]}
                    for e in v.get("extra", [])
                ]
                if extra_rows:
                    st.warning(
                        f"Also found {len(extra_rows)} student(s) in this marksheet "
                        "not listed in the master file for this subject:"
                    )
                    st.table(extra_rows)
                continue

            if r.get("error"):
                st.error(f"❌ **{r.get('name', 'File')}** — failed: {r.get('error')}")
                continue

            st.success(
                f"✅ **{r.get('name', '')}** → `{Path(r.get('name', '')).stem}.pdf` "
                f"({r.get('replaced_count', 0)} cell(s) updated)"
            )
            if verification and verification.get("verified") and verification.get("extra"):
                extra_rows = [
                    {"SAP No": e["sap"], "Roll No": e["roll"], "Name": e["name"]}
                    for e in verification.get("extra", [])
                ]
                st.warning(
                    f"⚠️ {len(extra_rows)} student(s) in this marksheet are not "
                    "listed in the master file for this subject (PDF still "
                    "generated):"
                )
                st.table(extra_rows)
            st.download_button(
                label=f"Download {Path(r.get('name', '')).stem}.pdf",
                data=r.get("pdf_bytes"),
                file_name=f"{Path(r.get('name', '')).stem}.pdf",
                mime="application/pdf",
                key=f"dl_{r.get('name', '')}",
            )

        if len(successes) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for r in successes:
                    zf.writestr(f"{Path(r.get('name', '')).stem}.pdf", r.get("pdf_bytes"))
            st.download_button(
                label=f"⬇️ Download all {len(successes)} PDFs as ZIP",
                data=zip_buffer.getvalue(),
                file_name="converted_marksheets.zip",
                mime="application/zip",
                key="bm_dl_all_zip",
            )


main = show

if __name__ == "__main__":
    st.set_page_config(page_title="Marksheet Excel -> PDF Converter", page_icon="📄")
    show()