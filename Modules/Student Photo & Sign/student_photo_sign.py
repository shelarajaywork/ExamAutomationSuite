"""
Hall Ticket Extractor
----------------------
Streamlit tool that reads SVKM-style Hall Ticket PDFs (one hall ticket per page,
one or more pages per PDF) and produces ONE Excel workbook PER UPLOADED PDF:

    Source File | Student Number | Roll No. | Student Name | Program Name |
    Academic Year | Academic Session | Photo | Signature

Photo and Signature are embedded directly as images in the output rows, using
a two-cell anchor so each image snaps to (and resizes with) its own cell
instead of floating independently on top of the grid.

Because each PDF can contain many students, and each student row carries two
embedded images, a single combined workbook can get very large. To keep files
a manageable size, every uploaded PDF gets its own workbook, named after the
Program Name / Academic Session / Academic Year found in it, and all the
resulting workbooks are bundled into a single downloadable ZIP.

Designed to be dropped into the Exam Tools Streamlit suite as its own module
(same pattern as merge_files.py) — import it from app.py behind a try/except
guard, e.g.:

    try:
        import hall_ticket_extractor
        TOOLS["Hall Ticket Extractor"] = hall_ticket_extractor.run
    except Exception as e:
        st.error(f"Hall Ticket Extractor failed to load: {e}")

and call hall_ticket_extractor.run() from the dashboard. It also runs
standalone via `streamlit run hall_ticket_extractor.py`.
"""

import io
import os
import re
import tempfile
import zipfile
from collections import Counter

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Alignment, Font
from openpyxl.utils.units import pixels_to_EMU

# --------------------------------------------------------------------------
# Field extraction
# --------------------------------------------------------------------------

# Each hall ticket prints its fields as "Label" on one line and the value on
# the next line (this is how PyMuPDF reads the underlying text layer).
FIELD_PATTERNS = {
    "Student Number": r"Student Number\s*\n\s*([^\n]+)",
    "Roll No.": r"Roll No\.\s*\n\s*([^\n]+)",
    "Student Name": r"Student Name\s*\n\s*([^\n]+)",
    "Program Name": r"Program Name\s*\n\s*([^\n]+)",
    "Academic Year": r"Academic Year\s*\n\s*([^\n]+)",
    "Academic Session": r"Academic Session\s*\n\s*([^\n]+)",
}


def extract_fields(page_text: str) -> dict:
    """Pull the labelled fields out of one page's plain text."""
    data = {}
    for field, pattern in FIELD_PATTERNS.items():
        m = re.search(pattern, page_text)
        data[field] = m.group(1).strip() if m else ""
    return data


# --------------------------------------------------------------------------
# Photo / signature separation
# --------------------------------------------------------------------------
# Every hall ticket page embeds four possible images:
#   1. SVKM logo (top-left)        -> near y0 ~ 0
#   2. UPG logo (top-right)        -> near y0 ~ 0
#   3. Controller of Examinations
#      signature stamp (fixed,
#      identical on every page)    -> around y0 ~ 600-650, x0 ~ 440-540
#   4. Student photo               -> around y0 ~ 160-230
#   5. Student signature (if any)  -> around y0 ~ 225-265, just under photo
#
# Rather than hard-coding exact coordinates (which could shift if the
# template changes slightly), we filter out the two logos (near the top of
# the page) and the fixed Controller-of-Examinations stamp (identified by
# its distinctive position), then take whatever images remain and sort them
# top-to-bottom: the first is the photo, the second (if present) is the
# student's signature. Students who did not upload a signature simply leave
# a blank/placeholder there, so no second image exists.

LOGO_Y_CUTOFF = 90
STAMP_Y_RANGE = (590, 660)
STAMP_X_RANGE = (440, 540)


def is_logo_or_stamp(rect) -> bool:
    x0, y0 = rect.x0, rect.y0
    if y0 < LOGO_Y_CUTOFF:
        return True
    if STAMP_Y_RANGE[0] < y0 < STAMP_Y_RANGE[1] and STAMP_X_RANGE[0] < x0 < STAMP_X_RANGE[1]:
        return True
    return False


def find_photo_and_signature_xrefs(page):
    """Return (photo_xref, signature_xref_or_None) for a hall ticket page."""
    candidates = []
    for img in page.get_images(full=True):
        xref = img[0]
        rects = page.get_image_rects(xref)
        if not rects:
            continue
        rect = rects[0]
        if is_logo_or_stamp(rect):
            continue
        candidates.append((rect.y0, xref))

    candidates.sort(key=lambda c: c[0])
    photo_xref = candidates[0][1] if len(candidates) >= 1 else None
    sig_xref = candidates[1][1] if len(candidates) >= 2 else None
    return photo_xref, sig_xref


def save_xref_as_png(doc, xref, out_path):
    pix = fitz.Pixmap(doc, xref)
    if pix.n - pix.alpha >= 4:  # CMYK or other non-RGB colorspace
        pix = fitz.Pixmap(fitz.csRGB, pix)
    pix.save(out_path)
    pix = None


# --------------------------------------------------------------------------
# Core processing
# --------------------------------------------------------------------------

def process_single_pdf(fname, file_bytes, tmp_dir, progress_callback=None,
                        base_done=0, total_pages=None):
    """
    Walk every page of ONE PDF's bytes and return a list of row dicts:
        {Source File, Student Number, Roll No., Student Name, Program Name,
         Academic Year, Academic Session, _photo_path, _sig_path}
    """
    rows = []
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    safe_stub = re.sub(r"[^A-Za-z0-9_.-]", "_", os.path.splitext(fname)[0])

    npages = len(doc)
    for pno in range(npages):
        page = doc[pno]
        text = page.get_text()
        fields = extract_fields(text)

        photo_xref, sig_xref = find_photo_and_signature_xrefs(page)

        photo_path = None
        sig_path = None
        if photo_xref is not None:
            photo_path = os.path.join(tmp_dir, f"{safe_stub}_{pno}_photo.png")
            try:
                save_xref_as_png(doc, photo_xref, photo_path)
            except Exception:
                photo_path = None
        if sig_xref is not None:
            sig_path = os.path.join(tmp_dir, f"{safe_stub}_{pno}_sig.png")
            try:
                save_xref_as_png(doc, sig_xref, sig_path)
            except Exception:
                sig_path = None

        row = {"Source File": fname, **fields, "_photo_path": photo_path, "_sig_path": sig_path}
        rows.append(row)

        if progress_callback:
            progress_callback(base_done + pno + 1, total_pages, fname, pno + 1, npages)

    doc.close()
    return rows


# --------------------------------------------------------------------------
# Output filename derivation
# --------------------------------------------------------------------------

def sanitize_filename(text: str) -> str:
    """Make a string safe to use as a (Windows-friendly) file name."""
    text = text.strip()
    text = re.sub(r'[\\/:*?"<>|]', "", text)   # strip characters illegal in filenames
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace(" ", "_")
    return text[:120] if text else ""


def most_common_value(rows, field):
    values = [r.get(field, "").strip() for r in rows if r.get(field, "").strip()]
    if not values:
        return ""
    return Counter(values).most_common(1)[0][0]


def derive_output_filename(fname, rows, used_names):
    """
    Build 'ProgramName_AcademicSession_AcademicYear.xlsx' from the rows
    belonging to one PDF. Falls back to the source PDF's own name for any
    part that couldn't be read, and de-duplicates against names already
    used in this batch.
    """
    program = sanitize_filename(most_common_value(rows, "Program Name"))
    session = sanitize_filename(most_common_value(rows, "Academic Session"))
    year = sanitize_filename(most_common_value(rows, "Academic Year"))

    parts = [p for p in (program, session, year) if p]
    if parts:
        base = "_".join(parts)
    else:
        base = sanitize_filename(os.path.splitext(fname)[0]) or "hall_tickets"

    candidate = f"{base}.xlsx"
    n = 2
    while candidate in used_names:
        candidate = f"{base}_{n}.xlsx"
        n += 1
    used_names.add(candidate)
    return candidate


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------

HEADERS = [
    "Source File",
    "Student Number",
    "Roll No.",
    "Student Name",
    "Program Name",
    "Academic Year",
    "Academic Session",
    "Photo",
    "Signature",
]

COLUMN_WIDTHS = {
    "A": 22,  # Source File
    "B": 16,  # Student Number
    "C": 10,  # Roll No.
    "D": 30,  # Student Name
    "E": 50,  # Program Name
    "F": 14,  # Academic Year
    "G": 18,  # Academic Session
    "H": 15,  # Photo
    "I": 18,  # Signature
}

# 0-indexed column positions matching COLUMN_WIDTHS, needed for anchoring.
COL_INDEX = {"Photo": 7, "Signature": 8}  # H=7, I=8

ROW_HEIGHT = 95  # points, tall enough for both images

# Small inset (in pixels) so images don't butt right up against the
# gridlines — purely cosmetic, keeps the "snapped to cell" look tidy.
CELL_IMAGE_PADDING_PX = 3


def add_image_snapped_to_cell(ws, image_path, row_idx_0based, col_idx_0based):
    """
    Embed an image with a TwoCellAnchor spanning exactly one cell
    (row_idx_0based, col_idx_0based). Because it's anchored to both the
    top-left AND bottom-right corners of that single cell (editAs="twoCell"),
    Excel stretches/moves the image together with the cell whenever the row
    height or column width changes — this is the closest openpyxl can get
    to Excel 365's native "Place in Cell" picture-as-value feature, which
    openpyxl itself does not support (see conversation for why).
    """
    img = XLImage(image_path)
    pad = pixels_to_EMU(CELL_IMAGE_PADDING_PX)

    marker_from = AnchorMarker(col=col_idx_0based, colOff=pad, row=row_idx_0based, rowOff=pad)
    marker_to = AnchorMarker(col=col_idx_0based + 1, colOff=-pad, row=row_idx_0based + 1, rowOff=-pad)

    img.anchor = TwoCellAnchor(editAs="twoCell", _from=marker_from, to=marker_to)
    ws.add_image(img)


def build_excel(rows) -> io.BytesIO:
    """Build one workbook (as bytes) for the given list of row dicts."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Hall Tickets"

    ws.append(HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

    for col, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width

    for i, row in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=row.get("Source File", ""))
        ws.cell(row=i, column=2, value=row.get("Student Number", ""))
        ws.cell(row=i, column=3, value=row.get("Roll No.", ""))
        ws.cell(row=i, column=4, value=row.get("Student Name", ""))
        ws.cell(row=i, column=5, value=row.get("Program Name", ""))
        ws.cell(row=i, column=6, value=row.get("Academic Year", ""))
        ws.cell(row=i, column=7, value=row.get("Academic Session", ""))
        for col_idx in range(1, 8):
            ws.cell(row=i, column=col_idx).alignment = Alignment(vertical="center", wrap_text=True)

        ws.row_dimensions[i].height = ROW_HEIGHT

        row_idx_0based = i - 1  # openpyxl anchors are 0-indexed

        photo_path = row.get("_photo_path")
        if photo_path and os.path.exists(photo_path):
            add_image_snapped_to_cell(ws, photo_path, row_idx_0based, COL_INDEX["Photo"])

        sig_path = row.get("_sig_path")
        if sig_path and os.path.exists(sig_path):
            add_image_snapped_to_cell(ws, sig_path, row_idx_0based, COL_INDEX["Signature"])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

def run():
    st.title("📋 Extract Student Photo & Sign from Halltickets")
    st.write(
        "Upload one or more Hall Ticket PDFs (each PDF may contain many students, "
        "one hall ticket per page). This tool extracts **Student Number, Roll No., "
        "Student Name, Program Name, Academic Year, Academic Session**, plus the "
        "**Photo** and **Signature** images.\n\n"
        "Each PDF gets its **own** Excel workbook (named after its Program Name, "
        "Academic Session and Academic Year), and all the workbooks are bundled "
        "into a single ZIP for download — this keeps individual files a "
        "manageable size despite the embedded images."
    )

    uploaded_files = st.file_uploader(
        "Upload Hall Ticket PDF(s)", type=["pdf"], accept_multiple_files=True
    )

    if not uploaded_files:
        st.info("👆 Upload one or more Hall Ticket PDFs to get started.")
        return

    if st.button("🚀 Extract to Excel (ZIP)", type="primary"):
        tmp_dir = tempfile.mkdtemp()
        progress_bar = st.progress(0.0)
        status = st.empty()

        # Read all files up front so we know total page count for the progress bar.
        file_bytes_list = [(f.name, f.read()) for f in uploaded_files]
        total_pages = 0
        for fname, fbytes in file_bytes_list:
            with fitz.open(stream=fbytes, filetype="pdf") as d:
                total_pages += len(d)

        def on_progress(done, total, fname, pno, npages):
            progress_bar.progress(done / total if total else 1.0)
            status.text(f"Processing {fname} — page {pno}/{npages}")

        used_names = set()
        workbook_entries = []  # (arcname, BytesIO)
        all_rows_preview = []
        done_pages = 0

        with st.spinner("Reading PDFs and pulling out photos & signatures..."):
            for fname, fbytes in file_bytes_list:
                rows = process_single_pdf(
                    fname, fbytes, tmp_dir,
                    progress_callback=on_progress,
                    base_done=done_pages, total_pages=total_pages,
                )
                done_pages += len(rows)
                all_rows_preview.extend(rows)

                if rows:
                    excel_buffer = build_excel(rows)
                    out_name = derive_output_filename(fname, rows, used_names)
                    workbook_entries.append((out_name, excel_buffer))

        progress_bar.empty()
        status.empty()

        if not workbook_entries:
            st.warning("No hall tickets were found in the uploaded file(s).")
            return

        st.success(
            f"Extracted {len(all_rows_preview)} hall ticket(s) from "
            f"{len(uploaded_files)} file(s) into {len(workbook_entries)} workbook(s)."
        )

        preview_df = pd.DataFrame(
            [{k: v for k, v in r.items() if not k.startswith("_")} for r in all_rows_preview]
        )
        st.dataframe(preview_df, use_container_width=True)

        missing_photo = sum(1 for r in all_rows_preview if not r.get("_photo_path"))
        missing_sig = sum(1 for r in all_rows_preview if not r.get("_sig_path"))
        if missing_photo:
            st.caption(f"⚠️ {missing_photo} record(s) had no photo detected.")
        if missing_sig:
            st.caption(f"ℹ️ {missing_sig} record(s) had no signature on file (placeholder only).")

        st.write("**Workbooks in this ZIP:**")
        for name, _ in workbook_entries:
            st.caption(f"• {name}")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, buf in workbook_entries:
                zf.writestr(name, buf.getvalue())
        zip_buffer.seek(0)

        st.download_button(
            label="⬇️ Download ZIP (one Excel per PDF, with photos & signatures)",
            data=zip_buffer,
            file_name="hall_tickets_extracted.zip",
            mime="application/zip",
        )


show = run


if __name__ == "__main__":
    st.set_page_config(page_title="Hall Ticket Extractor", layout="wide")
    run()