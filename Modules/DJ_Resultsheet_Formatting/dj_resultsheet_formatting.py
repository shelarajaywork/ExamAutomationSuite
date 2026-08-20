import io
import re
import xml.etree.ElementTree as ET

import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Flowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

st.set_page_config(
    page_title="SAP Result Gazette PDF Converter", page_icon="🎓", layout="wide"
)

st.title("🎓 SAP Result Gazette to A3 PDF Converter")
st.markdown(
    "Upload raw `.XLS` files generated from SAP to convert them into **Formatted A3 Landscape PDF Documents** "
    "with exact headers, footers, full column preservation (up to `Eligibility Status`), and clean table formatting."
)

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
SS = "{urn:schemas-microsoft-com:office:spreadsheet}"

HEADER_ROWS = 7
NO_WRAP_FIXED_COLS = {0, 1, 2, 3}

# Tall summary column headers that rotate 90 degrees upwards
TALL_SUMMARY_ROTATE_KEYWORDS = {
    "grand total",
    "out of marks",
    "percentage",
    "credits earned",
    "total c*g",
    "sgpa",
    "overall grade",
    "overall credits",
    "overall c*g",
    "cgpa",
    "cumulative percentage",    
}


class RotatedText(Flowable):
    """
    Renders tall summary column headers rotated 90 degrees upwards.
    """

    def __init__(self, text, font_name="Times-Bold", font_size=3.8):
        super().__init__()
        self.text = text
        self.font_name = font_name
        self.font_size = font_size
        self.text_w = stringWidth(text, font_name, font_size)
        self.text_h = font_size

    def wrap(self, availWidth, availHeight):
        return self.text_h, self.text_w

    def draw(self):
        self.canv.saveState()
        self.canv.setFont(self.font_name, self.font_size)
        self.canv.translate(self.text_h * 0.85, 0)
        self.canv.rotate(90)
        self.canv.drawString(0, 0, self.text)
        self.canv.restoreState()


def is_tall_summary_header(text, span_cols, span_rows):
    """Rotate only if header is single-column and vertically spans >= 3 rows."""
    if span_cols > 1 or span_rows < 3:
        return False
    t = text.strip().lower()
    return any(k in t for k in TALL_SUMMARY_ROTATE_KEYWORDS)


class NumberedCanvas(canvas.Canvas):
    """
    Two-pass canvas to dynamically render running header lines at top
    in Times-Bold (size 11) and footer lines at bottom of every A3 Landscape page.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()

        # A3 Landscape size: 1190.55 x 841.89 points
        page_w, page_h = landscape(A3)
        margin = 12.0

        # Running Header at Top (Times-Bold, font size 11 for all rows)
        header_text = getattr(self, "pdf_header_text", "")
        if header_text:
            self.setFillColor(colors.HexColor("#111111"))
            lines = [l.strip() for l in header_text.split("\n") if l.strip()]
            y_pos = page_h - 15.0
            for line in lines:
                self.setFont("Times-Bold", 11.0)
                self.drawCentredString(page_w / 2.0, y_pos, line)
                y_pos -= 12.5

        # Running Footer at Bottom with PRINCIPAL replaced with Principal
        footer_text = getattr(self, "pdf_footer_text", "")
        if footer_text:
            self.setFont("Times-Bold", 8.0)
            self.setFillColor(colors.HexColor("#222222"))
            clean_footer = re.sub(
                r"\s*Page\s+\d+\s+of\s+\d+\s*$", "", footer_text
            ).strip()
            clean_footer = clean_footer.replace("PRINCIPAL", "Principal")
            segments = [s for s in re.split(r"\s{2,}", clean_footer) if s]

            if segments:
                usable = page_w - 2 * margin
                n = len(segments)
                slot = usable / (n + 0.6)
                for i, seg in enumerate(segments):
                    x = margin + i * slot
                    self.drawString(x, 9, seg)

            page_str = f"Page {self._pageNumber} of {page_count}"
            self.drawRightString(page_w - margin, 9, page_str)

        self.restoreState()


def _parse_column_widths(table_elem):
    widths = {}
    ptr = 1
    for c in table_elem.findall("ss:Column", NS):
        idx = c.attrib.get(SS + "Index")
        if idx:
            ptr = int(idx)
        span = int(c.attrib.get(SS + "Span", 0))
        w = float(c.attrib.get(SS + "Width", 40.0))
        for k in range(ptr, ptr + span + 1):
            widths[k] = w
        ptr += span + 1
    return widths


def parse_sap_xml(file_bytes):
    try:
        tree = ET.parse(io.BytesIO(file_bytes))
    except ET.ParseError as e:
        raise ValueError(
            f"This file doesn't look like a valid SAP-exported SpreadsheetML (.xls) file: {e}"
        )

    root = tree.getroot()
    ws = root.find("ss:Worksheet", NS)
    if ws is None:
        raise ValueError("No <Worksheet> found in file.")
    table = ws.find("ss:Table", NS)
    if table is None:
        raise ValueError("No <Table> found inside worksheet.")

    col_widths = _parse_column_widths(table)
    rows_xml = table.findall("ss:Row", NS)

    raw_rows = []
    for r in rows_xml:
        row_data = []
        c_idx = 0
        cells = r.findall("ss:Cell", NS)
        for c in cells:
            idx = c.attrib.get(SS + "Index")
            if idx:
                c_idx = int(idx) - 1
            while len(row_data) < c_idx:
                row_data.append({"text": "", "ma": 0, "md": 0})

            data_elem = c.find("ss:Data", NS)
            text = (
                data_elem.text.strip()
                if data_elem is not None and data_elem.text
                else ""
            )
            ma = int(c.attrib.get(SS + "MergeAcross", 0))
            md = int(c.attrib.get(SS + "MergeDown", 0))

            row_data.append({"text": text, "ma": ma, "md": md})
            c_idx += 1 + ma
        raw_rows.append(row_data)

    header_lines = []
    blocks = []
    footer_text = ""

    i = 0
    N = len(raw_rows)
    while i < N:
        row_str = " ".join([c["text"] for c in raw_rows[i]])
        if (
            "Shri Vile Parle" in row_str
            or "DWARKADAS" in row_str
            or "RESULT GAZETTE" in row_str
        ):
            if not header_lines:
                for h_r in raw_rows[i : i + 5]:
                    h_t = " ".join([c["text"] for c in h_r if c["text"]])
                    if h_t:
                        header_lines.append(h_t)
            i += 5

            middle_rows = []
            while i < N:
                r_str = " ".join([c["text"] for c in raw_rows[i]])
                if "Result Declaration Date" in r_str:
                    if not footer_text:
                        footer_text = r_str
                    i += 1
                    if middle_rows:
                        blocks.append(middle_rows)
                    break
                else:
                    middle_rows.append(raw_rows[i])
                    i += 1
        else:
            i += 1

    if not blocks:
        raise ValueError("Could not find result rows in file.")

    return header_lines, blocks, footer_text, col_widths


def _split_trailer(block_rows):
    note_idx = None
    for i, row in enumerate(block_rows):
        joined = " ".join(c["text"] for c in row if c["text"]).strip()
        if joined.startswith("Note:"):
            note_idx = i
            break

    if note_idx is None:
        data_rows = [r for r in block_rows if any(c["text"].strip() for c in r)]
        return data_rows, "", []

    data_rows = [r for r in block_rows[:note_idx] if any(c["text"].strip() for c in r)]
    note_text = " ".join(
        c["text"] for c in block_rows[note_idx] if c["text"]
    ).strip()
    trailer = block_rows[note_idx + 1 :]
    legend_rows = [r for r in trailer if any(c["text"].strip() for c in r)]
    return data_rows, note_text, legend_rows


def _consolidate_candidate_columns(data_rows):
    """
    Merges Roll No. (col 1), Student No. (col 2), and APAAR ID (col 3)
    into a single column to give maximum space for Student Name.
    """
    new_rows = []
    for r_idx, row in enumerate(data_rows):
        if r_idx < HEADER_ROWS:
            if r_idx == 0:
                cand_cell = {"text": "CANDIDATE", "ma": 2, "md": 0}
                new_row = [cand_cell] + [
                    {"text": "", "ma": 0, "md": 0},
                    {"text": "", "ma": 0, "md": 0},
                ]
                new_row.extend(row[5:])
            elif r_idx == 1:
                c0 = {"text": "Sr. No.", "ma": 0, "md": 5}
                c1 = {
                    "text": "Roll No.<br/>Student No.<br/>APAAR ID",
                    "ma": 0,
                    "md": 5,
                }
                c2 = {"text": "Student Name", "ma": 0, "md": 5}
                new_row = [c0, c1, c2] + row[5:]
            else:
                c0 = {"text": "", "ma": 0, "md": 0}
                c1 = {"text": "", "ma": 0, "md": 0}
                c2 = {"text": "", "ma": 0, "md": 0}
                new_row = [c0, c1, c2] + row[5:]
        else:
            sr_no = row[0]["text"].strip() if len(row) > 0 else ""
            roll_no = row[1]["text"].strip() if len(row) > 1 else ""
            student_no = row[2]["text"].strip() if len(row) > 2 else ""
            apaar_id = row[3]["text"].strip() if len(row) > 3 else ""
            student_name = row[4]["text"].strip() if len(row) > 4 else ""

            id_parts = [p for p in [roll_no, student_no, apaar_id] if p]
            combined_id = "<br/>".join(id_parts)

            c0 = {"text": sr_no, "ma": 0, "md": 0}
            c1 = {"text": combined_id, "ma": 0, "md": 0}
            c2 = {"text": student_name, "ma": 0, "md": 0}

            new_row = [c0, c1, c2] + (row[5:] if len(row) > 5 else [])

        new_rows.append(new_row)
    return new_rows


def _compute_min_widths(
    all_data_rows, max_cols, header_font, header_size, body_font, body_size
):
    mins = [0.0] * max_cols
    for rows in all_data_rows:
        for r_idx in range(min(HEADER_ROWS, len(rows))):
            row = rows[r_idx]
            for idx, cell in enumerate(row):
                if idx >= max_cols:
                    break
                text = cell["text"]
                if not text:
                    continue
                span_c = cell.get("ma", 0) + 1
                span_r = cell.get("md", 0) + 1
                if is_tall_summary_header(text, span_c, span_r):
                    per_col = (header_size + 2.0) / span_c
                else:
                    words = text.replace("<br/>", " ").split()
                    max_word_w = max(
                        (stringWidth(w, header_font, header_size) for w in words),
                        default=0.0,
                    )
                    per_col = (max_word_w + 1.5) / span_c

                for k in range(idx, min(idx + span_c, max_cols)):
                    if per_col > mins[k]:
                        mins[k] = per_col

    for rows in all_data_rows:
        for r_idx in range(HEADER_ROWS, len(rows)):
            row = rows[r_idx]
            for c_idx in range(min(len(row), max_cols)):
                text = row[c_idx]["text"]
                if not text or c_idx == 2:
                    continue
                lines = text.split("<br/>")
                w = (
                    max(
                        stringWidth(l, body_font, body_size) for l in lines
                    )
                    + 1.5
                )
                if w > mins[c_idx]:
                    mins[c_idx] = w

    return mins


def _waterfill_widths(raw_weights, mins, printable_w):
    n = len(raw_weights)
    sum_w = sum(raw_weights) or 1.0
    total_min = sum(mins)

    if total_min >= printable_w:
        scale = printable_w / total_min if total_min else 1.0
        return [m * scale for m in mins]

    remaining = printable_w - total_min
    result = [mins[i] + remaining * (raw_weights[i] / sum_w) for i in range(n)]
    drift = printable_w - sum(result)
    if result:
        result[-1] += drift
    return result


def convert_to_a3_pdf(file_bytes):
    header_lines, middle_blocks, footer_text, col_widths = parse_sap_xml(
        file_bytes
    )

    split_blocks = [_split_trailer(b) for b in middle_blocks]

    # Global Note and Grade Table to repeat on every page
    global_note_text = ""
    global_legend_rows = []
    for _, n_txt, l_rows in split_blocks:
        if n_txt and not global_note_text:
            global_note_text = n_txt
        if l_rows and not global_legend_rows:
            global_legend_rows = l_rows

    all_data_rows = [
        _consolidate_candidate_columns(d) for d, _, _ in split_blocks
    ]

    # Determine rightmost column bound
    max_cols = 0
    for data_rows in all_data_rows:
        for row in data_rows:
            for c_idx, cell in enumerate(row):
                if "Eligibility Status" in cell["text"]:
                    max_cols = max(max_cols, c_idx + 1)
            max_cols = max(max_cols, len(row))

    if max_cols == 0:
        max_cols = 70

    # A3 Landscape PDF Document Setup
    page_w, page_h = landscape(A3)  # 1190.55 x 841.89 pt
    margin = 12.0
    printable_w = page_w - (2 * margin)  # ~1166.55 pt

    output_stream = io.BytesIO()
    doc = SimpleDocTemplate(
        output_stream,
        pagesize=landscape(A3),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=78.0,  # Top margin for font 11 running header
        bottomMargin=18.0,  # Reserved for Running Footer
    )

    header_str = "\n".join(header_lines)

    class DynamicPDFCanvas(NumberedCanvas):
        pdf_header_text = header_str
        pdf_footer_text = footer_text

    # Consistent Times New Roman Typography & Styles
    styles = getSampleStyleSheet()

    BODY_FONT = "Times-Roman"
    BODY_FONT_BOLD = "Times-Bold"
    BODY_SIZE = 3.8
    HEADER_SIZE = 3.8

    cell_style_center = ParagraphStyle(
        "CellCenter",
        parent=styles["Normal"],
        fontName=BODY_FONT,
        fontSize=BODY_SIZE,
        leading=BODY_SIZE + 0.5,
        spaceBefore=0,
        spaceAfter=0,
        alignment=1,  # Center aligned
    )

    cell_style_center_bold = ParagraphStyle(
        "CellCenterBold",
        parent=cell_style_center,
        fontName=BODY_FONT_BOLD,
        fontSize=BODY_SIZE,
        leading=BODY_SIZE + 0.5,
        spaceBefore=0,
        spaceAfter=0,
        alignment=1,  # Center aligned
    )

    note_style = ParagraphStyle(
        "NoteStyle",
        parent=styles["Normal"],
        fontName=BODY_FONT,
        fontSize=5.0,
        leading=5.8,
        spaceBefore=0,
        spaceAfter=0,
        alignment=0,
    )

    # Proportional Column Widths: Give Maximum Dedicated Room to Student Name (col 2)
    raw_weights = []
    for c in range(max_cols):
        if c == 0:  # Sr. No.
            w = 15.0
        elif c == 1:  # Combined Roll No. / Student No. / APAAR ID
            w = 58.0
        elif c == 2:  # Student Name gets maximum room
            w = 160.0
        else:
            w = col_widths.get(c + 3, 11.0)
        raw_weights.append(w)

    min_widths = _compute_min_widths(
        all_data_rows,
        max_cols,
        BODY_FONT_BOLD,
        HEADER_SIZE,
        BODY_FONT,
        BODY_SIZE,
    )
    calc_col_widths = _waterfill_widths(raw_weights, min_widths, printable_w)

    elements = []

    # Build Table for Each Gazette Block
    for block_idx, data_rows in enumerate(all_data_rows):
        if block_idx > 0:
            elements.append(PageBreak())

        data_rows = [r for r in data_rows if any(c["text"].strip() for c in r)]
        _, note_text, legend_rows = split_blocks[block_idx]

        table_data = []
        table_styles = [
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 0.3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0.3),
            ("LEFTPADDING", (0, 0), (-1, -1), 0.3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0.3),
        ]

        # STEP 1: Guaranteed Comprehensive Header Spans
        covered = set()
        cell_spans_map = {}

        # 1A. First register row 0 multi-column group headers (Semester Grade Point Average, Cumulative Credits & Grade, Count of Failures, CANDIDATE, Subjects)
        for c_idx in range(min(max_cols, len(data_rows[0]))):
            if (c_idx, 0) in covered:
                continue
            cell_info = data_rows[0][c_idx]
            ma = cell_info.get("ma", 0)
            text = cell_info.get("text", "").strip()
            t_lower = text.lower()

            if "count of failures" in t_lower:
                ma = max(ma, 1)
                md = 5  # Spans rows 0 to 5 across 2 columns
            elif "semester grade point average" in t_lower:
                ma = max(ma, 2)
                md = 0
            elif "cumulative credits & grade" in t_lower or "cumulative credits" in t_lower:
                ma = max(ma, 3)
                md = 0
            elif "candidate" in t_lower:
                ma = max(ma, 2)
                md = 0
            else:
                md = cell_info.get("md", 0)

            if ma > 0 or md > 0:
                end_c = min(c_idx + ma, max_cols - 1)
                end_r = min(0 + md, HEADER_ROWS - 1)
                can_span = True
                for r_k in range(0, end_r + 1):
                    for c_k in range(c_idx, end_c + 1):
                        if (c_k, r_k) in covered:
                            can_span = False
                            break
                if can_span:
                    table_styles.append(("SPAN", (c_idx, 0), (end_c, end_r)))
                    cell_spans_map[(c_idx, 0)] = (
                        end_c - c_idx + 1,
                        end_r - 0 + 1,
                    )
                    for r_k in range(0, end_r + 1):
                        for c_k in range(c_idx, end_c + 1):
                            covered.add((c_k, r_k))

        # 1B. Process all other header cells (Rows 0 to 6)
        for r_idx in range(min(HEADER_ROWS, len(data_rows))):
            row_cells = data_rows[r_idx]
            for c_idx in range(min(max_cols, len(row_cells))):
                if (c_idx, r_idx) in covered:
                    continue
                cell_info = row_cells[c_idx]
                ma = cell_info.get("ma", 0)
                md = cell_info.get("md", 0)
                text = cell_info.get("text", "").strip()
                t_lower = text.lower()

                # Headers that must span vertically down to Row 6
                if t_lower in [
                    "sr. no.",
                    "student name",
                    "grand total",
                    "out of marks",
                    "percentage",
                    "eligibility status",
                    "remark",
                    "overall grade",
                    "overall credits",
                    "overall c*g",
                    "cgpa",
                    "cumulative percentage",
                ] or "roll no." in t_lower:
                    r_end = HEADER_ROWS - 1
                    for r_k in range(r_idx + 1, HEADER_ROWS):
                        if (c_idx, r_k) in covered:
                            r_end = r_k - 1
                            break
                        if (
                            r_k < len(data_rows)
                            and c_idx < len(data_rows[r_k])
                            and data_rows[r_k][c_idx]["text"].strip() != ""
                        ):
                            r_end = r_k - 1
                            break
                    md = max(md, r_end - r_idx)

                # Subheaders under Semester GPA that span down to Row 5 (Row 6 has ΣC, ΣCG, GPA)
                elif t_lower in ["credits earned", "total c*g", "sgpa"]:
                    r_end = 5
                    for r_k in range(r_idx + 1, 6):
                        if (c_idx, r_k) in covered:
                            r_end = r_k - 1
                            break
                    md = max(md, r_end - r_idx)

                if ma > 0 or md > 0:
                    end_c = min(c_idx + ma, max_cols - 1)
                    end_r = min(r_idx + md, HEADER_ROWS - 1)
                    can_span = True
                    for r_k in range(r_idx, end_r + 1):
                        for c_k in range(c_idx, end_c + 1):
                            if (c_k, r_k) in covered:
                                can_span = False
                                break
                    if can_span:
                        table_styles.append(
                            ("SPAN", (c_idx, r_idx), (end_c, end_r))
                        )
                        cell_spans_map[(c_idx, r_idx)] = (
                            end_c - c_idx + 1,
                            end_r - r_idx + 1,
                        )
                        for r_k in range(r_idx, end_r + 1):
                            for c_k in range(c_idx, end_c + 1):
                                covered.add((c_k, r_k))

        # STEP 2: Populate Table Rows
        for r_idx, row_cells in enumerate(data_rows):
            row_data = []
            is_header_row = r_idx < HEADER_ROWS

            for c_idx in range(max_cols):
                if is_header_row:
                    table_styles.append(
                        (
                            "BACKGROUND",
                            (c_idx, r_idx),
                            (c_idx, r_idx),
                            colors.HexColor("#EAEAEA"),
                        )
                    )

                if c_idx < len(row_cells):
                    cell_info = row_cells[c_idx]
                    text = cell_info["text"]

                    is_bold_marker = (
                        "CANDIDATE" in text
                        or "Grade" in text
                        or "Successful" in text
                        or is_header_row
                    )

                    if is_header_row:
                        span_info = cell_spans_map.get((c_idx, r_idx), (1, 1))
                        span_c, span_r = span_info

                        if text and is_tall_summary_header(text, span_c, span_r):
                            flowable = RotatedText(
                                text,
                                font_name=BODY_FONT_BOLD,
                                font_size=HEADER_SIZE,
                            )
                            row_data.append(flowable)
                        else:
                            st_cell = (
                                cell_style_center_bold
                                if is_bold_marker
                                else cell_style_center
                            )
                            p = Paragraph(text, st_cell) if text else ""
                            row_data.append(p)
                    else:
                        st_cell = (
                            cell_style_center_bold
                            if is_bold_marker
                            else cell_style_center
                        )
                        p = Paragraph(text, st_cell) if text else ""
                        row_data.append(p)
                else:
                    row_data.append("")

            table_data.append(row_data)

        # Header row heights reset to exact order: [10, 16, 10, 10, 12, 10, 10]
        header_heights = [10.0, 16.0, 10.0, 10.0, 12.0, 10.0, 10.0]
        num_data_rows = max(0, len(data_rows) - HEADER_ROWS)
        # Data row height set to 21.0 pt
        data_row_height = 21.0
        row_heights = header_heights[: len(data_rows)] + [
            data_row_height
        ] * num_data_rows

        # Create Table with repeated top header rows
        table = Table(
            table_data,
            colWidths=calc_col_widths,
            rowHeights=row_heights,
            repeatRows=HEADER_ROWS,
        )
        table.setStyle(TableStyle(table_styles))
        elements.append(table)

        # Repeated Note Line on every page
        cur_note_text = note_text or global_note_text
        if cur_note_text:
            elements.append(Spacer(1, 2.0))
            elements.append(Paragraph(cur_note_text, note_style))

        # Repeated Marks / Grade Point / Letter Grade Legend Table on every page
        cur_legend_rows = legend_rows or global_legend_rows
        cur_legend_rows = [
            r for r in cur_legend_rows if any(c["text"].strip() for c in r)
        ]

        if cur_legend_rows:
            elements.append(Spacer(1, 2.0))
            used_indices = sorted(
                {
                    i
                    for row in cur_legend_rows
                    for i, c in enumerate(row)
                    if c["text"].strip()
                }
            )
            if used_indices:
                lo, hi = used_indices[0], used_indices[-1]
                legend_data = []
                for row in cur_legend_rows:
                    cells = row[lo : hi + 1] if len(row) > lo else []
                    vals = [
                        cells[i]["text"].strip() if i < len(cells) else ""
                        for i in range(hi - lo + 1)
                    ]
                    legend_data.append(vals)

                n_legend_cols = hi - lo + 1
                label_w = 68.0
                value_w = 38.0
                legend_col_widths = [label_w] + [value_w] * (n_legend_cols - 1)

                legend_style = TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#999999")),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # Centered horizontally
                        ("FONTNAME", (0, 0), (0, -1), "Times-Bold"),
                        ("FONTNAME", (1, 0), (-1, -1), "Times-Roman"),
                        ("FONTSIZE", (0, 0), (-1, -1), 5.2),
                        ("TOPPADDING", (0, 0), (-1, -1), 0.8),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 0.8),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0.8),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 0.8),
                        (
                            "BACKGROUND",
                            (0, 0),
                            (0, -1),
                            colors.HexColor("#EAEAEA"),
                        ),
                    ]
                )
                legend_table = Table(
                    legend_data, colWidths=legend_col_widths, hAlign="LEFT"
                )
                legend_table.setStyle(legend_style)
                elements.append(legend_table)

    # Build Document
    doc.build(elements, canvasmaker=DynamicPDFCanvas)
    output_stream.seek(0)
    return output_stream


# --- STREAMLIT UI ---
uploaded_file = st.file_uploader(
    "Upload raw SAP Excel file (.XLS)", type=["xls", "xml"]
)

if uploaded_file is not None:
    st.success(f"Uploaded file: **{uploaded_file.name}**")

    if st.button("Convert to A3 PDF Gazette", type="primary"):
        with st.spinner("Generating formatted A3 PDF Document..."):
            try:
                pdf_buffer = convert_to_a3_pdf(uploaded_file.getvalue())
                out_name = (
                    uploaded_file.name.rsplit(".", 1)[0] + "_A3_Gazette.pdf"
                )

                st.success("PDF Generated Successfully!")
                st.download_button(
                    label="⬇️ Download A3 PDF File",
                    data=pdf_buffer,
                    file_name=out_name,
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"Error generating PDF: {str(e)}")