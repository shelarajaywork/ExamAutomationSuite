import io
import os
import smtplib
import urllib.parse
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

# ReportLab imports for dynamic PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY


# ==========================================
# COLLEGE LETTERHEAD CONFIGURATION
# ==========================================
# One block per SVKM college. Add/edit freely — this is the single place
# that controls how each college's letters look. "logo" is optional; if the
# file doesn't exist on disk it's silently skipped (no crash).

COLLEGES = {
    "NMCCE": {
        "display_name": "Narsee Monjee College of Commerce & Economics",
        "line1": "Shri Vile Parle Kelavani Mandal's",
        "line2": "NARSEE MONJEE COLLEGE OF COMMERCE & ECONOMICS",
        "line3": "(Autonomous Affiliated to University of Mumbai)",
        "address": "Vile Parle (West), Mumbai — 400 056.",
        "contact": "Tel.: 26136039 - Website: www.nmcce.edu.in - Email: exam@nmcce.edu.in",
        "accreditation": "NAAC RE-ACCREDITED 'A++' GRADE",
        "logo": "logos/nmcce.png",
        "filename_keywords": ["narsee monjee", "nmcce"],
    },
    "MITHIBAI": {
        "display_name": "Mithibai College of Arts, Commerce & Science",
        "line1": "Shri Vile Parle Kelavani Mandal's",
        "line2": "MITHIBAI COLLEGE OF ARTS, COMMERCE & SCIENCE",
        "line3": "(Autonomous Affiliated to University of Mumbai)",
        "address": "Vile Parle (West), Mumbai — 400 056.",
        "contact": "Tel.: 26183183 - Website: www.mithibai.ac.in - Email: exam@mithibai.ac.in",
        "accreditation": "NAAC RE-ACCREDITED 'A+' GRADE",
        "logo": "logos/mithibai.png",
        "filename_keywords": ["mithibai"],
    },
    "SVKM_EXAM_CELL": {
        "display_name": "SVKM Colleges Examination Cell (default)",
        "line1": "Shri Vile Parle Kelavani Mandal's",
        "line2": "SVKM COLLEGES EXAMINATION CELL",
        "line3": "(Autonomous Affiliated to University of Mumbai)",
        "address": "Bhaktivedanta Swami Marg, Juhu Scheme, Vile Parle (West), Mumbai — 400 056.",
        "contact": "Tel.: 42332041/42 - Website: www.svkm.ac.in - Email: exam@svkm.ac.in",
        "accreditation": "NAAC RE-ACCREDITED 'A+' GRADE",
        "logo": "logos/svkm.png",
        "filename_keywords": [],
    },
    # Add more colleges here, e.g.:
    # "PILLAI": { ... },
}

DEFAULT_COLLEGE_KEY = "SVKM_EXAM_CELL"


def guess_college_from_filename(filename):
    """Best-effort match of an uploaded filename to a college config key."""
    if not filename:
        return DEFAULT_COLLEGE_KEY
    lower = filename.lower()
    for key, cfg in COLLEGES.items():
        for kw in cfg.get("filename_keywords", []):
            if kw in lower:
                return key
    return DEFAULT_COLLEGE_KEY


# ==========================================
# PDF GENERATOR FUNCTIONS
# ==========================================

def build_pdf_letterhead(styles, college_key):
    """Creates letterhead elements for the given college config key."""
    cfg = COLLEGES.get(college_key, COLLEGES[DEFAULT_COLLEGE_KEY])
    elements = []

    header_style = ParagraphStyle(
        'LetterheadSub', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=9, alignment=TA_CENTER, textColor=colors.HexColor("#003366"), leading=11
    )
    title_style = ParagraphStyle(
        'LetterheadTitle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=11, alignment=TA_CENTER, textColor=colors.HexColor("#003366"), leading=14
    )
    small_style = ParagraphStyle(
        'LetterheadSmall', parent=styles['Normal'], fontName='Helvetica',
        fontSize=7.5, alignment=TA_CENTER, textColor=colors.HexColor("#333333"), leading=9
    )

    # Optional logo, centered, only if the file actually exists
    logo_path = cfg.get("logo")
    if logo_path and os.path.isfile(logo_path):
        img = Image(logo_path, width=0.8 * inch, height=0.8 * inch)
        img.hAlign = 'CENTER'
        elements.append(img)
        elements.append(Spacer(1, 0.05 * inch))

    elements.append(Paragraph(cfg["line1"], header_style))
    elements.append(Paragraph(cfg["line2"], title_style))
    elements.append(Paragraph(cfg["line3"], header_style))
    elements.append(Paragraph(cfg["address"], small_style))
    elements.append(Paragraph(cfg["contact"], small_style))
    elements.append(Paragraph(f"<b>{cfg['accreditation']}</b>", small_style))
    elements.append(Spacer(1, 0.15 * inch))

    return elements


def generate_appointment_pdf(data, college_key):
    """Generates the official Appointment Letter PDF for the given college."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=54, leftMargin=54, topMargin=36, bottomMargin=36
    )

    styles = getSampleStyleSheet()

    body_style = ParagraphStyle(
        'BodyTextCustom', parent=styles['Normal'], fontName='Helvetica',
        fontSize=10, leading=14, alignment=TA_JUSTIFY, spaceAfter=10
    )
    bold_body = ParagraphStyle('BoldBody', parent=body_style, fontName='Helvetica-Bold')
    right_date = ParagraphStyle(
        'RightDate', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=10, alignment=TA_RIGHT
    )

    story = []
    story.extend(build_pdf_letterhead(styles, college_key))

    current_date = datetime.now().strftime("%dth %B %Y")
    story.append(Paragraph(current_date, right_date))
    story.append(Spacer(1, 0.15 * inch))

    address_text = f"To,<br/><b>{data['examiner_name']}</b><br/>{data['college_name']}"
    story.append(Paragraph(address_text, body_style))
    story.append(Spacer(1, 0.1 * inch))

    salutation = "Sir" if "Mr." in data['examiner_name'] else "Madam"
    subj_text = (
        f"<b>Sub: Appointment as {data['role']} for Examination of {data['course']}. "
        f"Year: {data['year']}, Semester: {data['semester']} ({data['ayear']})</b>"
    )
    story.append(Paragraph(subj_text, body_style))
    story.append(Spacer(1, 0.1 * inch))

    p1 = (
        f"Dear {salutation},<br/><br/>We are pleased to appoint you as <b>{data['role']}</b> "
        f"for the End Semester Examination of the above-mentioned program, "
        f"for the subject <b>{data['subject']}</b>."
    )
    story.append(Paragraph(p1, body_style))

    p2 = (
        f"The answer books are to be evaluated/moderated as per the examination guidelines "
        f"of our institution. A total count of <b>{data['count']}</b> answer books is assigned "
        f"for this task, as per the rules."
    )
    story.append(Paragraph(p2, body_style))

    p3 = "Evaluation/Moderation is in online mode (Onscreen Marking) for which you will receive system credentials upon task assignment."
    story.append(Paragraph(p3, body_style))

    p4 = "An honorarium, as per university rules, will be paid towards the assignment."
    story.append(Paragraph(p4, body_style))
    story.append(Spacer(1, 0.2 * inch))

    story.append(Paragraph("Thanking you,", body_style))
    story.append(Spacer(1, 0.4 * inch))
    story.append(Paragraph("<b>Controller of Examinations / Assistant Registrar</b>", bold_body))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def generate_thanking_pdf(data, college_key):
    """Generates the official Thanking Letter PDF for the given college."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=54, leftMargin=54, topMargin=36, bottomMargin=36
    )

    styles = getSampleStyleSheet()

    body_style = ParagraphStyle(
        'BodyTextCustom', parent=styles['Normal'], fontName='Helvetica',
        fontSize=10, leading=14, alignment=TA_LEFT, spaceAfter=10
    )
    right_date = ParagraphStyle(
        'RightDate', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=10, alignment=TA_RIGHT
    )

    story = []
    story.extend(build_pdf_letterhead(styles, college_key))

    current_date = datetime.now().strftime("%dth %B %Y")
    story.append(Paragraph(current_date, right_date))
    story.append(Spacer(1, 0.15 * inch))

    address_text = f"To,<br/><b>{data['examiner_name']}</b><br/>{data['college_name']}"
    story.append(Paragraph(address_text, body_style))
    story.append(Spacer(1, 0.15 * inch))

    salutation = "Sir" if "Mr." in data['examiner_name'] else "Madam"
    p1 = (
        f"Dear {salutation},<br/><br/>We are thankful to you for agreeing to be the "
        f"<b>{data['role']}</b> for the following subject at our college for the "
        f"End Semester Examination:"
    )
    story.append(Paragraph(p1, body_style))
    story.append(Spacer(1, 0.1 * inch))

    table_data = [
        ["Program", "Semester", "Subject", "Count"],
        [data['course'], data['semester'], data['subject'], str(data['count'])]
    ]
    t = Table(table_data, colWidths=[2.2 * inch, 1.2 * inch, 2.3 * inch, 0.8 * inch])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#003366")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (3, 0), (3, -1), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.25 * inch))

    p2 = (
        "The efforts that you have taken and the time you have spared for this task are "
        "truly appreciated.<br/>Thank you once again and looking forward to your continued support."
    )
    story.append(Paragraph(p2, body_style))
    story.append(Spacer(1, 0.4 * inch))

    story.append(Paragraph("Sincerely,", body_style))
    story.append(Spacer(1, 0.4 * inch))
    story.append(Paragraph("<b>Controller of Examinations / Assistant Registrar</b>", body_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ==========================================
# SMTP SEND-WITH-ATTACHMENTS
# ==========================================

def send_email_with_attachments(
    smtp_host, smtp_port, smtp_user, smtp_pass,
    to_addr, subject, body, attachments
):
    """
    Sends an email with one or more PDF attachments via SMTP (STARTTLS).
    attachments: list of (filename, bytes) tuples.
    Raises on failure — caller should catch and show the error in the UI.
    """
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    for filename, file_bytes in attachments:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, [to_addr], msg.as_string())


def get_smtp_credentials():
    """
    Pulls SMTP config from st.secrets if present (recommended for shared/
    deployed use), otherwise falls back to a manual entry expander so the
    tool still works without a secrets.toml during local testing.

    Expected st.secrets structure:

        [smtp]
        host = "smtp.office365.com"
        port = 587
        user = "exam-cell@yourcollege.ac.in"
        password = "app-password-or-account-password"
    """
    if "smtp" in st.secrets:
        cfg = st.secrets["smtp"]
        return cfg["host"], int(cfg["port"]), cfg["user"], cfg["password"], True

    with st.expander("🔐 SMTP credentials (not found in secrets — enter manually)"):
        host = st.text_input("SMTP Host", value="smtp.office365.com", key="smtp_host_manual")
        port = st.number_input("SMTP Port", value=587, key="smtp_port_manual")
        user = st.text_input("SMTP Username (sender email)", key="smtp_user_manual")
        pw = st.text_input("SMTP Password / App Password", type="password", key="smtp_pass_manual")
    return host, int(port), user, pw, False


# ==========================================
# MAIN STREAMLIT UI MODULE
# ==========================================

def show():
    st.title("✉️ Letters to External Examiners")
    st.caption(
        "Auto-generate official PDF Appointment & Thanking letters on the correct "
        "college letterhead, then email them out with attachments — no manual "
        "download/attach step required."
    )

    st.markdown("---")

    # ------------------------------------------------------------------
    # 1. Excel Upload
    # ------------------------------------------------------------------
    excel_file = st.file_uploader("Upload Evaluation Dashboard Excel File (.xlsx)", type=["xlsx", "xls"])

    if excel_file is None:
        st.info("👆 Upload the Evaluation Dashboard file to auto-generate examiner letters.")
        return

    try:
        df = pd.read_excel(excel_file)
        st.success(f"Loaded {len(df)} records from evaluation dashboard!")
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return

    # ------------------------------------------------------------------
    # 2. College selection (drives the letterhead)
    # ------------------------------------------------------------------
    st.markdown("---")
    st.subheader("1. Select College Letterhead")

    guessed_key = guess_college_from_filename(getattr(excel_file, "name", ""))
    college_keys = list(COLLEGES.keys())
    college_labels = [COLLEGES[k]["display_name"] for k in college_keys]

    default_idx = college_keys.index(guessed_key) if guessed_key in college_keys else 0
    chosen_label = st.selectbox(
        "College (auto-detected from filename where possible — override if wrong):",
        college_labels,
        index=default_idx,
    )
    college_key = college_keys[college_labels.index(chosen_label)]

    # ------------------------------------------------------------------
    # 3. Data Filter & Selection
    # ------------------------------------------------------------------
    st.markdown("---")
    st.subheader("2. Select Examiner Record")

    c1, c2 = st.columns(2)

    roles = df["RoleName"].dropna().unique().tolist() if "RoleName" in df.columns else []
    selected_role = c1.selectbox("Filter by Role:", ["All"] + roles)

    filtered_df = df.copy()
    if selected_role != "All":
        filtered_df = filtered_df[filtered_df["RoleName"] == selected_role]

    examiners = filtered_df["ExaminerName"].dropna().unique().tolist() if "ExaminerName" in filtered_df.columns else []
    selected_examiner = c2.selectbox("Select Examiner:", examiners)

    if not selected_examiner:
        st.warning("No examiner selected.")
        return

    examiner_rows = filtered_df[filtered_df["ExaminerName"] == selected_examiner]

    paper_list = examiner_rows["CategoryName"].tolist() if "CategoryName" in examiner_rows.columns else [
        f"Record {i + 1}" for i in range(len(examiner_rows))
    ]

    if len(paper_list) > 1:
        sel_paper_idx = st.selectbox(
            "Select Paper / Subject Assignment:", range(len(paper_list)),
            format_func=lambda x: paper_list[x]
        )
        row = examiner_rows.iloc[sel_paper_idx]
    else:
        row = examiner_rows.iloc[0]

    record_data = {
        "examiner_name": str(row.get("ExaminerName", "External Examiner")),
        "examiner_email": str(row.get("ExaminerEmail", "")) if pd.notna(row.get("ExaminerEmail")) else "",
        "college_name": str(row.get("CampusName", "Affiliated College/University")),
        "role": str(row.get("RoleName", "Moderator")),
        "course": str(row.get("CourseName", "Degree Program")),
        "subject": str(row.get("CategoryName", "Subject")),
        "semester": str(row.get("Semester/Trimester", "Semester I")),
        "year": "I",
        "ayear": "2025-26",
        "count": int(row.get("CheckCount", 0)) if pd.notna(row.get("CheckCount")) else 0,
    }

    st.dataframe(pd.DataFrame([row]), use_container_width=True)

    # ------------------------------------------------------------------
    # 4. Email configuration + preview
    # ------------------------------------------------------------------
    st.markdown("---")
    st.subheader("3. Configure Email")

    email_to = st.text_input("Recipient Email Address:", value=record_data["examiner_email"])
    email_subject = f"Official Communication: {record_data['subject']} - {record_data['role']} Assignment"

    email_body = (
        f"Dear {record_data['examiner_name']},\n\n"
        f"Greetings from the Examinations Department.\n\n"
        f"Please find attached the official correspondence regarding your appointment as "
        f"{record_data['role']} for the course {record_data['course']} "
        f"(Paper: {record_data['subject']}).\n\n"
        f"If you have any queries, please feel free to reach out to the Examination Cell.\n\n"
        f"Warm regards,\n"
        f"Assistant Registrar of Examinations\n"
        f"{COLLEGES[college_key]['display_name']}"
    )

    email_body = st.text_area("Email Content (editable before sending):", value=email_body, height=180)

    # ------------------------------------------------------------------
    # 5. Generate PDFs
    # ------------------------------------------------------------------
    st.markdown("---")
    st.subheader("4. Generate Letters")

    if st.button("⚡ Generate PDF Letters", type="primary"):
        st.session_state["appt_pdf"] = generate_appointment_pdf(record_data, college_key)
        st.session_state["thank_pdf"] = generate_thanking_pdf(record_data, college_key)
        st.session_state["clean_name"] = record_data["examiner_name"].replace(" ", "_")
        st.success("✅ Appointment Letter PDF & Thanking Letter PDF generated on the selected letterhead.")

    if "appt_pdf" in st.session_state:
        clean_name = st.session_state["clean_name"]

        d_col1, d_col2 = st.columns(2)
        d_col1.download_button(
            label="📄 Download Appointment Letter PDF",
            data=st.session_state["appt_pdf"],
            file_name=f"Appointment_Letter_{clean_name}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        d_col2.download_button(
            label="📄 Download Thanking Letter PDF",
            data=st.session_state["thank_pdf"],
            file_name=f"Thank_You_Letter_{clean_name}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

        # --------------------------------------------------------------
        # 6. Send via SMTP (real attachments, no manual step) or fall back
        #    to opening a mail client / Outlook Web for manual review.
        # --------------------------------------------------------------
        st.markdown("---")
        st.subheader("5. Send Email")

        send_mode = st.radio(
            "How do you want to send this?",
            [
                "Send automatically now (SMTP, attachments included)",
                "Open in Outlook Web / Mail client instead (manual attach required)",
            ],
        )

        which_attachments = st.multiselect(
            "Attach which letters?",
            ["Appointment Letter", "Thanking Letter"],
            default=["Appointment Letter", "Thanking Letter"],
        )

        attachments = []
        if "Appointment Letter" in which_attachments:
            attachments.append((f"Appointment_Letter_{clean_name}.pdf", st.session_state["appt_pdf"]))
        if "Thanking Letter" in which_attachments:
            attachments.append((f"Thank_You_Letter_{clean_name}.pdf", st.session_state["thank_pdf"]))

        if send_mode.startswith("Send automatically"):
            smtp_host, smtp_port, smtp_user, smtp_pass, from_secrets = get_smtp_credentials()
            if from_secrets:
                st.caption(f"Using configured sender: {smtp_user}")

            confirm = st.checkbox(f"Confirm: send to {email_to or '(no address entered)'}")
            if st.button("📤 Send Email Now", type="primary", disabled=not (email_to and confirm)):
                try:
                    send_email_with_attachments(
                        smtp_host, smtp_port, smtp_user, smtp_pass,
                        email_to, email_subject, email_body, attachments,
                    )
                    st.success(f"✅ Email sent to {email_to} with {len(attachments)} attachment(s).")
                except Exception as e:
                    st.error(f"❌ Send failed: {e}")
                    st.info(
                        "Common causes: wrong host/port, account needs an app password "
                        "(Office 365 with MFA), or SMTP-AUTH is disabled for this mailbox."
                    )
        else:
            st.info(
                "⚠️ Browser mail links can't attach files automatically — download the "
                "PDF(s) above first, then attach them manually in the compose window "
                "that opens below."
            )
            encoded_subject = urllib.parse.quote(email_subject)
            encoded_body = urllib.parse.quote(email_body)
            mailto_url = f"mailto:{email_to}?subject={encoded_subject}&body={encoded_body}"
            owa_url = (
                f"https://outlook.office.com/mail/deeplink/compose?"
                f"to={urllib.parse.quote(email_to)}&subject={encoded_subject}&body={encoded_body}"
            )

            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                st.link_button("🌐 Open in Outlook Web", owa_url, type="primary", use_container_width=True)
            with btn_col2:
                st.link_button("✉️ Open in Default Mail Client", mailto_url, use_container_width=True)


if __name__ == "__main__":
    show()