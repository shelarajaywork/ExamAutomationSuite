import io
import urllib.parse
import pandas as pd
import streamlit as st
from datetime import datetime

# ReportLab imports for dynamic PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY


# ==========================================
# PDF GENERATOR FUNCTIONS
# ==========================================

def build_pdf_letterhead(styles):
    """Creates standard institutional letterhead elements."""
    elements = []
    
    header_style = ParagraphStyle(
        'LetterheadSub',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#003366"),
        leading=11
    )
    
    title_style = ParagraphStyle(
        'LetterheadTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#003366"),
        leading=14
    )
    
    small_style = ParagraphStyle(
        'LetterheadSmall',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7.5,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#333333"),
        leading=9
    )

    elements.append(Paragraph("Shri Vile Parle Kelavani Mandal's", header_style))
    elements.append(Paragraph("SV K M COLLEGES EXAMINATION CELL", title_style))
    elements.append(Paragraph("(Autonomous Affiliated to University of Mumbai)", header_style))
    elements.append(Paragraph("Bhaktivedanta Swami Marg, Juhu Scheme, Vile Parle (West), Mumbai — 400 056.", small_style))
    elements.append(Paragraph("Tel.: 42332041/42 - Website: www.svkm.ac.in - Email: exam@svkm.ac.in", small_style))
    elements.append(Paragraph("<b>NAAC RE-ACCREDITED 'A+' GRADE</b>", small_style))
    elements.append(Spacer(1, 0.15 * inch))
    
    return elements


def generate_appointment_pdf(data):
    """Generates the official Appointment Letter PDF."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=54,
        leftMargin=54,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    body_style = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        alignment=TA_JUSTIFY,
        spaceAfter=10
    )
    
    bold_body = ParagraphStyle(
        'BoldBody',
        parent=body_style,
        fontName='Helvetica-Bold'
    )
    
    right_date = ParagraphStyle(
        'RightDate',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        alignment=TA_RIGHT
    )

    story = []
    
    # 1. Letterhead
    story.extend(build_pdf_letterhead(styles))
    
    # 2. Date
    current_date = datetime.now().strftime("%dth %B %Y")
    story.append(Paragraph(current_date, right_date))
    story.append(Spacer(1, 0.15 * inch))
    
    # 3. Addressee Block
    address_text = f"To,<br/><b>{data['examiner_name']}</b><br/>{data['college_name']}"
    story.append(Paragraph(address_text, body_style))
    story.append(Spacer(1, 0.1 * inch))
    
    # 4. Subject
    salutation = "Sir" if "Mr." in data['examiner_name'] else "Madam"
    subj_text = f"<b>Sub: Appointment as {data['role']} for Examination of {data['course']}. Year: {data['year']}, Semester: {data['semester']} ({data['ayear']})</b>"
    story.append(Paragraph(subj_text, body_style))
    story.append(Spacer(1, 0.1 * inch))
    
    # 5. Body
    p1 = f"Dear {salutation},<br/><br/>We are pleased to appoint you as <b>{data['role']}</b> for the End Semester Examination of the above-mentioned program, for the subject <b>{data['subject']}</b>."
    story.append(Paragraph(p1, body_style))
    
    p2 = f"The answer books are to be evaluated/moderated as per the examination guidelines of our institution. A total count of <b>{data['count']}</b> answer books is assigned for this task, as per the rules."
    story.append(Paragraph(p2, body_style))
    
    p3 = "Evaluation/Moderation is in online mode (Onscreen Marking) for which you will receive system credentials upon task assignment."
    story.append(Paragraph(p3, body_style))
    
    p4 = "An honorarium, as per university rules, will be paid towards the assignment."
    story.append(Paragraph(p4, body_style))
    story.append(Spacer(1, 0.2 * inch))
    
    # 6. Signature Block
    story.append(Paragraph("Thanking you,", body_style))
    story.append(Spacer(1, 0.4 * inch))
    story.append(Paragraph("<b>Controller of Examinations / Assistant Registrar</b>", bold_body))
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def generate_thanking_pdf(data):
    """Generates the official Thanking Letter PDF."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=54,
        leftMargin=54,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    body_style = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        alignment=TA_LEFT,
        spaceAfter=10
    )
    
    right_date = ParagraphStyle(
        'RightDate',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        alignment=TA_RIGHT
    )

    story = []
    
    # 1. Letterhead
    story.extend(build_pdf_letterhead(styles))
    
    # 2. Date
    current_date = datetime.now().strftime("%dth %B %Y")
    story.append(Paragraph(current_date, right_date))
    story.append(Spacer(1, 0.15 * inch))
    
    # 3. Addressee Block
    address_text = f"To,<br/><b>{data['examiner_name']}</b><br/>{data['college_name']}"
    story.append(Paragraph(address_text, body_style))
    story.append(Spacer(1, 0.15 * inch))
    
    # 4. Salutation & Opening
    salutation = "Sir" if "Mr." in data['examiner_name'] else "Madam"
    p1 = f"Dear {salutation},<br/><br/>We are thankful to you for agreeing to be the <b>{data['role']}</b> for the following subject at our college for the End Semester Examination:"
    story.append(Paragraph(p1, body_style))
    story.append(Spacer(1, 0.1 * inch))
    
    # 5. Structured Table
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
    
    # 6. Closing & Signature
    p2 = "The efforts that you have taken and the time you have spared for this task are truly appreciated.<br/>Thank you once again and looking forward to your continued support."
    story.append(Paragraph(p2, body_style))
    story.append(Spacer(1, 0.4 * inch))
    
    story.append(Paragraph("Sincerely,", body_style))
    story.append(Spacer(1, 0.4 * inch))
    story.append(Paragraph("<b>Controller of Examinations / Assistant Registrar</b>", body_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ==========================================
# MAIN STREAMLIT UI MODULE
# ==========================================

def show():
    st.title("✉️ Letters to External Examiners")
    st.caption("Auto-generate official PDF Appointment & Thanking letters from Excel and trigger pre-filled Outlook Web drafts.")

    st.markdown("---")

    # 1. Excel File Upload
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

    # Standardize Column Names
    col_map = {
        "ExaminerName": "examiner_name",
        "ExaminerEmail": "examiner_email",
        "RoleName": "role",
        "CourseName": "course",
        "CategoryName": "subject",
        "Semester/Trimester": "semester",
        "AYear": "ayear",
        "CheckCount": "count"
    }

    # Data Filter & Selection UI
    st.markdown("---")
    st.subheader("1. Select Examiner Record")
    
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
    
    paper_list = examiner_rows["CategoryName"].tolist() if "CategoryName" in examiner_rows.columns else [f"Record {i+1}" for i in range(len(examiner_rows))]
    
    if len(paper_list) > 1:
        sel_paper_idx = st.selectbox("Select Paper / Subject Assignment:", range(len(paper_list)), format_func=lambda x: paper_list[x])
        row = examiner_rows.iloc[sel_paper_idx]
    else:
        row = examiner_rows.iloc[0]

    # Map details safely
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
        "count": int(row.get("CheckCount", 0)) if pd.notna(row.get("CheckCount")) else 0
    }

    st.dataframe(pd.DataFrame([row]), use_container_width=True)

    # Email & PDF Options
    st.markdown("---")
    st.subheader("2. Configure & Auto-Generate Letters")

    email_to = st.text_input("Recipient Email Address:", value=record_data["examiner_email"])
    email_subject = f"Official Communication: {record_data['subject']} - {record_data['role']} Assignment"
    
    email_body = (
        f"Dear {record_data['examiner_name']},\n\n"
        f"Greetings from the Examinations Department.\n\n"
        f"Please find attached the official correspondence regarding your appointment as {record_data['role']} "
        f"for the course {record_data['course']} (Paper: {record_data['subject']}).\n\n"
        f"If you have any queries, please feel free to reach out to the Examination Cell.\n\n"
        f"Warm regards,\n"
        f"Assistant Registrar of Examinations\n"
        f"SVKM Examination Automation Suite"
    )

    st.text_area("Email Content Preview:", value=email_body, height=150, disabled=True)

    # Generation Trigger
    if st.button("⚡ Auto-Create PDF Letters & Prepare Email", type="primary"):
        # Auto-create PDFs in memory
        appt_pdf_bytes = generate_appointment_pdf(record_data)
        thank_pdf_bytes = generate_thanking_pdf(record_data)
        
        st.success("✅ **Appointment Letter PDF & Thanking Letter PDF generated successfully!**")

        st.markdown("---")
        st.subheader("3. Action & Outlook Web Redirection")

        # Downloads
        d_col1, d_col2 = st.columns(2)
        
        clean_name = record_data['examiner_name'].replace(' ', '_')
        
        d_col1.download_button(
            label="📄 Download Appointment Letter PDF",
            data=appt_pdf_bytes,
            file_name=f"Appointment_Letter_{clean_name}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
        d_col2.download_button(
            label="📄 Download Thanking Letter PDF",
            data=thank_pdf_bytes,
            file_name=f"Thank_You_Letter_{clean_name}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

        # Build Deep-Link URLs
        # 1. Standard Mailto Link (Opens default web/desktop email client pre-filled)
        encoded_subject = urllib.parse.quote(email_subject)
        encoded_body = urllib.parse.quote(email_body)
        mailto_url = f"mailto:{email_to}?subject={encoded_subject}&body={encoded_body}"

        # 2. Direct Outlook Web (Office 365) Deep-Link
        owa_url = f"https://outlook.office.com/mail/deeplink/compose?to={urllib.parse.quote(email_to)}&subject={encoded_subject}&body={encoded_body}"

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Display Action Buttons
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            st.link_button("🌐 Open in Outlook Web (Office 365)", owa_url, type="primary", use_container_width=True)
            
        with btn_col2:
            st.link_button("✉️ Open in Default Mail Client", mailto_url, use_container_width=True)

        st.info("💡 **Workflow Note:** Click the download buttons above to save the generated PDF letters, then click **'Open in Outlook Web'**. Outlook Web will open with a draft email pre-filled with the recipient, subject, and text. Simply attach the downloaded PDFs and hit Send!")


if __name__ == "__main__":
    show()