#!/usr/bin/env python3
"""
Generate the exact 2-page PDF for BloodBridge matching the reference PDF layout and indentation.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch, mm
pt = 1.0
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, KeepTogether
)
import os

BASE_DIR = "/Users/raghavmaheshwari/Desktop/Blood Donor System"
DOCS_DIR = os.path.join(BASE_DIR, "docs")
SRM_LOGO = "/tmp/docx_image1.png"
OUTPUT_PDF = os.path.join(BASE_DIR, "RA2511026010666.pdf")

# Page Dimensions (A4: 595.27 x 841.89 pt)
PAGE_W, PAGE_H = A4
LEFT_MARGIN = 45 * pt
RIGHT_MARGIN = 45 * pt
TOP_MARGIN = 32 * pt
BOTTOM_MARGIN = 32 * pt
USABLE_WIDTH = PAGE_W - (LEFT_MARGIN + RIGHT_MARGIN)  # 505.27 pt

def create_pdf():
    doc = SimpleDocTemplate(
        OUTPUT_PDF,
        pagesize=A4,
        leftMargin=LEFT_MARGIN,
        rightMargin=RIGHT_MARGIN,
        topMargin=TOP_MARGIN,
        bottomMargin=BOTTOM_MARGIN,
    )

    styles = getSampleStyleSheet()

    # Fonts & Styles matching reference PDF
    style_main_title = ParagraphStyle(
        "MainTitle",
        fontName="Helvetica-Bold",
        fontSize=12.5,
        leading=15,
        textColor=black,
        alignment=TA_CENTER,
        spaceAfter=4,
    )

    style_header_center = ParagraphStyle(
        "HeaderCenter",
        fontName="Helvetica",
        fontSize=9.5,
        leading=12.5,
        textColor=black,
        alignment=TA_CENTER,
    )

    style_submitted_by = ParagraphStyle(
        "SubmittedBy",
        fontName="Helvetica",
        fontSize=9.5,
        leading=13.5,
        textColor=black,
        alignment=TA_RIGHT,
    )

    style_section_heading = ParagraphStyle(
        "SectionHeading",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=black,
        alignment=TA_LEFT,
        spaceBefore=8,
        spaceAfter=4,
    )

    style_body = ParagraphStyle(
        "Body",
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        textColor=black,
        alignment=TA_JUSTIFY,
        spaceAfter=4,
    )

    style_table_cell = ParagraphStyle(
        "TableCell",
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=black,
    )

    style_table_cell_bold = ParagraphStyle(
        "TableCellBold",
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        textColor=black,
    )

    style_bullet = ParagraphStyle(
        "Bullet",
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        textColor=black,
        leftIndent=14,
        firstLineIndent=-10,
        spaceAfter=2,
    )

    style_caption = ParagraphStyle(
        "Caption",
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=black,
        alignment=TA_CENTER,
        spaceBefore=3,
        spaceAfter=4,
    )

    style_github_link = ParagraphStyle(
        "GithubLink",
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        textColor=HexColor("#0D47A1"),
        alignment=TA_LEFT,
    )

    story = []

    # =========================================================================
    # PAGE 1
    # =========================================================================
    
    # 1. Header with Logo on Top Left
    logo_img = Image(SRM_LOGO, width=1.65 * inch, height=0.55 * inch)
    logo_img.hAlign = 'LEFT'
    
    # Table to place logo on left cleanly without pushing title down
    hdr_table = Table([[logo_img]], colWidths=[USABLE_WIDTH])
    hdr_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(hdr_table)

    # 2. Main Title (Centered)
    story.append(Paragraph("BLOODBRIDGE: A BLOOD DONOR REGISTRATION AND SEARCH SYSTEM", style_main_title))

    # 3. Sub-header (Centered academic details)
    sub_text = (
        "College of Engineering and Technology<br/>"
        "School of Computing<br/>"
        "Department of Computational Intelligence<br/>"
        "Academic Year 2026 – 2027, Nov 2026.<br/>"
        "21CSC203P : Advanced Programming Practice<br/>"
        "II Year – III Semester, Specialization – AIML, Section – AK1"
    )
    story.append(Paragraph(sub_text, style_header_center))
    story.append(Spacer(1, 4))

    # 4. Submitted By (Right-aligned, matching reference format)
    sub_by_text = (
        "Submitted by<br/>"
        "1. RA2511026010666 – Raghav Maheshwari,<br/>"
        "2. RA2511026010618 – Shorya Sharma,<br/>"
        "3. RA2511026010676 – Yana Butala."
    )
    story.append(Paragraph(sub_by_text, style_submitted_by))
    story.append(Spacer(1, 6))

    # 5. Section 1: ABSTRACT
    story.append(Paragraph("1. ABSTRACT", style_section_heading))
    abstract_text = (
        "BloodBridge is a blood donor registration and emergency search system front end built with Python "
        "and Tkinter. It brings the main features of blood donation drives and blood bank registries into one desktop "
        "window with a clean medical design: deep crimson accents with green and amber used for eligibility indicators. "
        "All data is simulated and managed in memory, with zero backend setup required, allowing rapid testing and demonstration."
    )
    story.append(Paragraph(abstract_text, style_body))
    story.append(Spacer(1, 4))

    # 6. Section 2: FRONT END DESIGN DESCRIPTION
    story.append(Paragraph("2. FRONT END DESIGN DESCRIPTION", style_section_heading))
    intro_desc = (
        "The app has three tabs: Search & Directory, Donor Registration, and Compatibility Matrix & Stats. "
        "It includes the core features found in modern blood bank management and donor matching systems:"
    )
    story.append(Paragraph(intro_desc, style_body))
    story.append(Spacer(1, 4))

    # Feature Table
    table_data = [
        [Paragraph("Feature", style_table_cell_bold), Paragraph("Tab", style_table_cell_bold)],
        [Paragraph("Emergency compatibility search (ABO and Rh)", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Real-time donor eligibility tracking (90-day cooldown)", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Multi-parameter filters (blood group, city, keyword)", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Interactive donor directory with color-coded status tags", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Selected donor inspection card and quick actions", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Record donation today with automatic cooldown reset", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Export visible donor search results to CSV", style_table_cell), Paragraph("Search &amp; Directory", style_table_cell)],
        [Paragraph("Registration form with validation (age 18–65, weight ≥ 50kg)", style_table_cell), Paragraph("Donor Registration", style_table_cell)],
        [Paragraph("Inline editing mode for existing donor records", style_table_cell), Paragraph("Donor Registration", style_table_cell)],
        [Paragraph("ABO &amp; Rh red blood cell compatibility matrix table", style_table_cell), Paragraph("Compatibility Matrix", style_table_cell)],
        [Paragraph("Current registered donor pool distribution by blood group", style_table_cell), Paragraph("Compatibility Matrix", style_table_cell)],
    ]

    col_widths = [345 * pt, 160 * pt]
    feat_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    feat_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.7, black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2.2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.2),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(feat_table)
    story.append(Spacer(1, 5))

    # How it helps text & bullets
    story.append(Paragraph("How it helps a donor coordinator or healthcare worker:", style_body))
    bullets = [
        "• Fast emergency matching shows all compatible donor blood types instantly for a recipient.",
        "• Automated eligibility rules prevent contacting donors who are currently in the 90-day cooldown.",
        "• Registration form validation enforces clinical donation requirements (age, weight, contact info).",
        "• Simulated in-memory storage allows complete workflow demonstration without external database setup.",
    ]
    for b in bullets:
        story.append(Paragraph(b, style_bullet))

    # Force Page Break to Page 2
    story.append(PageBreak())

    # =========================================================================
    # PAGE 2
    # =========================================================================
    
    # 1. Top Logo on Page 2 (matching reference PDF)
    logo_img_p2 = Image(SRM_LOGO, width=1.65 * inch, height=0.55 * inch)
    logo_img_p2.hAlign = 'LEFT'
    hdr_table_p2 = Table([[logo_img_p2]], colWidths=[USABLE_WIDTH])
    hdr_table_p2.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(hdr_table_p2)

    # 2. Section 3: DIAGRAM
    story.append(Paragraph("3. DIAGRAM", style_section_heading))
    story.append(Spacer(1, 1))

    # 3. Figure 1: Structure of the Application
    fig1_path = os.path.join(DOCS_DIR, "fig1_structure.png")
    fig1_w = USABLE_WIDTH
    fig1_h = fig1_w * (270.0 / 800.0)  # ~170.5 pt
    fig1_img = Image(fig1_path, width=fig1_w, height=fig1_h)
    fig1_img.hAlign = 'CENTER'
    story.append(fig1_img)
    story.append(Paragraph("Figure 1. Structure of the application.", style_caption))
    story.append(Spacer(1, 4))

    # 4. Figure 2: Four Pictures of Design (2x2 Grid)
    pic1_path = os.path.join(DOCS_DIR, "pic1_search.png")
    pic2_path = os.path.join(DOCS_DIR, "pic2_compat.png")
    pic3_path = os.path.join(DOCS_DIR, "pic3_register.png")
    pic4_path = os.path.join(DOCS_DIR, "pic4_matrix.png")

    grid_cell_w = (USABLE_WIDTH - 6 * pt) / 2.0  # ~249.6 pt
    grid_cell_h = grid_cell_w * (580.0 / 900.0)  # ~160.8 pt

    im1 = Image(pic1_path, width=grid_cell_w, height=grid_cell_h)
    im2 = Image(pic2_path, width=grid_cell_w, height=grid_cell_h)
    im3 = Image(pic3_path, width=grid_cell_w, height=grid_cell_h)
    im4 = Image(pic4_path, width=grid_cell_w, height=grid_cell_h)

    pics_grid = Table(
        [
            [im1, im2],
            [im3, im4]
        ],
        colWidths=[grid_cell_w, grid_cell_w]
    )
    pics_grid.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 1),
        ('RIGHTPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(pics_grid)

    caption_text = (
        "Figure 2. Search &amp; Directory, Emergency Compatibility Match, Donor Registration, "
        "and Compatibility Matrix views.<br/>Screenshots of the running Tkinter application."
    )
    story.append(Paragraph(caption_text, style_caption))
    story.append(Spacer(1, 5))

    # 5. Section 4: GITHUB LINK OF YOUR DESIGN
    story.append(Paragraph("4. GITHUB LINK OF YOUR DESIGN", style_section_heading))
    repo_url = "https://github.com/raghavmaheshwari-web/Blood-Bridge"
    story.append(Paragraph(f'<link href="{repo_url}">{repo_url}</link>', style_github_link))

    # Build PDF
    doc.build(story)
    print("PDF build complete:", OUTPUT_PDF)

if __name__ == "__main__":
    create_pdf()
