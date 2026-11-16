#!/usr/bin/env python3
"""
Generate BloodBridge Presentation (PPTX) with an executive Dark Mode Theme.
- Dark theme background (#0F172A / #0B0F19)
- High contrast light text (#FFFFFF / #F1F5F9 / #CBD5E1)
- Large readable text sizes (13-16pt body, 20-22pt card titles, 28-32pt slide titles)
- Removed "Student 1, 2, 3" headers from Slide 1
- Names and register numbers prominent, no topics underneath
- No presenter badges on slides
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = "/Users/raghavmaheshwari/Desktop/Blood Donor System"
DOCS_DIR = os.path.join(BASE_DIR, "docs")
OUTPUT_PPTX = os.path.join(BASE_DIR, "BloodBridge_Presentation.pptx")
SRM_LOGO = os.path.join(DOCS_DIR, "srm_logo.png")

# Premium Dark Mode Palette
COLOR_BG_DARK = RGBColor(11, 15, 25)          # #0B0F19 (Obsidian Navy background)
COLOR_HEADER_BAR = RGBColor(15, 23, 42)       # #0F172A (Deep Slate Header)
COLOR_CARD_DARK = RGBColor(24, 33, 54)        # #182136 (Deep Card Slate)
COLOR_CARD_INNER = RGBColor(30, 41, 59)       # #1E293B (Card Section Slate)
COLOR_BORDER = RGBColor(51, 65, 85)           # #334155 (Subtle Slate Border)

# Accent Colors
COLOR_ROSE_ACCENT = RGBColor(244, 63, 94)     # #F43F5E (Vibrant Rose/Crimson)
COLOR_EMERALD = RGBColor(52, 211, 153)        # #34D399 (Mint Emerald)
COLOR_BLUE_ACCENT = RGBColor(96, 165, 250)    # #60A5FA (Vivid Sky Blue)
COLOR_AMBER = RGBColor(251, 191, 36)          # #FBBF24 (Soft Gold)

# Typography Colors (High contrast light text)
COLOR_TEXT_WHITE = RGBColor(255, 255, 255)    # #FFFFFF
COLOR_TEXT_PRIMARY = RGBColor(241, 245, 249)  # #F1F5F9 (95% White)
COLOR_TEXT_SECONDARY = RGBColor(203, 213, 225)# #CBD5E1 (Soft Silver)
COLOR_TEXT_MUTED = RGBColor(148, 163, 184)    # #94A3B8 (Slate Muted)
COLOR_WHITE = RGBColor(255, 255, 255)

def apply_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_BG_DARK

def add_header(slide, title_text, category_tag="BLOODBRIDGE • ADVANCED PROGRAMMING PRACTICE"):
    # Dark Header Banner
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_HEADER_BAR
    bar.line.fill.background()

    # Glowing Rose Accent Line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.2), Inches(13.333), Inches(0.06))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_ROSE_ACCENT
    line.line.fill.background()

    # Category breadcrumb
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.5), Inches(0.35))
    p_tag = tag_box.text_frame.paragraphs[0]
    p_tag.text = category_tag.upper()
    p_tag.font.name = "Helvetica"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_ROSE_ACCENT

    # Main slide title (large, 28pt bold white)
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.42), Inches(11.8), Inches(0.7))
    p = title_box.text_frame.paragraphs[0]
    p.text = title_text
    p.font.name = "Helvetica"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_WHITE

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD_DARK, border_color=COLOR_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1.4)
    else:
        card.line.fill.background()
    return card

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: TITLE & TEAM SLIDE (DARK MODE)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    apply_background(slide1)

    # Top thin accent strip
    top_strip = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    top_strip.fill.solid()
    top_strip.fill.fore_color.rgb = COLOR_ROSE_ACCENT
    top_strip.line.fill.background()

    # SRM Logo container card (white pill so official logo renders cleanly on dark bg)
    logo_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(3.4), Inches(1.15))
    logo_card.fill.solid()
    logo_card.fill.fore_color.rgb = COLOR_WHITE
    logo_card.line.color.rgb = COLOR_ROSE_ACCENT
    logo_card.line.width = Pt(1.5)

    if os.path.exists(SRM_LOGO):
        slide1.shapes.add_picture(SRM_LOGO, Inches(0.9), Inches(0.48), width=Inches(3.2))

    # University & Department text (Right of logo)
    dept_box = slide1.shapes.add_textbox(Inches(4.5), Inches(0.45), Inches(8.0), Inches(1.1))
    dept_tf = dept_box.text_frame
    p_d1 = dept_tf.paragraphs[0]
    p_d1.text = "SRM Institute of Science and Technology • Kattankulathur"
    p_d1.font.bold = True
    p_d1.font.size = Pt(14)
    p_d1.font.color.rgb = COLOR_BLUE_ACCENT
    p_d2 = dept_tf.add_paragraph()
    p_d2.text = "Department of Computational Intelligence  |  21CSC203P Advanced Programming Practice"
    p_d2.font.size = Pt(12)
    p_d2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Hero Banner Card
    create_card(slide1, Inches(0.8), Inches(1.75), Inches(11.733), Inches(2.45), COLOR_CARD_DARK, COLOR_BORDER)

    # Rose vertical accent bar
    hero_accent = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.75), Inches(0.22), Inches(2.45))
    hero_accent.fill.solid()
    hero_accent.fill.fore_color.rgb = COLOR_ROSE_ACCENT
    hero_accent.line.fill.background()

    title_box = slide1.shapes.add_textbox(Inches(1.2), Inches(1.9), Inches(11.0), Inches(2.1))
    t_frame = title_box.text_frame
    p_t1 = t_frame.paragraphs[0]
    p_t1.text = "BloodBridge 🩸"
    p_t1.font.size = Pt(40)
    p_t1.font.bold = True
    p_t1.font.color.rgb = COLOR_ROSE_ACCENT

    p_t2 = t_frame.add_paragraph()
    p_t2.text = "Blood Donor Registration and Emergency Search System"
    p_t2.font.size = Pt(21)
    p_t2.font.bold = True
    p_t2.font.color.rgb = COLOR_TEXT_WHITE

    p_t3 = t_frame.add_paragraph()
    p_t3.text = "A zero-dependency standalone desktop front-end in Python & Tkinter with real-time ABO/Rh compatibility matching."
    p_t3.font.size = Pt(13.5)
    p_t3.font.color.rgb = COLOR_TEXT_SECONDARY

    # 3 Team Member Cards - NO "Student 1,2,3" headers, NO topics underneath!
    team_members = [
        ("Yana Butala", "RA2511026010676"),
        ("Raghav Maheshwari", "RA2511026010666"),
        ("Shorya Sharma", "RA25111026010618"),
    ]

    for idx, (name, reg_no) in enumerate(team_members):
        card_x = Inches(0.8 + idx * 4.0)
        card_w = Inches(3.733)
        
        # Sleek dark card
        create_card(slide1, card_x, Inches(4.45), card_w, Inches(2.5), COLOR_CARD_DARK, COLOR_BORDER)

        # Thin glowing accent line on top of each member card
        accent_top = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, Inches(4.45), card_w, Inches(0.08))
        accent_top.fill.solid()
        accent_top.fill.fore_color.rgb = COLOR_ROSE_ACCENT if idx == 1 else COLOR_BLUE_ACCENT
        accent_top.line.fill.background()

        # Name and Registration Number Box (Large, Clear Typography)
        n_box = slide1.shapes.add_textbox(card_x + Inches(0.3), Inches(4.8), card_w - Inches(0.6), Inches(1.8))
        ntf = n_box.text_frame
        ntf.word_wrap = True
        
        # Name: 22pt Bold White
        p1 = ntf.paragraphs[0]
        p1.text = name
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_TEXT_WHITE
        p1.space_after = Pt(8)

        # Register Number: 16pt Bold Rose
        p2 = ntf.add_paragraph()
        p2.text = reg_no
        p2.font.size = Pt(16)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_ROSE_ACCENT
        p2.space_after = Pt(8)

        # Class details: 13pt Muted Silver
        p3 = ntf.add_paragraph()
        p3.text = "II Year • III Semester • AIML (AK1)"
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_SECONDARY

    # =========================================================================
    # SLIDE 2: PROBLEM STATEMENT & OBJECTIVES (DARK MODE)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    apply_background(slide2)
    add_header(slide2, "1. Problem Statement & Motivation", "OVERVIEW & CLINICAL CONTEXT")

    # Left Card: The Challenge
    create_card(slide2, Inches(0.8), Inches(1.45), Inches(5.65), Inches(5.6), COLOR_CARD_DARK)
    
    # Left Card Header Pill
    h1 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.05), Inches(1.7), Inches(5.15), Inches(0.6))
    h1.fill.solid()
    h1.fill.fore_color.rgb = COLOR_CARD_INNER
    h1.line.color.rgb = COLOR_ROSE_ACCENT
    p_h1 = h1.text_frame.paragraphs[0]
    p_h1.text = "⚠️ THE CLINICAL CHALLENGE"
    p_h1.font.bold = True
    p_h1.font.size = Pt(14)
    p_h1.font.color.rgb = COLOR_ROSE_ACCENT
    p_h1.alignment = PP_ALIGN.CENTER

    body1 = slide2.shapes.add_textbox(Inches(1.05), Inches(2.45), Inches(5.15), Inches(4.3))
    b1_tf = body1.text_frame
    b1_tf.word_wrap = True

    pts_problem = [
        ("Time-Critical Emergencies", "In trauma and surgical emergencies, finding compatible blood within minutes directly impacts patient survival."),
        ("Flawed Exact-Match Systems", "Existing hospital portals often filter only for identical blood groups, missing universal donors (O−) and compatible types (e.g. B+ can receive B−, O+, and O−)."),
        ("Ineligible Donor Waste", "Contacting donors who donated within the past 90 days wastes critical triage time."),
        ("Heavy Database Dependency", "Complex server-dependent systems crash during local power cuts, network drops, or rural emergency camps."),
    ]

    for i, (title, desc) in enumerate(pts_problem):
        p_t = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p_t.text = f"• {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(13.5)
        p_t.font.color.rgb = COLOR_TEXT_WHITE
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(12.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p_t.space_after = Pt(10)

    # Right Card: The Solution
    create_card(slide2, Inches(6.88), Inches(1.45), Inches(5.65), Inches(5.6), COLOR_CARD_DARK)

    h2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.13), Inches(1.7), Inches(5.15), Inches(0.6))
    h2.fill.solid()
    h2.fill.fore_color.rgb = COLOR_CARD_INNER
    h2.line.color.rgb = COLOR_EMERALD
    p_h2 = h2.text_frame.paragraphs[0]
    p_h2.text = "🎯 THE BLOODBRIDGE SOLUTION"
    p_h2.font.bold = True
    p_h2.font.size = Pt(14)
    p_h2.font.color.rgb = COLOR_EMERALD
    p_h2.alignment = PP_ALIGN.CENTER

    body2 = slide2.shapes.add_textbox(Inches(7.13), Inches(2.45), Inches(5.15), Inches(4.3))
    b2_tf = body2.text_frame
    b2_tf.word_wrap = True

    pts_solution = [
        ("Instant Zero-Latency Desktop UI", "Engineered in Python Tkinter for immediate launch with zero startup delay on any laptop."),
        ("Smart Emergency Matching Engine", "Implements full ABO & Rh compatibility rules, instantly identifying all clinically safe donor types for urgent patients."),
        ("Automated Cooldown Engine", "Enforces the mandatory 90-day cooldown interval and clinical age boundaries (18–65) dynamically."),
        ("100% Zero-Dependency Standalone", "Runs natively on the Python standard library with no pip packages or database setup required."),
    ]

    for i, (title, desc) in enumerate(pts_solution):
        p_t = b2_tf.paragraphs[0] if i == 0 else b2_tf.add_paragraph()
        p_t.text = f"✓ {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(13.5)
        p_t.font.color.rgb = COLOR_EMERALD
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(12.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p_t.space_after = Pt(10)

    # =========================================================================
    # SLIDE 3: SYSTEM ARCHITECTURE & STRUCTURE
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    apply_background(slide3)
    add_header(slide3, "2. System Architecture & Component Hierarchy", "TECHNICAL ARCHITECTURE")

    # Left: Architecture Diagram Card
    create_card(slide3, Inches(0.8), Inches(1.45), Inches(7.5), Inches(5.6), COLOR_CARD_DARK)
    
    # White inner container so the flowchart diagram displays with perfect clarity
    diag_container = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.7), Inches(7.1), Inches(4.1))
    diag_container.fill.solid()
    diag_container.fill.fore_color.rgb = COLOR_WHITE
    diag_container.line.fill.background()

    fig1_path = os.path.join(DOCS_DIR, "fig1_structure.png")
    if os.path.exists(fig1_path):
        slide3.shapes.add_picture(fig1_path, Inches(1.25), Inches(1.95), width=Inches(6.6))

    c_box = slide3.shapes.add_textbox(Inches(1.05), Inches(6.0), Inches(7.0), Inches(0.8))
    p_c = c_box.text_frame.paragraphs[0]
    p_c.text = "Figure 1: Event-driven architecture connecting user inputs, cooldown logic & in-memory store to all 3 tabs."
    p_c.font.size = Pt(12)
    p_c.font.italic = True
    p_c.font.color.rgb = COLOR_TEXT_SECONDARY
    p_c.alignment = PP_ALIGN.CENTER

    # Right: Modules Explanation Card
    create_card(slide3, Inches(8.6), Inches(1.45), Inches(3.933), Inches(5.6), COLOR_CARD_DARK)
    
    h3 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.8), Inches(1.7), Inches(3.533), Inches(0.6))
    h3.fill.solid()
    h3.fill.fore_color.rgb = COLOR_CARD_INNER
    h3.line.color.rgb = COLOR_BLUE_ACCENT
    p_h3 = h3.text_frame.paragraphs[0]
    p_h3.text = "⚙️ CORE MODULES"
    p_h3.font.bold = True
    p_h3.font.size = Pt(14)
    p_h3.font.color.rgb = COLOR_BLUE_ACCENT
    p_h3.alignment = PP_ALIGN.CENTER

    rh_box = slide3.shapes.add_textbox(Inches(8.8), Inches(2.45), Inches(3.533), Inches(4.4))
    rh_tf = rh_box.text_frame
    rh_tf.word_wrap = True

    modules = [
        ("BloodDonorApp (Controller)", "Main Tkinter class managing ttk.Notebook tabs, window layout, and global theme styling."),
        ("InMemoryStore (Model)", "Dictionary & list-based simulated repository pre-loaded with realistic donor records across 10 cities."),
        ("Eligibility Engine (Logic)", "Real-time date arithmetic computing days elapsed since last donation against the 90-day cooldown limit."),
        ("Compatibility Engine (Algorithm)", "Hash map matrix resolving universal donor and recipient relationships in O(1) time."),
    ]

    for i, (title, desc) in enumerate(modules):
        p_m = rh_tf.paragraphs[0] if i == 0 else rh_tf.add_paragraph()
        p_m.text = f"• {title}: "
        p_m.font.bold = True
        p_m.font.size = Pt(13)
        p_m.font.color.rgb = COLOR_TEXT_WHITE
        run = p_m.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p_m.space_after = Pt(10)

    # =========================================================================
    # SLIDE 4: FEATURE DEEP-DIVE – SEARCH & EMERGENCY MATCH
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    apply_background(slide4)
    add_header(slide4, "3. Feature Deep-Dive: Emergency Search & Directory", "TAB 1 SPOTLIGHT")

    # Left: Tab 1 Screenshot (Directory View)
    create_card(slide4, Inches(0.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_CARD_DARK)
    pic1_path = os.path.join(DOCS_DIR, "pic1_search.png")
    if os.path.exists(pic1_path):
        slide4.shapes.add_picture(pic1_path, Inches(0.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Right: Tab 1 Emergency Match Screenshot
    create_card(slide4, Inches(6.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_CARD_DARK)
    pic2_path = os.path.join(DOCS_DIR, "pic2_compat.png")
    if os.path.exists(pic2_path):
        slide4.shapes.add_picture(pic2_path, Inches(6.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Bottom Banner: Feature Talking Points (Large readable 13pt font)
    create_card(slide4, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.4), COLOR_CARD_DARK)
    b_box4 = slide4.shapes.add_textbox(Inches(1.05), Inches(5.9), Inches(11.2), Inches(1.25))
    b_tf4 = b_box4.text_frame
    b_tf4.word_wrap = True

    p4_1 = b_tf4.paragraphs[0]
    p4_1.text = "🔍 Key Search Capabilities Demonstrated Above:"
    p4_1.font.size = Pt(14)
    p4_1.font.bold = True
    p4_1.font.color.rgb = COLOR_ROSE_ACCENT
    p4_1.space_after = Pt(4)

    pts_s4 = (
        "• Real-Time KPI Ribbon: Live dashboard cards calculating Total Donors (10), Ready Donors (8), Cooldown (2), and Universal O− (2).\n"
        "• Emergency Recipient Mode: Searching for Patient B+ instantly returns all 4 clinically compatible types (B+, B−, O+, O−).\n"
        "• Interactive ttk.Treeview Table: Color-coded tags (Green = Ready, Amber = Cooldown, Red = Unavailable) with sortable columns.\n"
        "• Instant Actions: Quick donor dispatch panel with one-click 'Record Donation Today' (resets 90d cooldown) and CSV export."
    )
    p4_2 = b_tf4.add_paragraph()
    p4_2.text = pts_s4
    p4_2.font.size = Pt(12)
    p4_2.font.color.rgb = COLOR_TEXT_PRIMARY

    # =========================================================================
    # SLIDE 5: FEATURE DEEP-DIVE – REGISTRATION & MATRIX
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    apply_background(slide5)
    add_header(slide5, "4. Feature Deep-Dive: Registration & Transfusion Matrix", "TABS 2 & 3 SPOTLIGHT")

    # Left: Tab 2 Registration Portal Screenshot
    create_card(slide5, Inches(0.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_CARD_DARK)
    pic3_path = os.path.join(DOCS_DIR, "pic3_register.png")
    if os.path.exists(pic3_path):
        slide5.shapes.add_picture(pic3_path, Inches(0.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Right: Tab 3 Compatibility Matrix Screenshot
    create_card(slide5, Inches(6.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_CARD_DARK)
    pic4_path = os.path.join(DOCS_DIR, "pic4_matrix.png")
    if os.path.exists(pic4_path):
        slide5.shapes.add_picture(pic4_path, Inches(6.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Bottom Banner: Feature Talking Points (Large readable 13pt font)
    create_card(slide5, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.4), COLOR_CARD_DARK)
    b_box5 = slide5.shapes.add_textbox(Inches(1.05), Inches(5.9), Inches(11.2), Inches(1.25))
    b_tf5 = b_box5.text_frame
    b_tf5.word_wrap = True

    p5_1 = b_tf5.paragraphs[0]
    p5_1.text = "📝 Registration & Clinical Transfusion Matrix:"
    p5_1.font.size = Pt(14)
    p5_1.font.bold = True
    p5_1.font.color.rgb = COLOR_ROSE_ACCENT
    p5_1.space_after = Pt(4)

    pts_s5 = (
        "• Clinical Validation Rules: Enforces medical standards: Age (18–65), Weight (≥ 50 kg), and 10-digit phone verification.\n"
        "• First-Time vs Repeat Donors: Dedicated toggle disables date entry for first-timers; validates past dates for repeat donors.\n"
        "• Inline Edit Mode: Double-clicking any record in the directory loads existing data into the form for instant updates.\n"
        "• Transfusion Reference Matrix: Full ABO/Rh red blood cell rules + live donor pool supply counters across all 8 blood groups."
    )
    p5_2 = b_tf5.add_paragraph()
    p5_2.text = pts_s5
    p5_2.font.size = Pt(12)
    p5_2.font.color.rgb = COLOR_TEXT_PRIMARY

    # =========================================================================
    # SLIDE 6: TECHNICAL IMPLEMENTATION & CONCLUSION
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    apply_background(slide6)
    add_header(slide6, "5. Technical Implementation & Conclusion", "CONCLUSION & DEMO")

    card_w = Inches(3.733)

    # Card 1: Technical Stack
    create_card(slide6, Inches(0.8), Inches(1.5), card_w, Inches(4.3), COLOR_CARD_DARK)
    c1_hdr = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), card_w, Inches(0.5))
    c1_hdr.fill.solid()
    c1_hdr.fill.fore_color.rgb = COLOR_CARD_INNER
    c1_hdr.line.color.rgb = COLOR_BLUE_ACCENT
    p_c1h = c1_hdr.text_frame.paragraphs[0]
    p_c1h.text = "💻 TECHNICAL HIGHLIGHTS"
    p_c1h.font.bold = True
    p_c1h.font.size = Pt(12.5)
    p_c1h.font.color.rgb = COLOR_BLUE_ACCENT
    p_c1h.alignment = PP_ALIGN.CENTER

    b1_box = slide6.shapes.add_textbox(Inches(1.0), Inches(2.15), Inches(3.333), Inches(3.5))
    b1_tf = b1_box.text_frame
    b1_tf.word_wrap = True
    t_tech = [
        ("Python Standard Library", "Engineered purely with tkinter, ttk, datetime, csv, and re. Zero external pip packages needed."),
        ("Single-File Standalone", "Entire architecture contained in a clean 650-line bloodbridge.py module."),
        ("Lightweight & Cross-Platform", "Instant boot time (<0.2s), minimal RAM footprint (<25 MB), runs on Windows, macOS, Linux."),
    ]
    for i, (t, d) in enumerate(t_tech):
        p = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(12.5)
        p.font.color.rgb = COLOR_TEXT_WHITE
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p.space_after = Pt(10)

    # Card 2: Course Outcomes
    create_card(slide6, Inches(4.8), Inches(1.5), card_w, Inches(4.3), COLOR_CARD_DARK)
    c2_hdr = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.8), Inches(1.5), card_w, Inches(0.5))
    c2_hdr.fill.solid()
    c2_hdr.fill.fore_color.rgb = COLOR_CARD_INNER
    c2_hdr.line.color.rgb = COLOR_EMERALD
    p_c2h = c2_hdr.text_frame.paragraphs[0]
    p_c2h.text = "🏆 COURSE OUTCOMES"
    p_c2h.font.bold = True
    p_c2h.font.size = Pt(12.5)
    p_c2h.font.color.rgb = COLOR_EMERALD
    p_c2h.alignment = PP_ALIGN.CENTER

    b2_box = slide6.shapes.add_textbox(Inches(5.0), Inches(2.15), Inches(3.333), Inches(3.5))
    b2_tf = b2_box.text_frame
    b2_tf.word_wrap = True
    t_out = [
        ("Advanced GUI Programming", "Mastered Tkinter event binding, PanedWindow, Treeview manipulation, and custom theme styling."),
        ("Medical Domain Logic", "Translated biological ABO & Rh compatibility charts into efficient algorithmic search filters."),
        ("Collaborative Engineering", "Clean Git workflow, professional documentation, report generation, and presentation synthesis."),
    ]
    for i, (t, d) in enumerate(t_out):
        p = b2_tf.paragraphs[0] if i == 0 else b2_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(12.5)
        p.font.color.rgb = COLOR_TEXT_WHITE
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p.space_after = Pt(10)

    # Card 3: Future Roadmap
    create_card(slide6, Inches(8.8), Inches(1.5), card_w, Inches(4.3), COLOR_CARD_DARK)
    c3_hdr = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.8), Inches(1.5), card_w, Inches(0.5))
    c3_hdr.fill.solid()
    c3_hdr.fill.fore_color.rgb = COLOR_CARD_INNER
    c3_hdr.line.color.rgb = COLOR_ROSE_ACCENT
    p_c3h = c3_hdr.text_frame.paragraphs[0]
    p_c3h.text = "🚀 FUTURE ROADMAP"
    p_c3h.font.bold = True
    p_c3h.font.size = Pt(12.5)
    p_c3h.font.color.rgb = COLOR_ROSE_ACCENT
    p_c3h.alignment = PP_ALIGN.CENTER

    b3_box = slide6.shapes.add_textbox(Inches(9.0), Inches(2.15), Inches(3.333), Inches(3.5))
    b3_tf = b3_box.text_frame
    b3_tf.word_wrap = True
    t_fut = [
        ("Cloud Database Integration", "Adding SQLite/PostgreSQL connectors for multi-hospital synchronization."),
        ("Emergency SMS / WhatsApp Alerts", "Integrating Twilio API to broadcast emergency donor dispatch alerts instantly."),
        ("GIS Distance Matching", "Geocoding donor addresses to sort candidates by closest driving distance to emergency hospital."),
    ]
    for i, (t, d) in enumerate(t_fut):
        p = b3_tf.paragraphs[0] if i == 0 else b3_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(12.5)
        p.font.color.rgb = COLOR_TEXT_WHITE
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p.space_after = Pt(10)

    # Bottom Banner: GitHub & Thank You
    create_card(slide6, Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.9), COLOR_CARD_DARK)
    bot_box = slide6.shapes.add_textbox(Inches(1.0), Inches(6.08), Inches(11.333), Inches(0.75))
    bot_tf = bot_box.text_frame
    p_b1 = bot_tf.paragraphs[0]
    p_b1.text = "🔗 GitHub Repository: https://github.com/raghavmaheshwari-web/Blood-Bridge"
    p_b1.font.size = Pt(13)
    p_b1.font.bold = True
    p_b1.font.color.rgb = COLOR_ROSE_ACCENT
    p_b2 = bot_tf.add_paragraph()
    p_b2.text = "Thank you! Open for Questions & Live Demonstration."
    p_b2.font.size = Pt(12)
    p_b2.font.color.rgb = COLOR_TEXT_WHITE

    prs.save(OUTPUT_PPTX)
    print("Presentation saved:", OUTPUT_PPTX)

if __name__ == "__main__":
    build_presentation()
