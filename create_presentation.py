#!/usr/bin/env python3
"""
Generate an executive, highly readable 6-slide presentation for BloodBridge.
Modern design: Deep Slate Navy with Ruby Crimson accents (not harsh red & white).
Clean, large readable typography (13-16pt body), no presenter names on slides,
no topics under student names.
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

# Sophisticated modern color palette
COLOR_NAVY = RGBColor(15, 23, 42)          # #0F172A (Deep Slate Navy)
COLOR_SLATE_HEADER = RGBColor(30, 41, 59)  # #1E293B
COLOR_ACCENT_RED = RGBColor(225, 29, 72)   # #E11D48 (Modern Ruby Accent)
COLOR_EMERALD = RGBColor(16, 185, 129)     # #10B981 (Subtle Green)
COLOR_TEXT_MAIN = RGBColor(15, 23, 42)     # #0F172A (High-contrast dark text)
COLOR_TEXT_MUTED = RGBColor(71, 85, 105)   # #475569 (Clean slate secondary)
COLOR_BG = RGBColor(248, 250, 252)         # #F8FAFC (Soft executive background)
COLOR_CARD = RGBColor(255, 255, 255)       # Pure white
COLOR_BORDER = RGBColor(226, 232, 240)     # #E2E8F0 (Subtle card border)
COLOR_WHITE = RGBColor(255, 255, 255)

def apply_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_BG

def add_header(slide, title_text, category_tag="BLOODBRIDGE • FRONT-END SYSTEM"):
    # Dark Slate Header Band
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_NAVY
    bar.line.fill.background()

    # Slim crimson accent line beneath header
    accent_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(13.333), Inches(0.06))
    accent_line.fill.solid()
    accent_line.fill.fore_color.rgb = COLOR_ACCENT_RED
    accent_line.line.fill.background()

    # Category tag / breadcrumb
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.0), Inches(0.35))
    p_tag = tag_box.text_frame.paragraphs[0]
    p_tag.text = category_tag.upper()
    p_tag.font.name = "Helvetica"
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = RGBColor(244, 63, 94)  # Soft rose accent

    # Main slide title (large, clear, readable)
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(11.5), Inches(0.7))
    p = title_box.text_frame.paragraphs[0]
    p.text = title_text
    p.font.name = "Helvetica"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD, border_color=COLOR_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)
    else:
        card.line.fill.background()
    return card

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: TITLE & TEAM SLIDE
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    apply_background(slide1)

    # Top thin slate accent
    top_strip = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    top_strip.fill.solid()
    top_strip.fill.fore_color.rgb = COLOR_ACCENT_RED
    top_strip.line.fill.background()

    # SRM Logo (top-left)
    if os.path.exists(SRM_LOGO):
        slide1.shapes.add_picture(SRM_LOGO, Inches(0.8), Inches(0.4), width=Inches(3.3))

    # University & Department text (top-right)
    dept_box = slide1.shapes.add_textbox(Inches(4.5), Inches(0.42), Inches(8.0), Inches(1.0))
    dept_tf = dept_box.text_frame
    p_d1 = dept_tf.paragraphs[0]
    p_d1.text = "SRM Institute of Science and Technology • Kattankulathur"
    p_d1.font.bold = True
    p_d1.font.size = Pt(13)
    p_d1.font.color.rgb = COLOR_NAVY
    p_d2 = dept_tf.add_paragraph()
    p_d2.text = "Department of Computational Intelligence  |  21CSC203P Advanced Programming Practice"
    p_d2.font.size = Pt(11)
    p_d2.font.color.rgb = COLOR_TEXT_MUTED

    # Hero Banner Card (Modern Deep Slate Navy)
    hero_card = create_card(slide1, Inches(0.8), Inches(1.65), Inches(11.733), Inches(2.5), COLOR_NAVY, None)
    
    # Left vertical accent
    hero_accent = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.65), Inches(0.2), Inches(2.5))
    hero_accent.fill.solid()
    hero_accent.fill.fore_color.rgb = COLOR_ACCENT_RED
    hero_accent.line.fill.background()

    title_box = slide1.shapes.add_textbox(Inches(1.2), Inches(1.85), Inches(11.0), Inches(2.0))
    t_frame = title_box.text_frame
    p_t1 = t_frame.paragraphs[0]
    p_t1.text = "BloodBridge 🩸"
    p_t1.font.size = Pt(38)
    p_t1.font.bold = True
    p_t1.font.color.rgb = RGBColor(251, 113, 133) # Rose tint for title

    p_t2 = t_frame.add_paragraph()
    p_t2.text = "Blood Donor Registration and Emergency Search System"
    p_t2.font.size = Pt(19)
    p_t2.font.bold = True
    p_t2.font.color.rgb = COLOR_WHITE

    p_t3 = t_frame.add_paragraph()
    p_t3.text = "A zero-dependency standalone front-end built in Python & Tkinter with real-time ABO/Rh compatibility matching."
    p_t3.font.size = Pt(13)
    p_t3.font.color.rgb = RGBColor(203, 213, 225) # Soft slate text

    # 3 Team Member Cards - Clean, NO topics, exact names and reg numbers
    team_members = [
        ("Yana Butala", "RA2511026010676"),
        ("Raghav Maheshwari", "RA2511026010666"),
        ("Shorya Sharma", "RA25111026010618"),
    ]

    for idx, (name, reg_no) in enumerate(team_members):
        card_x = Inches(0.8 + idx * 4.0)
        card_w = Inches(3.733)
        create_card(slide1, card_x, Inches(4.5), card_w, Inches(2.4), COLOR_WHITE)

        # Slate header strip on each card
        c_hdr = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, Inches(4.5), card_w, Inches(0.38))
        c_hdr.fill.solid()
        c_hdr.fill.fore_color.rgb = COLOR_SLATE_HEADER
        c_hdr.line.fill.background()
        tf_ch = c_hdr.text_frame
        p_ch = tf_ch.paragraphs[0]
        p_ch.text = f"STUDENT {idx+1}"
        p_ch.font.size = Pt(10)
        p_ch.font.bold = True
        p_ch.font.color.rgb = RGBColor(226, 232, 240)
        p_ch.alignment = PP_ALIGN.CENTER

        # Name and Register Number box (large, clean, no topics)
        n_box = slide1.shapes.add_textbox(card_x + Inches(0.25), Inches(5.1), card_w - Inches(0.5), Inches(1.5))
        ntf = n_box.text_frame
        ntf.word_wrap = True
        
        p1 = ntf.paragraphs[0]
        p1.text = name
        p1.font.size = Pt(18)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_NAVY
        p1.space_after = Pt(6)

        p2 = ntf.add_paragraph()
        p2.text = reg_no
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_ACCENT_RED
        p2.space_after = Pt(4)

        p3 = ntf.add_paragraph()
        p3.text = "II Year • III Sem • AIML (Sec AK1)"
        p3.font.size = Pt(11)
        p3.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 2: PROBLEM STATEMENT & OBJECTIVES
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    apply_background(slide2)
    add_header(slide2, "1. Problem Statement & Motivation", "OVERVIEW & CLINICAL CONTEXT")

    # Left Card: The Problem
    create_card(slide2, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.5), COLOR_WHITE)
    
    # Left Card Header Banner
    h1 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(0.55))
    h1.fill.solid()
    h1.fill.fore_color.rgb = COLOR_SLATE_HEADER
    h1.line.fill.background()
    p_h1 = h1.text_frame.paragraphs[0]
    p_h1.text = "⚠️ THE CLINICAL CHALLENGE"
    p_h1.font.bold = True
    p_h1.font.size = Pt(13)
    p_h1.font.color.rgb = COLOR_WHITE
    p_h1.alignment = PP_ALIGN.CENTER

    body1 = slide2.shapes.add_textbox(Inches(1.05), Inches(2.2), Inches(5.1), Inches(4.5))
    b1_tf = body1.text_frame
    b1_tf.word_wrap = True

    pts_problem = [
        ("Time-Critical Delays", "In trauma emergencies and critical surgeries, finding safe compatible blood within minutes is life-or-death."),
        ("Flawed Exact-Match Queries", "Traditional systems search only for identical blood groups, missing universal donors (O−) and compatible groups (e.g. B+ can safely receive B−, O+, and O−)."),
        ("Ineligible Donor Waste", "Calling donors currently in the 90-day cooldown period wastes precious emergency time."),
        ("Deployment Overhead", "Heavy hospital management software requires database credentials and servers that fail during local power or network outages."),
    ]

    for i, (title, desc) in enumerate(pts_problem):
        p_t = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p_t.text = f"• {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(12)
        p_t.font.color.rgb = COLOR_NAVY
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_TEXT_MUTED
        p_t.space_after = Pt(10)

    # Right Card: The BloodBridge Solution
    create_card(slide2, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.5), COLOR_WHITE)

    h2 = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(0.55))
    h2.fill.solid()
    h2.fill.fore_color.rgb = COLOR_NAVY
    h2.line.fill.background()
    p_h2 = h2.text_frame.paragraphs[0]
    p_h2.text = "🎯 THE BLOODBRIDGE SOLUTION"
    p_h2.font.bold = True
    p_h2.font.size = Pt(13)
    p_h2.font.color.rgb = COLOR_WHITE
    p_h2.alignment = PP_ALIGN.CENTER

    body2 = slide2.shapes.add_textbox(Inches(7.05), Inches(2.2), Inches(5.2), Inches(4.5))
    b2_tf = body2.text_frame
    b2_tf.word_wrap = True

    pts_solution = [
        ("Instant Zero-Latency Desktop UI", "Built with Python Tkinter for instant boot time and zero server latency on any computer."),
        ("Smart Emergency Matching Engine", "Implements full ABO & Rh compatibility logic to immediately expand candidate donor pools for urgent patients."),
        ("Automated Cooldown Engine", "Enforces the mandatory 90-day cooldown interval and clinical age limits (18–65) automatically."),
        ("Zero External Dependencies", "Runs directly on Python standard library without requiring pip installs or database setup."),
    ]

    for i, (title, desc) in enumerate(pts_solution):
        p_t = b2_tf.paragraphs[0] if i == 0 else b2_tf.add_paragraph()
        p_t.text = f"✓ {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(12)
        p_t.font.color.rgb = COLOR_ACCENT_RED
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_TEXT_MUTED
        p_t.space_after = Pt(10)

    # =========================================================================
    # SLIDE 3: SYSTEM ARCHITECTURE & STRUCTURE
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    apply_background(slide3)
    add_header(slide3, "2. System Architecture & Component Hierarchy", "TECHNICAL ARCHITECTURE")

    # Left: Architecture Diagram Card
    create_card(slide3, Inches(0.8), Inches(1.5), Inches(7.5), Inches(5.5), COLOR_WHITE)
    fig1_path = os.path.join(DOCS_DIR, "fig1_structure.png")
    if os.path.exists(fig1_path):
        slide3.shapes.add_picture(fig1_path, Inches(1.1), Inches(2.1), width=Inches(6.9))

    c_box = slide3.shapes.add_textbox(Inches(1.0), Inches(6.1), Inches(7.1), Inches(0.6))
    p_c = c_box.text_frame.paragraphs[0]
    p_c.text = "Figure 1: Event-driven architecture connecting user inputs, cooldown logic & in-memory store to all 3 tabs."
    p_c.font.size = Pt(11)
    p_c.font.italic = True
    p_c.font.color.rgb = COLOR_TEXT_MUTED
    p_c.alignment = PP_ALIGN.CENTER

    # Right: Modules Explanation Card
    create_card(slide3, Inches(8.6), Inches(1.5), Inches(3.933), Inches(5.5), COLOR_WHITE)
    
    h3 = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.6), Inches(1.5), Inches(3.933), Inches(0.55))
    h3.fill.solid()
    h3.fill.fore_color.rgb = COLOR_SLATE_HEADER
    h3.line.fill.background()
    p_h3 = h3.text_frame.paragraphs[0]
    p_h3.text = "⚙️ CORE MODULES"
    p_h3.font.bold = True
    p_h3.font.size = Pt(13)
    p_h3.font.color.rgb = COLOR_WHITE
    p_h3.alignment = PP_ALIGN.CENTER

    rh_box = slide3.shapes.add_textbox(Inches(8.8), Inches(2.2), Inches(3.533), Inches(4.6))
    rh_tf = rh_box.text_frame
    rh_tf.word_wrap = True

    modules = [
        ("BloodDonorApp (Controller)", "Main Tkinter class managing ttk.Notebook tabs, window layout, and global theme styling."),
        ("InMemoryStore (Model)", "Dictionary & list-based simulated repository pre-loaded with realistic donor records across 10 cities."),
        ("Eligibility Engine (Logic)", "Real-time date arithmetic computing days elapsed since last donation against the 90-day limit."),
        ("Compatibility Engine (Algorithm)", "Hash map matrix resolving universal donor and recipient relationships in O(1) time."),
    ]

    for i, (title, desc) in enumerate(modules):
        p_m = rh_tf.paragraphs[0] if i == 0 else rh_tf.add_paragraph()
        p_m.text = f"• {title}: "
        p_m.font.bold = True
        p_m.font.size = Pt(11.5)
        p_m.font.color.rgb = COLOR_NAVY
        run = p_m.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(10.5)
        run.font.color.rgb = COLOR_TEXT_MUTED
        p_m.space_after = Pt(10)

    # =========================================================================
    # SLIDE 4: FEATURE DEEP-DIVE – SEARCH & EMERGENCY MATCH
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    apply_background(slide4)
    add_header(slide4, "3. Feature Deep-Dive: Emergency Search & Directory", "TAB 1 SPOTLIGHT")

    # Left: Tab 1 Screenshot (Directory View)
    create_card(slide4, Inches(0.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_WHITE)
    pic1_path = os.path.join(DOCS_DIR, "pic1_search.png")
    if os.path.exists(pic1_path):
        slide4.shapes.add_picture(pic1_path, Inches(0.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Right: Tab 1 Emergency Match Screenshot
    create_card(slide4, Inches(6.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_WHITE)
    pic2_path = os.path.join(DOCS_DIR, "pic2_compat.png")
    if os.path.exists(pic2_path):
        slide4.shapes.add_picture(pic2_path, Inches(6.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Bottom Banner: Feature Talking Points (High readability, 11pt)
    create_card(slide4, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.4), COLOR_WHITE)
    b_box4 = slide4.shapes.add_textbox(Inches(1.05), Inches(5.92), Inches(11.2), Inches(1.25))
    b_tf4 = b_box4.text_frame
    b_tf4.word_wrap = True

    p4_1 = b_tf4.paragraphs[0]
    p4_1.text = "🔍 Key Search Capabilities Demonstrated Above:"
    p4_1.font.size = Pt(12)
    p4_1.font.bold = True
    p4_1.font.color.rgb = COLOR_ACCENT_RED
    p4_1.space_after = Pt(3)

    pts_s4 = (
        "• Real-Time KPI Ribbon: Live dashboard cards calculating Total Donors (10), Ready Donors (8), Cooldown (2), and Universal O− (2).\n"
        "• Emergency Recipient Mode: Searching for Patient B+ instantly returns all 4 clinically compatible types (B+, B−, O+, O−).\n"
        "• Interactive ttk.Treeview Table: Color-coded tags (Green = Ready, Amber = Cooldown, Red = Unavailable) with sortable columns.\n"
        "• Instant Actions: Quick donor dispatch panel with one-click 'Record Donation Today' (resets 90d cooldown) and CSV export."
    )
    p4_2 = b_tf4.add_paragraph()
    p4_2.text = pts_s4
    p4_2.font.size = Pt(10.5)
    p4_2.font.color.rgb = COLOR_TEXT_MAIN

    # =========================================================================
    # SLIDE 5: FEATURE DEEP-DIVE – REGISTRATION & MATRIX
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    apply_background(slide5)
    add_header(slide5, "4. Feature Deep-Dive: Registration & Transfusion Matrix", "TABS 2 & 3 SPOTLIGHT")

    # Left: Tab 2 Registration Portal Screenshot
    create_card(slide5, Inches(0.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_WHITE)
    pic3_path = os.path.join(DOCS_DIR, "pic3_register.png")
    if os.path.exists(pic3_path):
        slide5.shapes.add_picture(pic3_path, Inches(0.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Right: Tab 3 Compatibility Matrix Screenshot
    create_card(slide5, Inches(6.8), Inches(1.4), Inches(5.7), Inches(4.3), COLOR_WHITE)
    pic4_path = os.path.join(DOCS_DIR, "pic4_matrix.png")
    if os.path.exists(pic4_path):
        slide5.shapes.add_picture(pic4_path, Inches(6.9), Inches(1.48), width=Inches(5.5), height=Inches(4.14))

    # Bottom Banner: Feature Talking Points (High readability, 11pt)
    create_card(slide5, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.4), COLOR_WHITE)
    b_box5 = slide5.shapes.add_textbox(Inches(1.05), Inches(5.92), Inches(11.2), Inches(1.25))
    b_tf5 = b_box5.text_frame
    b_tf5.word_wrap = True

    p5_1 = b_tf5.paragraphs[0]
    p5_1.text = "📝 Registration & Clinical Transfusion Matrix:"
    p5_1.font.size = Pt(12)
    p5_1.font.bold = True
    p5_1.font.color.rgb = COLOR_ACCENT_RED
    p5_1.space_after = Pt(3)

    pts_s5 = (
        "• Clinical Validation Rules: Enforces medical standards: Age (18–65), Weight (≥ 50 kg), and 10-digit phone verification.\n"
        "• First-Time vs Repeat Donors: Dedicated toggle disables date entry for first-timers; validates past dates for repeat donors.\n"
        "• Inline Edit Mode: Double-clicking any record in the directory loads existing data into the form for instant updates.\n"
        "• Transfusion Reference Matrix: Full ABO/Rh red blood cell rules + live donor pool supply counters across all 8 blood groups."
    )
    p5_2 = b_tf5.add_paragraph()
    p5_2.text = pts_s5
    p5_2.font.size = Pt(10.5)
    p5_2.font.color.rgb = COLOR_TEXT_MAIN

    # =========================================================================
    # SLIDE 6: TECHNICAL IMPLEMENTATION & CONCLUSION
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    apply_background(slide6)
    add_header(slide6, "5. Technical Implementation & Conclusion", "CONCLUSION & DEMO")

    # 3 Summary Cards
    card_w = Inches(3.733)

    # Card 1: Technical Stack
    create_card(slide6, Inches(0.8), Inches(1.5), card_w, Inches(4.3), COLOR_WHITE)
    c1_hdr = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), card_w, Inches(0.48))
    c1_hdr.fill.solid()
    c1_hdr.fill.fore_color.rgb = COLOR_NAVY
    c1_hdr.line.fill.background()
    p_c1h = c1_hdr.text_frame.paragraphs[0]
    p_c1h.text = "💻 TECHNICAL HIGHLIGHTS"
    p_c1h.font.bold = True
    p_c1h.font.size = Pt(11)
    p_c1h.font.color.rgb = COLOR_WHITE
    p_c1h.alignment = PP_ALIGN.CENTER

    b1_box = slide6.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(3.333), Inches(3.5))
    b1_tf = b1_box.text_frame
    b1_tf.word_wrap = True
    t_tech = [
        ("Python Standard Library", "Built using only tkinter, ttk, datetime, csv, and re. Zero external pip packages needed."),
        ("Single-File Standalone", "Entire architecture contained in a clean 650-line bloodbridge.py module."),
        ("Lightweight & Fast", "Instant boot time (<0.2s), minimal RAM footprint (<25 MB), works across Windows, macOS, Linux."),
    ]
    for i, (t, d) in enumerate(t_tech):
        p = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_NAVY
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT_MUTED
        p.space_after = Pt(10)

    # Card 2: Course Outcomes
    create_card(slide6, Inches(4.8), Inches(1.5), card_w, Inches(4.3), COLOR_WHITE)
    c2_hdr = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.8), Inches(1.5), card_w, Inches(0.48))
    c2_hdr.fill.solid()
    c2_hdr.fill.fore_color.rgb = COLOR_SLATE_HEADER
    c2_hdr.line.fill.background()
    p_c2h = c2_hdr.text_frame.paragraphs[0]
    p_c2h.text = "🏆 COURSE OUTCOMES"
    p_c2h.font.bold = True
    p_c2h.font.size = Pt(11)
    p_c2h.font.color.rgb = COLOR_WHITE
    p_c2h.alignment = PP_ALIGN.CENTER

    b2_box = slide6.shapes.add_textbox(Inches(5.0), Inches(2.1), Inches(3.333), Inches(3.5))
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
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_NAVY
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT_MUTED
        p.space_after = Pt(10)

    # Card 3: Future Roadmap
    create_card(slide6, Inches(8.8), Inches(1.5), card_w, Inches(4.3), COLOR_WHITE)
    c3_hdr = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.8), Inches(1.5), card_w, Inches(0.48))
    c3_hdr.fill.solid()
    c3_hdr.fill.fore_color.rgb = COLOR_ACCENT_RED
    c3_hdr.line.fill.background()
    p_c3h = c3_hdr.text_frame.paragraphs[0]
    p_c3h.text = "🚀 FUTURE ROADMAP"
    p_c3h.font.bold = True
    p_c3h.font.size = Pt(11)
    p_c3h.font.color.rgb = COLOR_WHITE
    p_c3h.alignment = PP_ALIGN.CENTER

    b3_box = slide6.shapes.add_textbox(Inches(9.0), Inches(2.1), Inches(3.333), Inches(3.5))
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
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_NAVY
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_TEXT_MUTED
        p.space_after = Pt(10)

    # Bottom Banner: GitHub & Thank You
    create_card(slide6, Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.9), COLOR_WHITE)
    bot_box = slide6.shapes.add_textbox(Inches(1.0), Inches(6.1), Inches(11.333), Inches(0.7))
    bot_tf = bot_box.text_frame
    p_b1 = bot_tf.paragraphs[0]
    p_b1.text = "🔗 GitHub Repository: https://github.com/raghavmaheshwari-web/Blood-Bridge"
    p_b1.font.size = Pt(12)
    p_b1.font.bold = True
    p_b1.font.color.rgb = COLOR_ACCENT_RED
    p_b2 = bot_tf.add_paragraph()
    p_b2.text = "Thank you! Open for Questions & Live Demonstration."
    p_b2.font.size = Pt(11)
    p_b2.font.color.rgb = COLOR_NAVY

    prs.save(OUTPUT_PPTX)
    print("Presentation saved:", OUTPUT_PPTX)

if __name__ == "__main__":
    build_presentation()
