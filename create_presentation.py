#!/usr/bin/env python3
"""
Generate BloodBridge Presentation (PPTX) with 6 executive widescreen slides for 3 presenters.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = "/Users/raghavmaheshwari/Desktop/Blood Donor System"
DOCS_DIR = os.path.join(BASE_DIR, "docs")
OUTPUT_PPTX = os.path.join(BASE_DIR, "BloodBridge_Presentation.pptx")
SRM_LOGO = os.path.join(DOCS_DIR, "srm_logo.png")

# Colors
COLOR_CRIMSON = RGBColor(198, 40, 40)      # #C62828
COLOR_DARK_CRIMSON = RGBColor(142, 0, 0)   # #8E0000
COLOR_DARK = RGBColor(33, 33, 33)          # #212121
COLOR_GREY = RGBColor(90, 100, 110)        # #5A646E
COLOR_LIGHT_BG = RGBColor(245, 247, 250)   # #F5F7FA
COLOR_CARD_BG = RGBColor(255, 255, 255)    # #FFFFFF
COLOR_GREEN = RGBColor(46, 125, 50)        # #2E7D32
COLOR_SRM_BLUE = RGBColor(26, 76, 139)     # #1A4C8B
COLOR_WHITE = RGBColor(255, 255, 255)

def apply_background(slide, color=COLOR_LIGHT_BG):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, presenter_text=""):
    # Crimson top bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_CRIMSON
    bar.line.fill.background()

    # Title text
    tx_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.12), Inches(8.5), Inches(0.85))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Helvetica"
    p.font.size = Pt(23)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    # Presenter badge on top-right
    if presenter_text:
        tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.2), Inches(0.25), Inches(3.5), Inches(0.6))
        tag_box.fill.solid()
        tag_box.fill.fore_color.rgb = COLOR_DARK_CRIMSON
        tag_box.line.color.rgb = RGBColor(255, 180, 180)
        tag_box.line.width = Pt(1)
        tf_tag = tag_box.text_frame
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = presenter_text
        p_tag.font.name = "Helvetica"
        p_tag.font.size = Pt(11.5)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_WHITE
        p_tag.alignment = PP_ALIGN.CENTER

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=RGBColor(218, 224, 233)):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
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
    apply_background(slide1, COLOR_LIGHT_BG)

    top_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_CRIMSON
    top_bar.line.fill.background()

    # SRM Logo
    if os.path.exists(SRM_LOGO):
        slide1.shapes.add_picture(SRM_LOGO, Inches(0.8), Inches(0.4), width=Inches(3.2))

    dept_box = slide1.shapes.add_textbox(Inches(4.5), Inches(0.45), Inches(8.0), Inches(1.0))
    dept_tf = dept_box.text_frame
    p_d1 = dept_tf.paragraphs[0]
    p_d1.text = "SRM Institute of Science and Technology • Kattankulathur"
    p_d1.font.bold = True
    p_d1.font.size = Pt(13)
    p_d1.font.color.rgb = COLOR_SRM_BLUE
    p_d2 = dept_tf.add_paragraph()
    p_d2.text = "Department of Computational Intelligence • 21CSC203P Advanced Programming Practice"
    p_d2.font.size = Pt(11)
    p_d2.font.color.rgb = COLOR_GREY

    # Project Title Hero Box
    create_card(slide1, Inches(0.8), Inches(1.75), Inches(11.733), Inches(2.3), COLOR_WHITE, COLOR_CRIMSON)

    left_stripe = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.75), Inches(0.2), Inches(2.3))
    left_stripe.fill.solid()
    left_stripe.fill.fore_color.rgb = COLOR_CRIMSON
    left_stripe.line.fill.background()

    title_box = slide1.shapes.add_textbox(Inches(1.2), Inches(1.95), Inches(11.0), Inches(1.9))
    t_frame = title_box.text_frame
    p_t1 = t_frame.paragraphs[0]
    p_t1.text = "BloodBridge 🩸"
    p_t1.font.size = Pt(38)
    p_t1.font.bold = True
    p_t1.font.color.rgb = COLOR_CRIMSON

    p_t2 = t_frame.add_paragraph()
    p_t2.text = "Blood Donor Registration and Emergency Search System"
    p_t2.font.size = Pt(18)
    p_t2.font.bold = True
    p_t2.font.color.rgb = COLOR_DARK

    p_t3 = t_frame.add_paragraph()
    p_t3.text = "A zero-dependency standalone front-end built with Python & Tkinter with real-time ABO/Rh compatibility matching."
    p_t3.font.size = Pt(12)
    p_t3.font.color.rgb = COLOR_GREY

    # 3 Presenters Cards
    presenters = [
        ("Presenter 1", "Raghav Maheshwari", "RA2511026010666", "UI Framework, Architecture & In-Memory Store"),
        ("Presenter 2", "Shorya Sharma", "RA2511026010618", "Emergency Compatibility Engine & Search Directory"),
        ("Presenter 3", "Yana Butala", "RA2511026010676", "Donor Registration Validation & Transfusion Matrix"),
    ]

    for idx, (role, name, reg_no, domain) in enumerate(presenters):
        card_x = Inches(0.8 + idx * 4.0)
        card_w = Inches(3.733)
        create_card(slide1, card_x, Inches(4.35), card_w, Inches(2.6), COLOR_WHITE)

        p_hdr = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, Inches(4.35), card_w, Inches(0.42))
        p_hdr.fill.solid()
        p_hdr.fill.fore_color.rgb = COLOR_DARK_CRIMSON if idx == 0 else (COLOR_CRIMSON if idx == 1 else COLOR_SRM_BLUE)
        p_hdr.line.fill.background()
        tf_ph = p_hdr.text_frame
        p_ph = tf_ph.paragraphs[0]
        p_ph.text = role.upper()
        p_ph.font.size = Pt(10.5)
        p_ph.font.bold = True
        p_ph.font.color.rgb = COLOR_WHITE
        p_ph.alignment = PP_ALIGN.CENTER

        p_box = slide1.shapes.add_textbox(card_x + Inches(0.2), Inches(4.9), card_w - Inches(0.4), Inches(1.9))
        ptf = p_box.text_frame
        p1 = ptf.paragraphs[0]
        p1.text = name
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_DARK

        p2 = ptf.add_paragraph()
        p2.text = f"Reg No: {reg_no}"
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_CRIMSON

        p3 = ptf.add_paragraph()
        p3.text = f"Specialization: AIML (Sec AK1)"
        p3.font.size = Pt(10)
        p3.font.color.rgb = COLOR_GREY

        p4 = ptf.add_paragraph()
        p4.text = f"Topic: {domain}"
        p4.font.size = Pt(9.5)
        p4.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 2: PROBLEM STATEMENT & MOTIVATION (Raghav)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    apply_background(slide2)
    add_header(slide2, "1. Problem Statement & Motivation", "Presenter: Raghav Maheshwari")

    create_card(slide2, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.6), COLOR_WHITE)
    lbl1 = slide2.shapes.add_textbox(Inches(1.1), Inches(1.6), Inches(5.0), Inches(0.6))
    lbl1_tf = lbl1.text_frame
    p_l1 = lbl1_tf.paragraphs[0]
    p_l1.text = "⚠️ The Clinical Challenge"
    p_l1.font.size = Pt(18)
    p_l1.font.bold = True
    p_l1.font.color.rgb = COLOR_CRIMSON

    body1 = slide2.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.4))
    b1_tf = body1.text_frame
    b1_tf.word_wrap = True

    pts_problem = [
        ("Time-Critical Delays", "In emergency trauma and surgeries, finding compatible blood donors within minutes is often life-or-death."),
        ("Flawed Exact-Match Systems", "Most existing registries only filter for the exact identical blood group, completely missing universal donors (O−) and compatible groups (e.g. B+ can receive B−, O+, O−)."),
        ("Donor Ineligibility Waste", "Calling donors who donated blood recently (<90 days cooldown) wastes critical response time."),
        ("Complex Setup Overhead", "Most hospital systems require complex databases, servers, and network dependencies that fail during local clinic outages."),
    ]

    for i, (title, desc) in enumerate(pts_problem):
        p_t = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p_t.text = f"• {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(11)
        p_t.font.color.rgb = COLOR_DARK
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(10.5)
        run.font.color.rgb = COLOR_GREY
        p_t.space_after = Pt(8)

    create_card(slide2, Inches(6.8), Inches(1.4), Inches(5.7), Inches(5.6), COLOR_WHITE)
    lbl2 = slide2.shapes.add_textbox(Inches(7.1), Inches(1.6), Inches(5.1), Inches(0.6))
    lbl2_tf = lbl2.text_frame
    p_l2 = lbl2_tf.paragraphs[0]
    p_l2.text = "🎯 BloodBridge Solution & Objectives"
    p_l2.font.size = Pt(18)
    p_l2.font.bold = True
    p_l2.font.color.rgb = COLOR_GREEN

    body2 = slide2.shapes.add_textbox(Inches(7.1), Inches(2.3), Inches(5.1), Inches(4.4))
    b2_tf = body2.text_frame
    b2_tf.word_wrap = True

    pts_solution = [
        ("Instant Zero-Latency Desktop UI", "Built with Python Tkinter for rapid desktop execution with zero startup lag on any machine."),
        ("Smart Emergency Matching Engine", "Implements full ABO & Rh compatibility rules, instantly expanding candidate donors for urgent recipients."),
        ("Automated Cooldown Engine", "Enforces clinical 90-day cooldown interval and age bounds (18–65) dynamically based on donation dates."),
        ("100% Zero-Dependency Standalone", "Runs out-of-the-box on vanilla Python standard library without requiring pip installs or database servers."),
    ]

    for i, (title, desc) in enumerate(pts_solution):
        p_t = b2_tf.paragraphs[0] if i == 0 else b2_tf.add_paragraph()
        p_t.text = f"✓ {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(11)
        p_t.font.color.rgb = COLOR_GREEN
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(10.5)
        run.font.color.rgb = COLOR_GREY
        p_t.space_after = Pt(8)

    # =========================================================================
    # SLIDE 3: SYSTEM ARCHITECTURE & STRUCTURE (Raghav & Shorya)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    apply_background(slide3)
    add_header(slide3, "2. System Architecture & Component Hierarchy", "Presenters: Raghav & Shorya")

    # Left: Architecture Diagram Card
    create_card(slide3, Inches(0.8), Inches(1.4), Inches(7.5), Inches(5.6), COLOR_WHITE)
    fig1_path = os.path.join(DOCS_DIR, "fig1_structure.png")
    if os.path.exists(fig1_path):
        # Display cleanly centered inside card with proper aspect ratio
        slide3.shapes.add_picture(fig1_path, Inches(1.1), Inches(2.0), width=Inches(6.9))

    c_box = slide3.shapes.add_textbox(Inches(1.0), Inches(5.9), Inches(7.1), Inches(0.7))
    c_tf = c_box.text_frame
    p_c = c_tf.paragraphs[0]
    p_c.text = "Figure 1: Event-driven architecture connecting user actions, cooldown engine & in-memory store to all 3 tabs."
    p_c.font.size = Pt(10.5)
    p_c.font.italic = True
    p_c.font.color.rgb = COLOR_GREY
    p_c.alignment = PP_ALIGN.CENTER

    # Right: Modules Explanation Card
    create_card(slide3, Inches(8.6), Inches(1.4), Inches(3.933), Inches(5.6), COLOR_WHITE)
    rh_box = slide3.shapes.add_textbox(Inches(8.8), Inches(1.6), Inches(3.533), Inches(5.2))
    rh_tf = rh_box.text_frame
    rh_tf.word_wrap = True

    p_rh1 = rh_tf.paragraphs[0]
    p_rh1.text = "⚙️ Core Modules"
    p_rh1.font.size = Pt(16)
    p_rh1.font.bold = True
    p_rh1.font.color.rgb = COLOR_SRM_BLUE
    p_rh1.space_after = Pt(10)

    modules = [
        ("BloodDonorApp (Controller)", "Main Tkinter class managing ttk.Notebook tabs, window layout, and global theme styling."),
        ("InMemoryStore (Model)", "Dictionary & list-based simulated repository with preloaded realistic donor records across 10 cities."),
        ("Eligibility Engine (Logic)", "Real-time date arithmetic computing days elapsed since last donation vs 90-day cooldown limit."),
        ("Compatibility Engine (Algorithm)", "Hash map matrix resolving universal donor and recipient relationships in O(1) time."),
    ]

    for title, desc in modules:
        p_m = rh_tf.add_paragraph()
        p_m.text = f"• {title}: "
        p_m.font.bold = True
        p_m.font.size = Pt(10.5)
        p_m.font.color.rgb = COLOR_DARK
        run = p_m.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(9.5)
        run.font.color.rgb = COLOR_GREY
        p_m.space_after = Pt(8)

    # =========================================================================
    # SLIDE 4: FEATURE DEEP-DIVE – SEARCH & EMERGENCY MATCH (Shorya)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    apply_background(slide4)
    add_header(slide4, "3. Feature Deep-Dive: Emergency Search & Directory", "Presenter: Shorya Sharma")

    # Left: Tab 1 Screenshot (Directory View)
    create_card(slide4, Inches(0.8), Inches(1.35), Inches(5.7), Inches(4.35), COLOR_WHITE)
    pic1_path = os.path.join(DOCS_DIR, "pic1_search.png")
    if os.path.exists(pic1_path):
        slide4.shapes.add_picture(pic1_path, Inches(0.9), Inches(1.45), width=Inches(5.5), height=Inches(4.21))

    # Right: Tab 1 Emergency Match Screenshot
    create_card(slide4, Inches(6.8), Inches(1.35), Inches(5.7), Inches(4.35), COLOR_WHITE)
    pic2_path = os.path.join(DOCS_DIR, "pic2_compat.png")
    if os.path.exists(pic2_path):
        slide4.shapes.add_picture(pic2_path, Inches(6.9), Inches(1.45), width=Inches(5.5), height=Inches(4.21))

    # Bottom Banner: Feature Talking Points
    bot_card_s4 = create_card(slide4, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.35), COLOR_WHITE)
    b_box4 = slide4.shapes.add_textbox(Inches(1.0), Inches(5.92), Inches(11.333), Inches(1.2))
    b_tf4 = b_box4.text_frame
    b_tf4.word_wrap = True

    p4_1 = b_tf4.paragraphs[0]
    p4_1.text = "🔍 Key Search Features Demonstrated Above:"
    p4_1.font.size = Pt(11)
    p4_1.font.bold = True
    p4_1.font.color.rgb = COLOR_CRIMSON

    pts_s4 = (
        "• Real-Time KPI Ribbon: Live dashboard cards for Total Donors (10), Ready Donors (8), Cooldown (2), and Universal O− (2).\n"
        "• Emergency Recipient Mode: Searching for Patient B+ instantly returns all 4 clinically compatible types (B+, B−, O+, O−).\n"
        "• Interactive ttk.Treeview Table: Color tags (Green = Ready, Amber = Cooldown, Red = Unavailable) with sortable columns.\n"
        "• Instant Actions: Quick donor dispatch panel with one-click 'Record Donation Today' (resets 90d cooldown) and CSV export."
    )
    p4_2 = b_tf4.add_paragraph()
    p4_2.text = pts_s4
    p4_2.font.size = Pt(9.5)
    p4_2.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 5: FEATURE DEEP-DIVE – REGISTRATION & MATRIX (Yana)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    apply_background(slide5)
    add_header(slide5, "4. Feature Deep-Dive: Registration & Transfusion Matrix", "Presenter: Yana Butala")

    # Left: Tab 2 Registration Portal Screenshot
    create_card(slide5, Inches(0.8), Inches(1.35), Inches(5.7), Inches(4.35), COLOR_WHITE)
    pic3_path = os.path.join(DOCS_DIR, "pic3_register.png")
    if os.path.exists(pic3_path):
        slide5.shapes.add_picture(pic3_path, Inches(0.9), Inches(1.45), width=Inches(5.5), height=Inches(4.21))

    # Right: Tab 3 Compatibility Matrix Screenshot
    create_card(slide5, Inches(6.8), Inches(1.35), Inches(5.7), Inches(4.35), COLOR_WHITE)
    pic4_path = os.path.join(DOCS_DIR, "pic4_matrix.png")
    if os.path.exists(pic4_path):
        slide5.shapes.add_picture(pic4_path, Inches(6.9), Inches(1.45), width=Inches(5.5), height=Inches(4.21))

    # Bottom Banner: Feature Talking Points
    bot_card_s5 = create_card(slide5, Inches(0.8), Inches(5.85), Inches(11.733), Inches(1.35), COLOR_WHITE)
    b_box5 = slide5.shapes.add_textbox(Inches(1.0), Inches(5.92), Inches(11.333), Inches(1.2))
    b_tf5 = b_box5.text_frame
    b_tf5.word_wrap = True

    p5_1 = b_tf5.paragraphs[0]
    p5_1.text = "📝 Registration & Clinical Matrix Capabilities:"
    p5_1.font.size = Pt(11)
    p5_1.font.bold = True
    p5_1.font.color.rgb = COLOR_CRIMSON

    pts_s5 = (
        "• Clinical Validation Rules: Enforces medical standards: Age (18–65), Weight (≥ 50 kg), and 10-digit phone verification.\n"
        "• First-Time vs Repeat Donors: Dedicated toggle disables date entry for first-timers; validates past dates for repeat donors.\n"
        "• Inline Edit Mode: Double-clicking any record in the directory loads existing data into the form for instant updates.\n"
        "• Transfusion Reference Matrix: Full ABO/Rh red blood cell rules + live donor pool supply counters across all 8 blood groups."
    )
    p5_2 = b_tf5.add_paragraph()
    p5_2.text = pts_s5
    p5_2.font.size = Pt(9.5)
    p5_2.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 6: TECHNICAL IMPLEMENTATION & CONCLUSION (All)
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    apply_background(slide6)
    add_header(slide6, "5. Technical Implementation & Conclusion", "Summary & Conclusion")

    # 3 Summary Cards
    create_card(slide6, Inches(0.8), Inches(1.5), Inches(3.733), Inches(4.3), COLOR_WHITE)
    c1_bar = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(3.733), Inches(0.4))
    c1_bar.fill.solid()
    c1_bar.fill.fore_color.rgb = COLOR_CRIMSON
    c1_bar.line.fill.background()
    p_c1h = c1_bar.text_frame.paragraphs[0]
    p_c1h.text = "💻 TECHNICAL HIGHLIGHTS"
    p_c1h.font.bold = True
    p_c1h.font.size = Pt(11)
    p_c1h.font.color.rgb = COLOR_WHITE
    p_c1h.alignment = PP_ALIGN.CENTER

    b1_box = slide6.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(3.333), Inches(3.6))
    b1_tf = b1_box.text_frame
    b1_tf.word_wrap = True
    t_tech = [
        ("Python Standard Library", "Built using only tkinter, ttk, datetime, csv, and re. Zero external pip packages needed."),
        ("Single-File Standalone", "Entire architecture contained in clean 650-line bloodbridge.py module."),
        ("Lightweight & Fast", "Instant boot time (<0.2s), minimal RAM footprint (<25 MB), works across Windows, macOS, Linux."),
    ]
    for i, (t, d) in enumerate(t_tech):
        p = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_DARK
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(9.5)
        run.font.color.rgb = COLOR_GREY
        p.space_after = Pt(8)

    # Card 2: Course Outcomes
    create_card(slide6, Inches(4.8), Inches(1.5), Inches(3.733), Inches(4.3), COLOR_WHITE)
    c2_bar = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.8), Inches(1.5), Inches(3.733), Inches(0.4))
    c2_bar.fill.solid()
    c2_bar.fill.fore_color.rgb = COLOR_SRM_BLUE
    c2_bar.line.fill.background()
    p_c2h = c2_bar.text_frame.paragraphs[0]
    p_c2h.text = "🏆 COURSE OUTCOMES"
    p_c2h.font.bold = True
    p_c2h.font.size = Pt(11)
    p_c2h.font.color.rgb = COLOR_WHITE
    p_c2h.alignment = PP_ALIGN.CENTER

    b2_box = slide6.shapes.add_textbox(Inches(5.0), Inches(2.0), Inches(3.333), Inches(3.6))
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
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_DARK
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(9.5)
        run.font.color.rgb = COLOR_GREY
        p.space_after = Pt(8)

    # Card 3: Future Roadmap
    create_card(slide6, Inches(8.8), Inches(1.5), Inches(3.733), Inches(4.3), COLOR_WHITE)
    c3_bar = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.8), Inches(1.5), Inches(3.733), Inches(0.4))
    c3_bar.fill.solid()
    c3_bar.fill.fore_color.rgb = COLOR_GREEN
    c3_bar.line.fill.background()
    p_c3h = c3_bar.text_frame.paragraphs[0]
    p_c3h.text = "🚀 FUTURE ROADMAP"
    p_c3h.font.bold = True
    p_c3h.font.size = Pt(11)
    p_c3h.font.color.rgb = COLOR_WHITE
    p_c3h.alignment = PP_ALIGN.CENTER

    b3_box = slide6.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(3.333), Inches(3.6))
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
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_DARK
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(9.5)
        run.font.color.rgb = COLOR_GREY
        p.space_after = Pt(8)

    # Bottom Banner
    bot_card = create_card(slide6, Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.9), COLOR_WHITE)
    bot_box = slide6.shapes.add_textbox(Inches(1.0), Inches(6.1), Inches(11.333), Inches(0.7))
    bot_tf = bot_box.text_frame
    p_b1 = bot_tf.paragraphs[0]
    p_b1.text = "🔗 GitHub Repository: https://github.com/raghavmaheshwari-web/Blood-Bridge"
    p_b1.font.size = Pt(12)
    p_b1.font.bold = True
    p_b1.font.color.rgb = COLOR_SRM_BLUE
    p_b2 = bot_tf.add_paragraph()
    p_b2.text = "Thank you! Open for Questions & Live Demonstration."
    p_b2.font.size = Pt(11)
    p_b2.font.color.rgb = COLOR_DARK

    prs.save(OUTPUT_PPTX)
    print("Presentation saved:", OUTPUT_PPTX)

if __name__ == "__main__":
    build_presentation()
