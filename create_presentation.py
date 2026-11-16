#!/usr/bin/env python3
"""
Generate BloodBridge Presentation (PPTX) with an executive Dark Mode Theme.
- Tightly co-related with the BloodBridge Tkinter desktop interface
- Prominent UI screenshots on Slides 2, 4, 5, and 6
- Architecture mapped directly to GUI components on Slide 3
- Dark theme background (#0B0F19) with BloodBridge Crimson & Emerald accents
- High contrast light text (#FFFFFF / #F1F5F9 / #CBD5E1)
- Large readable text sizes (12.5-16pt body, 18-22pt headers, 28-36pt slide titles)
- Clean 3-member roster on Slide 1 (no "Student 1,2,3" tags, no individual topics)
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
COLOR_BG_DARK = RGBColor(11, 15, 25)          # #0B0F19 (Deep Obsidian Navy background)
COLOR_HEADER_BAR = RGBColor(15, 23, 42)       # #0F172A (Deep Slate Header)
COLOR_CARD_DARK = RGBColor(22, 30, 48)        # #161E30 (Card Slate)
COLOR_CARD_INNER = RGBColor(30, 41, 59)       # #1E293B (Inner Slate)
COLOR_BORDER = RGBColor(51, 65, 85)           # #334155 (Subtle Slate Border)

# BloodBridge Identity Accents (matching Tkinter interface)
COLOR_CRIMSON = RGBColor(198, 40, 40)         # #C62828 (BloodBridge Primary Red)
COLOR_ROSE_ACCENT = RGBColor(244, 63, 94)     # #F43F5E (Vibrant Rose)
COLOR_EMERALD = RGBColor(52, 211, 153)        # #34D399 (Eligible / Ready Green)
COLOR_BLUE_ACCENT = RGBColor(96, 165, 250)    # #60A5FA (Total Donors Sky Blue)
COLOR_AMBER = RGBColor(251, 191, 36)          # #FBBF24 (Cooldown Amber Gold)

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
    cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(11.733), Inches(0.35))
    cat_tf = cat_box.text_frame
    p_cat = cat_tf.paragraphs[0]
    p_cat.text = category_tag.upper()
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = COLOR_ROSE_ACCENT

    t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.75))
    t_tf = t_box.text_frame
    p_title = t_tf.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(28)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_WHITE

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.35), Inches(13.333), Inches(0.04))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_ROSE_ACCENT
    line.line.fill.background()

def create_card(slide, left, top, width, height, bg_color=COLOR_CARD_DARK, border_color=COLOR_BORDER):
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
    # SLIDE 1: TITLE & TEAM SLIDE (DARK MODE)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    apply_background(slide1)

    top_strip = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    top_strip.fill.solid()
    top_strip.fill.fore_color.rgb = COLOR_ROSE_ACCENT
    top_strip.line.fill.background()

    logo_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(3.4), Inches(1.15))
    logo_card.fill.solid()
    logo_card.fill.fore_color.rgb = COLOR_WHITE
    logo_card.line.color.rgb = COLOR_ROSE_ACCENT
    logo_card.line.width = Pt(1.5)

    if os.path.exists(SRM_LOGO):
        slide1.shapes.add_picture(SRM_LOGO, Inches(0.9), Inches(0.48), width=Inches(3.2))

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

    create_card(slide1, Inches(0.8), Inches(1.75), Inches(11.733), Inches(2.45), COLOR_CARD_DARK, COLOR_BORDER)

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

    team_members = [
        ("Yana Butala", "RA2511026010676"),
        ("Raghav Maheshwari", "RA2511026010666"),
        ("Shorya Sharma", "RA25111026010618"),
    ]

    for idx, (name, reg_no) in enumerate(team_members):
        card_x = Inches(0.8 + idx * 4.0)
        card_w = Inches(3.733)
        
        create_card(slide1, card_x, Inches(4.45), card_w, Inches(2.5), COLOR_CARD_DARK, COLOR_BORDER)

        accent_top = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, Inches(4.45), card_w, Inches(0.08))
        accent_top.fill.solid()
        accent_top.fill.fore_color.rgb = COLOR_ROSE_ACCENT if idx == 1 else COLOR_BLUE_ACCENT
        accent_top.line.fill.background()

        n_box = slide1.shapes.add_textbox(card_x + Inches(0.3), Inches(4.8), card_w - Inches(0.6), Inches(1.8))
        ntf = n_box.text_frame
        ntf.word_wrap = True
        
        p1 = ntf.paragraphs[0]
        p1.text = name
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_TEXT_WHITE
        p1.space_after = Pt(8)

        p2 = ntf.add_paragraph()
        p2.text = reg_no
        p2.font.size = Pt(16)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_ROSE_ACCENT
        p2.space_after = Pt(8)

        p3 = ntf.add_paragraph()
        p3.text = "II Year • III Semester • AIML (AK1)"
        p3.font.size = Pt(13)
        p3.font.color.rgb = COLOR_TEXT_SECONDARY

    # =========================================================================
    # SLIDE 2: PROBLEM STATEMENT & SOLUTION WITH INTERFACE PREVIEW
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    apply_background(slide2)
    add_header(slide2, "1. Problem Statement & The BloodBridge Solution", "CLINICAL CONTEXT & LIVE INTERFACE")

    # Left Column: Challenge & Solution Cards
    left_w = Inches(5.6)
    
    # 1. Challenge Card (Top Left)
    create_card(slide2, Inches(0.8), Inches(1.5), left_w, Inches(2.65), COLOR_CARD_DARK)
    h1 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.62), left_w - Inches(0.4), Inches(0.45))
    h1.fill.solid()
    h1.fill.fore_color.rgb = COLOR_CARD_INNER
    h1.line.color.rgb = COLOR_ROSE_ACCENT
    p_h1 = h1.text_frame.paragraphs[0]
    p_h1.text = "⚠️ THE CLINICAL CHALLENGE"
    p_h1.font.bold = True
    p_h1.font.size = Pt(12.5)
    p_h1.font.color.rgb = COLOR_ROSE_ACCENT
    p_h1.alignment = PP_ALIGN.CENTER

    body1 = slide2.shapes.add_textbox(Inches(1.0), Inches(2.15), left_w - Inches(0.4), Inches(1.9))
    b1_tf = body1.text_frame
    b1_tf.word_wrap = True

    pts_prob = [
        ("Time-Critical Emergencies", "Finding compatible blood in trauma situations directly impacts patient survival."),
        ("Flawed Exact-Match Systems", "Portals filter only identical blood groups, missing universal (O−) and compatible types."),
        ("Ineligible Donor Waste", "Calling donors in their 90-day cooldown period wastes precious triage time."),
    ]
    for i, (title, desc) in enumerate(pts_prob):
        p_t = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p_t.text = f"• {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(12)
        p_t.font.color.rgb = COLOR_TEXT_WHITE
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p_t.space_after = Pt(4)

    # 2. Solution Card (Bottom Left)
    create_card(slide2, Inches(0.8), Inches(4.3), left_w, Inches(2.75), COLOR_CARD_DARK)
    h2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.42), left_w - Inches(0.4), Inches(0.45))
    h2.fill.solid()
    h2.fill.fore_color.rgb = COLOR_CARD_INNER
    h2.line.color.rgb = COLOR_EMERALD
    p_h2 = h2.text_frame.paragraphs[0]
    p_h2.text = "🎯 THE BLOODBRIDGE SOLUTION"
    p_h2.font.bold = True
    p_h2.font.size = Pt(12.5)
    p_h2.font.color.rgb = COLOR_EMERALD
    p_h2.alignment = PP_ALIGN.CENTER

    body2 = slide2.shapes.add_textbox(Inches(1.0), Inches(4.95), left_w - Inches(0.4), Inches(2.0))
    b2_tf = body2.text_frame
    b2_tf.word_wrap = True

    pts_sol = [
        ("Zero-Latency Desktop UI", "Python Tkinter GUI boots in <0.2s with zero external dependencies."),
        ("Smart Emergency Matching", "Instantly resolves all clinically safe donor types (ABO & Rh rules)."),
        ("Automated 90-Day Cooldown", "Live date arithmetic automatically tags eligible vs cooldown donors."),
    ]
    for i, (title, desc) in enumerate(pts_sol):
        p_t = b2_tf.paragraphs[0] if i == 0 else b2_tf.add_paragraph()
        p_t.text = f"✓ {title}: "
        p_t.font.bold = True
        p_t.font.size = Pt(12)
        p_t.font.color.rgb = COLOR_EMERALD
        run = p_t.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p_t.space_after = Pt(4)

    # Right Column: Hero Screenshot Card showing the Interface in Action
    right_x = Inches(6.65)
    right_w = Inches(5.88)
    create_card(slide2, right_x, Inches(1.5), right_w, Inches(5.55), COLOR_CARD_DARK)

    # Header Pill on Right Card
    h_preview = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x + Inches(0.2), Inches(1.62), right_w - Inches(0.4), Inches(0.45))
    h_preview.fill.solid()
    h_preview.fill.fore_color.rgb = COLOR_CARD_INNER
    h_preview.line.color.rgb = COLOR_BLUE_ACCENT
    p_hp = h_preview.text_frame.paragraphs[0]
    p_hp.text = "🖥️ FRONT-END INTERFACE: TAB 1 DIRECTORY & SEARCH"
    p_hp.font.bold = True
    p_hp.font.size = Pt(12)
    p_hp.font.color.rgb = COLOR_BLUE_ACCENT
    p_hp.alignment = PP_ALIGN.CENTER

    # Framed Screenshot of pic1_search.png
    pic1_path = os.path.join(DOCS_DIR, "pic1_search.png")
    if os.path.exists(pic1_path):
        slide2.shapes.add_picture(pic1_path, right_x + Inches(0.25), Inches(2.2), width=right_w - Inches(0.5), height=Inches(3.95))

    # Caption / Badges below Screenshot
    lbl_s2 = slide2.shapes.add_textbox(right_x + Inches(0.2), Inches(6.25), right_w - Inches(0.4), Inches(0.7))
    l_tf2 = lbl_s2.text_frame
    p_l2 = l_tf2.paragraphs[0]
    p_l2.text = "⚡ Real-time KPI Ribbon (10 Total, 8 Ready, 2 Cooldown)  •  Live ttk.Treeview Table"
    p_l2.font.size = Pt(11)
    p_l2.font.bold = True
    p_l2.font.color.rgb = COLOR_TEXT_SECONDARY
    p_l2.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 3: SYSTEM ARCHITECTURE & GUI COMPONENT MAPPING
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    apply_background(slide3)
    add_header(slide3, "2. System Architecture & Component Hierarchy", "TECHNICAL ARCHITECTURE & GUI MAPPING")

    # Left: Architecture Flowchart Card
    create_card(slide3, Inches(0.8), Inches(1.5), Inches(7.4), Inches(5.55), COLOR_CARD_DARK)
    
    # White container so flowchart has crisp contrast
    diag_container = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.7), Inches(7.0), Inches(4.2))
    diag_container.fill.solid()
    diag_container.fill.fore_color.rgb = COLOR_WHITE
    diag_container.line.fill.background()

    fig1_path = os.path.join(DOCS_DIR, "fig1_structure.png")
    if os.path.exists(fig1_path):
        slide3.shapes.add_picture(fig1_path, Inches(1.15), Inches(1.85), width=Inches(6.7))

    c_box = slide3.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(7.0), Inches(0.8))
    p_c = c_box.text_frame.paragraphs[0]
    p_c.text = "Figure 1: Event-driven architecture connecting user inputs, cooldown logic & in-memory store to all 3 tabs."
    p_c.font.size = Pt(11.5)
    p_c.font.italic = True
    p_c.font.color.rgb = COLOR_TEXT_SECONDARY
    p_c.alignment = PP_ALIGN.CENTER

    # Right: Core Modules Card Mapped to GUI
    create_card(slide3, Inches(8.45), Inches(1.5), Inches(4.08), Inches(5.55), COLOR_CARD_DARK)
    
    h3 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.65), Inches(1.65), Inches(3.68), Inches(0.5))
    h3.fill.solid()
    h3.fill.fore_color.rgb = COLOR_CARD_INNER
    h3.line.color.rgb = COLOR_BLUE_ACCENT
    p_h3 = h3.text_frame.paragraphs[0]
    p_h3.text = "⚙️ ARCHITECTURE TO GUI MAPPING"
    p_h3.font.bold = True
    p_h3.font.size = Pt(12.5)
    p_h3.font.color.rgb = COLOR_BLUE_ACCENT
    p_h3.alignment = PP_ALIGN.CENTER

    rh_box = slide3.shapes.add_textbox(Inches(8.65), Inches(2.25), Inches(3.68), Inches(4.6))
    rh_tf = rh_box.text_frame
    rh_tf.word_wrap = True

    modules = [
        ("Controller (BloodBridge)", "Main Tkinter class managing ttk.Notebook tabs, window layout, and theme styling."),
        ("Model (InMemoryStore)", "Pre-loaded in-memory dictionary repository with 10 realistic donor records across 10 cities."),
        ("Eligibility Engine", "Real-time date arithmetic computing days elapsed against the mandatory 90-day cooldown limit."),
        ("Compatibility Engine", "ABO/Rh hash map resolving universal donor and recipient relationships in O(1) time."),
    ]

    for i, (title, desc) in enumerate(modules):
        p_m = rh_tf.paragraphs[0] if i == 0 else rh_tf.add_paragraph()
        p_m.text = f"• {title}: "
        p_m.font.bold = True
        p_m.font.size = Pt(12.5)
        p_m.font.color.rgb = COLOR_TEXT_WHITE
        run = p_m.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p_m.space_after = Pt(8)

    # =========================================================================
    # SLIDE 4: FEATURE DEEP-DIVE 1 – EMERGENCY SEARCH & COMPATIBILITY
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    apply_background(slide4)
    add_header(slide4, "3. Feature Deep-Dive: Emergency Search & Compatibility", "TAB 1 SPOTLIGHT • REAL-TIME ABO/Rh MATCHING")

    # Left: Hero High-Res Screenshot of pic2_compat.png (Emergency Match Mode)
    hero_w = Inches(7.0)
    create_card(slide4, Inches(0.8), Inches(1.5), hero_w, Inches(5.55), COLOR_CARD_DARK)

    # Red Emergency Header Badge inside screenshot frame
    h_em = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.62), hero_w - Inches(0.4), Inches(0.45))
    h_em.fill.solid()
    h_em.fill.fore_color.rgb = COLOR_CRIMSON
    h_em.line.fill.background()
    p_hem = h_em.text_frame.paragraphs[0]
    p_hem.text = "🚨 LIVE EMERGENCY MATCH: PATIENT NEEDS BLOOD GROUP B+"
    p_hem.font.bold = True
    p_hem.font.size = Pt(12)
    p_hem.font.color.rgb = COLOR_TEXT_WHITE
    p_hem.alignment = PP_ALIGN.CENTER

    pic2_path = os.path.join(DOCS_DIR, "pic2_compat.png")
    if os.path.exists(pic2_path):
        slide4.shapes.add_picture(pic2_path, Inches(1.0), Inches(2.15), width=hero_w - Inches(0.4), height=Inches(4.15))

    lbl_s4 = slide4.shapes.add_textbox(Inches(1.0), Inches(6.35), hero_w - Inches(0.4), Inches(0.6))
    l_tf4 = lbl_s4.text_frame
    p_l4 = l_tf4.paragraphs[0]
    p_l4.text = "✓ 4 Clinically Safe Matches Found (B+, B−, O+, O−)  •  Emergency Dispatch Card Active"
    p_l4.font.size = Pt(11)
    p_l4.font.bold = True
    p_l4.font.color.rgb = COLOR_EMERALD
    p_l4.alignment = PP_ALIGN.CENTER

    # Right: Feature Callout Cards Directly Explaining the Interface
    right_x4 = Inches(8.05)
    right_w4 = Inches(4.48)
    create_card(slide4, right_x4, Inches(1.5), right_w4, Inches(5.55), COLOR_CARD_DARK)

    h_fc = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x4 + Inches(0.2), Inches(1.62), right_w4 - Inches(0.4), Inches(0.45))
    h_fc.fill.solid()
    h_fc.fill.fore_color.rgb = COLOR_CARD_INNER
    h_fc.line.color.rgb = COLOR_ROSE_ACCENT
    p_hfc = h_fc.text_frame.paragraphs[0]
    p_hfc.text = "🔍 KEY CAPABILITIES DEMONSTRATED"
    p_hfc.font.bold = True
    p_hfc.font.size = Pt(12)
    p_hfc.font.color.rgb = COLOR_ROSE_ACCENT
    p_hfc.alignment = PP_ALIGN.CENTER

    box_s4 = slide4.shapes.add_textbox(right_x4 + Inches(0.2), Inches(2.2), right_w4 - Inches(0.4), Inches(4.7))
    tf_s4 = box_s4.text_frame
    tf_s4.word_wrap = True

    pts_s4 = [
        ("Intelligent ABO/Rh Matching", "When searching for Patient B+, BloodBridge instantly returns all clinically safe types: B+, B−, O+, and O−. Prevents missing universal donors in crises."),
        ("Live 90-Day Cooldown Triage", "Treeview tags automatically color-code status: Green (Ready), Amber (In Cooldown with remaining days), Red (Unavailable). Prevents calling ineligible donors."),
        ("One-Click Emergency Dispatch", "Selecting any candidate displays the bottom dispatch panel with verified contact details, location, and a direct 'Record Donation Today' action button."),
        ("Multi-Parameter Filtering", "Combines recipient mode, city/district filter, and search keywords for precise emergency coordination."),
    ]

    for i, (title, desc) in enumerate(pts_s4):
        p = tf_s4.paragraphs[0] if i == 0 else tf_s4.add_paragraph()
        p.text = f"• {title}: "
        p.font.bold = True
        p.font.size = Pt(12.5)
        p.font.color.rgb = COLOR_TEXT_WHITE
        run = p.add_run()
        run.text = desc
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 5: FEATURE DEEP-DIVE 2 – REGISTRATION & TRANSFUSION MATRIX
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    apply_background(slide5)
    add_header(slide5, "4. Feature Deep-Dive: Registration & Transfusion Matrix", "TABS 2 & 3 SPOTLIGHT • FORM VALIDATION & CLINICAL MATRIX")

    col_w5 = Inches(5.7)
    
    # Left Column: Tab 2 Registration Portal
    create_card(slide5, Inches(0.8), Inches(1.5), col_w5, Inches(5.55), COLOR_CARD_DARK)
    
    h_reg = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.62), col_w5 - Inches(0.4), Inches(0.45))
    h_reg.fill.solid()
    h_reg.fill.fore_color.rgb = COLOR_CARD_INNER
    h_reg.line.color.rgb = COLOR_EMERALD
    p_hr = h_reg.text_frame.paragraphs[0]
    p_hr.text = "📋 TAB 2: VOLUNTARY DONOR REGISTRATION"
    p_hr.font.bold = True
    p_hr.font.size = Pt(12)
    p_hr.font.color.rgb = COLOR_EMERALD
    p_hr.alignment = PP_ALIGN.CENTER

    pic3_path = os.path.join(DOCS_DIR, "pic3_register.png")
    if os.path.exists(pic3_path):
        slide5.shapes.add_picture(pic3_path, Inches(1.0), Inches(2.15), width=col_w5 - Inches(0.4), height=Inches(2.75))

    box_r = slide5.shapes.add_textbox(Inches(1.0), Inches(5.0), col_w5 - Inches(0.4), Inches(1.9))
    tf_r = box_r.text_frame
    tf_r.word_wrap = True

    pts_r = [
        ("Clinical Validation Rules", "Enforces age (18–65), weight (≥50 kg), and 10-digit mobile number format."),
        ("First-Time vs Repeat Donors", "Dedicated toggle disables date field for first-timers; validates past dates for repeat donors."),
        ("Inline Edit Mode", "Double-clicking any directory row pre-populates form for quick updates."),
    ]
    for i, (t, d) in enumerate(pts_r):
        p = tf_r.paragraphs[0] if i == 0 else tf_r.add_paragraph()
        p.text = f"✓ {t}: "
        p.font.bold = True
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_EMERALD
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p.space_after = Pt(3)

    # Right Column: Tab 3 Transfusion Matrix & Pool Stats
    create_card(slide5, Inches(6.833), Inches(1.5), col_w5, Inches(5.55), COLOR_CARD_DARK)

    h_mat = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.033), Inches(1.62), col_w5 - Inches(0.4), Inches(0.45))
    h_mat.fill.solid()
    h_mat.fill.fore_color.rgb = COLOR_CARD_INNER
    h_mat.line.color.rgb = COLOR_BLUE_ACCENT
    p_hm = h_mat.text_frame.paragraphs[0]
    p_hm.text = "🩸 TAB 3: TRANSFUSION MATRIX & POOL STATS"
    p_hm.font.bold = True
    p_hm.font.size = Pt(12)
    p_hm.font.color.rgb = COLOR_BLUE_ACCENT
    p_hm.alignment = PP_ALIGN.CENTER

    pic4_path = os.path.join(DOCS_DIR, "pic4_matrix.png")
    if os.path.exists(pic4_path):
        slide5.shapes.add_picture(pic4_path, Inches(7.033), Inches(2.15), width=col_w5 - Inches(0.4), height=Inches(2.75))

    box_m = slide5.shapes.add_textbox(Inches(7.033), Inches(5.0), col_w5 - Inches(0.4), Inches(1.9))
    tf_m = box_m.text_frame
    tf_m.word_wrap = True

    pts_m = [
        ("Clinical Reference Grid", "Comprehensive reference for red blood cells & whole blood across all 8 blood groups."),
        ("Live Inventory Counters", "Real-time supply tally across each blood group in the active donor repository."),
        ("Reset Demo Data", "One-click in-memory repository reset for reliable, repeatable evaluations."),
    ]
    for i, (t, d) in enumerate(pts_m):
        p = tf_m.paragraphs[0] if i == 0 else tf_m.add_paragraph()
        p.text = f"✓ {t}: "
        p.font.bold = True
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_BLUE_ACCENT
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_TEXT_SECONDARY
        p.space_after = Pt(3)

    # =========================================================================
    # SLIDE 6: TECHNICAL IMPLEMENTATION & LIVE DEMO PREVIEW
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    apply_background(slide6)
    add_header(slide6, "5. Technical Implementation & Live Demonstration", "CONCLUSION & DEMO PREPARATION")

    # Left Column: Tech Stack & Course Outcomes (Width 6.0 in)
    col_w6 = Inches(5.7)

    # Card 1: Tech Specs (Top Left)
    create_card(slide6, Inches(0.8), Inches(1.5), col_w6, Inches(2.15), COLOR_CARD_DARK)
    c1_h = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), col_w6, Inches(0.38))
    c1_h.fill.solid()
    c1_h.fill.fore_color.rgb = COLOR_CARD_INNER
    c1_h.line.color.rgb = COLOR_BLUE_ACCENT
    p_c1 = c1_h.text_frame.paragraphs[0]
    p_c1.text = "💻 TECHNICAL SPECIFICATIONS"
    p_c1.font.bold = True
    p_c1.font.size = Pt(11.5)
    p_c1.font.color.rgb = COLOR_BLUE_ACCENT
    p_c1.alignment = PP_ALIGN.CENTER

    b1_box = slide6.shapes.add_textbox(Inches(0.95), Inches(1.95), col_w6 - Inches(0.3), Inches(1.6))
    b1_tf = b1_box.text_frame
    b1_tf.word_wrap = True
    t_tech = [
        ("Python Standard Library", "Pure tkinter, ttk, datetime, csv, re. Zero pip dependencies."),
        ("Single-File Standalone", "Full MVC architecture contained in 650-line bloodbridge.py."),
        ("Lightweight & Cross-Platform", "Instant boot (<0.2s), minimal RAM (<25 MB), runs everywhere."),
    ]
    for i, (t, d) in enumerate(t_tech):
        p = b1_tf.paragraphs[0] if i == 0 else b1_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_WHITE
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Course Outcomes (Bottom Left)
    create_card(slide6, Inches(0.8), Inches(3.8), col_w6, Inches(2.05), COLOR_CARD_DARK)
    c2_h = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.8), col_w6, Inches(0.38))
    c2_h.fill.solid()
    c2_h.fill.fore_color.rgb = COLOR_CARD_INNER
    c2_h.line.color.rgb = COLOR_EMERALD
    p_c2 = c2_h.text_frame.paragraphs[0]
    p_c2.text = "🏆 COURSE OUTCOMES ACHIEVED"
    p_c2.font.bold = True
    p_c2.font.size = Pt(11.5)
    p_c2.font.color.rgb = COLOR_EMERALD
    p_c2.alignment = PP_ALIGN.CENTER

    b2_box = slide6.shapes.add_textbox(Inches(0.95), Inches(4.25), col_w6 - Inches(0.3), Inches(1.5))
    b2_tf = b2_box.text_frame
    b2_tf.word_wrap = True
    t_out = [
        ("Advanced GUI Programming", "ttk.Notebook, PanedWindow, Treeview, and dynamic event binding."),
        ("Medical Domain Logic", "Translated biological ABO & Rh compatibility rules into O(1) hash maps."),
        ("Software Engineering Lifecycle", "Clean Git history, structured documentation, and 2-page report."),
    ]
    for i, (t, d) in enumerate(t_out):
        p = b2_tf.paragraphs[0] if i == 0 else b2_tf.add_paragraph()
        p.text = f"• {t}: "
        p.font.bold = True
        p.font.size = Pt(11.5)
        p.font.color.rgb = COLOR_TEXT_WHITE
        run = p.add_run()
        run.text = d
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Column: Live Demonstration Preview with Framed Screenshot
    right_w6 = Inches(5.8)
    create_card(slide6, Inches(6.75), Inches(1.5), right_w6, Inches(4.35), COLOR_CARD_DARK)

    c3_h = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.75), Inches(1.5), right_w6, Inches(0.45))
    c3_h.fill.solid()
    c3_h.fill.fore_color.rgb = COLOR_CARD_INNER
    c3_h.line.color.rgb = COLOR_ROSE_ACCENT
    p_c3 = c3_h.text_frame.paragraphs[0]
    p_c3.text = "🚀 READY FOR LIVE DEMONSTRATION"
    p_c3.font.bold = True
    p_c3.font.size = Pt(12)
    p_c3.font.color.rgb = COLOR_ROSE_ACCENT
    p_c3.alignment = PP_ALIGN.CENTER

    # Framed Screenshot
    if os.path.exists(pic1_path):
        slide6.shapes.add_picture(pic1_path, Inches(6.95), Inches(2.05), width=right_w6 - Inches(0.4), height=Inches(2.45))

    b3_box = slide6.shapes.add_textbox(Inches(6.95), Inches(4.55), right_w6 - Inches(0.4), Inches(1.2))
    b3_tf = b3_box.text_frame
    b3_tf.word_wrap = True
    p_demo_steps = b3_tf.paragraphs[0]
    p_demo_steps.text = "Live Demo Walkthrough:\n1. Emergency Search for urgent patient (Patient B+)  •  2. Donor Registration with validation\n3. One-click donation record & 90d cooldown reset  •  4. Dynamic inventory updates in Tab 3"
    p_demo_steps.font.size = Pt(11)
    p_demo_steps.font.color.rgb = COLOR_TEXT_SECONDARY

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
    p_b2.text = "Presented by Yana Butala • Raghav Maheshwari • Shorya Sharma | Thank you!"
    p_b2.font.size = Pt(12)
    p_b2.font.color.rgb = COLOR_TEXT_WHITE

    prs.save(OUTPUT_PPTX)
    print("Presentation saved:", OUTPUT_PPTX)

if __name__ == "__main__":
    build_presentation()
