import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # LIGHT BRAND PALETTE (Crisp, High-End Enterprise Tech)
    COLOR_BG = RGBColor(246, 248, 252)          # Ultra-clean Ice Blue-White (#f6f8fc)
    COLOR_CARD_BG = RGBColor(255, 255, 255)     # Crisp Pure White Card (#ffffff)
    COLOR_CARD_BORDER = RGBColor(222, 230, 242) # Subtle Silver-Slate Border (#dee6f2)
    COLOR_CARD_HOVER = RGBColor(238, 244, 255)  # Very Soft Tinted Accent Card (#eef4ff)
    
    # BRAND COLORS FROM MOBRILZ LOGO
    COLOR_BRAND_NAVY = RGBColor(9, 20, 63)      # Deep Navy from logo (#09143f)
    COLOR_BRAND_PRIMARY = RGBColor(37, 99, 235) # Vibrant Royal Blue (#2563eb / #3779cb)
    COLOR_BRAND_ACCENT = RGBColor(2, 132, 199)  # Tech Cyan Blue (#0284c7)
    
    # TEXT COLORS
    COLOR_TEXT_PRIMARY = RGBColor(15, 23, 42)   # Deep Obsidian Slate (#0f172a)
    COLOR_TEXT_MUTED = RGBColor(71, 85, 105)    # Refined Slate Gray (#475569)
    COLOR_TEXT_LIGHT = RGBColor(100, 116, 139)  # Light Slate (#64748b)

    LOGO_PATH = r'C:\Users\Manoj\.gemini\antigravity\scratch\mobrilz_transparent_logo.png'
    MOCKUP_DIR = r'C:\Users\Manoj\.gemini\antigravity\scratch\processed_assets'

    def get_mockup(filename):
        p = os.path.join(MOCKUP_DIR, filename)
        return p if os.path.exists(p) else None

    BG_WAVES_SOFT = [
        os.path.join(r'C:\Users\Manoj\.gemini\antigravity\scratch', "bg_wave_1_soft.jpg"),
        os.path.join(r'C:\Users\Manoj\.gemini\antigravity\scratch', "bg_wave_2_soft.jpg"),
        os.path.join(r'C:\Users\Manoj\.gemini\antigravity\scratch', "bg_wave_3_soft.jpg"),
    ]
    BG_COVER = os.path.join(r'C:\Users\Manoj\.gemini\antigravity\scratch', "bg_wave_1.jpg")
    bg_counter = 0

    def set_bg(slide):
        nonlocal bg_counter
        # Slide 1 (Cover) uses full opacity; all subsequent slides use turned-down soft opacity
        if bg_counter == 0:
            bg_img_path = BG_COVER
        else:
            bg_img_path = BG_WAVES_SOFT[bg_counter % 3]
        bg_counter += 1
        if os.path.exists(bg_img_path):
            pic = slide.shapes.add_picture(bg_img_path, Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
            # Move picture to the back of the shape z-order
            slide.shapes._spTree.remove(pic._element)
            slide.shapes._spTree.insert(2, pic._element)
        else:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = COLOR_BG

    def add_header(slide, kicker, title, slide_num):
        # Kicker pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.46), Inches(2.9), Inches(0.36))
        pill.fill.solid()
        pill.fill.fore_color.rgb = COLOR_CARD_HOVER
        pill.line.color.rgb = RGBColor(191, 219, 254) # Soft blue border
        pill.line.width = Pt(1)
        tf = pill.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = Inches(0.04)
        p = tf.paragraphs[0]
        p.text = kicker.upper()
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_BRAND_PRIMARY
        p.font.name = "Arial"
        p.alignment = PP_ALIGN.CENTER

        # Title
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.88), Inches(9.8), Inches(0.68))
        tf2 = txBox.text_frame
        tf2.word_wrap = True
        tf2.margin_left = tf2.margin_top = Inches(0)
        p2 = tf2.paragraphs[0]
        p2.text = title
        p2.font.size = Pt(22)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_BRAND_NAVY
        p2.font.name = "Arial"

        # Transparent Logo Watermark
        if os.path.exists(LOGO_PATH):
            slide.shapes.add_picture(LOGO_PATH, Inches(10.7), Inches(0.46), width=Inches(1.85))

        # Footer
        footer = slide.shapes.add_textbox(Inches(12.0), Inches(7.02), Inches(1.0), Inches(0.3))
        ftf = footer.text_frame
        fp = ftf.paragraphs[0]
        fp.text = f"{slide_num:02d} / 24"
        fp.font.size = Pt(9.5)
        fp.font.color.rgb = COLOR_TEXT_LIGHT
        fp.alignment = PP_ALIGN.RIGHT

    def create_card(slide, left, top, width, height, title="", border=COLOR_CARD_BORDER, fill=COLOR_CARD_BG, is_highlight=False):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_HOVER if is_highlight else fill
        card.line.color.rgb = COLOR_BRAND_PRIMARY if is_highlight else border
        card.line.width = Pt(1.5) if is_highlight else Pt(1)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.18)
        if title:
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = COLOR_BRAND_NAVY
            p.space_after = Pt(6)
        return card, tf

    # =========================================================================
    # SLIDE 01: COVER (LIGHT MODERN ENTERPRISE)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_bg(s1)

    bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(0.1), Inches(4.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_BRAND_PRIMARY
    bar.line.fill.background()

    pill1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(1.2), Inches(3.2), Inches(0.42))
    pill1.fill.solid()
    pill1.fill.fore_color.rgb = COLOR_CARD_HOVER
    pill1.line.color.rgb = RGBColor(191, 219, 254)
    pill1.line.width = Pt(1.2)
    pp = pill1.text_frame.paragraphs[0]
    pp.text = "MOBRILZ PRIVATE LIMITED"
    pp.font.size = Pt(10)
    pp.font.bold = True
    pp.font.color.rgb = COLOR_BRAND_PRIMARY
    pp.alignment = PP_ALIGN.CENTER

    tx1 = s1.shapes.add_textbox(Inches(1.1), Inches(1.85), Inches(6.8), Inches(3.4))
    stf = tx1.text_frame
    stf.word_wrap = True
    ph = stf.paragraphs[0]
    ph.text = "Building Digital Solutions\nEnabling Business Growth"
    ph.font.size = Pt(36)
    ph.font.bold = True
    ph.font.color.rgb = COLOR_BRAND_NAVY
    ph.space_after = Pt(14)

    pd = stf.add_paragraph()
    pd.text = "Mobrilz is your technology and digital transformation partner helping businesses build, scale, and thrive through software engineering, intuitive product design, cloud operations, and performance digital marketing."
    pd.font.size = Pt(12)
    pd.font.color.rgb = COLOR_TEXT_MUTED
    pd.line_spacing = 1.3

    hero_mock = get_mockup('mockup_childcare.png')
    if hero_mock:
        s1.shapes.add_picture(hero_mock, Inches(8.2), Inches(1.6), width=Inches(4.35))

    bot1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(5.65), Inches(11.45), Inches(1.0))
    bot1.fill.solid()
    bot1.fill.fore_color.rgb = COLOR_CARD_BG
    bot1.line.color.rgb = COLOR_CARD_BORDER
    btf = bot1.text_frame
    btf.margin_top = Inches(0.14)
    bp = btf.paragraphs[0]
    bp.text = "WEB & E-COMMERCE  •  MOBILE APPS  •  CUSTOM SOFTWARE  •  UI/UX DESIGN  •  CLOUD & DEVOPS  •  DIGITAL MARKETING  •  AI AUTOMATIONS"
    bp.font.size = Pt(9.5)
    bp.font.bold = True
    bp.font.color.rgb = COLOR_BRAND_PRIMARY
    bp.alignment = PP_ALIGN.CENTER

    if os.path.exists(LOGO_PATH):
        s1.shapes.add_picture(LOGO_PATH, Inches(10.2), Inches(0.9), width=Inches(2.3))

    # =========================================================================
    # SLIDE 02: YOUR PARTNER IN DIGITAL TRANSFORMATION
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_bg(s2)
    add_header(s2, "COMPANY PROFILE", "Your Partner in Digital Transformation", 2)

    top_tx = s2.shapes.add_textbox(Inches(0.8), Inches(1.65), Inches(11.7), Inches(0.8))
    ttf = top_tx.text_frame
    ttf.word_wrap = True
    tp = ttf.paragraphs[0]
    tp.text = "Mobrilz turns ideas and operational bottlenecks into market ready digital products. By uniting software engineering, product design, cloud operations, and marketing, we support organizations across their complete digital journey."
    tp.font.size = Pt(12)
    tp.font.color.rgb = COLOR_TEXT_MUTED

    cards_s2 = [
        ("01 · ESTABLISH", "Digital Presence & Brand", [
            "Architect high converting corporate websites and e-commerce platforms.",
            "Establish unified brand visual language and frictionless UI/UX touchpoints.",
            "Equip internal marketing teams with modern headless and CMS capabilities."
        ]),
        ("02 · DEVELOP", "Custom Software & Mobile", [
            "Engineer bespoke cloud applications and native iOS/Android mobile solutions.",
            "Automate core business processes and eliminate manual operational drag.",
            "Modernize legacy software systems with resilient backend microservices."
        ]),
        ("03 · SCALE", "Cloud, Marketing & AI", [
            "Deploy secure cloud infrastructure with automated CI/CD and DevOps pipelines.",
            "Drive qualified customer acquisition via technical SEO and Google Ads.",
            "Integrate intelligent AI workflows and automations for exponential efficiency."
        ])
    ]

    for i, (kick, title, points) in enumerate(cards_s2):
        l = Inches(0.8) + i * Inches(4.0)
        c, tf = create_card(s2, l, Inches(2.7), Inches(3.7), Inches(4.1), is_highlight=(i == 1))
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(12)
        for p_str in points:
            pb = tf.add_paragraph()
            pb.text = f"•  {p_str}"
            pb.font.size = Pt(11)
            pb.font.color.rgb = COLOR_TEXT_MUTED
            pb.space_after = Pt(8)

    # =========================================================================
    # SLIDE 03: OUR JOURNEY IN NUMBERS
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_bg(s3)
    add_header(s3, "PROVEN SCALE", "Our Journey in Numbers", 3)

    metrics = [
        ("2018", "ESTABLISHED", "Pioneering transformative digital solutions, agile engineering, and scalable tech execution."),
        ("250+", "PROJECTS DELIVERED", "Proven track record across consumer mobile apps, enterprise web portals, and SaaS platforms."),
        ("200+", "GLOBAL CLIENTS", "Long term client partnerships spanning startups, scaleups, and established enterprises."),
        ("25+", "COUNTRIES SERVED", "International presence spanning India, UAE (Dubai), Kenya, France, and West Africa.")
    ]

    for i, (val, label, desc) in enumerate(metrics):
        l = Inches(0.8) + i * Inches(2.98)
        c, tf = create_card(s3, l, Inches(1.75), Inches(2.78), Inches(4.95), is_highlight=(i in [1, 2]))
        pv = tf.paragraphs[0]
        pv.text = val
        pv.font.size = Pt(38)
        pv.font.bold = True
        pv.font.color.rgb = COLOR_BRAND_PRIMARY
        pl = tf.add_paragraph()
        pl.text = label
        pl.font.size = Pt(11)
        pl.font.bold = True
        pl.font.color.rgb = COLOR_BRAND_NAVY
        pl.space_after = Pt(12)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.25

    # =========================================================================
    # SLIDE 04: MISSION, VISION AND VALUES
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_bg(s4)
    add_header(s4, "CORE PURPOSE", "Our Mission, Vision and Values", 4)

    c_m, tf_m = create_card(s4, Inches(0.8), Inches(1.8), Inches(5.6), Inches(2.35), is_highlight=True)
    pmk = tf_m.paragraphs[0]
    pmk.text = "OUR MISSION"
    pmk.font.size = Pt(10)
    pmk.font.bold = True
    pmk.font.color.rgb = COLOR_BRAND_PRIMARY
    pm_t = tf_m.add_paragraph()
    pm_t.text = "Solve Meaningful Challenges Through Tech"
    pm_t.font.size = Pt(15)
    pm_t.font.bold = True
    pm_t.font.color.rgb = COLOR_BRAND_NAVY
    pm_t.space_after = Pt(8)
    pm_d = tf_m.add_paragraph()
    pm_d.text = "To help businesses solve real world operational and commercial challenges through reliable digital solutions, thoughtful user experience design, and modern engineering practices."
    pm_d.font.size = Pt(11)
    pm_d.font.color.rgb = COLOR_TEXT_MUTED

    c_v, tf_v = create_card(s4, Inches(0.8), Inches(4.35), Inches(5.6), Inches(2.35))
    pvk = tf_v.paragraphs[0]
    pvk.text = "OUR VISION"
    pvk.font.size = Pt(10)
    pvk.font.bold = True
    pvk.font.color.rgb = COLOR_BRAND_PRIMARY
    pv_t = tf_v.add_paragraph()
    pv_t.text = "The Trusted Global Technology Partner"
    pv_t.font.size = Pt(15)
    pv_t.font.bold = True
    pv_t.font.color.rgb = COLOR_BRAND_NAVY
    pv_t.space_after = Pt(8)
    pv_d = tf_v.add_paragraph()
    pv_d.text = "To be the most reliable, adaptive technology partner for organizations worldwide, enabling them to continuously innovate, outpace competitors, and lead in a digital first economy."
    pv_d.font.size = Pt(11)
    pv_d.font.color.rgb = COLOR_TEXT_MUTED

    values = [
        ("Client Partnership", "We treat your objectives as our own, fostering radical transparency, direct collaboration, and long term shared success."),
        ("Quality & Reliability", "Zero compromise on code integrity, security standards, usability, and rigorous end to end quality assurance."),
        ("Continuous Innovation", "Rapidly evaluating and integrating modern technologies, from cross platform frameworks to generative AI automations."),
        ("Accountability & Ownership", "Clear ownership over delivery timelines, transparent sprints, proactive communication, and measured outcomes.")
    ]

    for idx, (v_name, v_desc) in enumerate(values):
        rx = Inches(6.7) + (idx % 2) * Inches(2.9)
        ry = Inches(1.8) + (idx // 2) * Inches(2.45)
        c_val, tf_val = create_card(s4, rx, ry, Inches(2.75), Inches(2.35))
        pv_name = tf_val.paragraphs[0]
        pv_name.text = v_name
        pv_name.font.size = Pt(13)
        pv_name.font.bold = True
        pv_name.font.color.rgb = COLOR_BRAND_NAVY
        pv_name.space_after = Pt(8)
        pvd = tf_val.add_paragraph()
        pvd.text = v_desc
        pvd.font.size = Pt(10.5)
        pvd.font.color.rgb = COLOR_TEXT_MUTED
        pvd.line_spacing = 1.2

    # =========================================================================
    # SLIDE 05: THE PROBLEM
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_bg(s5)
    add_header(s5, "THE PROBLEM & OPPORTUNITY", "Technology Should Solve Real Business Problems", 5)

    tx5 = s5.shapes.add_textbox(Inches(0.8), Inches(1.65), Inches(11.7), Inches(0.75))
    t5_f = tx5.text_frame
    t5_f.word_wrap = True
    p5 = t5_f.paragraphs[0]
    p5.text = "Organizations hit critical growth plateaus when their digital infrastructure fails to keep up. Mobrilz directly resolves the four major friction points holding modern companies back:"
    p5.font.size = Pt(12)
    p5.font.color.rgb = COLOR_TEXT_MUTED

    problems = [
        ("01 · OUTDATED DIGITAL EXPERIENCES", "The Conversion Barrier", 
         "Slow, clunky, non-responsive websites and apps erode customer trust, drive bounce rates up, and hand market share to tech-savvy competitors. We rebuild digital touchpoints with modern UI/UX and lightning fast performance."),
        ("02 · INEFFICIENT MANUAL PROCESSES", "The Capacity Bottleneck",
         "Teams bogged down by repetitive manual data entry, disconnected spreadsheets, and fragmented tools waste costly hours. We engineer custom automation workflows and business systems that save hundreds of labor hours."),
        ("03 · DISCONNECTED SYSTEMS & SILOS", "The Data Blindspot",
         "When CRM, inventory, ERP, and payment systems cannot communicate, operations stall and decisions are made blind. We build resilient API integrations and centralized dashboards that provide unified real time visibility."),
        ("04 · LIMITED DIGITAL REACH & TRAFFIC", "The Growth Ceiling",
         "Building a great product means nothing if qualified prospects never discover it. We deploy high intent SEO, targeted paid acquisition (Google Ads/Social), and digital PR to build an ongoing inbound customer pipeline.")
    ]

    for i, (kick, title, desc) in enumerate(problems):
        x = Inches(0.8) + (i % 2) * Inches(5.95)
        y = Inches(2.6) + (i // 2) * Inches(2.1)
        c, tf = create_card(s5, x, y, Inches(5.75), Inches(2.0), is_highlight=False)
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(6)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.2

    # =========================================================================
    # SLIDE 06: MULTIPLE DIGITAL CAPABILITIES
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_bg(s6)
    add_header(s6, "OUR CAPABILITIES", "One Partner. Multiple Digital Capabilities.", 6)

    services_s6 = [
        ("Web & E-Commerce", "High performance corporate websites, global e-commerce stores, custom web apps, and modern CMS architectures built for conversion.", "Next.js · React · Node · Laravel"),
        ("Mobile Applications", "Native iOS & Android, plus high velocity cross platform Flutter and React Native apps with real time sync and offline support.", "iOS (Swift) · Android · Flutter"),
        ("Custom Software & UX", "Tailored ERP/CRM business systems, SaaS platforms, intuitive UI/UX design systems, interactive prototypes, and usability audits.", "Figma · Design Systems · SaaS"),
        ("Cloud & Technical Ops", "Scalable cloud infrastructure, DevOps pipelines (CI/CD), API integrations, 24/7 proactive maintenance, and security hardening.", "AWS · Azure · Docker · CI/CD"),
        ("Digital Marketing", "Technical and e-commerce SEO, high ROI Google & Meta ad campaigns, social media management, brand positioning, and content.", "SEO · PPC · Social · Branding"),
        ("AI & Intelligent Automation", "Generative AI, machine learning, automated customer SDR agents, voice bots, and intelligent workflow automation tools.", "LLMs · NLP · Custom AI Agents")
    ]

    for i, (title, desc, stack) in enumerate(services_s6):
        x = Inches(0.8) + (i % 3) * Inches(3.98)
        y = Inches(1.8) + (i // 3) * Inches(2.55)
        c, tf = create_card(s6, x, y, Inches(3.75), Inches(2.4), is_highlight=False)
        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(14.5)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.25
        pd.space_after = Pt(10)
        ps = tf.add_paragraph()
        ps.text = stack
        ps.font.size = Pt(9.5)
        ps.font.bold = True
        ps.font.color.rgb = COLOR_BRAND_PRIMARY

    # =========================================================================
    # SLIDE 07: WEB & E-COMMERCE (WITH MAROLIMA REAL MOCKUP)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_bg(s7)
    add_header(s7, "CORE SERVICE · 01", "Digital Experiences Built for Business: Web & E-Commerce", 7)

    c_w1, tf_w1 = create_card(s7, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.9), is_highlight=True)
    pwk = tf_w1.paragraphs[0]
    pwk.text = "THE MOBRILZ WEB ADVANTAGE"
    pwk.font.size = Pt(10)
    pwk.font.bold = True
    pwk.font.color.rgb = COLOR_BRAND_PRIMARY
    pwt = tf_w1.add_paragraph()
    pwt.text = "Engineered for Conversion & Speed"
    pwt.font.size = Pt(16)
    pwt.font.bold = True
    pwt.font.color.rgb = COLOR_BRAND_NAVY
    pwt.space_after = Pt(10)
    pwd = tf_w1.add_paragraph()
    pwd.text = "A modern website must do far more than look impressive, acting as an automated engine for qualified leads, smooth checkout journeys, and operational efficiency.\n\nFrom high traffic corporate portals to headless e-commerce architectures, we engineer websites built for blazing speed (95+ Core Web Vitals), ironclad security, and maximum ROI."
    pwd.font.size = Pt(11)
    pwd.font.color.rgb = COLOR_TEXT_MUTED
    pwd.line_spacing = 1.3
    pwd.space_after = Pt(12)

    web_bullets = [
        "Corporate Portals & Enterprise Sites",
        "High Conversion E-Commerce Engines",
        "Dynamic Custom Web Applications & SaaS",
        "Headless CMS Architectures (Strapi, WP)",
        "Progressive Web Apps (PWA) with Offline Sync"
    ]
    for b in web_bullets:
        pb = tf_w1.add_paragraph()
        pb.text = f"•  {b}"
        pb.font.size = Pt(10.5)
        pb.font.color.rgb = COLOR_BRAND_NAVY
        pb.space_after = Pt(4)

    maro_mock = get_mockup('mockup_marolima.png')
    if maro_mock:
        s7.shapes.add_picture(maro_mock, Inches(6.9), Inches(1.8), width=Inches(5.6))
        cap = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(5.85), Inches(5.6), Inches(0.85))
        cap.fill.solid()
        cap.fill.fore_color.rgb = COLOR_CARD_BG
        cap.line.color.rgb = COLOR_CARD_BORDER
        ctf = cap.text_frame
        ctf.margin_top = Inches(0.12)
        cp1 = ctf.paragraphs[0]
        cp1.text = "FEATURED CASE STUDY: MAROLIMA ASSET MANAGEMENT"
        cp1.font.size = Pt(9.5)
        cp1.font.bold = True
        cp1.font.color.rgb = COLOR_BRAND_PRIMARY
        cp2 = ctf.add_paragraph()
        cp2.text = "Institutional corporate web platform engineered for security, high-net-worth investor engagement, and compliance."
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 08: MOBILE APPLICATIONS (WITH TAXISCOUT24 MOCKUP)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_bg(s8)
    add_header(s8, "CORE SERVICE · 02", "Bringing Digital Products to Mobile: iOS & Android", 8)

    # Left: Mockup Image
    taxi_mock = get_mockup('mockup_taxiscout.png')
    if taxi_mock:
        s8.shapes.add_picture(taxi_mock, Inches(0.8), Inches(2.0), width=Inches(5.6))

    # Right: Content Card
    c_m1, tf_m1 = create_card(s8, Inches(6.7), Inches(1.8), Inches(5.8), Inches(4.9), is_highlight=True)
    pmk = tf_m1.paragraphs[0]
    pmk.text = "MOBILE EXCELLENCE"
    pmk.font.size = Pt(10)
    pmk.font.bold = True
    pmk.font.color.rgb = COLOR_BRAND_PRIMARY
    pmt = tf_m1.add_paragraph()
    pmt.text = "Native Speed, Unified Multi Platform"
    pmt.font.size = Pt(16)
    pmt.font.bold = True
    pmt.font.color.rgb = COLOR_BRAND_NAVY
    pmt.space_after = Pt(10)
    pmd = tf_m1.add_paragraph()
    pmd.text = "Mobile apps give businesses a direct line to customers' pockets. We architect iOS and Android apps with striking UI, resilient backends, real time sync, and fluid 60fps animations.\n\nWhether developing native Swift/Kotlin or high velocity Flutter/React Native solutions, we ensure maximum performance and five-star user ratings."
    pmd.font.size = Pt(11)
    pmd.font.color.rgb = COLOR_TEXT_MUTED
    pmd.line_spacing = 1.3
    pmd.space_after = Pt(12)

    mob_bullets = [
        "Customer-Facing eCommerce & Account Apps",
        "On Demand Booking & Dispatch with GPS Tracking",
        "Enterprise Mobility & Field Staff Tools",
        "Real Time Chat, Push Alerts & Streaming",
        "End to End Store Publishing (App Store & Play Store)"
    ]
    for b in mob_bullets:
        pb = tf_m1.add_paragraph()
        pb.text = f"•  {b}"
        pb.font.size = Pt(10.5)
        pb.font.color.rgb = COLOR_BRAND_NAVY
        pb.space_after = Pt(4)

    # =========================================================================
    # SLIDE 09: CUSTOM SOFTWARE SOLUTIONS (WITH MINDTO MOCKUP)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_bg(s9)
    add_header(s9, "CORE SERVICE · 03", "Software Designed Around Your Business: Tailored Systems", 9)

    c_s1, tf_s1 = create_card(s9, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.9), is_highlight=True)
    psk = tf_s1.paragraphs[0]
    psk.text = "BESPOKE ENGINEERING"
    psk.font.size = Pt(10)
    psk.font.bold = True
    psk.font.color.rgb = COLOR_BRAND_PRIMARY
    pst = tf_s1.add_paragraph()
    pst.text = "Own Your Tech. Eliminate Bottlenecks."
    pst.font.size = Pt(16)
    pst.font.bold = True
    pst.font.color.rgb = COLOR_BRAND_NAVY
    pst.space_after = Pt(10)
    psd = tf_s1.add_paragraph()
    psd.text = "Standard software forces your business into someone else's rigid mold, creating expensive per seat fees, disjointed spreadsheets, and manual workarounds.\n\nMobrilz architects bespoke software applications built 100% around your proprietary rules, team workflows, and competitive edge. You own the IP, eliminate recurring SaaS seat tax, and scale infinitely."
    psd.font.size = Pt(11)
    psd.font.color.rgb = COLOR_TEXT_MUTED
    psd.line_spacing = 1.3
    psd.space_after = Pt(12)

    soft_bullets = [
        "Business Process Management & Workflow Automation",
        "Custom Multi Tenant SaaS Platforms from MVP to Scale",
        "Centralized Enterprise ERP/CRM Platforms",
        "Legacy Code Refactoring & Microservice Modernization",
        "Secure Enterprise API Integrations & Webhooks"
    ]
    for b in soft_bullets:
        pb = tf_s1.add_paragraph()
        pb.text = f"•  {b}"
        pb.font.size = Pt(10.5)
        pb.font.color.rgb = COLOR_BRAND_NAVY
        pb.space_after = Pt(4)

    mind_mock = get_mockup('mockup_mindto.png')
    if mind_mock:
        s9.shapes.add_picture(mind_mock, Inches(6.9), Inches(1.8), width=Inches(5.6))
        cap = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(5.85), Inches(5.6), Inches(0.85))
        cap.fill.solid()
        cap.fill.fore_color.rgb = COLOR_CARD_BG
        cap.line.color.rgb = COLOR_CARD_BORDER
        ctf = cap.text_frame
        ctf.margin_top = Inches(0.12)
        cp1 = ctf.paragraphs[0]
        cp1.text = "FEATURED CASE STUDY: MINDTO"
        cp1.font.size = Pt(9.5)
        cp1.font.bold = True
        cp1.font.color.rgb = COLOR_BRAND_PRIMARY
        cp2 = ctf.add_paragraph()
        cp2.text = "Tailored digital platform built with modern full-stack architecture, dynamic role-based access, and automated workflow logic."
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 10: THOUGHTFUL DESIGN & UI/UX (WITH LIVING ART MOCKUP)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_bg(s10)
    add_header(s10, "CORE SERVICE · 04", "Thoughtful Design. Better Digital Experiences: UI/UX", 10)

    ux_steps = [
        ("01 · DISCOVER", "UX Research & Strategy", "User interviews, persona definition, competitor benchmarking, and behavioral analytics mapping."),
        ("02 · STRUCTURE", "Information Architecture", "Sitemaps, navigation hierarchies, user flows, and core customer journey optimization."),
        ("03 · PROTOTYPE", "Interactive Figma Models", "Rapid clickable prototypes to test usability, validate hypotheses, and align stakeholders before coding."),
        ("04 · POLISH", "UI Design & Design Systems", "High-fidelity visual design, WCAG accessibility, cohesive component libraries, and design tokens.")
    ]

    for i, (kick, title, desc) in enumerate(ux_steps):
        l = Inches(0.8) + i * Inches(2.98)
        c, tf = create_card(s10, l, Inches(1.8), Inches(2.78), Inches(2.4), is_highlight=(i == 3))
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.2

    art_mock = get_mockup('mockup_livingart.png')
    if art_mock:
        s10.shapes.add_picture(art_mock, Inches(0.8), Inches(4.45), width=Inches(3.8))

    bot_card10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(4.45), Inches(7.7), Inches(2.4))
    bot_card10.fill.solid()
    bot_card10.fill.fore_color.rgb = COLOR_CARD_BG
    bot_card10.line.color.rgb = COLOR_BRAND_PRIMARY
    bot10_tf = bot_card10.text_frame
    bot10_tf.margin_left = Inches(0.3)
    bot10_tf.margin_top = Inches(0.2)
    bp1 = bot10_tf.paragraphs[0]
    bp1.text = "THE CRO FIRST DESIGN PHILOSOPHY"
    bp1.font.size = Pt(10.5)
    bp1.font.bold = True
    bp1.font.color.rgb = COLOR_BRAND_PRIMARY
    bp2 = bot10_tf.add_paragraph()
    bp2.text = "We don't design just for surface-level aesthetics; every interface, form, button, and checkout funnel is engineered to reduce cognitive friction, accelerate user engagement, and lift conversion rates.\n\nFrom interactive wireframing to comprehensive component libraries, our design systems ensure that your brand looks cohesive and authoritative across mobile, tablet, and desktop viewports."
    bp2.font.size = Pt(11)
    bp2.font.color.rgb = COLOR_BRAND_NAVY
    bp2.line_spacing = 1.3

    # =========================================================================
    # SLIDE 11: CLOUD SOLUTIONS & DEVOPS (WITH TESTING MOCKUP)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_bg(s11)
    add_header(s11, "CORE SERVICE · 05", "Keeping Digital Products Resilient: Cloud & DevOps", 11)

    c_c1, tf_c1 = create_card(s11, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.9), is_highlight=True)
    pck = tf_c1.paragraphs[0]
    pck.text = "INFRASTRUCTURE & RELIABILITY"
    pck.font.size = Pt(10)
    pck.font.bold = True
    pck.font.color.rgb = COLOR_BRAND_PRIMARY
    pct = tf_c1.add_paragraph()
    pct.text = "Zero Downtime and Elastic Scalability"
    pct.font.size = Pt(16)
    pct.font.bold = True
    pct.font.color.rgb = COLOR_BRAND_NAVY
    pct.space_after = Pt(10)
    pcd = tf_c1.add_paragraph()
    pcd.text = "Great software requires bulletproof hosting infrastructure. Mobrilz architects elastic cloud environments, automates CI/CD deployment pipelines, and manages 24/7 security hardening.\n\nOur DevOps practices eliminate deployment bottlenecks, maintain 99.9% uptime SLAs, and ensure your system handles viral traffic surges with ease."
    pcd.font.size = Pt(11)
    pcd.font.color.rgb = COLOR_TEXT_MUTED
    pcd.line_spacing = 1.3
    pcd.space_after = Pt(12)

    cloud_bullets = [
        "Cloud Infrastructure on AWS, Microsoft Azure & GCP",
        "Automated CI/CD Deployment Pipelines (Docker, Jenkins)",
        "Microservices & Decoupled REST/GraphQL Architectures",
        "24/7 Server Monitoring, Security Patches & Backups",
        "Database Sharding, Caching (Redis) & Load Balancing"
    ]
    for b in cloud_bullets:
        pb = tf_c1.add_paragraph()
        pb.text = f"•  {b}"
        pb.font.size = Pt(10.5)
        pb.font.color.rgb = COLOR_BRAND_NAVY
        pb.space_after = Pt(4)

    test_mock = get_mockup('mockup_testing.png')
    if test_mock:
        s11.shapes.add_picture(test_mock, Inches(6.9), Inches(1.8), width=Inches(5.6))
        cap = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(5.85), Inches(5.6), Inches(0.85))
        cap.fill.solid()
        cap.fill.fore_color.rgb = COLOR_CARD_BG
        cap.line.color.rgb = COLOR_CARD_BORDER
        ctf = cap.text_frame
        ctf.margin_top = Inches(0.12)
        cp1 = ctf.paragraphs[0]
        cp1.text = "CONTINUOUS AUTOMATION TESTING & DEVOPS"
        cp1.font.size = Pt(9.5)
        cp1.font.bold = True
        cp1.font.color.rgb = COLOR_BRAND_PRIMARY
        cp2 = ctf.add_paragraph()
        cp2.text = "Automated test suites, load testing, containerization, and automated rollback safeguarding your mission-critical applications."
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 12: DIGITAL MARKETING & SEO (WITH AYRTON COFFEE MOCKUP)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_bg(s12)
    add_header(s12, "CORE SERVICE · 06", "Helping Businesses Reach the Right Audience: Digital Growth", 12)

    c_d1, tf_d1 = create_card(s12, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.9), is_highlight=True)
    pdk = tf_d1.paragraphs[0]
    pdk.text = "PERFORMANCE MARKETING"
    pdk.font.size = Pt(10)
    pdk.font.bold = True
    pdk.font.color.rgb = COLOR_BRAND_PRIMARY
    pdt = tf_d1.add_paragraph()
    pdt.text = "Sustainable Organic Rank. High ROAS Ads."
    pdt.font.size = Pt(16)
    pdt.font.bold = True
    pdt.font.color.rgb = COLOR_BRAND_NAVY
    pdt.space_after = Pt(10)
    pdd = tf_d1.add_paragraph()
    pdd.text = "Building a stunning digital platform is only half the battle. We drive qualified buyer traffic through rigorous technical SEO, data driven Google Ads (PPC), social advertising, and brand campaigns.\n\nEvery marketing dollar is tracked against concrete commercial KPIs, measuring Cost Per Lead (CPL), Customer Acquisition Cost (CAC), and Return on Ad Spend (ROAS)."
    pdd.font.size = Pt(11)
    pdd.font.color.rgb = COLOR_TEXT_MUTED
    pdd.line_spacing = 1.3
    pdd.space_after = Pt(12)

    mkt_bullets = [
        "Technical, On-Page & Local SEO with Core Web Vitals Focus",
        "High Intent Google Ads, Search, Shopping & Performance Max",
        "Paid Social Media Campaigns on Meta, LinkedIn & TikTok",
        "Landing Page Conversion Rate Optimization (A/B Testing)",
        "Brand Strategy, Digital PR & High-Authority Link Building"
    ]
    for b in mkt_bullets:
        pb = tf_d1.add_paragraph()
        pb.text = f"•  {b}"
        pb.font.size = Pt(10.5)
        pb.font.color.rgb = COLOR_BRAND_NAVY
        pb.space_after = Pt(4)

    ayr_mock = get_mockup('mockup_ayrton.png')
    if ayr_mock:
        s12.shapes.add_picture(ayr_mock, Inches(6.9), Inches(1.8), width=Inches(5.6))
        cap = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(5.85), Inches(5.6), Inches(0.85))
        cap.fill.solid()
        cap.fill.fore_color.rgb = COLOR_CARD_BG
        cap.line.color.rgb = COLOR_CARD_BORDER
        ctf = cap.text_frame
        ctf.margin_top = Inches(0.12)
        cp1 = ctf.paragraphs[0]
        cp1.text = "FEATURED CASE STUDY: AYRTON COFFEE COMPANY"
        cp1.font.size = Pt(9.5)
        cp1.font.bold = True
        cp1.font.color.rgb = COLOR_BRAND_PRIMARY
        cp2 = ctf.add_paragraph()
        cp2.text = "Full digital branding, e-commerce storefront, and ongoing organic search visibility lifting direct-to-consumer sales."
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 13: ARTIFICIAL INTELLIGENCE & AUTOMATION
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_bg(s13)
    add_header(s13, "CORE SERVICE · 07", "Exploring New Possibilities: AI & Intelligent Automation", 13)

    ai_cards = [
        ("Generative AI & LLMs", "Integrate state-of-the-art language models for smart enterprise search, dynamic document summarization, and automated content generation.", "Custom LLM & RAG Pipelines"),
        ("AI SDRs & Chat Agents", "Deploy 24/7 intelligent sales and customer support agents that qualify incoming leads, answer technical queries, and book appointments automatically.", "Voice & Chat Lead Capture"),
        ("Intelligent Process Automation", "Automate complex decision trees, invoice data extraction, document classification, and multi-system syncing without manual touch.", "Zero-Manual Data Handling"),
        ("Predictive Analytics & ML", "Harness customer behavioral data and past sales metrics to forecast demand, reduce churn, and personalize product recommendations.", "Data-Driven Forecasting")
    ]

    for i, (title, desc, tag) in enumerate(ai_cards):
        x = Inches(0.8) + (i % 2) * Inches(5.95)
        y = Inches(1.8) + (i // 2) * Inches(2.55)
        c, tf = create_card(s13, x, y, Inches(5.75), Inches(2.4), is_highlight=(i == 1))
        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(14.5)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.25
        pd.space_after = Pt(12)
        ps = tf.add_paragraph()
        ps.text = f"CAPABILITY: {tag.upper()}"
        ps.font.size = Pt(9.5)
        ps.font.bold = True
        ps.font.color.rgb = COLOR_BRAND_PRIMARY

    # =========================================================================
    # SLIDE 14: DIGITAL SOLUTIONS ACROSS INDUSTRIES
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_bg(s14)
    add_header(s14, "DOMAIN EXPERTISE", "Digital Solutions Across Key Industries", 14)

    industries = [
        ("Finance & Fintech", "Secure investment portals, wealth management dashboards, AML/KYC compliant workflows, and transaction systems."),
        ("Healthcare & Fitness", "HIPAA conscious patient scheduling portals, wellness mobile apps, telehealth consultations, and clinic management."),
        ("E-Commerce & Retail", "Omnichannel retail engines, B2B wholesale portals, multi vendor marketplaces, and integrated payment gateways."),
        ("Logistics & Supply Chain", "Fleet tracking apps, dispatch dispatchers, automated shipment tracking, and warehouse inventory control systems."),
        ("Manufacturing & Industrial", "Custom ERP suites, equipment maintenance logging, supply procurement automation, and production monitoring."),
        ("Education & E-Learning", "Interactive learning management systems (LMS), student progress analytics, course delivery, and live video sessions."),
        ("Professional Consulting", "Lead capture engines, client onboarding portals, automated contract signing, and digital brand platforms."),
        ("Hospitality & Services", "Direct booking engines, dynamic pricing integrations, reservation management, and interactive guest apps.")
    ]

    for i, (title, desc) in enumerate(industries):
        x = Inches(0.8) + (i % 4) * Inches(2.98)
        y = Inches(1.8) + (i // 4) * Inches(2.55)
        c, tf = create_card(s14, x, y, Inches(2.78), Inches(2.4), is_highlight=(i in [0, 2]))
        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.2

    # =========================================================================
    # SLIDE 15: THE AGILE DELIVERY PROCESS
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_bg(s15)
    add_header(s15, "HOW WE WORK", "From Business Challenge to Digital Solution: 5 Step Lifecycle", 15)

    stages = [
        ("01 · DISCOVER", "Scope & Alignment", "Deep stakeholder interviews, technical discovery, KPI definition, and requirements documentation."),
        ("02 · DEFINE", "Architecture & Plan", "Sprint roadmap, cloud tech stack selection, data modeling, and security framework specification."),
        ("03 · DESIGN", "Prototypes & Flows", "Interactive Figma wireframes, UI design systems, and rapid user testing before a line of code is written."),
        ("04 · DEVELOP", "Agile Sprints & QA", "Bi weekly sprint demos, automated unit/integration testing, security audits, and continuous peer reviews."),
        ("05 · DEPLOY", "Launch & Evolution", "Production rollout, CDN configuration, real time telemetry, team training, and continuous sprint iteration.")
    ]

    for i, (kick, title, desc) in enumerate(stages):
        l = Inches(0.8) + i * Inches(2.38)
        c, tf = create_card(s15, l, Inches(1.8), Inches(2.2), Inches(4.9), is_highlight=(i in [2, 3]))
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(12)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.25

    # =========================================================================
    # SLIDE 16: FLEXIBLE WAYS TO WORK WITH MOBRILZ
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_bg(s16)
    add_header(s16, "COMMERCIAL MODELS", "Flexible Ways to Work With Mobrilz", 16)

    models = [
        ("PATH 01 · PROJECT-BASED", "Fixed Scope Delivery", [
            "Ideal for clearly defined websites, mobile apps, MVPs, and software builds.",
            "Guaranteed delivery milestones, fixed pricing, and predetermined launch dates.",
            "Dedicated Project Manager leading complete end to end execution."
        ], "Fixed Timeline & Budget"),
        ("PATH 02 · DEDICATED TALENT", "Staff Augmentation", [
            "Embed dedicated senior developers, designers, or SEO specialists into your team.",
            "Full time commitment (160 hrs/mo) working directly under your workflows.",
            "70% cost reduction vs. local onshore hires with zero HR/payroll management."
        ], "Plugged Straight Into Your Team"),
        ("PATH 03 · MANAGED SERVICES", "Ongoing Support & Growth", [
            "Continuous technical maintenance, security updates, and infrastructure monitoring.",
            "Ongoing monthly SEO, Google Ads management, and conversion rate optimization.",
            "Guaranteed response SLAs with dedicated communication via Slack & Asana."
        ], "Continuous Long-Term Evolution")
    ]

    for i, (kick, title, points, tag) in enumerate(models):
        l = Inches(0.8) + i * Inches(4.0)
        c, tf = create_card(s16, l, Inches(1.8), Inches(3.7), Inches(4.9), is_highlight=(i == 1))
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(14)
        for p_str in points:
            pb = tf.add_paragraph()
            pb.text = f"•  {p_str}"
            pb.font.size = Pt(11)
            pb.font.color.rgb = COLOR_TEXT_MUTED
            pb.space_after = Pt(10)
        ps = tf.add_paragraph()
        ps.text = f"VALUE: {tag.upper()}"
        ps.font.size = Pt(9.5)
        ps.font.bold = True
        ps.font.color.rgb = COLOR_BRAND_PRIMARY
        ps.space_before = Pt(14)

    # =========================================================================
    # SLIDE 17: SELECTED PORTFOLIO (WITH 4 REAL PROJECT MOCKUPS)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_bg(s17)
    add_header(s17, "PROVEN CASE STUDIES", "Turning Ideas Into Digital Experiences: Selected Client Work", 17)

    projects_s17 = [
        ("MAROLIMA ASSET MANAGEMENT", "Corporate Fintech Web Portal", "mockup_marolima.png"),
        ("TAXISCOUT24", "On Demand Ride Hailing App", "mockup_taxiscout.png"),
        ("AYRTON COFFEE COMPANY", "Direct-to-Consumer E-Commerce", "mockup_ayrton.png"),
        ("SF50 FLIGHT SUPPORT", "Aviation Services & Corporate Web", "mockup_sf50.png")
    ]

    for i, (p_title, p_sub, p_img) in enumerate(projects_s17):
        col = i % 2
        row = i // 2
        card_x = Inches(0.8) + col * Inches(5.95)
        card_y = Inches(1.8) + row * Inches(2.55)
        
        c, tf = create_card(s17, card_x, card_y, Inches(5.75), Inches(2.4), is_highlight=False)
        
        img_p = get_mockup(p_img)
        if img_p:
            s17.shapes.add_picture(img_p, card_x + Inches(0.15), card_y + Inches(0.18), width=Inches(2.5))
            
        tx_card = s17.shapes.add_textbox(card_x + Inches(2.8), card_y + Inches(0.18), Inches(2.8), Inches(2.0))
        ctf = tx_card.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_top = Inches(0)
        
        cp1 = ctf.paragraphs[0]
        cp1.text = p_title
        cp1.font.size = Pt(12)
        cp1.font.bold = True
        cp1.font.color.rgb = COLOR_BRAND_NAVY
        cp1.space_after = Pt(4)
        
        cp2 = ctf.add_paragraph()
        cp2.text = p_sub
        cp2.font.size = Pt(10)
        cp2.font.bold = True
        cp2.font.color.rgb = COLOR_BRAND_PRIMARY
        cp2.space_after = Pt(6)
        
        cp3 = ctf.add_paragraph()
        cp3.text = "Custom-engineered by Mobrilz delivering enterprise security, high velocity UX, and scalable cloud performance."
        cp3.font.size = Pt(9.5)
        cp3.font.color.rgb = COLOR_TEXT_MUTED
        cp3.line_spacing = 1.15

    # =========================================================================
    # SLIDE 18: TRUSTED CLIENT PORTFOLIO (CLIENT LOGOS & APPS) - HIDDEN SLIDE
    # =========================================================================
    s_clients = prs.slides.add_slide(blank_layout)
    # Mark slide as hidden in PowerPoint presentation mode (show="0")
    s_clients.element.set('show', '0')
    set_bg(s_clients)
    add_header(s_clients, "GLOBAL CLIENTELE", "Trusted by Fast Growing Businesses & Enterprises", 18)

    clients_data = [
        ("AI & SAAS", "Mindto", "mindto.io", "AI productivity & knowledge management SaaS"),
        ("ENTERPRISE SAAS", "Flowysis", "flowysis.com", "Automated workflow & business analytics platform"),
        ("COMMUNITY", "Tula Family", "tula-family.com", "Digital community & modern family portal"),
        ("SWISS MOBILITY", "LocaFri", "locafri.ch", "Swiss vehicle rental & fleet management software"),
        ("MATCHMAKING", "Niyyah Marriage", "niyyah-marriage.com", "Verified matchmaking & privacy-first social web app"),
        ("DISPLAY & EVENTS", "Team1 Display", "team1displayservices.com", "Exhibition booth & event operations portal"),
        ("REAL ESTATE", "El Manantial", "elmanantial.eu", "Official residential community management in Marbella"),
        ("HEALTHCARE APP", "KoreMe Longevity", "koremeantiaging.com", "Patient wellness tracking & mobile app (iOS & Android)")
    ]

    for i, (kicker, name, url, desc) in enumerate(clients_data):
        col = i % 4
        row = i // 4
        x = Inches(0.8) + col * Inches(2.98)
        y = Inches(1.8) + row * Inches(2.05)
        
        c, tf = create_card(s_clients, x, y, Inches(2.8), Inches(1.9), is_highlight=False)
        pk = tf.paragraphs[0]
        pk.text = kicker
        pk.font.size = Pt(8.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        
        pn = tf.add_paragraph()
        pn.text = name
        pn.font.size = Pt(13)
        pn.font.bold = True
        pn.font.color.rgb = COLOR_BRAND_NAVY
        pn.space_after = Pt(2)

        pu = tf.add_paragraph()
        pu.text = url
        pu.font.size = Pt(9)
        pu.font.color.rgb = COLOR_BRAND_PRIMARY
        pu.space_after = Pt(6)

        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(9.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.2

    # Bottom Banner: Chalo-Reisen Travel App Showcase
    chalo_card = s_clients.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.05), Inches(11.7), Inches(0.85))
    chalo_card.fill.solid()
    chalo_card.fill.fore_color.rgb = COLOR_CARD_BG
    chalo_card.line.color.rgb = COLOR_BRAND_PRIMARY
    chalo_tf = chalo_card.text_frame
    chalo_tf.margin_left = Inches(0.25)
    chalo_tf.margin_top = Inches(0.12)
    
    cp1 = chalo_tf.paragraphs[0]
    cp1.text = "FEATURED TRAVEL APP: CHALO REISEN (iOS & ANDROID)"
    cp1.font.size = Pt(9.5)
    cp1.font.bold = True
    cp1.font.color.rgb = COLOR_BRAND_PRIMARY
    
    cp2 = chalo_tf.add_paragraph()
    cp2.text = "End to end travel companion published across Apple App Store and Google Play Store with real time booking and interactive travel guides."
    cp2.font.size = Pt(10)
    cp2.font.color.rgb = COLOR_BRAND_NAVY

    # =========================================================================
    # SLIDE 19: TECHNOLOGY STACK
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    set_bg(s18)
    add_header(s18, "TECHNICAL CAPABILITIES", "The Right Technologies for the Right Requirements", 19)

    tech_stacks = [
        ("Frontend Development", "React.js · Next.js · Vue.js · TypeScript · Tailwind CSS · Vite", "High speed, responsive, SEO-ready user interfaces."),
        ("Backend Engineering", "Node.js · Python (FastAPI, Django) · PHP (Laravel) · Express", "Scalable, secure, high-throughput microservice backends."),
        ("Mobile Engineering", "Flutter · React Native · iOS (Swift) · Android (Kotlin)", "Fluid native animations and multi-platform velocity."),
        ("Databases & Storage", "PostgreSQL · MongoDB · MySQL · Redis · Cloud Firestore", "ACID-compliant relational systems and distributed document stores."),
        ("Cloud & DevOps", "Amazon Web Services (AWS) · Google Cloud · Microsoft Azure · Docker", "Elastic cloud scaling, CI/CD pipelines, and zero-downtime deploys."),
        ("CMS & Headless", "WordPress · Strapi · Shopify · WooCommerce · Headless CMS", "Flexible content management tailored for marketing independence.")
    ]

    for i, (title, stack, desc) in enumerate(tech_stacks):
        x = Inches(0.8) + (i % 3) * Inches(3.98)
        y = Inches(1.8) + (i // 3) * Inches(2.55)
        c, tf = create_card(s18, x, y, Inches(3.75), Inches(2.4), is_highlight=False)
        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(6)
        ps = tf.add_paragraph()
        ps.text = stack
        ps.font.size = Pt(10.5)
        ps.font.bold = True
        ps.font.color.rgb = COLOR_BRAND_PRIMARY
        ps.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.2

    # =========================================================================
    # SLIDE 19: WHY MOBRILZ (WITH FLS MOCKUP)
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    set_bg(s19)
    add_header(s19, "OUR VALUE PROPOSITION", "A Technology Partner for the Entire Digital Journey", 20)

    c_why, tf_why = create_card(s19, Inches(0.8), Inches(1.8), Inches(6.8), Inches(4.9), is_highlight=True)
    pwk = tf_why.paragraphs[0]
    pwk.text = "THE MOBRILZ ADVANTAGE"
    pwk.font.size = Pt(10)
    pwk.font.bold = True
    pwk.font.color.rgb = COLOR_BRAND_PRIMARY

    pillars19 = [
        ("01 · Integrated Capabilities", "No need to juggle separate agencies for mobile, web, design, and SEO. Mobrilz handles your full stack under one synchronized team."),
        ("02 · Commercial Orientation", "We don't build software for technology's sake. Every line of code and user journey is engineered directly around your revenue and operational KPIs."),
        ("03 · Maximum Flexibility", "From fixed scope turnkey project delivery to dedicated full time talent embedding, our engagement models adapt to your exact stage of growth."),
        ("04 · Lifecycle Commitment", "We don't disappear after launch. We provide ongoing security patches, performance scaling, technical maintenance, and growth marketing.")
    ]

    for p_title, p_desc in pillars19:
        pt = tf_why.add_paragraph()
        pt.text = p_title
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_before = Pt(8)
        pd = tf_why.add_paragraph()
        pd.text = p_desc
        pd.font.size = Pt(10)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.15

    fls_mock = get_mockup('mockup_fls.png')
    if fls_mock:
        s19.shapes.add_picture(fls_mock, Inches(7.9), Inches(1.8), width=Inches(4.6))
        cap = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.9), Inches(5.1), Inches(4.6), Inches(1.6))
        cap.fill.solid()
        cap.fill.fore_color.rgb = COLOR_CARD_BG
        cap.line.color.rgb = COLOR_CARD_BORDER
        ctf = cap.text_frame
        ctf.margin_top = Inches(0.14)
        cp1 = ctf.paragraphs[0]
        cp1.text = "REAL RESULTS: FAST LIGHTING SUPPLY"
        cp1.font.size = Pt(9.5)
        cp1.font.bold = True
        cp1.font.color.rgb = COLOR_BRAND_PRIMARY
        cp2 = ctf.add_paragraph()
        cp2.text = "Engineered custom B2B e-commerce platform with automated catalog indexing, high-volume order processing, and wholesale tier pricing."
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = COLOR_TEXT_MUTED

    # =========================================================================
    # SLIDE 20: LEADERSHIP & TEAM
    # =========================================================================
    s20 = prs.slides.add_slide(blank_layout)
    set_bg(s20)
    add_header(s20, "OUR LEADERSHIP", "Driven by People and United by Technology", 21)

    leaders = [
        ("EXECUTIVE LEADERSHIP", "Chief Executive Officer", "Sets corporate vision, global expansion strategy, and high level client relationship standards."),
        ("ENGINEERING LEADERSHIP", "Chief Technology Officer", "Directs software architecture, engineering standards, cloud security, and R&D adoption."),
        ("PRODUCT & UX", "Head of Product Design", "Ensures world class user experience, visual elegance, accessibility standards, and CRO principles."),
        ("DELIVERY & OPERATIONS", "Director of Project Delivery", "Guarantees on time agile delivery, sprint velocity, transparent communication, and QA governance.")
    ]

    for i, (kicker, title, desc) in enumerate(leaders):
        l = Inches(0.8) + i * Inches(2.98)
        c, tf = create_card(s20, l, Inches(1.8), Inches(2.78), Inches(4.9), is_highlight=False)
        pk = tf.paragraphs[0]
        pk.text = kicker
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(14)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.3

    # =========================================================================
    # SLIDE 21: DELIVERY PRINCIPLES
    # =========================================================================
    s21 = prs.slides.add_slide(blank_layout)
    set_bg(s21)
    add_header(s21, "ENGINEERING GOVERNANCE", "Built Around Clarity, Collaboration & Quality", 21)

    principles = [
        ("01 · CLEAR SCOPE & EXPECTATIONS", "No Ambiguity", "Detailed specifications, wireframes, and agreed acceptance criteria up front to ensure everyone is completely aligned before development begins."),
        ("02 · COLLABORATIVE DECISION-MAKING", "Partners, Not Vendors", "Regular sprint check ins, transparent backlog reviews, and collaborative problem solving to adapt dynamically as requirements evolve."),
        ("03 · RIGOROUS QUALITY ASSURANCE", "Zero Tolerance for Flaws", "Automated regression testing, peer code reviews, vulnerability scanning, and manual device testing across iOS, Android, and web."),
        ("04 · TRANSPARENT SPRINT COMMUNICATION", "You're Never Out of the Loop", "Real time updates via dedicated Slack/Teams channels, weekly video check-in calls, and fully documented tasks in Asana/Jira.")
    ]

    for i, (kick, title, desc) in enumerate(principles):
        x = Inches(0.8) + (i % 2) * Inches(5.95)
        y = Inches(1.8) + (i // 2) * Inches(2.55)
        c, tf = create_card(s21, x, y, Inches(5.75), Inches(2.4), is_highlight=False)
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(9.5)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.3

    # =========================================================================
    # SLIDE 22: CLIENT TRUST & PARTNERSHIP (WITH HYFIL GREEN MOCKUP)
    # =========================================================================
    s22 = prs.slides.add_slide(blank_layout)
    set_bg(s22)
    add_header(s22, "CLIENT SATISFACTION", "Building Relationships Beyond Delivery", 22)

    trust_cards = [
        ("98%", "CLIENT RETENTION", "Our clients stay because we deliver consistently, communicate proactively, and constantly hunt for optimizations."),
        ("250+", "DEPLOYMENTS", "Battle-tested experience across complex multi-stakeholder deployments, high-load e-commerce, and SaaS apps."),
        ("25+", "GLOBAL MARKETS", "Cross-cultural expertise serving businesses in the US, Europe, Middle East, Africa, and Asia-Pacific."),
        ("24/7", "SUPPORT AVAILABILITY", "Dedicated engineering and maintenance teams available to resolve critical issues and maintain seamless uptime.")
    ]

    for i, (val, label, desc) in enumerate(trust_cards):
        l = Inches(0.8) + i * Inches(2.98)
        c, tf = create_card(s22, l, Inches(1.8), Inches(2.78), Inches(2.5), is_highlight=(i == 0))
        pv = tf.paragraphs[0]
        pv.text = val
        pv.font.size = Pt(36)
        pv.font.bold = True
        pv.font.color.rgb = COLOR_BRAND_PRIMARY
        pl = tf.add_paragraph()
        pl.text = label
        pl.font.size = Pt(10.5)
        pl.font.bold = True
        pl.font.color.rgb = COLOR_BRAND_NAVY
        pl.space_after = Pt(8)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(10)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.2

    hyfil_mock = get_mockup('mockup_hyfil.png')
    if hyfil_mock:
        s22.shapes.add_picture(hyfil_mock, Inches(0.8), Inches(4.55), width=Inches(3.8))

    bot_card22 = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(4.55), Inches(7.7), Inches(2.3))
    bot_card22.fill.solid()
    bot_card22.fill.fore_color.rgb = COLOR_CARD_BG
    bot_card22.line.color.rgb = COLOR_BRAND_PRIMARY
    bot22_tf = bot_card22.text_frame
    bot22_tf.margin_left = Inches(0.3)
    bot22_tf.margin_top = Inches(0.18)
    bp1 = bot22_tf.paragraphs[0]
    bp1.text = "LONG TERM COLLABORATION & SHARED OUTCOMES"
    bp1.font.size = Pt(10.5)
    bp1.font.bold = True
    bp1.font.color.rgb = COLOR_BRAND_PRIMARY
    bp2 = bot22_tf.add_paragraph()
    bp2.text = "\"Mobrilz took complete ownership of our product architecture from the initial sprint through national deployment. Their engineers are transparent, proactive, and exceptionally skilled.\"\n\nWe measure our success purely by the business impact delivered to our partners, measured in uptime, customer acquisition, and technical resilience."
    bp2.font.size = Pt(11)
    bp2.font.color.rgb = COLOR_BRAND_NAVY
    bp2.line_spacing = 1.25

    # =========================================================================
    # SLIDE 23: READY FOR WHAT IS NEXT
    # =========================================================================
    s23 = prs.slides.add_slide(blank_layout)
    set_bg(s23)
    add_header(s23, "LOOKING FORWARD", "Ready for What Is Next: The Transformation Roadmap", 24)

    steps23 = [
        ("01 · DISCOVER", "Map Your Goals", "A focused strategy session to evaluate your current technology bottlenecks, product roadmap, and business targets."),
        ("02 · ARCHITECT", "Tailored Proposal", "A transparent delivery roadmap with clear architecture specs, milestone timelines, and resource packaging."),
        ("03 · EXECUTE", "Agile Deployment", "Seamless onboarding, dedicated Slack integration, and sprint based execution with working software in weeks.")
    ]

    for i, (kick, title, desc) in enumerate(steps23):
        l = Inches(0.8) + i * Inches(4.0)
        c, tf = create_card(s23, l, Inches(2.0), Inches(3.7), Inches(4.7), is_highlight=(i == 1))
        pk = tf.paragraphs[0]
        pk.text = kick
        pk.font.size = Pt(10)
        pk.font.bold = True
        pk.font.color.rgb = COLOR_BRAND_PRIMARY
        pt = tf.add_paragraph()
        pt.text = title
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_BRAND_NAVY
        pt.space_after = Pt(14)
        pd = tf.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.line_spacing = 1.35

    # =========================================================================
    # SLIDE 24: CONTACT / CALL TO ACTION
    # =========================================================================
    s24 = prs.slides.add_slide(blank_layout)
    set_bg(s24)
    add_header(s24, "START A CONVERSATION", "Let's Build Something Meaningful Together.", 25)

    c_cta, tf_cta = create_card(s24, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.9), is_highlight=True)
    pk_cta = tf_cta.paragraphs[0]
    pk_cta.text = "MOBRILZ PRIVATE LIMITED"
    pk_cta.font.size = Pt(10)
    pk_cta.font.bold = True
    pk_cta.font.color.rgb = COLOR_BRAND_PRIMARY
    
    pt_cta = tf_cta.add_paragraph()
    pt_cta.text = "Have a Project in Mind?\nLet's Discuss Your Vision."
    pt_cta.font.size = Pt(20)
    pt_cta.font.bold = True
    pt_cta.font.color.rgb = COLOR_BRAND_NAVY
    pt_cta.space_after = Pt(14)

    pd_cta = tf_cta.add_paragraph()
    pd_cta.text = "Whether you are looking to build a new mobile app, revamp an enterprise web portal, augment your engineering capacity, or accelerate customer acquisition, we are ready to partner with you.\n\nSchedule a consultation with our technology strategists and receive a tailored implementation plan."
    pd_cta.font.size = Pt(11.5)
    pd_cta.font.color.rgb = COLOR_TEXT_MUTED
    pd_cta.line_spacing = 1.35

    c_info, tf_info = create_card(s24, Inches(6.7), Inches(1.8), Inches(5.8), Inches(4.9))
    pk_info = tf_info.paragraphs[0]
    pk_info.text = "GET IN TOUCH"
    pk_info.font.size = Pt(10)
    pk_info.font.bold = True
    pk_info.font.color.rgb = COLOR_BRAND_PRIMARY
    
    contacts = [
        ("Official Website", "https://www.mobrilz.com/"),
        ("Sales & Inquiries", "sales@mobrilz.com"),
        ("Direct Line", "+91 8010355718"),
        ("Global Footprint", "Head Office: India\nRegional Offices: Dubai (UAE) · Kenya · France · Sénégal"),
        ("Preferred Consultation", "Video Call · Phone Call · On-Site Meeting")
    ]

    for label, val in contacts:
        pl = tf_info.add_paragraph()
        pl.text = label.upper()
        pl.font.size = Pt(9.5)
        pl.font.bold = True
        pl.font.color.rgb = COLOR_BRAND_PRIMARY
        pl.space_before = Pt(8)
        
        pv = tf_info.add_paragraph()
        pv.text = val
        pv.font.size = Pt(11)
        pv.font.color.rgb = COLOR_BRAND_NAVY
        pv.line_spacing = 1.15

    # Save
    prs.save(output_path)
    print(f"Light-themed presentation v2 successfully created with {len(prs.slides)} slides at: {output_path}")

if __name__ == '__main__':
    out_file = r'C:\Users\Manoj\.gemini\antigravity\scratch\Mobrilz_Company_Profile_Presentation_v2.pptx'
    build_presentation(out_file)
