import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck(filename="Operations_and_Human_Factors_in_Cyber_Security.pptx"):
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette Definitions
    DARK_BG = RGBColor(15, 23, 42)        # Slate 900
    DARK_CARD = RGBColor(30, 41, 59)      # Slate 800
    LIGHT_BG = RGBColor(248, 250, 252)    # Slate 50
    WHITE = RGBColor(255, 255, 255)
    TEXT_DARK = RGBColor(15, 23, 42)      # Primary text dark
    TEXT_MUTED = RGBColor(100, 116, 139)  # Slate 500
    TEXT_LIGHT = RGBColor(241, 245, 249)  # Slate 100
    ACCENT_CYAN = RGBColor(14, 165, 233)  # Sky 500
    ACCENT_BLUE = RGBColor(37, 99, 235)   # Blue 600
    ACCENT_RED = RGBColor(225, 29, 72)    # Rose 600
    ACCENT_GREEN = RGBColor(16, 185, 129) # Emerald 500
    ACCENT_AMBER = RGBColor(245, 158, 11) # Amber 500
    ACCENT_PURPLE = RGBColor(147, 51, 234)# Purple 600
    CARD_BG = RGBColor(255, 255, 255)
    CARD_BORDER = RGBColor(226, 232, 240) # Slate 200

    def set_slide_background(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.color.rgb = color
        return bg

    def add_header(slide, title_text, category_text="CYBER SECURITY MODULE • PAGES 165 - 184"):
        # Header banner text
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p_cat = tf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT_CYAN
        p_cat.font.name = "Segoe UI"
        
        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_DARK
        p_title.font.name = "Segoe UI"
        p_title.space_before = Pt(4)

    def add_card(slide, left, top, width, height, border_accent=None, bg_color=CARD_BG):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_accent:
            card.line.color.rgb = border_accent
            card.line.width = Pt(2)
        else:
            card.line.color.rgb = CARD_BORDER
            card.line.width = Pt(1)
        
        # Add colored top strip if border_accent provided
        if border_accent:
            strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.1))
            strip.fill.solid()
            strip.fill.fore_color.rgb = border_accent
            strip.line.color.rgb = border_accent

        return card

    # =========================================================================
    # SLIDE 1: Title Slide (Dark Theme)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, DARK_BG)
    
    # Title badge
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(3.8), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = DARK_CARD
    badge.line.color.rgb = ACCENT_CYAN
    badge.line.width = Pt(1.5)
    tf_b = badge.text_frame
    tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_b.paragraphs[0]
    p.text = "CYBER SECURITY SELF-STUDY • PAGES 165–184"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    p.font.name = "Segoe UI"
    p.alignment = PP_ALIGN.CENTER

    # Main Title
    tbox = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.2))
    tf = tbox.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "Operations & Human Factors"
    p1.font.size = Pt(44)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.font.name = "Segoe UI"

    p2 = tf.add_paragraph()
    p2.text = "Social Engineering, Threat Vectors & Defensive Experiments"
    p2.font.size = Pt(22)
    p2.font.bold = False
    p2.font.color.rgb = RGBColor(148, 163, 184) # Slate 400
    p2.font.name = "Segoe UI"
    p2.space_before = Pt(10)

    p3 = tf.add_paragraph()
    p3.text = "Even robust technical security fails when people are manipulated. This presentation covers core psychological attack vectors, real-world case scenarios, detection hygiene, and 8 structured laboratory experiments."
    p3.font.size = Pt(14)
    p3.font.color.rgb = RGBColor(203, 213, 225)
    p3.font.name = "Segoe UI"
    p3.space_before = Pt(20)

    # 3 Stat Cards on Title Slide
    stat_items = [
        ("Human Vector", "90%+ of breaches", "Triggered by human error & manipulation", ACCENT_CYAN),
        ("Threat Breadth", "6 Attack Vectors", "Phishing, Vishing, Smishing, Tailgating...", ACCENT_AMBER),
        ("Lab Framework", "8 Core Experiments", "Rigorous analysis, procedures & conclusions", ACCENT_GREEN)
    ]
    for idx, (label, val, desc, color) in enumerate(stat_items):
        cx = Inches(1.0 + idx * 3.85)
        cy = Inches(5.4)
        cbox = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, Inches(3.6), Inches(1.3))
        cbox.fill.solid()
        cbox.fill.fore_color.rgb = DARK_CARD
        cbox.line.color.rgb = color
        cbox.line.width = Pt(1.5)
        ctf = cbox.text_frame
        ctf.margin_left = Inches(0.2)
        ctf.margin_top = Inches(0.15)
        
        cp1 = ctf.paragraphs[0]
        cp1.text = label.upper()
        cp1.font.size = Pt(10)
        cp1.font.bold = True
        cp1.font.color.rgb = color
        cp1.font.name = "Segoe UI"

        cp2 = ctf.add_paragraph()
        cp2.text = val
        cp2.font.size = Pt(16)
        cp2.font.bold = True
        cp2.font.color.rgb = WHITE
        cp2.font.name = "Segoe UI"

        cp3 = ctf.add_paragraph()
        cp3.text = desc
        cp3.font.size = Pt(10)
        cp3.font.color.rgb = RGBColor(148, 163, 184)
        cp3.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 2: Agenda & Module Architecture
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, LIGHT_BG)
    add_header(s2, "Module Structure & Learning Roadmap")

    agenda_cards = [
        ("PART 1: ATTACK VECTORS", "Exploiting Human Psychology", [
            "Operations & Human Factors foundations",
            "Social Engineering mechanics & triggers",
            "Phishing, Spear Phishing & targeting",
            "Vishing (Voice) & Smishing (SMS) vectors",
            "Tailgating & physical security bypass"
        ], ACCENT_BLUE, Inches(0.8)),
        ("PART 2: DEFENSE & HYGIENE", "Mitigation & Awareness", [
            "Insider Threats (Malicious vs Negligent)",
            "Enterprise Security Awareness training",
            "Password Hygiene & complexity rules",
            "Phishing email identification markers",
            "Password validation matrix"
        ], ACCENT_PURPLE, Inches(4.8)),
        ("PART 3: LAB EXPERIMENTS", "Hands-on Practical Study", [
            "Experiments 1-3: Social Eng., Phishing, Spear",
            "Experiments 4-5: Vishing/Smishing, Tailgating",
            "Experiment 6: Insider Threat analysis",
            "Experiments 7-8: Awareness & Password Lab",
            "Core conclusions & practical findings"
        ], ACCENT_GREEN, Inches(8.8))
    ]

    for title, subtitle, bullets, color, left in agenda_cards:
        add_card(s2, left, Inches(1.8), Inches(3.733), Inches(5.1), border_accent=color)
        tb = s2.shapes.add_textbox(left + Inches(0.25), Inches(2.0), Inches(3.233), Inches(4.7))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = color
        p.font.name = "Segoe UI"

        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.size = Pt(16)
        p_sub.font.bold = True
        p_sub.font.color.rgb = TEXT_DARK
        p_sub.font.name = "Segoe UI"
        p_sub.space_after = Pt(16)

        for b in bullets:
            bp = tf.add_paragraph()
            bp.text = f"•  {b}"
            bp.font.size = Pt(12)
            bp.font.color.rgb = RGBColor(51, 65, 85)
            bp.font.name = "Segoe UI"
            bp.space_before = Pt(8)

    # =========================================================================
    # SLIDE 3: Operations and Human Factors (Page 164-165)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, LIGHT_BG)
    add_header(s3, "Operations & Human Factors: The Vulnerability Core", "PAGES 164 - 165 • FOUNDATION")

    # Left Card: Core Philosophy
    add_card(s3, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_BLUE)
    tb = s3.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "Why Focus on Human Behavior?"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    bullets_l = [
        ("The Technological Illusion:", "Organizations spend millions on firewalls, encryption, and EDR systems. Yet a single employee clicking an infected link bypasses these controls."),
        ("Hacking the Human:", "Attackers exploit cognitive biases—fear, obedience to authority, curiosity, and urgency—rather than brute-forcing cryptography."),
        ("Core Takeaway:", "'Even the most sophisticated system will collapse if human operators make fatal mistakes.'")
    ]
    for btitle, bdesc in bullets_l:
        p_t = tf.add_paragraph()
        p_t.text = btitle
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = ACCENT_BLUE
        p_t.font.name = "Segoe UI"
        p_t.space_before = Pt(12)

        p_d = tf.add_paragraph()
        p_d.text = bdesc
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = RGBColor(71, 85, 105)
        p_d.font.name = "Segoe UI"

    # Right: The Triad of Breaches
    add_card(s3, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_RED)
    tb_r = s3.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    p = tf_r.paragraphs[0]
    p.text = "The Triad of Breach Drivers (Page 165)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    triad = [
        ("1. Human Error", "Mistakes made under stress, cognitive fatigue, or distraction (e.g., misdirected emails with PII, misconfigured cloud storage S3 buckets).", ACCENT_AMBER),
        ("2. Social Manipulation", "Intentional psychological coercion by adversaries impersonating executives, authorities, or vendor support to extract secrets.", ACCENT_RED),
        ("3. Lack of Awareness", "Inadequate training leaving staff unable to recognize deceptive cues, suspicious domains, spoofed headers, or smishing lures.", ACCENT_PURPLE)
    ]
    for num_t, desc, c in triad:
        p_t = tf_r.add_paragraph()
        p_t.text = num_t
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = c
        p_t.font.name = "Segoe UI"
        p_t.space_before = Pt(12)

        p_d = tf_r.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = RGBColor(71, 85, 105)
        p_d.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 4: Social Engineering Deep Dive (Pages 165-166)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, LIGHT_BG)
    add_header(s4, "Social Engineering: The Mechanics of Deception", "PAGES 165 - 166 • ATTACK MECHANICS")

    cards_se = [
        ("CONCEPT & DEFINITION", "Manipulating Human Trust", [
            ("Definition:", "The technique of psychologically manipulating individuals into revealing confidential data."),
            ("Attacker Mindset:", "'Hacking the human mind' instead of breaking system cryptography."),
            ("Key Triggers:", "Urgency, fear, intimidation, greed, or willingness to help.")
        ], ACCENT_BLUE, Inches(0.8)),
        ("PRIMARY ATTACK METHODS", "Delivery Channels", [
            ("Fake Emails:", "Forged notifications impersonating trusted brands or vendors."),
            ("Fake Phone Calls:", "Urgent voice inquiries posing as fraud specialists or admins."),
            ("Fake Websites:", "Cloned corporate portals designed to harvest login credentials.")
        ], ACCENT_AMBER, Inches(4.8)),
        ("CASE SCENARIO & DEFENSE", "IT Support Impersonation", [
            ("Attack Scenario:", "Attacker calls an employee claiming to be IT Support requiring their password for an urgent patch."),
            ("Rule 1:", "Legitimate IT will NEVER ask for user passwords."),
            ("Rule 2:", "Always verify identity via independent internal callback channels.")
        ], ACCENT_GREEN, Inches(8.8))
    ]

    for title, subtitle, items, color, left in cards_se:
        add_card(s4, left, Inches(1.8), Inches(3.733), Inches(5.1), border_accent=color)
        tb = s4.shapes.add_textbox(left + Inches(0.25), Inches(2.0), Inches(3.233), Inches(4.7))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = color
        p.font.name = "Segoe UI"

        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.size = Pt(16)
        p_sub.font.bold = True
        p_sub.font.color.rgb = TEXT_DARK
        p_sub.font.name = "Segoe UI"
        p_sub.space_after = Pt(12)

        for heading, body in items:
            pt = tf.add_paragraph()
            pt.text = heading
            pt.font.size = Pt(13)
            pt.font.bold = True
            pt.font.color.rgb = color
            pt.font.name = "Segoe UI"
            pt.space_before = Pt(8)

            pb = tf.add_paragraph()
            pb.text = body
            pb.font.size = Pt(12)
            pb.font.color.rgb = RGBColor(71, 85, 105)
            pb.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 5: Phishing & Spear Phishing (Pages 166-167)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, LIGHT_BG)
    add_header(s5, "Phishing vs. Spear Phishing: Broadcast vs. Precision", "PAGES 166 - 167 • EMAIL ATTACKS")

    # Left: Standard Phishing
    add_card(s5, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_BLUE)
    tb_l = s5.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True

    p = tf_l.paragraphs[0]
    p.text = "Standard Phishing (Broadcast)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    phish_details = [
        ("Core Mechanism:", "Mass-distributed fake emails sent indiscriminately to thousands of recipients hoping a small percentage fall for it."),
        ("High-Value Targets:", "Account passwords, bank account numbers, OTPs, and credit card credentials."),
        ("Real-World Example:", "'Your bank account is blocked due to security issues. Click here immediately to login and restore access.'"),
        ("Key Defenses:", "Inspect sender domain, verify SSL certificate, check link destination URLs, deploy anti-phishing gateways.")
    ]
    for title, desc in phish_details:
        pt = tf_l.add_paragraph()
        pt.text = title
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_BLUE
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_l.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Spear Phishing
    add_card(s5, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_RED)
    tb_r = s5.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    p = tf_r.paragraphs[0]
    p.text = "Spear Phishing (Targeted & Personalized)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    spear_details = [
        ("Core Mechanism:", "Carefully customized attack targeting a specific person, department, or organization using prior intelligence."),
        ("Danger Level: HIGH", "Extremely convincing because it leverages the victim's actual name, job role, company projects, and internal terminology."),
        ("Reconnaissance (OSINT):", "Attackers harvest personal data from LinkedIn, public registries, corporate blogs, and social platforms."),
        ("Real-World Example:", "'Hi Kamal, attached is the revised Q4 Cyber audit report for your team. Please review the highlighted figures.'")
    ]
    for title, desc in spear_details:
        pt = tf_r.add_paragraph()
        pt.text = title
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_RED
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_r.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 6: Vishing and Smishing (Pages 168-169)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, LIGHT_BG)
    add_header(s6, "Voice & Mobile Attacks: Vishing & Smishing", "PAGES 168 - 169 • ALTERNATIVE CHANNELS")

    # Left: Vishing
    add_card(s6, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_PURPLE)
    tb_v = s6.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_v = tb_v.text_frame
    tf_v.word_wrap = True

    p = tf_v.paragraphs[0]
    p.text = "Vishing (Voice Phishing)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    vishing_items = [
        ("Definition:", "Phishing executed directly over phone calls or VoIP channels to manipulate victims in real-time."),
        ("Tactic / Modus Operandi:", "Caller ID spoofing makes the call appear to come from the bank or police. Scammers induce panic with claims of compromised accounts."),
        ("Classic Example:", "Caller pretending to be a bank fraud investigator: 'Read back the 6-digit OTP just sent to your phone to cancel the illegal charge.'"),
        ("Key Rule:", "Banks NEVER call asking for OTPs or PINs. Always hang up and dial official numbers from your physical card.")
    ]
    for t, d in vishing_items:
        pt = tf_v.add_paragraph()
        pt.text = t
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_PURPLE
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_v.add_paragraph()
        pd.text = d
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Smishing
    add_card(s6, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_AMBER)
    tb_s = s6.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_s = tb_s.text_frame
    tf_s.word_wrap = True

    p = tf_s.paragraphs[0]
    p.text = "Smishing (SMS Phishing)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    smishing_items = [
        ("Definition:", "Phishing conducted via Short Message Service (SMS) or instant messaging containing dangerous links."),
        ("Tactic / Modus Operandi:", "Exploits the high open-rate of SMS. Embeds shortened URLs (`bit.ly`, `tinyurl`) to hide malicious destinations."),
        ("Classic Example:", "'Your package delivery is on hold due to missing address details. Click link to reschedule: track-parcel-fast.info'"),
        ("Key Rule:", "Never click tracking links inside unprompted SMS messages. Verify parcels directly on official courier apps.")
    ]
    for t, d in smishing_items:
        pt = tf_s.add_paragraph()
        pt.text = t
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_AMBER
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_s.add_paragraph()
        pd.text = d
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 7: Tailgating & Insider Threats (Pages 169-170)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, LIGHT_BG)
    add_header(s7, "Physical Security & Internal Vulnerabilities", "PAGES 169 - 170 • PHYSICAL & INTERNAL")

    # Left: Tailgating
    add_card(s7, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_RED)
    tb_tg = s7.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_tg = tb_tg.text_frame
    tf_tg.word_wrap = True

    p = tf_tg.paragraphs[0]
    p.text = "Tailgating (Piggybacking)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    tailgating_items = [
        ("Physical Security Attack:", "An unauthorized person enters a restricted physical area by closely following an authorized employee through a door."),
        ("Human Exploitation:", "Leverages social politeness (holding doors open for someone holding coffee or carrying heavy boxes)."),
        ("Real Danger:", "Physical access allows attackers to plug in rogue USB devices, compromise server racks, or steal hardware."),
        ("Prevention:", "Access cards, biometric turnstiles, security guards, and mantraps (single-person interlocking doors).")
    ]
    for t, d in tailgating_items:
        pt = tf_tg.add_paragraph()
        pt.text = t
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_RED
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_tg.add_paragraph()
        pd.text = d
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Insider Threat
    add_card(s7, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_CYAN)
    tb_in = s7.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_in = tb_in.text_frame
    tf_in.word_wrap = True

    p = tf_in.paragraphs[0]
    p.text = "Insider Threats (Internal Risks)"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    insider_items = [
        ("Definition:", "A security threat originating from people inside the organization with legitimate access (employees, contractors, vendors)."),
        ("1. Malicious Insider:", "Intentional breach—a disgruntled employee exfiltrating customer records, selling secrets, or sabotaging infrastructure."),
        ("2. Negligent Insider:", "Accidental compromise—careless employees falling for phishing, leaking data, or disabling safety tools."),
        ("Prevention Framework:", "Principle of Least Privilege (PoLP), continuous audit logging, UEBA behavioral monitoring, and swift offboarding.")
    ]
    for t, d in insider_items:
        pt = tf_in.add_paragraph()
        pt.text = t
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_CYAN
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_in.add_paragraph()
        pd.text = d
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 8: Security Awareness & Password Hygiene (Pages 171-172)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8, LIGHT_BG)
    add_header(s8, "Building the Human Firewall: Awareness & Hygiene", "PAGES 171 - 172 • PROACTIVE DEFENSE")

    add_card(s8, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_GREEN)
    tb_aw = s8.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_aw = tb_aw.text_frame
    tf_aw.word_wrap = True

    p = tf_aw.paragraphs[0]
    p.text = "Security Awareness Programs"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    awareness_items = [
        ("Objective:", "Educating all staff on cyber risks, social engineering recognition, and safe daily operational practices."),
        ("Core Curriculum:", "Phishing drills, password safety, safe internet browsing, incident reporting, and data protection policies."),
        ("Measurable Impact:", "Reduces successful phishing click rates by over 70%, fosters an open reporting culture, and reduces human error."),
        ("Golden Rule:", "Security is not just an IT job—it is an organizational responsibility shared by every team member.")
    ]
    for t, d in awareness_items:
        pt = tf_aw.add_paragraph()
        pt.text = t
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_GREEN
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_aw.add_paragraph()
        pd.text = d
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Password Hygiene
    add_card(s8, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_BLUE)
    tb_pw = s8.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_pw = tb_pw.text_frame
    tf_pw.word_wrap = True

    p = tf_pw.paragraphs[0]
    p.text = "Password Hygiene Essentials"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    pw_rules = [
        ("Rule 1: Length & Complexity", "Combine uppercase, lowercase, numbers, and symbols (e.g. 14+ characters minimum)."),
        ("Rule 2: Never Reuse Passwords", "Credential stuffing ensures that one breached account unlocks all your services."),
        ("Rule 3: Use Password Managers", "Avoid writing credentials on sticky notes or spreadsheets; use encrypted vaults."),
        ("Rule 4: Multi-Factor Authentication", "Even if a password is leaked, MFA provides an essential second security barrier.")
    ]
    for t, d in pw_rules:
        pt = tf_pw.add_paragraph()
        pt.text = t
        pt.font.size = Pt(13)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_BLUE
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pd = tf_pw.add_paragraph()
        pd.text = d
        pd.font.size = Pt(12)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 9: Detection & Evaluation Checklist (Page 173)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9, LIGHT_BG)
    add_header(s9, "Practical Evaluation: Phishing & Password Hygiene", "PAGE 173 • OPERATIONAL CHECKLIST")

    # Left: Identify Phishing Email Table
    add_card(s9, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_AMBER)
    tb_t1 = s9.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_t1 = tb_t1.text_frame
    tf_t1.word_wrap = True

    p = tf_t1.paragraphs[0]
    p.text = "Phishing Email Red Flags (Page 173)"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    flags = [
        ("Fake Sender Domain", "bank-security@gmail.com", "Legitimate banks NEVER use public Gmail or Yahoo addresses."),
        ("Artificial Urgency", "'Account will close in 24 hours'", "Pressure forces impulsive reactions before logical thinking kicks in."),
        ("Suspicious Link", "click-here.ru / verify-login.cc", "Domains that don't match the company's authentic website URL."),
        ("Generic Greetings", "'Dear Customer / Valued Member'", "Attacker lacks personalized recipient identification.")
    ]
    for feat, ex, note in flags:
        pt = tf_t1.add_paragraph()
        pt.text = f"{feat}: {ex}"
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = ACCENT_RED
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(8)

        pn = tf_t1.add_paragraph()
        pn.text = note
        pn.font.size = Pt(11)
        pn.font.color.rgb = RGBColor(71, 85, 105)
        pn.font.name = "Segoe UI"

    # Right: Password Hygiene Status Table
    add_card(s9, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_GREEN)
    tb_t2 = s9.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_t2 = tb_t2.text_frame
    tf_t2.word_wrap = True

    p = tf_t2.paragraphs[0]
    p.text = "Password Strength Benchmark (Page 173)"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK
    p.font.name = "Segoe UI"

    pw_bench = [
        ("12345 / 123456", "WEAK ❌", "Instantly cracked in < 0.1 seconds via dictionary attack.", ACCENT_RED),
        ("password / admin", "WEAK ❌", "Most common defaults; first tested by automated botnets.", ACCENT_RED),
        ("Secure@123", "STRONG ✔", "Includes symbols, numbers, mixed casing; higher entropy.", ACCENT_GREEN),
        ("P@ssW0rd!2026", "STRONG ✔", "Exceeds 12 characters, complex character composition.", ACCENT_GREEN),
        ("S3cure#Ops_2026!", "EXCELLENT ★", "Passphrase model; highly resistant to GPU hash crackers.", ACCENT_CYAN)
    ]
    for pw, status, desc, c in pw_bench:
        pt = tf_t2.add_paragraph()
        pt.text = f"{pw}  —  [{status}]"
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = c
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_t2.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 10: Section Divider: Part 2 Laboratory Experiments
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10, DARK_BG)

    badge10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.8), Inches(3.5), Inches(0.45))
    badge10.fill.solid()
    badge10.fill.fore_color.rgb = DARK_CARD
    badge10.line.color.rgb = ACCENT_GREEN
    badge10.line.width = Pt(1.5)
    tf_b10 = badge10.text_frame
    tf_b10.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_b10.paragraphs[0]
    p.text = "HANDS-ON LABORATORY MODULE"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN
    p.font.name = "Segoe UI"
    p.alignment = PP_ALIGN.CENTER

    tbox10 = s10.shapes.add_textbox(Inches(1.0), Inches(2.5), Inches(11.333), Inches(3.0))
    tf10 = tbox10.text_frame
    tf10.word_wrap = True

    p = tf10.paragraphs[0]
    p.text = "Practical Cybersecurity Experiments"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = "Segoe UI"

    p_sub = tf10.add_paragraph()
    p_sub.text = "Analysis of Experiments 1 through 8 (Pages 174–184)"
    p_sub.font.size = Pt(22)
    p_sub.font.bold = False
    p_sub.font.color.rgb = RGBColor(148, 163, 184)
    p_sub.font.name = "Segoe UI"
    p_sub.space_before = Pt(8)

    p_desc = tf10.add_paragraph()
    p_desc.text = "Each experiment establishes a formal security study: defining the Aim, theoretical foundation, standardized procedure, observable results, and key conclusions."
    p_desc.font.size = Pt(14)
    p_desc.font.color.rgb = RGBColor(203, 213, 225)
    p_desc.font.name = "Segoe UI"
    p_desc.space_before = Pt(16)

    # 4 Quick experiment categories
    lab_chips = [
        ("Exp 1 & 2", "Social Eng. & Phishing"),
        ("Exp 3 & 4", "Spear Phishing & Vishing/Smishing"),
        ("Exp 5 & 6", "Tailgating & Insider Threat"),
        ("Exp 7 & 8", "Security Awareness & Password Lab")
    ]
    for idx, (code, title) in enumerate(lab_chips):
        chip = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0 + idx * 2.85), Inches(5.6), Inches(2.7), Inches(1.0))
        chip.fill.solid()
        chip.fill.fore_color.rgb = DARK_CARD
        chip.line.color.rgb = ACCENT_CYAN
        chip.line.width = Pt(1)
        ctf = chip.text_frame
        ctf.margin_top = Inches(0.12)
        cp1 = ctf.paragraphs[0]
        cp1.text = code
        cp1.font.size = Pt(13)
        cp1.font.bold = True
        cp1.font.color.rgb = ACCENT_CYAN
        cp1.font.name = "Segoe UI"
        cp2 = ctf.add_paragraph()
        cp2.text = title
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = WHITE
        cp2.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 11: Experiments 1 & 2 (Pages 174-176)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11, LIGHT_BG)
    add_header(s11, "Lab Studies 1 & 2: Social Engineering & Phishing", "PAGES 174 - 176 • EXPERIMENTS 1 & 2")

    # Left: Experiment 1
    add_card(s11, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_BLUE)
    tb_e1 = s11.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_e1 = tb_e1.text_frame
    tf_e1.word_wrap = True

    p = tf_e1.paragraphs[0]
    p.text = "EXPERIMENT 1: SOCIAL ENGINEERING"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE
    p.font.name = "Segoe UI"

    e1_details = [
        ("Aim:", "To study social engineering attacks in cyber security."),
        ("Theory:", "Attackers trick humans into revealing confidential information by exploiting psychological vulnerabilities rather than technical flaws."),
        ("Procedure:", "1. Study common human-based attacks.\n2. Identify how attackers manipulate users.\n3. List prevention and defense methods."),
        ("Result:", "Social engineering concepts and manipulation vectors are understood."),
        ("Conclusion:", "Humans are the weakest link in security systems.")
    ]
    for lbl, text in e1_details:
        pt = tf_e1.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e1.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Experiment 2
    add_card(s11, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_RED)
    tb_e2 = s11.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_e2 = tb_e2.text_frame
    tf_e2.word_wrap = True

    p = tf_e2.paragraphs[0]
    p.text = "EXPERIMENT 2: PHISHING ATTACKS"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_RED
    p.font.name = "Segoe UI"

    e2_details = [
        ("Aim:", "To study phishing attacks and data theft workflows."),
        ("Theory & Example:", "Fake emails or messages steal sensitive data.\nExample: 'Your bank account is blocked, click here to login.'"),
        ("Procedure:", "1. Observe fake email structure & headers.\n2. Identify suspicious hyperlinks & fake domains.\n3. Understand the data theft process."),
        ("Result:", "Phishing attack lifecycle is analyzed and understood."),
        ("Conclusion:", "Phishing is primarily utilized to steal passwords and sensitive personal credentials.")
    ]
    for lbl, text in e2_details:
        pt = tf_e2.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e2.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 12: Experiments 3 & 4 (Pages 176-179)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12, LIGHT_BG)
    add_header(s12, "Lab Studies 3 & 4: Spear Phishing, Vishing & Smishing", "PAGES 176 - 179 • EXPERIMENTS 3 & 4")

    # Left: Experiment 3
    add_card(s12, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_PURPLE)
    tb_e3 = s12.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_e3 = tb_e3.text_frame
    tf_e3.word_wrap = True

    p = tf_e3.paragraphs[0]
    p.text = "EXPERIMENT 3: SPEAR PHISHING"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_PURPLE
    p.font.name = "Segoe UI"

    e3_details = [
        ("Aim:", "To study targeted spear phishing attacks."),
        ("Theory:", "Spear phishing is a highly personalized attack directed against specific individuals or organizations using gathered intelligence."),
        ("Procedure:", "1. Identify targeted email communications.\n2. Check personalized content & tailored lures.\n3. Analyze attack methods and vectors."),
        ("Result:", "Spear phishing concept and reconnaissance mechanics understood."),
        ("Conclusion:", "Spear phishing is significantly more dangerous and harder to detect than generic phishing.")
    ]
    for lbl, text in e3_details:
        pt = tf_e3.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e3.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Experiment 4
    add_card(s12, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_AMBER)
    tb_e4 = s12.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_e4 = tb_e4.text_frame
    tf_e4.word_wrap = True

    p = tf_e4.paragraphs[0]
    p.text = "EXPERIMENT 4: VISHING & SMISHING"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_AMBER
    p.font.name = "Segoe UI"

    e4_details = [
        ("Aim:", "To study voice and SMS phishing attacks."),
        ("Theory & Examples:", "• Vishing = Voice call phishing (fake bank call asking OTP)\n• Smishing = SMS text phishing (fake package delay link)"),
        ("Procedure:", "1. Study call-based voice frauds & spoofing.\n2. Study SMS-based link redirection frauds.\n3. Identify common warning signs & urgency triggers."),
        ("Result:", "Vishing and smishing mechanics are analyzed."),
        ("Conclusion:", "Both attacks directly exploit human trust and spontaneous mobile reactions.")
    ]
    for lbl, text in e4_details:
        pt = tf_e4.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e4.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 13: Experiments 5 & 6 (Pages 179-181)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13, LIGHT_BG)
    add_header(s13, "Lab Studies 5 & 6: Tailgating & Insider Threats", "PAGES 179 - 181 • EXPERIMENTS 5 & 6")

    # Left: Experiment 5
    add_card(s13, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_RED)
    tb_e5 = s13.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_e5 = tb_e5.text_frame
    tf_e5.word_wrap = True

    p = tf_e5.paragraphs[0]
    p.text = "EXPERIMENT 5: STUDY OF TAILGATING"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_RED
    p.font.name = "Segoe UI"

    e5_details = [
        ("Aim:", "To study tailgating attacks in physical security environments."),
        ("Theory:", "Tailgating is a physical attack where an unauthorized person enters a restricted zone by following an authorized individual."),
        ("Procedure:", "1. Observe facility entry control systems.\n2. Identify unauthorized entry attempts & piggybacking.\n3. Study physical prevention methods."),
        ("Result:", "Tailgating vulnerabilities and access dynamics understood."),
        ("Conclusion:", "Physical security is foundational; cybersecurity fails if physical perimeters are breached.")
    ]
    for lbl, text in e5_details:
        pt = tf_e5.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e5.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Experiment 6
    add_card(s13, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_CYAN)
    tb_e6 = s13.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_e6 = tb_e6.text_frame
    tf_e6.word_wrap = True

    p = tf_e6.paragraphs[0]
    p.text = "EXPERIMENT 6: INSIDER THREATS"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    p.font.name = "Segoe UI"

    e6_details = [
        ("Aim:", "To study insider threats within organizations."),
        ("Theory & Types:", "Threat arises from employees/trusted users who misuse access.\n• Malicious Insider (intentional data theft)\n• Negligent Insider (careless errors)"),
        ("Procedure:", "1. Study user roles and privilege boundaries.\n2. Identify data misuse & exfiltration scenarios.\n3. Analyze risk matrices and monitoring policies."),
        ("Result:", "Insider threat risk patterns analyzed."),
        ("Conclusion:", "Internal trusted users can pose equal or greater risk than external hackers.")
    ]
    for lbl, text in e6_details:
        pt = tf_e6.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e6.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 14: Experiments 7 & 8 (Pages 182-184)
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14, LIGHT_BG)
    add_header(s14, "Lab Studies 7 & 8: Security Awareness & Password Hygiene", "PAGES 182 - 184 • EXPERIMENTS 7 & 8")

    # Left: Experiment 7
    add_card(s14, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.1), border_accent=ACCENT_GREEN)
    tb_e7 = s14.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.5))
    tf_e7 = tb_e7.text_frame
    tf_e7.word_wrap = True

    p = tf_e7.paragraphs[0]
    p.text = "EXPERIMENT 7: SECURITY AWARENESS"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_GREEN
    p.font.name = "Segoe UI"

    e7_details = [
        ("Aim:", "To study security awareness programs and training modules."),
        ("Theory:", "Security awareness trains users to recognize, evade, and report cyber attacks before compromise occurs."),
        ("Topics Covered:", "• Phishing identification drills\n• Password hygiene rules\n• Safe internet navigation"),
        ("Procedure:", "1. Study training materials and interactive simulations.\n2. Observe user behavior improvements & reporting rates."),
        ("Result & Conclusion:", "Result: Training framework understood.\nConclusion: Continuous awareness substantially reduces cyber incidents.")
    ]
    for lbl, text in e7_details:
        pt = tf_e7.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e7.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # Right: Experiment 8
    add_card(s14, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1), border_accent=ACCENT_BLUE)
    tb_e8 = s14.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.1), Inches(4.5))
    tf_e8 = tb_e8.text_frame
    tf_e8.word_wrap = True

    p = tf_e8.paragraphs[0]
    p.text = "EXPERIMENT 8: PASSWORD HYGIENE"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE
    p.font.name = "Segoe UI"

    e8_details = [
        ("Aim:", "To study password hygiene practices and entropy benchmarks."),
        ("Theory & Rules:", "Strong password practices defend against automated attacks:\n• Use strong, complex passphrases\n• Never reuse credentials\n• Rotate credentials systematically"),
        ("Comparative Example:", "❌ Insecure: 123456\n✔ Secure: S3cure@2026"),
        ("Procedure:", "1. Create sample passwords with varying entropy.\n2. Check strength level against cracking tools."),
        ("Result & Conclusion:", "Result: Password policies evaluated.\nConclusion: Strong passwords drastically harden systems.")
    ]
    for lbl, text in e8_details:
        pt = tf_e8.add_paragraph()
        pt.text = lbl
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_DARK
        pt.font.name = "Segoe UI"
        pt.space_before = Pt(6)

        pd = tf_e8.add_paragraph()
        pd.text = text
        pd.font.size = Pt(11)
        pd.font.color.rgb = RGBColor(71, 85, 105)
        pd.font.name = "Segoe UI"

    # =========================================================================
    # SLIDE 15: Laboratory Summary Matrix
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15, LIGHT_BG)
    add_header(s15, "Master Summary: Experiments 1 through 8", "LABORATORY SYNTHESIS MATRIX")

    # Table of all 8 experiments
    rows, cols = 9, 4
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(11.733)
    height = Inches(5.1)

    table_shape = s15.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # Column widths
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(3.2)
    table.columns[2].width = Inches(3.2)
    table.columns[3].width = Inches(3.533)

    headers = ["Experiment #", "Domain / Focus", "Key Attack / Mechanism", "Core Lab Conclusion"]
    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = h_text
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK_CARD
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(11)
            p.font.color.rgb = WHITE
            p.font.name = "Segoe UI"

    exp_rows = [
        ("Exp 1 (p. 174)", "Social Engineering", "Human psychological manipulation", "Humans are the weakest link in security"),
        ("Exp 2 (p. 175)", "Phishing Attacks", "Fake emails & credential harvesting", "Primary tool for password and identity theft"),
        ("Exp 3 (p. 176)", "Spear Phishing", "Personalized OSINT targeted emails", "Significantly more dangerous than generic spam"),
        ("Exp 4 (p. 178)", "Vishing & Smishing", "Voice call fraud & malicious SMS links", "Attacks directly exploit immediate human trust"),
        ("Exp 5 (p. 179)", "Tailgating Attack", "Piggybacking through secured doors", "Physical security is vital for digital defense"),
        ("Exp 6 (p. 180)", "Insider Threat", "Malicious sabotage or negligent errors", "Internal users can be as dangerous as hackers"),
        ("Exp 7 (p. 182)", "Security Awareness", "User training, phishing simulations", "Proactive education drastically reduces breaches"),
        ("Exp 8 (p. 183)", "Password Hygiene", "Entropy, uniqueness, password managers", "Strong credential hygiene prevents brute-forcing")
    ]

    for row_idx, data in enumerate(exp_rows, start=1):
        bg = RGBColor(255, 255, 255) if row_idx % 2 != 0 else RGBColor(241, 245, 249)
        for col_idx, text in enumerate(data):
            cell = table.cell(row_idx, col_idx)
            cell.text = text
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10)
                p.font.name = "Segoe UI"
                p.font.color.rgb = TEXT_DARK
                if col_idx == 0:
                    p.font.bold = True

    # =========================================================================
    # SLIDE 16: Defense-in-Depth Strategic Framework
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16, LIGHT_BG)
    add_header(s16, "Comprehensive Defense-in-Depth Strategy", "ARCHITECTURAL BLUEPRINT")

    pillars = [
        ("PEOPLE LAYER", "The Human Firewall", [
            "Continuous awareness training & simulations",
            "Clear, blame-free incident reporting culture",
            "Strict identity verification for voice requests",
            "Anti-tailgating culture ('Badge everyone')"
        ], ACCENT_GREEN, Inches(0.8)),
        ("PROCESS LAYER", "Operational Governance", [
            "Principle of Least Privilege (PoLP)",
            "Mandatory MFA on all accounts & endpoints",
            "Dual-authorization for financial wire transfers",
            "Automated employee offboarding protocols"
        ], ACCENT_BLUE, Inches(4.8)),
        ("TECHNOLOGY LAYER", "Automated Safeguards", [
            "SPF, DKIM, DMARC email authentication",
            "Endpoint Detection & Response (EDR)",
            "Data Loss Prevention (DLP) & monitoring",
            "Biometric mantraps & physical access gates"
        ], ACCENT_PURPLE, Inches(8.8))
    ]

    for title, subtitle, bullets, color, left in pillars:
        add_card(s16, left, Inches(1.8), Inches(3.733), Inches(5.1), border_accent=color)
        tb = s16.shapes.add_textbox(left + Inches(0.25), Inches(2.0), Inches(3.233), Inches(4.7))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = color
        p.font.name = "Segoe UI"

        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.size = Pt(16)
        p_sub.font.bold = True
        p_sub.font.color.rgb = TEXT_DARK
        p_sub.font.name = "Segoe UI"
        p_sub.space_after = Pt(16)

        for b in bullets:
            bp = tf.add_paragraph()
            bp.text = f"✔  {b}"
            bp.font.size = Pt(12)
            bp.font.color.rgb = RGBColor(51, 65, 85)
            bp.font.name = "Segoe UI"
            bp.space_before = Pt(10)

    # =========================================================================
    # SLIDE 17: Conclusion & Key Takeaways (Dark Theme)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_background(s17, DARK_BG)

    badge17 = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(3.2), Inches(0.45))
    badge17.fill.solid()
    badge17.fill.fore_color.rgb = DARK_CARD
    badge17.line.color.rgb = ACCENT_CYAN
    badge17.line.width = Pt(1.5)
    tf_b17 = badge17.text_frame
    tf_b17.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_b17.paragraphs[0]
    p.text = "MODULE CONCLUSION & SUMMARY"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    p.font.name = "Segoe UI"
    p.alignment = PP_ALIGN.CENTER

    tbox17 = s17.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(1.5))
    tf17 = tbox17.text_frame
    tf17.word_wrap = True
    p1 = tf17.paragraphs[0]
    p1.text = "Key Takeaways & Presentation Summary"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.font.name = "Segoe UI"

    takeaways = [
        ("1. The Human Factor Dominates", "Over 90% of security incidents involve human manipulation. Firewalls cannot defend against an employee willingly handing over an OTP or password.", ACCENT_AMBER),
        ("2. Multi-Vector Vigilance", "Adversaries pivot across email (phishing), phone (vishing), SMS (smishing), physical doors (tailgating), and internal access (insider threats).", ACCENT_CYAN),
        ("3. Culture Is the Strongest Defense", "Technical controls must be reinforced by continuous security awareness, strong password hygiene, zero trust, and a vigilant reporting culture.", ACCENT_GREEN)
    ]

    for idx, (head, body, color) in enumerate(takeaways):
        cx = Inches(1.0 + idx * 3.85)
        cy = Inches(3.6)
        cbox = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, Inches(3.6), Inches(2.9))
        cbox.fill.solid()
        cbox.fill.fore_color.rgb = DARK_CARD
        cbox.line.color.rgb = color
        cbox.line.width = Pt(1.5)
        ctf = cbox.text_frame
        ctf.margin_left = Inches(0.25)
        ctf.margin_top = Inches(0.25)
        ctf.margin_right = Inches(0.25)
        ctf.word_wrap = True

        cp1 = ctf.paragraphs[0]
        cp1.text = head
        cp1.font.size = Pt(15)
        cp1.font.bold = True
        cp1.font.color.rgb = color
        cp1.font.name = "Segoe UI"

        cp2 = ctf.add_paragraph()
        cp2.text = body
        cp2.font.size = Pt(12)
        cp2.font.color.rgb = RGBColor(203, 213, 225)
        cp2.font.name = "Segoe UI"
        cp2.space_before = Pt(10)

    prs.save(filename)
    print(f"Presentation successfully created and saved to: {os.path.abspath(filename)}")

if __name__ == "__main__":
    create_deck()
