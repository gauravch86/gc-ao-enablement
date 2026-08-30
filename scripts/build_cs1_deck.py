#!/usr/bin/env python3
"""Build Case Study 1 executive presentation (PPTX) — ~10 min storyline."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from pptx.oxml import parse_xml
from copy import deepcopy
from lxml import etree

# Brand palette (aligned with case-study HTML)
BG = RGBColor(0x0F, 0x14, 0x19)
SURFACE = RGBColor(0x1A, 0x23, 0x32)
SURFACE2 = RGBColor(0x24, 0x30, 0x44)
TEXT = RGBColor(0xE6, 0xED, 0xF3)
MUTED = RGBColor(0x8B, 0x9C, 0xB3)
ACCENT = RGBColor(0x58, 0xA6, 0xFF)
GOLD = RGBColor(0xD4, 0xA8, 0x53)
SUCCESS = RGBColor(0x3F, 0xB9, 0x50)
DANGER = RGBColor(0xF8, 0x71, 0x71)
PURPLE = RGBColor(0xA3, 0x71, 0xF7)

W = Inches(13.333)
H = Inches(7.5)


def set_run(run, size=14, bold=False, color=TEXT, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_text(tf, text, size=14, bold=False, color=TEXT, align=PP_ALIGN.LEFT, space_after=6):
    p = tf.paragraphs[0] if not tf.paragraphs[0].text else tf.add_paragraph()
    if not tf.paragraphs[0].text and len(tf.paragraphs) == 1:
        p = tf.paragraphs[0]
    else:
        # ensure we use empty first para if blank
        if tf.paragraphs[0].text == "" and len([x for x in tf.paragraphs if x.text]) == 0:
            p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def clear_tf(tf):
    tf.clear()
    # leave one empty paragraph
    p = tf.paragraphs[0]
    p.text = ""
    return tf


def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def rect(slide, left, top, width, height, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    fill_shape(s, fill)
    # softer corners
    try:
        s.adjustments[0] = 0.08
    except Exception:
        pass
    return s


def textbox(slide, left, top, width, height):
    return slide.shapes.add_textbox(left, top, width, height)


def add_para(tf, text, size=13, bold=False, color=TEXT, align=PP_ALIGN.LEFT, space_after=4, space_before=0):
    # Use first empty paragraph if available
    if tf.paragraphs[0].text == "" and all(not p.text for p in tf.paragraphs):
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    # full background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    fill_shape(bg, BG)
    # send to back
    spTree = slide.shapes._spTree
    sp = bg._element
    spTree.remove(sp)
    spTree.insert(2, sp)
    return slide


def label(slide, text, top=Inches(0.28)):
    tb = textbox(slide, Inches(0.55), top, Inches(12), Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, text.upper(), size=11, bold=True, color=ACCENT, space_after=0)


def title(slide, text, top=Inches(0.52)):
    tb = textbox(slide, Inches(0.55), top, Inches(12.2), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, text, size=28, bold=True, color=GOLD, space_after=0)


def footer(slide, page, total=6):
    tb = textbox(slide, Inches(0.55), Inches(7.1), Inches(10), Inches(0.28))
    tf = tb.text_frame
    add_para(tf, "Case Study 1 · Reposition underperforming offering · Gaurav Chaudhary", size=10, color=MUTED, space_after=0)
    tb2 = textbox(slide, Inches(11.6), Inches(7.1), Inches(1.2), Inches(0.28))
    tf2 = tb2.text_frame
    add_para(tf2, f"{page} / {total}", size=10, color=MUTED, align=PP_ALIGN.RIGHT, space_after=0)


def metric_card(slide, left, top, width, height, lbl, val, sub, val_color=GOLD):
    card = rect(slide, left, top, width, height, SURFACE)
    tb = textbox(slide, left + Inches(0.15), top + Inches(0.12), width - Inches(0.3), height - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, lbl.upper(), size=10, bold=True, color=MUTED, align=PP_ALIGN.CENTER, space_after=2)
    add_para(tf, val, size=26, bold=True, color=val_color, align=PP_ALIGN.CENTER, space_after=2)
    add_para(tf, sub, size=11, color=MUTED, align=PP_ALIGN.CENTER, space_after=0)


def card_block(slide, left, top, width, height, heading, lines, accent=GOLD):
    rect(slide, left, top, width, height, SURFACE)
    # accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), height)
    fill_shape(bar, accent)
    tb = textbox(slide, left + Inches(0.22), top + Inches(0.12), width - Inches(0.35), height - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, heading, size=14, bold=True, color=GOLD, space_after=6)
    for line in lines:
        add_para(tf, "•  " + line, size=12, color=TEXT, space_after=3)


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # ─── SLIDE 1: Situation ───────────────────────────────────────────
    s = blank_slide(prs)
    label(s, "Strategic Product Manager  ·  Case Study 1  ·  10-min brief")
    title(s, "Reposition an underperforming portfolio offering")

    tb = textbox(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Application Managed Services (AMS): strong technical delivery, but commercially stuck — "
        "flat net sales, declining margins, Sales cannot articulate value, regional pricing chaos.",
        size=15,
        color=MUTED,
        space_after=0,
    )

    metric_card(s, Inches(0.55), Inches(2.15), Inches(3.9), Inches(1.55),
                "Net sales (2 yrs)", "Flat", "Weak conversion · no attach", MUTED)
    metric_card(s, Inches(4.7), Inches(2.15), Inches(3.9), Inches(1.55),
                "Gross margin", "−4–6 pts", "Discount · scope creep · KT bleed", DANGER)
    metric_card(s, Inches(8.85), Inches(2.15), Inches(3.9), Inches(1.55),
                "Delivery NPS / C-SAT", "Strong", "Execution is not the problem", SUCCESS)

    # Hypothesis box
    rect(s, Inches(0.55), Inches(4.0), Inches(12.2), Inches(1.35), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(4.15), Inches(11.8), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "WORKING HYPOTHESIS", size=11, bold=True, color=ACCENT, space_after=4)
    add_para(
        tf,
        "Commercially mis-positioned — sold as FTE headcount while the market buys outcomes, "
        "autonomy, and TCO reduction. Delivery excellence exists; the product story, packaging, "
        "and pricing model do not.",
        size=14,
        color=TEXT,
        space_after=0,
    )

    tb = textbox(s, Inches(0.55), Inches(5.55), Inches(12.2), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "AGENDA  ·  10 minutes", size=11, bold=True, color=GOLD, space_after=6)
    add_para(
        tf,
        "1  Diagnose   →   2  Decide (evolve / reposition / sunset)   →   3  Package & price   →   "
        "4  Prioritize & align   →   5  Measure & close",
        size=14,
        color=TEXT,
        space_after=8,
    )
    add_para(
        tf,
        "Proof anchors:  H3G UK  $125m packaging & economics   ·   TEF Argentina  $100m renewal reposition",
        size=12,
        color=MUTED,
        space_after=0,
    )
    footer(s, 1)

    # ─── SLIDE 2: Diagnosis → Decision ────────────────────────────────
    s = blank_slide(prs)
    label(s, "Diagnosis framework  ·  Recommendation")
    title(s, "Strong delivery. Broken commercial model. → Reposition.")

    # Five lenses compact
    lenses = [
        ("Win / Loss", "15–20 deals · tag\nprice vs value vs unclear"),
        ("Margin", "GM waterfall · discount\n· KT overrun · CR leak"),
        ("Competitive", "FTE SI vs outcome /\nAO product MS shift"),
        ("Voice", "NPS high + willingness\nto pay low = story gap"),
        ("Portfolio", "18-mo GM path?\nAttach to AO / products?"),
    ]
    for i, (h, body) in enumerate(lenses):
        left = Inches(0.55) + i * Inches(2.5)
        rect(s, left, Inches(1.35), Inches(2.35), Inches(1.55), SURFACE)
        tb = textbox(s, left + Inches(0.12), Inches(1.45), Inches(2.1), Inches(1.35))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h.upper(), size=11, bold=True, color=ACCENT, space_after=4)
        for line in body.split("\n"):
            add_para(tf, line, size=11, color=TEXT, space_after=1)

    # Root causes
    causes = [
        ("Commoditized packaging", "Sold as L1/L2/L3 FTE towers — procurement benches us vs SIs"),
        ("Pricing chaos", "Regional ad-hoc discounts · underpriced KT · unpaid scope creep"),
        ("Value story gap", "Buyers want TCO & autonomy path; we pitch headcount replacement"),
    ]
    for i, (h, body) in enumerate(causes):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(3.15), Inches(3.95), Inches(1.35), SURFACE)
        bar = slide_bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(3.15), Inches(3.95), Inches(0.06))
        fill_shape(slide_bar, DANGER)
        tb = textbox(s, left + Inches(0.15), Inches(3.3), Inches(3.65), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=13, bold=True, color=DANGER, space_after=4)
        add_para(tf, body, size=12, color=TEXT, space_after=0)

    # Decision row
    decisions = [
        ("SUNSET", "Rejects delivery DNA &\nEricsson footprint.\nOnly if no 18-mo GM path.", MUTED, False),
        ("EVOLVE ONLY", "Tooling without SKU change\nfixes neither sales narrative\nnor margin guardrails.", MUTED, False),
        ("REPOSITION  ✓", "Outcome-tiered Application Ops\nStabilize → Optimize → Autonomize\nEvolve modules underneath.", SUCCESS, True),
    ]
    for i, (h, body, color, selected) in enumerate(decisions):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(4.75), Inches(3.95), Inches(1.85), SURFACE)
        if selected:
            border = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(4.75), Inches(3.95), Inches(1.85))
            border.fill.background()
            border.line.color.rgb = SUCCESS
            border.line.width = Pt(2)
            try:
                border.adjustments[0] = 0.08
            except Exception:
                pass
        tb = textbox(s, left + Inches(0.18), Inches(4.9), Inches(3.6), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=14, bold=True, color=color, space_after=6)
        for line in body.split("\n"):
            add_para(tf, line, size=12, color=TEXT if selected else MUTED, space_after=1)

    footer(s, 2)

    # ─── SLIDE 3: Packaging & Pricing ─────────────────────────────────
    s = blank_slide(prs)
    label(s, "Revised pricing & packaging  ·  Value proposition")
    title(s, "Sell tiers & outcomes — not FTE towers")

    tb = textbox(s, Inches(0.55), Inches(1.2), Inches(12.2), Inches(0.45))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Value prop:  Outcome-guaranteed Application Operations with a built-in ramp to autonomy — "
        "reducing customer TCO while protecting margin through standardized tiers & global guardrails.",
        size=13,
        color=MUTED,
        space_after=0,
    )

    # Tier table as cards
    tiers = [
        ("STABILIZE", GOLD,
         "Restore control",
         ["Backlog burn · SLA baseline", "Single accountability · CPI board", "Top-offender program",
          "Price: platform + app-class", "Floor GM ≥ 22%"]),
        ("OPTIMIZE", ACCENT,
         "TCO & MTTR proof",
         ["+ Unified observability / SPOG", "+ GenAI SOP · automation factory", "Outcome-band pricing",
          "Gain-share on TCO proof", "Floor GM ≥ 26%"]),
        ("AUTONOMIZE", SUCCESS,
         "Path to L3–L4 autonomy",
         ["+ Closed-loop · policy packs", "+ Soft-ramp CPI contract", "Milestone + performance",
          "On-ramp to reusable AO SKU", "Floor GM ≥ 30%"]),
    ]
    for i, (name, color, tagline, lines) in enumerate(tiers):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(1.8), Inches(3.95), Inches(3.55), SURFACE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.8), Inches(3.95), Inches(0.08))
        fill_shape(bar, color)
        tb = textbox(s, left + Inches(0.2), Inches(2.0), Inches(3.55), Inches(3.2))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, name, size=16, bold=True, color=color, space_after=2)
        add_para(tf, tagline, size=12, bold=True, color=GOLD, space_after=8)
        for line in lines:
            add_para(tf, "•  " + line, size=12, color=TEXT, space_after=4)

    # Attach packs strip
    rect(s, Inches(0.55), Inches(5.55), Inches(12.2), Inches(1.15), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(5.65), Inches(11.8), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "MODULAR ATTACH  ·  priced, not free", size=11, bold=True, color=ACCENT, space_after=4)
    add_para(
        tf,
        "Observability jump-start (90d)   ·   Top-offender sprint (6mo)   ·   Transition accelerator (mandatory KT)   ·   "
        "Vendor-consolidation bridge   ·   FTE sizing stays internal — customer buys service class & CPI attainment",
        size=13,
        color=TEXT,
        space_after=0,
    )
    footer(s, 3)

    # ─── SLIDE 4: Guardrails + Prioritization ─────────────────────────
    s = blank_slide(prs)
    label(s, "Commercial guardrails  ·  Investment sequencing")
    title(s, "Deal Desk rules that stop the bleed — fund ROI first")

    # Guardrails left
    rect(s, Inches(0.55), Inches(1.3), Inches(6.3), Inches(5.4), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.45), Inches(5.9), Inches(5.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "PRICING GUARDRAILS", size=13, bold=True, color=GOLD, space_after=10)
    rules = [
        ("Floor GM%", "Stabilize ≥22% · Optimize ≥26% · Autonomize ≥30%"),
        ("Max discount", "≤12% off list (≤8% renewals) — trade for scope / term / step-up"),
        ("Transition", "KT module mandatory on new logos · min 8% of Y1 ACV"),
        ("Scope CR", "Unpaid work cap 2% ACV · 5-day CR cycle"),
        ("Regional variance", "Like-for-like within ±15% of global rate card"),
        ("CPI / penalty", "Soft-ramp M1–6 · credit caps · Finance rev-rec sign-off"),
    ]
    for name, detail in rules:
        add_para(tf, name, size=13, bold=True, color=ACCENT, space_after=1)
        add_para(tf, detail, size=12, color=TEXT, space_after=8)

    # Priority right
    rect(s, Inches(7.05), Inches(1.3), Inches(5.7), Inches(3.15), SURFACE)
    tb = textbox(s, Inches(7.25), Inches(1.45), Inches(5.3), Inches(2.9))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "FUND  ·  HIGH ROI", size=13, bold=True, color=SUCCESS, space_after=8)
    for line in [
        "Q1: Global rate card · tier SKUs · ROI calc · Deal Desk rules",
        "Q2: CPI catalogue · transition factory · 2 regional pilots",
        "Q3–4: Observability pack · SPOG · automation library · renewals",
        "Y2: Autonomize scale → merge with reusable AO offering",
    ]:
        add_para(tf, "▸  " + line, size=12, color=TEXT, space_after=6)

    rect(s, Inches(7.05), Inches(4.65), Inches(5.7), Inches(2.05), SURFACE)
    tb = textbox(s, Inches(7.25), Inches(4.8), Inches(5.3), Inches(1.8))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "STOP / DEPRIORITIZE", size=13, bold=True, color=DANGER, space_after=8)
    for line in [
        "Per-account custom observability builds",
        "Below-floor PoCs to “win logo”",
        "Regional bespoke rate cards",
        "Rebrand without sales / Deal Desk kit",
    ]:
        add_para(tf, "×  " + line, size=12, color=TEXT, space_after=4)

    footer(s, 4)

    # ─── SLIDE 5: Alignment + Metrics ─────────────────────────────────
    s = blank_slide(prs)
    label(s, "Cross-functional alignment  ·  Success metrics & feedback loop")
    title(s, "One program — Sales, Marketing, Finance, Regions")

    funcs = [
        ("SALES / DEAL DESK", ACCENT,
         "≥80% bids on tier SKU by Q2",
         "Discovery kit · ROI · margin waterfall · weekly pilot deal review"),
        ("MARKETING", PURPLE,
         "One narrative: Apps Ops → Autonomy",
         "Tier one-pagers · proof stories · retire FTE-tower language"),
        ("FINANCE", GOLD,
         "Rev-rec compliant · GM floors in ERP",
         "Platform vs milestone memo · penalty lot accounting · bid sign-off"),
        ("REGIONS / DELIVERY", SUCCESS,
         "Pilots hit CPI catalogue",
         "KT playbook · SDM board · QBR tier step-ups · train-the-trainer"),
    ]
    for i, (name, color, outcome, artifacts) in enumerate(funcs):
        left = Inches(0.55) + (i % 4) * Inches(3.15)
        top = Inches(1.3)
        rect(s, left, top, Inches(3.0), Inches(2.55), SURFACE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), Inches(2.55))
        fill_shape(bar, color)
        tb = textbox(s, left + Inches(0.2), top + Inches(0.15), Inches(2.65), Inches(2.25))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, name, size=11, bold=True, color=color, space_after=6)
        add_para(tf, outcome, size=13, bold=True, color=TEXT, space_after=6)
        add_para(tf, artifacts, size=11, color=MUTED, space_after=0)

    # Metrics table header style
    rect(s, Inches(0.55), Inches(4.1), Inches(12.2), Inches(2.55), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(4.2), Inches(11.8), Inches(2.35))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "12-MONTH SCOREBOARD  ·  Quarterly Pricing Council adjusts tiers / floors / packs", size=12, bold=True, color=GOLD, space_after=8)
    metrics = [
        ("Net sales", "Flat → +8–12%", "Sales"),
        ("Gross margin", "Recover +3–4 pts", "Finance / SPM"),
        ("Tier SKU adoption", "≥80% of pipeline", "Deal Desk"),
        ("Price variance", "<15% like-for-like", "SPM"),
        ("Wins on value/TCO", "≥40% of wins", "Mkt / Sales"),
        ("Renewal step-up", "≥25% Optimize+", "Delivery"),
    ]
    # two columns of metrics
    for i, (m, t, o) in enumerate(metrics):
        col = 0 if i < 3 else 1
        # we'll just list them inline in two groups via paragraphs
        pass
    add_para(
        tf,
        "Net sales  Flat → +8–12%          Gross margin  recover +3–4 pts          Tier SKU adoption  ≥80%",
        size=13,
        color=TEXT,
        space_after=6,
    )
    add_para(
        tf,
        "Regional price variance  <15%          Wins citing value/TCO  ≥40%          Renewal step-up  ≥25% Optimize+",
        size=13,
        color=TEXT,
        space_after=10,
    )
    add_para(
        tf,
        "Feedback loop: monthly Deal Desk overrides + CR volume → quarterly pricing council (±3% rate card) → "
        "annual sunset re-test & Autonomize → AO SKU merge",
        size=12,
        color=MUTED,
        space_after=0,
    )
    footer(s, 5)

    # ─── SLIDE 6: Close / Ask ─────────────────────────────────────────
    s = blank_slide(prs)
    label(s, "Recommendation  ·  Ask")
    title(s, "Reposition AMS. Land with tiers. Expand to autonomy.")

    rect(s, Inches(0.55), Inches(1.35), Inches(12.2), Inches(1.7), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.5), Inches(11.8), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "THE CALL", size=11, bold=True, color=ACCENT, space_after=6)
    add_para(
        tf,
        "Reposition Application Managed Services as Outcome-Tiered Application Operations "
        "(Stabilize → Optimize → Autonomize). Do not sunset a delivery-capable franchise. "
        "Do not evolve tooling without fixing the commercial SKU. Fund guardrails & enablement in Q1; "
        "pilot in two regions; measure with a quarterly pricing council.",
        size=15,
        color=TEXT,
        space_after=0,
    )

    # Proof + Ask
    rect(s, Inches(0.55), Inches(3.3), Inches(6.0), Inches(2.4), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(3.45), Inches(5.6), Inches(2.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "WHY CREDIBLE", size=11, bold=True, color=GOLD, space_after=8)
    add_para(tf, "H3G UK — $125m", size=14, bold=True, color=TEXT, space_after=2)
    add_para(tf, "75 CPIs · priced transition · 474→336 economics — sold transformation, not FTEs", size=12, color=MUTED, space_after=8)
    add_para(tf, "TEF Argentina — $100m", size=14, bold=True, color=TEXT, space_after=2)
    add_para(tf, "Repositioned accountability & KPI transparency when the story was broken → renewal", size=12, color=MUTED, space_after=0)

    rect(s, Inches(6.8), Inches(3.3), Inches(5.95), Inches(2.4), SURFACE)
    tb = textbox(s, Inches(7.0), Inches(3.45), Inches(5.55), Inches(2.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "ASK OF LEADERSHIP", size=11, bold=True, color=SUCCESS, space_after=8)
    for line in [
        "Approve reposition mandate & global rate card",
        "Stand up Deal Desk guardrails this quarter",
        "Authorize 2 regional pilots (UK + LATAM)",
        "Name SPM owner + monthly steering forum",
    ]:
        add_para(tf, "▸  " + line, size=13, color=TEXT, space_after=6)

    tb = textbox(s, Inches(0.55), Inches(5.95), Inches(12.2), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Q&A  ·  5 minutes     ·     Detail appendix: full diagnosis scorecard, rate-card rules, CPI catalogue, enablement plan",
        size=14,
        bold=True,
        color=MUTED,
        align=PP_ALIGN.CENTER,
        space_after=0,
    )
    footer(s, 6)

    out = "/workspace/case-study-1-reposition-executive.pptx"
    prs.save(out)
    print(f"Wrote {out}")
    return out


if __name__ == "__main__":
    build()
