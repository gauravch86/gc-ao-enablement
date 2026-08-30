#!/usr/bin/env python3
"""Case Study 1 executive PPTX — debiased, full forms, detailed speaker notes."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BG = RGBColor(0x0F, 0x14, 0x19)
SURFACE = RGBColor(0x1A, 0x23, 0x32)
TEXT = RGBColor(0xE6, 0xED, 0xF3)
MUTED = RGBColor(0x8B, 0x9C, 0xB3)
ACCENT = RGBColor(0x58, 0xA6, 0xFF)
GOLD = RGBColor(0xD4, 0xA8, 0x53)
SUCCESS = RGBColor(0x3F, 0xB9, 0x50)
DANGER = RGBColor(0xF8, 0x71, 0x71)
PURPLE = RGBColor(0xA3, 0x71, 0xF7)

W = Inches(13.333)
H = Inches(7.5)


def set_run(run, size=14, bold=False, color=TEXT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"


def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def rect(slide, left, top, width, height, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    fill_shape(s, fill)
    try:
        s.adjustments[0] = 0.08
    except Exception:
        pass
    return s


def textbox(slide, left, top, width, height):
    return slide.shapes.add_textbox(left, top, width, height)


def add_para(tf, text, size=13, bold=False, color=TEXT, align=PP_ALIGN.LEFT, space_after=4):
    if tf.paragraphs[0].text == "" and all(not p.text for p in tf.paragraphs):
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def blank_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    fill_shape(bg, BG)
    spTree = slide.shapes._spTree
    sp = bg._element
    spTree.remove(sp)
    spTree.insert(2, sp)
    return slide


def label(slide, text):
    tb = textbox(slide, Inches(0.55), Inches(0.28), Inches(12), Inches(0.3))
    add_para(tb.text_frame, text.upper(), size=11, bold=True, color=ACCENT, space_after=0)


def title(slide, text):
    tb = textbox(slide, Inches(0.55), Inches(0.52), Inches(12.2), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, text, size=26, bold=True, color=GOLD, space_after=0)


def footer(slide, page, total=6):
    tb = textbox(slide, Inches(0.55), Inches(7.1), Inches(10), Inches(0.28))
    add_para(tb.text_frame, "Case Study 1 · Reposition underperforming offering · Gaurav Chaudhary", size=10, color=MUTED, space_after=0)
    tb2 = textbox(slide, Inches(11.6), Inches(7.1), Inches(1.2), Inches(0.28))
    add_para(tb2.text_frame, f"{page} / {total}", size=10, color=MUTED, align=PP_ALIGN.RIGHT, space_after=0)


def metric_card(slide, left, top, width, height, lbl, val, sub, val_color=GOLD):
    rect(slide, left, top, width, height, SURFACE)
    tb = textbox(slide, left + Inches(0.15), top + Inches(0.12), width - Inches(0.3), height - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, lbl.upper(), size=10, bold=True, color=MUTED, align=PP_ALIGN.CENTER, space_after=2)
    add_para(tf, val, size=24, bold=True, color=val_color, align=PP_ALIGN.CENTER, space_after=2)
    add_para(tf, sub, size=11, color=MUTED, align=PP_ALIGN.CENTER, space_after=0)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # ─── SLIDE 1 ───
    s = blank_slide(prs)
    label(s, "Strategic Product Manager  ·  Case Study 1  ·  10-minute commercial brief")
    title(s, "Reposition an underperforming portfolio offering")

    tb = textbox(s, Inches(0.55), Inches(1.2), Inches(12.2), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Application Managed Services: strong technical delivery, but commercially stuck — "
        "flat net sales, declining gross margin, Sales cannot articulate value, inconsistent regional pricing.",
        size=14, color=MUTED, space_after=0,
    )

    metric_card(s, Inches(0.55), Inches(2.05), Inches(3.9), Inches(1.5),
                "Net sales (2 years)", "Flat", "Weak conversion · little attach", MUTED)
    metric_card(s, Inches(4.7), Inches(2.05), Inches(3.9), Inches(1.5),
                "Gross margin (illustrative)", "28% → 22–24%", "−4 to −6 percentage points over 2 years", DANGER)
    metric_card(s, Inches(8.85), Inches(2.05), Inches(3.9), Inches(1.5),
                "Delivery satisfaction", "Strong", "Execution is not the failure mode", SUCCESS)

    rect(s, Inches(0.55), Inches(3.85), Inches(12.2), Inches(1.35), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(4.0), Inches(11.8), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "EXECUTIVE FRAMING", size=11, bold=True, color=ACCENT, space_after=4)
    add_para(
        tf,
        "This is a commercial and go-to-market problem — not a delivery turnaround. "
        "The offer is sold as headcount while buyers purchase outcomes, cost predictability, and risk reduction. "
        "Decide evolve / reposition / sunset with evidence, then fix packaging, pricing discipline, and Sales enablement.",
        size=13, color=TEXT, space_after=0,
    )

    tb = textbox(s, Inches(0.55), Inches(5.45), Inches(12.2), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "AGENDA  ·  10 MINUTES  (+ 5 MINUTES QUESTIONS)", size=11, bold=True, color=GOLD, space_after=6)
    add_para(
        tf,
        "1  Diagnose   →   2  Decide   →   3  Package & price   →   4  Prioritize & align   →   5  Measure & ask",
        size=14, color=TEXT, space_after=8,
    )
    add_para(
        tf,
        "Leadership question answered: evolve, reposition, or sunset — with a plan Sales, Deal Desk, Finance, and regions can run.",
        size=12, color=MUTED, space_after=0,
    )
    footer(s, 1)
    set_notes(s, """TIMING: ~1 minute 30 seconds

WHAT TO SAY:
Open with the paradox: delivery Net Promoter Score / customer satisfaction is strong, yet net sales are flat and gross margin has fallen illustratively from about 28% to 22–24% — that is a drop of four to six percentage points over two years. That means this is NOT a delivery crisis — it is a commercial and go-to-market failure.

Define the offering once: Application Managed Services — application support sold today as staffing (people and hours).

State the executive framing: buyers purchase outcomes, cost predictability, and risk reduction; we sell headcount. Leadership asked us to choose evolve, reposition, or sunset. Walk the agenda.

GLOSSARY (if asked):
• Net sales = booked revenue after discounts, before cost
• Gross margin = (net sales − direct delivery cost) ÷ net sales. Example 28% → 22–24% = −4 to −6 percentage points (not “percent of percent”)
• Deal Desk = the commercial approval function that reviews discounts, terms, and exceptions
• Attach = selling additional modules or higher tiers onto an existing deal or renewal

DO NOT: dive into technology, automation, or vendor consolidation on slide 1 — stay commercial.""")

    # ─── SLIDE 2 ───
    s = blank_slide(prs)
    label(s, "Diagnosis framework  ·  Recommendation")
    title(s, "Strong delivery. Broken commercial model. → Reposition.")

    lenses = [
        ("WIN / LOSS", "15–20 deals: tag losses as\nprice-only vs value unclear"),
        ("MARGIN", "Gross-margin waterfall:\ndiscount · transition · unpaid scope"),
        ("COMPETITIVE", "Staff augmentation vs\noutcome-packaged peers"),
        ("CUSTOMER VOICE", "High satisfaction + low\nwillingness to pay = story gap"),
        ("PORTFOLIO FIT", "18-month margin path?\nStrategic pull-through?"),
    ]
    for i, (h, body) in enumerate(lenses):
        left = Inches(0.55) + i * Inches(2.5)
        rect(s, left, Inches(1.3), Inches(2.35), Inches(1.5), SURFACE)
        tb = textbox(s, left + Inches(0.12), Inches(1.4), Inches(2.1), Inches(1.3))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=11, bold=True, color=ACCENT, space_after=4)
        for line in body.split("\n"):
            add_para(tf, line, size=11, color=TEXT, space_after=1)

    causes = [
        ("Commoditized packaging", "Sold as Level-1/2/3 staffing towers — procurement benches us against systems integrators"),
        ("Pricing chaos", "Regional ad-hoc discounts · underpriced knowledge transfer · unpaid scope creep"),
        ("Value story gap", "Buyers want outcomes and predictability; Sales pitches headcount replacement"),
    ]
    for i, (h, body) in enumerate(causes):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(3.05), Inches(3.95), Inches(1.35), SURFACE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(3.05), Inches(3.95), Inches(0.06))
        fill_shape(bar, DANGER)
        tb = textbox(s, left + Inches(0.15), Inches(3.2), Inches(3.65), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=13, bold=True, color=DANGER, space_after=4)
        add_para(tf, body, size=12, color=TEXT, space_after=0)

    decisions = [
        ("SUNSET", "Rejects delivery DNA &\naccount footprint.\nOnly if no 18-month margin path.", MUTED, False),
        ("EVOLVE ONLY", "More tooling, same sellable unit.\nFixes neither Sales narrative\nnor pricing discipline.", MUTED, False),
        ("REPOSITION  ✓", "Outcome-tiered Application Ops\nStabilize → Optimize → Transform\nEvolve modules selectively.", SUCCESS, True),
    ]
    for i, (h, body, color, selected) in enumerate(decisions):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(4.65), Inches(3.95), Inches(1.95), SURFACE)
        if selected:
            border = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(4.65), Inches(3.95), Inches(1.95))
            border.fill.background()
            border.line.color.rgb = SUCCESS
            border.line.width = Pt(2)
            try:
                border.adjustments[0] = 0.08
            except Exception:
                pass
        tb = textbox(s, left + Inches(0.18), Inches(4.8), Inches(3.6), Inches(1.65))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=14, bold=True, color=color, space_after=6)
        for line in body.split("\n"):
            add_para(tf, line, size=12, color=TEXT if selected else MUTED, space_after=1)

    footer(s, 2)
    set_notes(s, """TIMING: ~2 minutes

WHAT TO SAY:
Before any portfolio decision, run a four-week diagnostic across five lenses. Decision triggers matter: if losses are “value unclear,” positioning is broken; if margin erosion is mostly discount, guardrails are broken.

Name three root causes — packaging, pricing chaos, value story — all commercial.

Decision matrix: Sunset destroys footprint. Evolve-only (tooling without changing the sellable unit) will not fix Sales or Deal Desk. Recommendation is REPOSITION to Outcome-Tiered Application Operations: Stabilize → Optimize → Transform.

GLOSSARY:
• Sellable unit = what appears on the customer proposal (tier + packs), not internal staffing plan
• Knowledge transfer = paid transition from incumbent / customer teams into our operating model
• Systems integrator = competitor selling labour-led application support
• Transform tier = operating-model uplift (automation share, fewer fire drills, executive scorecard) — NOT a mandatory “autonomy product” dependency

TRADE-OFF TO VOLUNTEER: near-term average deal size may dip; attach and renewal step-up recover value while margin quality improves.""")

    # ─── SLIDE 3 ───
    s = blank_slide(prs)
    label(s, "Revised pricing and packaging  ·  Value proposition")
    title(s, "Sell tiers and outcomes — not staffing towers")

    tb = textbox(s, Inches(0.55), Inches(1.15), Inches(12.2), Inches(0.45))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Value proposition: outcome-tiered managed applications — restore control, improve total cost of ownership where contracted, "
        "and give Sales a globally consistent story with Deal Desk pricing discipline.",
        size=12, color=MUTED, space_after=0,
    )

    tiers = [
        ("STABILIZE", GOLD, "Restore control",
         ["Backlog reduction · service-level baseline", "Single accountability · monthly board", "Top-offender (highest-impact) program",
          "Price: platform fee + criticality class", "Floor gross margin ≥ 22%"]),
        ("OPTIMIZE", ACCENT, "Prove efficiency",
         ["+ Standard monitoring / correlation pack", "+ Runbook automation · proactive problems", "Outcome-band pricing vs service levels",
          "Optional gain-share on proven savings", "Floor gross margin ≥ 26%"]),
        ("TRANSFORM", SUCCESS, "Operating-model uplift",
         ["+ Continuous-improvement factory", "+ Policy-based standard changes", "Milestone elements (Finance sign-off)",
          "Executive scorecard · fewer fire drills", "Floor gross margin ≥ 28%"]),
    ]
    for i, (name, color, tagline, lines) in enumerate(tiers):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(1.75), Inches(3.95), Inches(3.5), SURFACE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.75), Inches(3.95), Inches(0.08))
        fill_shape(bar, color)
        tb = textbox(s, left + Inches(0.2), Inches(1.95), Inches(3.55), Inches(3.15))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, name, size=16, bold=True, color=color, space_after=2)
        add_para(tf, tagline, size=12, bold=True, color=GOLD, space_after=8)
        for line in lines:
            add_para(tf, "•  " + line, size=12, color=TEXT, space_after=4)

    rect(s, Inches(0.55), Inches(5.5), Inches(12.2), Inches(1.2), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(5.6), Inches(11.8), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "MODULAR ATTACH PACKS  ·  PRICED, NOT FREE", size=11, bold=True, color=ACCENT, space_after=4)
    add_para(
        tf,
        "Mandatory: Transition / knowledge-transfer module on new logos.  Optional: top-offender sprint · monitoring jump-start.  "
        "Multi-vendor governance pack — ONLY when win/loss shows consolidation demand (applicable possibility, not required strategy).",
        size=12, color=TEXT, space_after=0,
    )
    footer(s, 3)
    set_notes(s, """TIMING: ~2 minutes

WHAT TO SAY:
The customer-facing product is a tier, not a headcount pyramid. Staffing remains an internal delivery plan.

Walk Stabilize → Optimize → Transform as a commercial ladder Sales can land and expand.
• Stabilize: stop the bleeding — backlog, service levels, single accountability
• Optimize: prove efficiency against contracted service levels
• Transform: operating-model uplift with milestone reviews — Finance must approve revenue recognition before proposal

Stress: transition / knowledge-transfer is MANDATORY and priced — underpriced transition is the leading Year-1 margin killer.

On multi-vendor consolidation: say clearly — “That is an optional pack when win/loss shows buyers consolidating suppliers. It is not required for this reposition. We do not need consolidation to justify tiers.”

GLOSSARY:
• Platform fee = recurring base charge for the operating model / tooling baseline
• Criticality class = how business-critical the application is (drives unit price)
• Outcome-band pricing = price tied to service-level bands (availability, restore time), not hours
• Gain-share = optional commercial where we share proven customer cost reduction
• Total cost of ownership = customer’s full run cost including governance and failure cost, not only our invoice""")

    # ─── SLIDE 4 ───
    s = blank_slide(prs)
    label(s, "Commercial guardrails  ·  Investment sequencing")
    title(s, "Deal Desk rules that stop the bleed — fund return first")

    rect(s, Inches(0.55), Inches(1.25), Inches(6.3), Inches(5.45), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.4), Inches(5.9), Inches(5.15))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "PRICING GUARDRAILS", size=13, bold=True, color=GOLD, space_after=10)
    rules = [
        ("Floor gross margin", "Stabilize ≥22% · Optimize ≥26% · Transform ≥28%"),
        ("Maximum discount", "≤12% off list (≤8% renewals) — trade for scope / term / step-up"),
        ("Transition pricing", "Knowledge-transfer mandatory · min 8% of Year-1 contract value"),
        ("Scope change", "Unpaid work cap 2% of annual contract value · 5-day cycle"),
        ("Regional variance", "Like-for-like within ±15% of global rate card"),
        ("Credits / penalties", "Soft-ramp months 1–6 · Finance revenue-recognition sign-off"),
    ]
    for name, detail in rules:
        add_para(tf, name, size=13, bold=True, color=ACCENT, space_after=1)
        add_para(tf, detail, size=12, color=TEXT, space_after=8)

    rect(s, Inches(7.05), Inches(1.25), Inches(5.7), Inches(3.15), SURFACE)
    tb = textbox(s, Inches(7.25), Inches(1.4), Inches(5.3), Inches(2.9))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "FUND  ·  HIGH RETURN", size=13, bold=True, color=SUCCESS, space_after=8)
    for line in [
        "Q1: Rate card · tier definitions · Deal Desk toolkit",
        "Q2: Service-level catalogue · transition factory · 2 pilots",
        "Q3–4: Modular packs · renewal step-up · scale enablement",
        "Year 2: Retire headcount-only contracts on renewal",
    ]:
        add_para(tf, "▸  " + line, size=12, color=TEXT, space_after=6)

    rect(s, Inches(7.05), Inches(4.6), Inches(5.7), Inches(2.1), SURFACE)
    tb = textbox(s, Inches(7.25), Inches(4.75), Inches(5.3), Inches(1.85))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "STOP / DEPRIORITIZE", size=13, bold=True, color=DANGER, space_after=8)
    for line in [
        "Per-account custom tooling builds",
        "Below-floor proofs of concept to “win logo”",
        "Regional bespoke rate cards",
        "Rebrand without Sales / Deal Desk kit",
    ]:
        add_para(tf, "×  " + line, size=12, color=TEXT, space_after=4)

    footer(s, 4)
    set_notes(s, """TIMING: ~2 minutes

WHAT TO SAY:
Guardrails turn Deal Desk into a deal architect: discounts are traded for scope, term, or tier step-up — not given away.

Hit three hard rules interviewers will probe:
1) Floor gross margin by tier
2) Mandatory priced knowledge-transfer (minimum 8% of Year-1 annual contract value)
3) Regional variance capped at ±15% versus global rate card

Investment sequence shows product discipline: commercial assets first (they unblock Sales this quarter), then delivery standards, then harvest. Explicit stop list proves you will not fund vanity work.

GLOSSARY:
• Annual contract value = expected revenue for one contract year
• Soft-ramp = reduced service-credit exposure in early months while the operating model stabilizes
• Revenue recognition = Finance rules for when revenue can be booked (critical for milestone / credit elements)
• Rate card = approved global list prices and regional index bands
• Q1 / Q2 = calendar or fiscal quarters one and two of the program""")

    # ─── SLIDE 5 ───
    s = blank_slide(prs)
    label(s, "Cross-functional alignment  ·  Success metrics and feedback loop")
    title(s, "One program — Sales, Marketing, Finance, regions")

    funcs = [
        ("SALES / DEAL DESK", ACCENT,
         "≥80% of bids on tier + rate card by Quarter 2",
         "Discovery kit · return-on-investment calculator · weekly pilot deal review"),
        ("MARKETING", PURPLE,
         "One narrative: Application Operations by outcome tier",
         "Tier one-pagers · proof stories · retire staffing-tower language"),
        ("FINANCE", GOLD,
         "Revenue-recognition compliant · gross-margin floors in bid model",
         "Platform vs milestone memo · credit accounting · bid sign-off"),
        ("REGIONS / DELIVERY", SUCCESS,
         "Pilots hit standard service-level catalogue",
         "Knowledge-transfer playbook · performance board · renewal step-ups"),
    ]
    for i, (name, color, outcome, artifacts) in enumerate(funcs):
        left = Inches(0.55) + i * Inches(3.15)
        rect(s, left, Inches(1.25), Inches(3.0), Inches(2.55), SURFACE)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.25), Inches(0.08), Inches(2.55))
        fill_shape(bar, color)
        tb = textbox(s, left + Inches(0.2), Inches(1.4), Inches(2.65), Inches(2.25))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, name, size=11, bold=True, color=color, space_after=6)
        add_para(tf, outcome, size=13, bold=True, color=TEXT, space_after=6)
        add_para(tf, artifacts, size=11, color=MUTED, space_after=0)

    rect(s, Inches(0.55), Inches(4.05), Inches(12.2), Inches(2.6), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(4.2), Inches(11.8), Inches(2.3))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "12-MONTH SCOREBOARD  ·  Quarterly Pricing Council adjusts tiers / floors / packs", size=12, bold=True, color=GOLD, space_after=8)
    add_para(
        tf,
        "Net sales  Flat → +8–12%          Gross margin  recover 3–4 percentage points (e.g. toward 25–27%)          Tier adoption  ≥80% of pipeline",
        size=13, color=TEXT, space_after=6,
    )
    add_para(
        tf,
        "Regional price variance  <15%          Wins citing value / total-cost story  ≥40%          Renewal step-up  ≥25%",
        size=13, color=TEXT, space_after=10,
    )
    add_para(
        tf,
        "Feedback loop: monthly Deal Desk overrides + change-request volume → quarterly Pricing Council (±3% rate-card index) → "
        "annual sunset re-test and Transform demand check",
        size=12, color=MUTED, space_after=0,
    )
    footer(s, 5)
    set_notes(s, """TIMING: ~1 minute 30 seconds

WHAT TO SAY:
Repositioning is a portfolio program. Name owners:
• Sales / Deal Desk — adoption of the sellable unit
• Marketing — one narrative; kill staffing-tower language
• Finance — revenue recognition and gross-margin floors before any creative commercial construct
• Regions / Delivery — pilots prove the catalogue is deliverable

Scoreboard is how leadership knows it worked: sales growth, margin recovery, tier adoption, price consistency, win reasons, renewal step-up — without trading away delivery satisfaction.

Pricing Council is the continuous refinement mechanism the brief asks for.

GLOSSARY:
• Pricing Council = quarterly forum (Product, Sales, Finance, regions) that adjusts rate card and packs from field evidence
• Renewal step-up = moving a live account from Stabilize to Optimize/Transform at renewal
• Change request = formal scope change; unpaid change requests are a hidden margin leak""")

    # ─── SLIDE 6 ───
    s = blank_slide(prs)
    label(s, "Recommendation  ·  Ask")
    title(s, "Reposition. Land with tiers. Expand when value is proven.")

    rect(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(1.7), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.45), Inches(11.8), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "THE CALL", size=11, bold=True, color=ACCENT, space_after=6)
    add_para(
        tf,
        "Reposition Application Managed Services as Outcome-Tiered Application Operations "
        "(Stabilize → Optimize → Transform). Do not sunset a delivery-capable franchise. "
        "Do not evolve tooling without fixing the commercial sellable unit. "
        "Fund guardrails and enablement in Quarter 1; pilot in two regions; measure with a quarterly Pricing Council.",
        size=14, color=TEXT, space_after=0,
    )

    rect(s, Inches(0.55), Inches(3.25), Inches(6.0), Inches(2.45), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(3.4), Inches(5.6), Inches(2.15))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "THOUGHT PROCESS (WHY THIS IS EXECUTIVE)", size=11, bold=True, color=GOLD, space_after=8)
    add_para(tf, "Evidence before instinct — five-lens diagnostic", size=12, color=TEXT, space_after=4)
    add_para(tf, "Separate delivery quality from commercial failure", size=12, color=TEXT, space_after=4)
    add_para(tf, "Choose with trade-offs; fund Sales motion before tooling", size=12, color=TEXT, space_after=4)
    add_para(tf, "Optional packs (e.g. multi-vendor) only when data says so", size=12, color=TEXT, space_after=0)

    rect(s, Inches(6.8), Inches(3.25), Inches(5.95), Inches(2.45), SURFACE)
    tb = textbox(s, Inches(7.0), Inches(3.4), Inches(5.55), Inches(2.15))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "ASK OF LEADERSHIP", size=11, bold=True, color=SUCCESS, space_after=8)
    for line in [
        "Approve reposition mandate and global rate card",
        "Stand up Deal Desk guardrails this quarter",
        "Authorize two regional pilots before global cutover",
        "Name Strategic Product Manager owner + monthly steering",
    ]:
        add_para(tf, "▸  " + line, size=13, color=TEXT, space_after=6)

    tb = textbox(s, Inches(0.55), Inches(5.95), Inches(12.2), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "QUESTIONS  ·  5 MINUTES     ·     Deep dive: HTML case study (diagnosis scorecard, full guardrail table, enablement plan)",
        size=13, bold=True, color=MUTED, align=PP_ALIGN.CENTER, space_after=0,
    )
    footer(s, 6)
    set_notes(s, """TIMING: ~1 minute, then hand to questions (5 minutes)

WHAT TO SAY:
Restate the call in one breath. Emphasize thought process: evidence → separate delivery from commercial → decide with trade-offs → fund Sales motion before tooling → optional strategies only when data supports them.

Close with four concrete asks. Then: “Happy to go deep on any lens, guardrail, or pilot design.”

LIKELY QUESTIONS AND SHORT ANSWERS:
• Why not sunset? Delivery asset + footprint; 18-month margin path exists via reposition.
• Won’t smaller tiers cut revenue? Near term maybe; attach + renewal step-up + better margin quality.
• How fast on rate card? Quarter 1 mandate; two pilots before forced global cutover.
• Finance concern on outcome pricing? Platform fee first; milestone / credit elements only with revenue-recognition memo.
• Is vendor consolidation required? No — optional pack when win/loss shows consolidation demand.
• How is this different from your lived delivery experience? This case is a portfolio commercial problem — packaging, pricing governance, and Sales enablement — solved with product-management discipline, not an operating-model rebuild.""")

    out = "/workspace/case-study-1-reposition-executive.pptx"
    prs.save(out)
    print(f"Wrote {out}")
    return out


if __name__ == "__main__":
    build()
