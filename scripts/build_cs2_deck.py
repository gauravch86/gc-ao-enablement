#!/usr/bin/env python3
"""Case Study 2 executive PPTX — Autonomous Operations offering, 10-min brief + detailed notes."""

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
CYAN = RGBColor(0x38, 0xBD, 0xF8)

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
    tb = textbox(slide, Inches(0.55), Inches(0.28), Inches(12.2), Inches(0.3))
    add_para(tb.text_frame, text.upper(), size=11, bold=True, color=ACCENT, space_after=0)


def title(slide, text):
    tb = textbox(slide, Inches(0.55), Inches(0.52), Inches(12.2), Inches(0.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, text, size=24, bold=True, color=GOLD, space_after=0)


def footer(slide, page, total=6):
    tb = textbox(slide, Inches(0.55), Inches(7.1), Inches(10.5), Inches(0.28))
    add_para(tb.text_frame, "Case Study 2 · Reusable Autonomous Operations offering · Gaurav Chaudhary", size=10, color=MUTED, space_after=0)
    tb2 = textbox(slide, Inches(11.6), Inches(7.1), Inches(1.2), Inches(0.28))
    add_para(tb2.text_frame, f"{page} / {total}", size=10, color=MUTED, align=PP_ALIGN.RIGHT, space_after=0)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # 1 Brief
    s = blank_slide(prs)
    label(s, "Strategic Product Manager  ·  Case Study 2  ·  10-minute brief  ·  Lived experience")
    title(s, "Reusable Autonomous Operations on customer KPIs / OKRs")

    tb = textbox(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Large telecom operator · multi-vendor · multi-domain · margin and reliability pressure. "
        "Today: reactive operations, high incidents, manual troubleshooting, inconsistent service-level agreements. "
        "Propose a reusable managed-services offering: reactive → predictive → autonomous, anchored on customer-agreed key performance indicators.",
        size=14, color=MUTED, space_after=0,
    )

    rect(s, Inches(0.55), Inches(2.35), Inches(12.2), Inches(1.2), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(2.5), Inches(11.8), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "NORTH STAR", size=11, bold=True, color=ACCENT, space_after=4)
    add_para(
        tf,
        "End-user experience and Net Promoter Score · service-level / contractual performance-indicator attainment · market trust. "
        "Customer total cost of ownership and a reusable commercial offering are how we fund and scale that outcome — not the outcome itself.",
        size=13, color=TEXT, space_after=0,
    )

    boxes = [
        ("GIVEN", "Ericsson products + third-party portfolio under one Managed Services umbrella", GOLD),
        ("METHOD", "Design → Deploy → Operate → Assure → Improve  ·  pilots before scale", ACCENT),
        ("PROOF", "Telefónica Argentina renewal · Three UK consolidation & autonomy economics", SUCCESS),
    ]
    for i, (h, body, col) in enumerate(boxes):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(3.8), Inches(3.95), Inches(1.7), SURFACE)
        tb = textbox(s, left + Inches(0.2), Inches(3.95), Inches(3.55), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=12, bold=True, color=col, space_after=6)
        add_para(tf, body, size=13, color=TEXT, space_after=0)

    tb = textbox(s, Inches(0.55), Inches(5.8), Inches(12.2), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "AGENDA  ·  10 MINUTES", size=11, bold=True, color=GOLD, space_after=6)
    add_para(
        tf,
        "1  Economic unlock   →   2  Autonomy path & lifecycle   →   3  Measure & reuse   →   4  Enable & commercialize   →   5  Profit-and-loss & close",
        size=13, color=TEXT, space_after=0,
    )
    footer(s, 1)
    set_notes(s, """TIMING: ~1:30

WHAT TO SAY:
Paint the customer pain in one breath: multi-vendor chaos, reactive firefighting, inconsistent service-level agreements, margin pressure.

North star is end-user experience and trust — total cost of ownership is how we fund it, not the slogan.

Given: Ericsson + third-party estate under Managed Services — we productize operations across the mix.

State this is lived pattern (Telefónica Argentina, Three UK) productized into a reusable offer.

GLOSSARY:
• KPI = key performance indicator (what we measure)
• OKR = objectives and key results (outcome goals tied to KPIs)
• TCO = total cost of ownership (customer run cost including governance)
• SLA = service-level agreement
• CPI = contractual performance indicator (service levels with financial credits)
• 3PP = third-party products
• NPS = Net Promoter Score""")

    # 2 Economic unlock + autonomy
    s = blank_slide(prs)
    label(s, "Step 1–2  ·  Economic unlock and autonomy path")
    title(s, "Vendor consolidation unlocks autonomy — it is not the north star")

    for i, (h, v, sub, col) in enumerate([
        ("IT operations annual budget (assumed)", "$50m", "Multi-vendor run cost today", GOLD),
        ("Target after consolidation", "$40m", "~20% total-cost reduction", SUCCESS),
        ("Customer governance load", "−20%", "Vendor management / service-level overhead", ACCENT),
    ]):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(1.3), Inches(3.95), Inches(1.35), SURFACE)
        tb = textbox(s, left + Inches(0.15), Inches(1.4), Inches(3.65), Inches(1.15))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h.upper(), size=10, bold=True, color=MUTED, align=PP_ALIGN.CENTER, space_after=2)
        add_para(tf, v, size=22, bold=True, color=col, align=PP_ALIGN.CENTER, space_after=2)
        add_para(tf, sub, size=11, color=MUTED, align=PP_ALIGN.CENTER, space_after=0)

    stages = [
        ("NOW · REACTIVE", "Ticket storms · manual triage · multi-vendor blame", DANGER),
        ("PHASE A · OBSERVE", "One health view across infrastructure, cloud, application, network", ACCENT),
        ("PHASE B · PREDICT", "Anomaly detection · prevent before end-user impact", PURPLE),
        ("PHASE C · AUTONOMOUS", "Self-heal on known classes · human-in-loop → policy auto", SUCCESS),
    ]
    for i, (h, body, col) in enumerate(stages):
        left = Inches(0.55) + i * Inches(3.15)
        rect(s, left, Inches(2.95), Inches(3.0), Inches(1.85), SURFACE)
        tb = textbox(s, left + Inches(0.15), Inches(3.1), Inches(2.7), Inches(1.55))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=11, bold=True, color=col, space_after=6)
        add_para(tf, body, size=12, color=TEXT, space_after=0)

    rect(s, Inches(0.55), Inches(5.05), Inches(12.2), Inches(1.6), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(5.2), Inches(11.8), Inches(1.35))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "LIFECYCLE  ·  ONE PRODUCTIZED JOURNEY", size=11, bold=True, color=GOLD, space_after=6)
    add_para(
        tf,
        "Design (intent & policy · KPI contract)  →  Deploy (connectors · single pane of glass · shadow mode)  →  "
        "Operate (human-in-loop · self-heal)  →  Assure (digital Service Delivery Manager board · credits)  →  "
        "Improve (offender backlog · known-error sunset · library feedback)",
        size=13, color=TEXT, space_after=0,
    )
    footer(s, 2)
    set_notes(s, """TIMING: ~2:00

WHAT TO SAY:
Consolidation is the operating-model unlock — single accountability, cleaner baseline for observability and autonomy. Customer total-cost and governance reduction are proof, not the north star.

Walk Reactive → Observe → Predict → Autonomous with gates (no penalty cliffs).

Lifecycle is one productized journey — artificial intelligence, orchestration, and assurance plug into each stage; intent and policy replace tribal manual action over time.

GLOSSARY:
• MTTR = mean time to restore
• SPOG = single pane of glass (one executive / Service Delivery Manager view of health)
• HITL = human-in-the-loop (AI proposes; human approves until policy allows full auto)
• Shadow mode = run detection / proposals without acting in production yet
• Known error = documented root cause with workaround pending permanent fix""")

    # 3 Measure & reuse
    s = blank_slide(prs)
    label(s, "Step 3–4  ·  Data-driven KPIs and reuse library")
    title(s, "Size from incidents. Contract soft-ramp. Productize reuse.")

    rect(s, Inches(0.55), Inches(1.3), Inches(6.2), Inches(3.4), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.45), Inches(5.8), Inches(3.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "DISCOVERY → CONTRACT", size=13, bold=True, color=GOLD, space_after=8)
    for line in [
        "Ingest 6–12 months of incident history",
        "Subject-matter expert + AI enrichment → maturity per component",
        "Top-offender backlog (volume × business impact)",
        "Contractual performance indicators: availability, response/restore, problem aging, change success",
        "Soft-ramp windows — avoid penalty cliffs in early months",
        "Same scoreboard for roadmap priority and milestone profit-and-loss",
    ]:
        add_para(tf, "•  " + line, size=13, color=TEXT, space_after=5)

    rect(s, Inches(7.0), Inches(1.3), Inches(5.75), Inches(3.4), SURFACE)
    tb = textbox(s, Inches(7.2), Inches(1.45), Inches(5.35), Inches(3.1))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "REUSE LIBRARY (WHY IT SCALES)", size=13, bold=True, color=ACCENT, space_after=8)
    for line in [
        "Metrics / runbook / self-heal packs per domain",
        "Drag-drop into the next account — not rebuild",
        "Ericsson + third-party adapters under one operating model",
        "Trade-off: thin adapters beat rip-and-replace dogma",
        "Library feedback from Improve stage compounds globally",
        "Portfolio research and development absorbs long-term platform cost",
    ]:
        add_para(tf, "•  " + line, size=13, color=TEXT, space_after=5)

    rect(s, Inches(0.55), Inches(4.95), Inches(12.2), Inches(1.7), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(5.1), Inches(11.8), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "ILLUSTRATIVE CONTRACTUAL PERFORMANCE INDICATORS (FROM LIVE MANAGED-SERVICES PRACTICE)", size=11, bold=True, color=GOLD, space_after=6)
    add_para(
        tf,
        "Availability by service class (credits up to ~20%)  ·  Priority-0/1 restore ≤4 hours (credit ~10%)  ·  "
        "Change success ≥98%  ·  Proactive problems ≥20%  ·  Repeat Priority-0–2 ≤10%  ·  "
        "Credit % = financial exposure if breached — protects customer return and our margin narrative.",
        size=13, color=TEXT, space_after=0,
    )
    footer(s, 3)
    set_notes(s, """TIMING: ~2:00

WHAT TO SAY:
We do not invent service levels in a slide workshop — we size them from 6–12 months of incidents so maturity and risk are honest.

Soft-ramp protects both parties from early penalty cliffs.

Reuse library is the productization answer: next customer inherits packs; adapters for Ericsson and third-party; Improve stage feeds the global library.

Point at a few contractual performance indicators with credits — shows commercial gravity, not vanity metrics.

GLOSSARY:
• Top offender = highest volume × impact defect class
• Soft-ramp = reduced credit exposure in early contract months
• Credit = service credit / financial penalty if contractual performance indicator is missed
• P0/P1 = highest severity incidents
• Adapter = thin integration layer to vendor tools/APIs without full rip-and-replace""")

    # 4 Enable & commercialize
    s = blank_slide(prs)
    label(s, "Step 5–6  ·  Enablement and commercial construct")
    title(s, "Enable customer ops and regions. Sell autonomy maturity.")

    cols = [
        ("CUSTOMER OPERATIONS", [
            "Train on single pane of glass and performance board",
            "Shadow → co-pilot → autonomy coverage expand",
            "Resist shadow tooling and tribal workarounds",
            "Success: service-level attainment + adoption metrics",
        ], ACCENT),
        ("ERICSSON REGIONS", [
            "Playbooks from the reuse library",
            "Deal Desk guardrails by autonomy tier",
            "Demo: single pane + self-heal + performance board",
            "Executive quality gates at month 3 / 12 / 18",
        ], PURPLE),
        ("COMMERCIAL PACKAGING", [
            "Outcome bands mapped to autonomy phases",
            "Not effort-hours as the customer-facing product",
            "Pilot assumptions before multi-region scale",
            "One-pager: total-cost story + roadmap milestones",
        ], GOLD),
    ]
    for i, (h, lines, col) in enumerate(cols):
        left = Inches(0.55) + i * Inches(4.15)
        rect(s, left, Inches(1.3), Inches(3.95), Inches(3.6), SURFACE)
        tb = textbox(s, left + Inches(0.2), Inches(1.45), Inches(3.55), Inches(3.3))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=13, bold=True, color=col, space_after=8)
        for line in lines:
            add_para(tf, "•  " + line, size=13, color=TEXT, space_after=6)

    rect(s, Inches(0.55), Inches(5.15), Inches(12.2), Inches(1.5), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(5.3), Inches(11.8), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "ASSUMPTIONS → PILOTS BEFORE SCALE", size=11, bold=True, color=SUCCESS, space_after=6)
    add_para(
        tf,
        "Validate total-cost trajectory, soft-ramp credibility, and library reuse on a bounded estate. "
        "Pass-gates required before regional rollout — protects customer trust and our profit-and-loss.",
        size=13, color=TEXT, space_after=0,
    )
    footer(s, 4)
    set_notes(s, """TIMING: ~1:30

WHAT TO SAY:
Enablement is two audiences: customer operations (adoption) and Ericsson regions (repeatable sell/delivery).

Commercial packaging sells autonomy maturity and outcome bands — not a body shop.

Pilots with pass-gates are non-negotiable — assumptions must be proven before scale.

GLOSSARY:
• Deal Desk = commercial approval / deal architecture function
• Autonomy tier = commercial/packaging band aligned to Reactive / Predictive / Autonomous capability
• Pass-gate = explicit go/no-go criteria before expanding scope or geography""")

    # 5 P&L
    s = blank_slide(prs)
    label(s, "Commercial depth  ·  Five-year managed-services profit-and-loss")
    title(s, "Full-time-equivalent ramp · site-mix · tools · penalty lot")

    rect(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(1.15), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.4), Inches(11.8), Inches(0.9))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(
        tf,
        "Minimum five-year managed-services deal. Baseline multi-vendor headcount = B. "
        "Stand-up target X = B − 30%. Customer committed ~20% run-cost reduction; we size at 30% — "
        "the 10-point gap is structural earnings before interest and tax, not a sales promise.",
        size=13, color=TEXT, space_after=0,
    )

    # FTE path
    years = [("Y1", "X+20%", "Discovery / knowledge-transfer peak"),
             ("Y2", "X+10%", "Human-in-loop packs expanding"),
             ("Y3", "X", "Steady-state target model"),
             ("Y4", "X−10%", "Autonomy + reuse harvest"),
             ("Y5", "X−20%", "Closed-loop scale")]
    for i, (y, v, r) in enumerate(years):
        left = Inches(0.55) + i * Inches(2.5)
        rect(s, left, Inches(2.65), Inches(2.35), Inches(1.45), SURFACE)
        tb = textbox(s, left + Inches(0.1), Inches(2.75), Inches(2.15), Inches(1.25))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, y, size=12, bold=True, color=ACCENT, align=PP_ALIGN.CENTER, space_after=2)
        add_para(tf, v, size=18, bold=True, color=GOLD, align=PP_ALIGN.CENTER, space_after=2)
        add_para(tf, r, size=10, color=MUTED, align=PP_ALIGN.CENTER, space_after=0)

    levers = [
        ("① Labour gap", "+10%", "Size −30% vs commit −20%"),
        ("② Site-mix", "+10%", "Onsite leads · offshore factory"),
        ("③ Tools seed", "−5%", "Reusable IP start cost"),
        ("④ License recover", "+5%", "Customer redirects tooling budget"),
        ("⑤ Penalty lot", "+5% if met", "Parked envelope → earn-back"),
    ]
    for i, (h, v, sub) in enumerate(levers):
        left = Inches(0.55) + i * Inches(2.5)
        rect(s, left, Inches(4.35), Inches(2.35), Inches(1.75), SURFACE)
        tb = textbox(s, left + Inches(0.1), Inches(4.5), Inches(2.15), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        add_para(tf, h, size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER, space_after=2)
        add_para(tf, v, size=16, bold=True, color=SUCCESS if "+" in v else DANGER, align=PP_ALIGN.CENTER, space_after=2)
        add_para(tf, sub, size=10, color=MUTED, align=PP_ALIGN.CENTER, space_after=0)

    footer(s, 5)
    set_notes(s, """TIMING: ~2:00

WHAT TO SAY:
This is where Deal Desk and leadership lean in. Five-year construct. Target headcount X = baseline minus 30%. Customer story commits about 20% run-cost reduction — the gap is structural earnings.

Walk the year path: overstaff early for knowledge transfer and soft-ramp; harvest later via autonomy and reuse.

Five levers: labour gap, site-mix, tools seed, license recovery, penalty parking lot with earn-back if service levels hold.

Plan at 10–15% earnings before interest and tax with explicit envelope; upside toward ~20% is earned, not assumed in the bid.

GLOSSARY:
• FTE = full-time equivalent (headcount unit)
• B = baseline multi-vendor headcount
• X = Ericsson stand-up target headcount (B − 30%)
• EBIT = earnings before interest and tax
• Site-mix = onsite leadership / module leads versus offshore or near-shore delivery factory
• Penalty lot = reserved margin for potential service credits; converts to earnings if gates are met
• Earn-back = contractual path to recover credits when performance recovers""")

    # 6 Close
    s = blank_slide(prs)
    label(s, "Credibility and close")
    title(s, "Lived proof. One recommendation.")

    rect(s, Inches(0.55), Inches(1.3), Inches(6.0), Inches(2.5), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(1.45), Inches(5.6), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "TELEFÓNICA ARGENTINA — ~$100m RENEWAL", size=12, bold=True, color=GOLD, space_after=8)
    add_para(tf, "3,000+ backlog · ~500/day → offender taxonomy", size=13, color=TEXT, space_after=4)
    add_para(tf, "Executive key-performance-indicator transparency · 24×7 scale", size=13, color=TEXT, space_after=4)
    add_para(tf, "Accountability restored → renewal won on trust", size=13, color=TEXT, space_after=0)

    rect(s, Inches(6.8), Inches(1.3), Inches(5.95), Inches(2.5), SURFACE)
    tb = textbox(s, Inches(7.0), Inches(1.45), Inches(5.55), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "THREE UK — ~$125m VENDOR CONSOLIDATION", size=12, bold=True, color=ACCENT, space_after=8)
    add_para(tf, "Four vendors · 400+ components · ~75 contractual performance indicators", size=13, color=TEXT, space_after=4)
    add_para(tf, "Single pane of glass + self-heal + generative-AI runbooks", size=13, color=TEXT, space_after=4)
    add_para(tf, "474→336 full-time-equivalent economics · lowest Priority-1s", size=13, color=TEXT, space_after=0)

    rect(s, Inches(0.55), Inches(4.05), Inches(12.2), Inches(2.0), SURFACE)
    tb = textbox(s, Inches(0.75), Inches(4.2), Inches(11.8), Inches(1.7))
    tf = tb.text_frame
    tf.word_wrap = True
    add_para(tf, "RECOMMENDATION", size=11, bold=True, color=SUCCESS, space_after=6)
    add_para(
        tf,
        "Launch a reusable Autonomous Operations managed-services offering (Ericsson + third-party under Managed Services) "
        "that sells autonomy maturity + unified observability + outcome key performance indicators as one productized journey — "
        "with enablement for customer operations and Ericsson regions, and pilots that validate total cost of ownership, "
        "soft-ramp, and library reuse before scaling.",
        size=14, color=TEXT, space_after=0,
    )

    tb = textbox(s, Inches(0.55), Inches(6.25), Inches(12.2), Inches(0.5))
    add_para(
        tb.text_frame,
        "QUESTIONS  ·  5 MINUTES     ·     Deep dive: HTML case study (full contractual performance-indicator catalogue, profit-and-loss waterfall)",
        size=13, bold=True, color=MUTED, align=PP_ALIGN.CENTER, space_after=0,
    )
    footer(s, 6)
    set_notes(s, """TIMING: ~1:00 then questions

WHAT TO SAY:
Two proof anchors — renewal when trust was broken; consolidation when economics and autonomy had to be real.

Close with the recommendation sentence. Invite questions on contractual performance indicators, five-year economics, or enablement design.

LIKELY QUESTIONS:
• Why consolidation first? Operating-model unlock for single accountability and observability — not the north star.
• How avoid penalty shock? Soft-ramp + data-sized service levels + penalty lot in profit-and-loss.
• How reusable across customers? Metrics/runbook/self-heal library + thin adapters for Ericsson and third-party.
• What if pilots fail a gate? Do not scale; envelope absorbs; fix before regional rollout.""")

    out = "/workspace/case-study-2-ao-offering-executive.pptx"
    prs.save(out)
    print(f"Wrote {out}")
    return out


if __name__ == "__main__":
    build()
