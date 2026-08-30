#!/usr/bin/env python3
"""Case Study 1 executive PDF — 8-slide sequential interview flow."""

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

BG = HexColor("#0F1419")
SURFACE = HexColor("#1A2332")
TEXT = HexColor("#E6EDF3")
MUTED = HexColor("#8B9CB3")
ACCENT = HexColor("#58A6FF")
GOLD = HexColor("#D4A853")
SUCCESS = HexColor("#3FB950")
DANGER = HexColor("#F87171")
PURPLE = HexColor("#A371F7")
PAGE = landscape(A4)
W, H = PAGE
TOTAL = 8


def rr(c, x, y, w, h, fill, radius=8):
    c.setFillColor(fill)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=0)


def foot(c, page):
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8)
    c.drawString(0.55 * inch, 0.28 * inch, "Case Study 1 · Reposition underperforming offering · Gaurav Chaudhary")
    c.drawRightString(W - 0.55 * inch, 0.28 * inch, f"{page} / {TOTAL}")


def wrap(c, text, x, y, max_width, font="Helvetica", size=11, color=TEXT, leading=14):
    c.setFont(font, size)
    c.setFillColor(color)
    words, lines, cur = text.split(), [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if c.stringWidth(test, font, size) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    for i, line in enumerate(lines):
        c.drawString(x, y - i * leading, line)
    return len(lines) * leading


def bg(c):
    c.setFillColor(BG)
    c.rect(0, 0, W, H, fill=1, stroke=0)


def flow(c, active):
    steps = ["1 Diagnose", "2 Findings", "3 Decide", "4 Package", "5 Guardrails", "6 Align", "7 Ask"]
    ww = 1.45 * inch
    gap = 0.1 * inch
    y = 0.55 * inch
    for i, name in enumerate(steps):
        x = 0.55 * inch + i * (ww + gap)
        rr(c, x, y, ww, 0.28 * inch, SURFACE, 4)
        if i == active:
            c.setFillColor(GOLD)
            c.rect(x, y + 0.23 * inch, ww, 0.05 * inch, fill=1, stroke=0)
        c.setFillColor(GOLD if i == active else MUTED)
        c.setFont("Helvetica-Bold" if i == active else "Helvetica", 8)
        c.drawCentredString(x + ww / 2, y + 0.08 * inch, name)


def badge(c, num, name, color, y=None):
    if y is None:
        y = H - 1.55 * inch
    rr(c, 0.55 * inch, y, 2.5 * inch, 0.35 * inch, SURFACE)
    c.setFillColor(color)
    c.rect(0.55 * inch, y, 0.07 * inch, 0.35 * inch, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, y + 0.1 * inch, f"STAGE {num}  ·  {name}")


def build(path="/workspace/case-study-1-reposition-executive.pdf"):
    c = canvas.Canvas(path, pagesize=PAGE)

    # 1 Situation
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STRATEGIC PRODUCT MANAGER · CASE STUDY 1 · 10-MINUTE COMMERCIAL BRIEF")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 18)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Reposition an underperforming portfolio offering")
    wrap(c, "Application Managed Services: strong technical delivery, but commercially stuck — flat net sales, declining gross margin, Sales cannot articulate value, inconsistent regional pricing.",
         0.55 * inch, H - 1.15 * inch, W - 1.1 * inch, size=10, color=MUTED)
    cards = [("NET SALES (2 YRS)", "Flat", "Weak conversion · little attach", MUTED),
             ("GROSS MARGIN (ILLUSTRATIVE)", "28% → 22–24%", "−4 to −6 percentage points over 2 years", DANGER),
             ("DELIVERY SATISFACTION", "Strong", "Execution is not the failure mode", SUCCESS)]
    cw = (W - 1.3 * inch) / 3
    for i, (lbl, val, sub, col) in enumerate(cards):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rr(c, x, H - 3.2 * inch, cw, 1.35 * inch, SURFACE)
        c.setFillColor(MUTED); c.setFont("Helvetica-Bold", 7); c.drawCentredString(x + cw / 2, H - 2.1 * inch, lbl)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 16); c.drawCentredString(x + cw / 2, H - 2.5 * inch, val)
        c.setFillColor(MUTED); c.setFont("Helvetica", 8); c.drawCentredString(x + cw / 2, H - 2.85 * inch, sub)
    rr(c, 0.55 * inch, H - 4.5 * inch, W - 1.1 * inch, 1.0 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, H - 3.7 * inch, "EXECUTIVE FRAMING")
    wrap(c, "Commercial / go-to-market problem — not a delivery turnaround. Sold as headcount; buyers purchase outcomes, cost predictability, and risk reduction.",
         0.75 * inch, H - 4.0 * inch, W - 1.5 * inch, size=10)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(0.55 * inch, 1.55 * inch, "STORYLINE · ONE STAGE PER SLIDE")
    c.setFillColor(TEXT); c.setFont("Helvetica", 11)
    c.drawString(0.55 * inch, 1.25 * inch, "1 Five-lens diagnosis  →  2 Illustrative findings  →  3 Decision → Reposition")
    c.drawString(0.55 * inch, 1.0 * inch, "4 Package & price  ·  5 Guardrails & invest  ·  6 Align & measure  ·  7 Ask")
    foot(c, 1); c.showPage()

    # 2 Diagnose
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 1 OF 7 · GATHER EVIDENCE BEFORE ANY PORTFOLIO CALL")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Five-lens diagnosis — what we would gather")
    badge(c, "1", "FIVE-LENS DIAGNOSIS", ACCENT)
    lenses = [
        ("A  WIN / LOSS", "15–20 deals. Tag price-only vs value unclear. Trigger: value unclear >30% → positioning broken."),
        ("B  MARGIN", "Discount · transition overrun · unpaid scope. Trigger: discount-led bleed → guardrails."),
        ("C  COMPETITIVE", "Staff augmentation vs outcome peers. Trigger: buyers ask for tiers → reposition."),
        ("D  CUSTOMER VOICE", "High satisfaction + low willingness to pay = story gap."),
        ("E  PORTFOLIO FIT", "18-month margin path? Pull-through? Sunset only if both fail."),
    ]
    lw = (W - 1.3 * inch) / 5
    for i, (h, body) in enumerate(lenses):
        x = 0.55 * inch + i * (lw + 0.08 * inch)
        rr(c, x, 1.1 * inch, lw, 3.7 * inch, SURFACE)
        c.setFillColor(ACCENT); c.rect(x, 1.1 * inch + 3.7 * inch - 4, lw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 9); c.drawString(x + 0.1 * inch, 4.5 * inch, h)
        wrap(c, body, x + 0.1 * inch, 4.15 * inch, lw - 0.2 * inch, size=9, leading=12)
    flow(c, 0); foot(c, 2); c.showPage()

    # 3 Findings
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 2 OF 7 · ILLUSTRATIVE SYNTHESIS FROM THE DIAGNOSTIC")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Likely findings — commercial root causes, not delivery")
    badge(c, "2", "ILLUSTRATIVE FINDINGS", DANGER)
    causes = [
        ("1  COMMODITIZED PACKAGING", "Sold as L1/L2/L3 staffing towers. Procurement benches rate-per-person vs systems integrators. Symptom: flat sales."),
        ("2  PRICING CHAOS", "Regional ad-hoc discounts. Underpriced knowledge transfer. Unpaid scope. Symptom: 28% → 22–24% margin."),
        ("3  VALUE STORY GAP", "Buyers want outcomes and predictability. Sales pitches headcount. Symptom: cannot explain value."),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (h, body) in enumerate(causes):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rr(c, x, 1.1 * inch, cw, 3.7 * inch, SURFACE)
        c.setFillColor(DANGER); c.rect(x, 1.1 * inch + 3.7 * inch - 4, cw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 10); c.drawString(x + 0.15 * inch, 4.5 * inch, h)
        wrap(c, body, x + 0.15 * inch, 4.1 * inch, cw - 0.3 * inch, size=11, leading=15)
    flow(c, 1); foot(c, 3); c.showPage()

    # 4 Decide
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 3 OF 7 · PORTFOLIO CALL WITH EXPLICIT TRADE-OFFS")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Decision framework → Reposition")
    badge(c, "3", "DECISION FRAMEWORK", SUCCESS)
    decisions = [
        ("SUNSET", MUTED, False, "Rejects footprint. Only if no 18-month margin path AND no pull-through. Not indicated."),
        ("EVOLVE ONLY", MUTED, False, "More tooling, same sellable unit. Fixes neither Sales narrative nor pricing chaos."),
        ("REPOSITION  ✓", SUCCESS, True, "Outcome-tiered Application Ops. Stabilize → Optimize → Transform. Evolve modules selectively."),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (h, col, sel, body) in enumerate(decisions):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rr(c, x, 1.1 * inch, cw, 3.7 * inch, SURFACE)
        if sel:
            c.setStrokeColor(SUCCESS); c.setLineWidth(2.5)
            c.roundRect(x, 1.1 * inch, cw, 3.7 * inch, 8, fill=0, stroke=1)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 14); c.drawString(x + 0.15 * inch, 4.5 * inch, h)
        wrap(c, body, x + 0.15 * inch, 4.05 * inch, cw - 0.3 * inch, size=11, color=TEXT if sel else MUTED, leading=15)
    flow(c, 2); foot(c, 4); c.showPage()

    # 5 Package
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 4 OF 7 · REVISED PACKAGING AND PRICING")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Sell tiers and outcomes — not staffing towers")
    badge(c, "4", "PACKAGE & PRICE", GOLD)
    tiers = [
        ("STABILIZE", GOLD, "Restore control", ["Service-level baseline", "Platform + criticality", "Floor GM ≥ 22%"]),
        ("OPTIMIZE", ACCENT, "Prove efficiency", ["Monitoring / automation", "Outcome-band pricing", "Floor GM ≥ 26%"]),
        ("TRANSFORM", SUCCESS, "Operating-model uplift", ["Continuous improvement", "Milestone (Finance OK)", "Floor GM ≥ 28%"]),
    ]
    tw = (W - 1.3 * inch) / 3
    for i, (name, col, tag, lines) in enumerate(tiers):
        x = 0.55 * inch + i * (tw + 0.1 * inch)
        rr(c, x, 1.7 * inch, tw, 3.1 * inch, SURFACE)
        c.setFillColor(col); c.rect(x, 1.7 * inch + 3.1 * inch - 4, tw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 12); c.drawString(x + 0.15 * inch, 4.5 * inch, name)
        c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(x + 0.15 * inch, 4.2 * inch, tag)
        c.setFillColor(TEXT); c.setFont("Helvetica", 10)
        for j, line in enumerate(lines):
            c.drawString(x + 0.15 * inch, 3.8 * inch - j * 18, "•  " + line)
    rr(c, 0.55 * inch, 1.05 * inch, W - 1.1 * inch, 0.5 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 8)
    c.drawString(0.75 * inch, 1.3 * inch, "ATTACH: mandatory knowledge-transfer · optional packs · multi-vendor ONLY when win/loss shows demand")
    flow(c, 3); foot(c, 5); c.showPage()

    # 6 Guardrails
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 5 OF 7 · COMMERCIAL DISCIPLINE AND INVESTMENT SEQUENCE")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Deal Desk rules that stop the bleed — fund return first")
    badge(c, "5", "GUARDRAILS & INVEST", PURPLE)
    rr(c, 0.55 * inch, 1.1 * inch, 5.6 * inch, 3.7 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 11); c.drawString(0.75 * inch, 4.5 * inch, "PRICING GUARDRAILS")
    rules = [("Floor GM", "≥22% / 26% / 28% by tier"), ("Max discount", "≤12% list · trade scope/term"),
             ("Transition", "KT mandatory · min 8% Y1 value"), ("Scope CR", "Unpaid cap 2% ACV"),
             ("Regional variance", "±15% of global rate card"), ("Credits", "Soft-ramp · Finance rev-rec")]
    y = 4.15 * inch
    for name, detail in rules:
        c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, y, name)
        c.setFillColor(TEXT); c.setFont("Helvetica", 9); c.drawString(0.75 * inch, y - 12, detail)
        y -= 32
    rr(c, 6.4 * inch, 3.0 * inch, 4.7 * inch, 1.8 * inch, SURFACE)
    c.setFillColor(SUCCESS); c.setFont("Helvetica-Bold", 11); c.drawString(6.6 * inch, 4.5 * inch, "FUND FIRST")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for i, line in enumerate(["Q1 Rate card · Deal Desk", "Q2 Catalogue · 2 pilots", "Q3–4 Packs · renewals", "Y2 Retire headcount-only"]):
        c.drawString(6.6 * inch, 4.15 * inch - i * 18, "▸  " + line)
    rr(c, 6.4 * inch, 1.1 * inch, 4.7 * inch, 1.7 * inch, SURFACE)
    c.setFillColor(DANGER); c.setFont("Helvetica-Bold", 11); c.drawString(6.6 * inch, 2.5 * inch, "STOP")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for i, line in enumerate(["Custom tooling per account", "Below-floor PoCs", "Regional bespoke rates", "Rebrand without Sales kit"]):
        c.drawString(6.6 * inch, 2.15 * inch - i * 16, "×  " + line)
    flow(c, 4); foot(c, 6); c.showPage()

    # 7 Align
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 6 OF 7 · CROSS-FUNCTIONAL PLAN AND FEEDBACK LOOP")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "One program — Sales, Marketing, Finance, regions")
    badge(c, "6", "ALIGN & MEASURE", ACCENT)
    funcs = [
        ("SALES / DEAL DESK", ACCENT, "≥80% bids on tier by Q2"),
        ("MARKETING", PURPLE, "One narrative by outcome tier"),
        ("FINANCE", GOLD, "Rev-rec · margin floors"),
        ("REGIONS / DELIVERY", SUCCESS, "Pilots hit catalogue"),
    ]
    fw = (W - 1.35 * inch) / 4
    for i, (name, col, outcome) in enumerate(funcs):
        x = 0.55 * inch + i * (fw + 0.08 * inch)
        rr(c, x, 2.6 * inch, fw, 2.0 * inch, SURFACE)
        c.setFillColor(col); c.rect(x, 2.6 * inch, 4, 2.0 * inch, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 8); c.drawString(x + 0.12 * inch, 4.3 * inch, name)
        wrap(c, outcome, x + 0.12 * inch, 3.9 * inch, fw - 0.25 * inch, size=10, color=TEXT, leading=12)
    rr(c, 0.55 * inch, 1.05 * inch, W - 1.1 * inch, 1.35 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, 2.1 * inch, "12-MONTH SCOREBOARD · QUARTERLY PRICING COUNCIL")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    c.drawString(0.75 * inch, 1.75 * inch, "Net sales Flat→+8–12%  ·  GM recover 3–4 percentage points (e.g. →25–27%)  ·  Tier adoption ≥80%")
    c.setFillColor(MUTED); c.setFont("Helvetica", 9)
    c.drawString(0.75 * inch, 1.4 * inch, "Feedback: monthly Deal Desk overrides → quarterly Pricing Council → annual sunset re-test")
    flow(c, 5); foot(c, 7); c.showPage()

    # 8 Ask
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 7 OF 7 · CLOSE AND HAND TO QUESTIONS")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Reposition. Land with tiers. Expand when value is proven.")
    badge(c, "7", "ASK", SUCCESS)
    rr(c, 0.55 * inch, H - 3.3 * inch, W - 1.1 * inch, 1.35 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, H - 2.15 * inch, "THE CALL")
    wrap(c, "Reposition AMS as Outcome-Tiered Application Operations (Stabilize → Optimize → Transform). Do not sunset. Do not evolve tooling without fixing the sellable unit. Fund guardrails in Q1; pilot two regions; measure with Pricing Council.",
         0.75 * inch, H - 2.5 * inch, W - 1.5 * inch, size=10)
    rr(c, 0.55 * inch, 1.15 * inch, 5.4 * inch, 2.0 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, 2.85 * inch, "THOUGHT PROCESS")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for i, line in enumerate(["1 Diagnose with five lenses", "2 Synthesize commercial root causes", "3 Decide with trade-offs → Reposition", "4 Package / guardrail / enable / measure"]):
        c.drawString(0.75 * inch, 2.5 * inch - i * 18, line)
    rr(c, 6.2 * inch, 1.15 * inch, 5.0 * inch, 2.0 * inch, SURFACE)
    c.setFillColor(SUCCESS); c.setFont("Helvetica-Bold", 9); c.drawString(6.4 * inch, 2.85 * inch, "ASK OF LEADERSHIP")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for i, line in enumerate(["Approve reposition + global rate card", "Stand up Deal Desk guardrails", "Authorize two regional pilots", "Name Product owner + steering"]):
        c.drawString(6.4 * inch, 2.5 * inch - i * 18, "▸  " + line)
    c.setFillColor(MUTED); c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, 0.95 * inch, "QUESTIONS · 5 MINUTES · Deep dive on any stage: HTML case study")
    flow(c, 6); foot(c, 8); c.showPage()

    c.save()
    print(f"Wrote {path}")


if __name__ == "__main__":
    build()
