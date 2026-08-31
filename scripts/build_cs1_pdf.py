#!/usr/bin/env python3
"""Case Study 1 executive PDF — 8-slide sequential interview flow (no timers)."""

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
    rr(c, 0.55 * inch, y, 2.7 * inch, 0.35 * inch, SURFACE)
    c.setFillColor(color)
    c.rect(0.55 * inch, y, 0.07 * inch, 0.35 * inch, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(color)
    c.drawString(0.75 * inch, y + 0.1 * inch, f"STAGE {num}  ·  {name}")


def build(path="/workspace/case-study-1-reposition-executive.pdf"):
    c = canvas.Canvas(path, pagesize=PAGE)

    # 1 Situation
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STRATEGIC PRODUCT MANAGER · CASE STUDY 1 · DUAL MANDATE: NEW REVENUE + MARGIN")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 17)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Grow topline. Protect EBIT. Reposition the offering.")
    wrap(
        c,
        "Application Managed Services: strong delivery, but commercially stuck — flat net sales, gross margin 28% → 22–24%, Sales cannot sell value, regions discount without trading scope.",
        0.55 * inch,
        H - 1.15 * inch,
        W - 1.1 * inch,
        size=10,
        color=MUTED,
    )
    cards = [
        ("SPM OWNS — NEW SALES", "Topline", "Net sales · attach · expansion", ACCENT),
        ("SPM OWNS — PROFITABILITY", "Margin / EBIT", "Gross margin · deal quality", GOLD),
        ("TODAY’S PARADOX", "Strong NPS", "Flat sales · eroding GM", DANGER),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (lbl, val, sub, col) in enumerate(cards):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rr(c, x, H - 3.15 * inch, cw, 1.3 * inch, SURFACE)
        c.setFillColor(MUTED)
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(x + cw / 2, H - 2.1 * inch, lbl)
        c.setFillColor(col)
        c.setFont("Helvetica-Bold", 15)
        c.drawCentredString(x + cw / 2, H - 2.45 * inch, val)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8)
        c.drawCentredString(x + cw / 2, H - 2.8 * inch, sub)
    rr(c, 0.55 * inch, H - 4.55 * inch, W - 1.1 * inch, 1.15 * inch, SURFACE)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, H - 3.6 * inch, "THE JOB IN THIS ROLE")
    wrap(
        c,
        "Accountable for new revenue (new sales / topline) and profitability (gross margin / EBIT). Discounts sustain revenue only when traded for additional scope, term, or customer investment — never free margin bleed.",
        0.75 * inch,
        H - 3.9 * inch,
        W - 1.5 * inch,
        size=10,
    )
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, 1.55 * inch, "WHAT LEADERSHIP GETS IN 10 MINUTES")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    c.drawString(0.55 * inch, 1.25 * inch, "1 Diagnose → 2 Findings → 3 Reposition  ·  4 Packaging & value prop  ·  5 Guardrails + invest ROI")
    c.drawString(0.55 * inch, 1.0 * inch, "6 Sales · Marketing · Finance · regions + metrics feedback loop  ·  7 Asks")
    foot(c, 1)
    c.showPage()

    # 2 Diagnose
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 1 OF 7 · EVIDENCE BEFORE ANY PORTFOLIO CALL")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Five-lens diagnosis — what we would gather")
    badge(c, "1", "FIVE-LENS DIAGNOSIS", ACCENT)
    lenses = [
        ("A  WIN / LOSS", "Top 10 accounts. Tag price-only vs value unclear. Trigger: value unclear >30% → positioning broken."),
        ("B  MARGIN", "GM waterfall. Was discount traded for scope / term / attach? Bleed without trade → fix."),
        ("C  COMPETITIVE", "Staff aug vs outcome peers. Buyers ask for tiers → reposition."),
        ("D  CUSTOMER VOICE", "High satisfaction + low willingness to pay = story gap."),
        ("E  PORTFOLIO FIT", "18-month path to topline + margin? Pull-through? Sunset only if both fail."),
    ]
    lw = (W - 1.3 * inch) / 5
    for i, (h, body) in enumerate(lenses):
        x = 0.55 * inch + i * (lw + 0.08 * inch)
        rr(c, x, 1.1 * inch, lw, 3.7 * inch, SURFACE)
        c.setFillColor(ACCENT)
        c.rect(x, 1.1 * inch + 3.7 * inch - 4, lw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 0.08 * inch, 4.5 * inch, h)
        wrap(c, body, x + 0.08 * inch, 4.15 * inch, lw - 0.16 * inch, size=8, leading=11)
    flow(c, 0)
    foot(c, 2)
    c.showPage()

    # 3 Findings
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 2 OF 7 · ILLUSTRATIVE SYNTHESIS")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Root causes — commercial, not delivery")
    badge(c, "2", "ILLUSTRATIVE FINDINGS", DANGER)
    causes = [
        ("1  WRONG SELLABLE UNIT", "Sold as L1/L2/L3 staffing. Procurement benches rate-per-person. Hits topline: flat sales."),
        ("2  DISCOUNT WITHOUT TRADE", "Regions discount with no floor and no scope trade. Should trade discount for scope/term/investment. Hits margin: 28%→22–24%."),
        ("3  VALUE STORY GAP", "Buyers want outcomes; Sales pitches headcount. Hits both new revenue and price defence."),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (h, body) in enumerate(causes):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rr(c, x, 1.1 * inch, cw, 3.7 * inch, SURFACE)
        c.setFillColor(DANGER)
        c.rect(x, 1.1 * inch + 3.7 * inch - 4, cw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x + 0.15 * inch, 4.5 * inch, h)
        wrap(c, body, x + 0.15 * inch, 4.1 * inch, cw - 0.3 * inch, size=11, leading=15)
    flow(c, 1)
    foot(c, 3)
    c.showPage()

    # 4 Decide
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 3 OF 7 · PORTFOLIO CALL")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Decision → Reposition (grow sales + restore margin)")
    badge(c, "3", "DECISION FRAMEWORK", SUCCESS)
    decisions = [
        ("SUNSET", MUTED, False, "Rejects footprint. Only if no 18-month path to sales + margin AND no pull-through. Not indicated."),
        ("EVOLVE ONLY", MUTED, False, "More tooling, same sellable unit. Fixes neither Sales narrative nor discount-without-trade."),
        ("REPOSITION  ✓", SUCCESS, True, "Outcome-tiered Application Ops. Stabilize → Optimize → Transform. New value prop + Deal Desk rules."),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (h, col, sel, body) in enumerate(decisions):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rr(c, x, 1.1 * inch, cw, 3.7 * inch, SURFACE)
        if sel:
            c.setStrokeColor(SUCCESS)
            c.setLineWidth(2.5)
            c.roundRect(x, 1.1 * inch, cw, 3.7 * inch, 8, fill=0, stroke=1)
        c.setFillColor(col)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x + 0.15 * inch, 4.5 * inch, h)
        wrap(c, body, x + 0.15 * inch, 4.05 * inch, cw - 0.3 * inch, size=11, color=TEXT if sel else MUTED, leading=15)
    flow(c, 2)
    foot(c, 4)
    c.showPage()

    # 5 Package
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 4 OF 7 · REVISED PACKAGING, PRICING & VALUE PROPOSITION")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "What Sales sells — tiers, outcomes, clearer price")
    badge(c, "4", "PACKAGE & PRICE", GOLD)
    tiers = [
        ("STABILIZE", GOLD, "Land — restore control", ["Service-level baseline", "Platform + criticality", "Floor GM ≥ 22% · new logos"]),
        ("OPTIMIZE", ACCENT, "Expand — prove efficiency", ["Monitoring / automation", "Outcome-band pricing", "Floor GM ≥ 26% · attach"]),
        ("TRANSFORM", SUCCESS, "Deepen — operating uplift", ["Continuous improvement", "Milestone (Finance OK)", "Floor GM ≥ 28% · renewals"]),
    ]
    tw = (W - 1.3 * inch) / 3
    for i, (name, col, tag, lines) in enumerate(tiers):
        x = 0.55 * inch + i * (tw + 0.1 * inch)
        rr(c, x, 1.85 * inch, tw, 2.95 * inch, SURFACE)
        c.setFillColor(col)
        c.rect(x, 1.85 * inch + 2.95 * inch - 4, tw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x + 0.15 * inch, 4.5 * inch, name)
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 0.15 * inch, 4.2 * inch, tag)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 10)
        for j, line in enumerate(lines):
            c.drawString(x + 0.15 * inch, 3.8 * inch - j * 18, "·  " + line)
    rr(c, 0.55 * inch, 1.05 * inch, W - 1.1 * inch, 0.65 * inch, SURFACE)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(0.75 * inch, 1.45 * inch, "VALUE PROP: Outcome-tiered Application Operations — restore control, improve TCO where contracted, global Sales story with Deal Desk discipline that grows revenue without giving margin away.")
    flow(c, 3)
    foot(c, 5)
    c.showPage()

    # 6 Guardrails
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 5 OF 7 · GUARDRAILS + INVESTMENT ROI SEQUENCE")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Deal Desk rules + what we fund (and stop)")
    badge(c, "5", "GUARDRAILS & INVEST", PURPLE)
    rr(c, 0.55 * inch, 1.1 * inch, 5.6 * inch, 3.7 * inch, SURFACE)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(0.75 * inch, 4.5 * inch, "COMMERCIAL GUARDRAILS")
    rules = [
        ("Floor GM", "≥22% / 26% / 28% by tier"),
        ("Discount = trade", "≤12% only if traded for scope, term ≥24 mo, or attach"),
        ("Sustain topline", "Discount to keep/win revenue must raise contracted value"),
        ("Transition", "KT mandatory · min 8% Y1 value"),
        ("Scope CR", "Unpaid cap 2% ACV"),
        ("Regional variance", "±15% of global rate card"),
    ]
    y = 4.15 * inch
    for name, detail in rules:
        c.setFillColor(ACCENT)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(0.75 * inch, y, name)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 9)
        c.drawString(0.75 * inch, y - 12, detail)
        y -= 30
    rr(c, 6.4 * inch, 3.0 * inch, 4.7 * inch, 1.8 * inch, SURFACE)
    c.setFillColor(SUCCESS)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(6.6 * inch, 4.5 * inch, "FUND FIRST · MAX ROI")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for i, line in enumerate(["Q1 Rate card · Deal Desk playbook", "Q2 Catalogue · 2 pilots", "Q3–4 Packs · renewals", "Y2 Retire headcount-only"]):
        c.drawString(6.6 * inch, 4.15 * inch - i * 18, "▸  " + line)
    rr(c, 6.4 * inch, 1.1 * inch, 4.7 * inch, 1.7 * inch, SURFACE)
    c.setFillColor(DANGER)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(6.6 * inch, 2.5 * inch, "DEPRIORITIZE / STOP")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for i, line in enumerate(["Custom tooling before global kit", "Below-floor PoCs", "Regional bespoke rates", "Rebrand without Sales kit"]):
        c.drawString(6.6 * inch, 2.15 * inch - i * 16, "×  " + line)
    flow(c, 4)
    foot(c, 6)
    c.showPage()

    # 7 Align
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 6 OF 7 · ALIGNMENT + METRICS + FEEDBACK LOOP")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "One program — Sales, Marketing, Finance, regions")
    badge(c, "6", "ALIGN & MEASURE", ACCENT)
    funcs = [
        ("SALES / DEAL DESK", ACCENT, "≥80% bids on tier by Q2"),
        ("MARKETING", PURPLE, "One narrative by outcome tier"),
        ("FINANCE", GOLD, "Rev-rec · margin floors · EBIT"),
        ("REGIONS / ENABLE", SUCCESS, "Pilots + renewal step-ups"),
    ]
    fw = (W - 1.35 * inch) / 4
    for i, (name, col, outcome) in enumerate(funcs):
        x = 0.55 * inch + i * (fw + 0.08 * inch)
        rr(c, x, 2.7 * inch, fw, 1.9 * inch, SURFACE)
        c.setFillColor(col)
        c.rect(x, 2.7 * inch, 4, 1.9 * inch, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 0.12 * inch, 4.3 * inch, name)
        wrap(c, outcome, x + 0.12 * inch, 3.9 * inch, fw - 0.25 * inch, size=10, color=TEXT, leading=12)
    rr(c, 0.55 * inch, 1.05 * inch, W - 1.1 * inch, 1.45 * inch, SURFACE)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, 2.2 * inch, "12-MONTH SCOREBOARD · BOTH SPM LEVERS · QUARTERLY PRICING COUNCIL")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 9)
    c.drawString(0.75 * inch, 1.85 * inch, "NEW REVENUE: Net sales Flat→+8–12% · Tier ≥80% · Value/TCO wins ≥40% · Renewal step-up ≥25%")
    c.drawString(0.75 * inch, 1.55 * inch, "MARGIN/EBIT: Recover 3–4 GM pts · Price variance <15% · ≥70% discounts have documented scope/term trade")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8)
    c.drawString(0.75 * inch, 1.25 * inch, "Feedback: monthly Deal Desk overrides → quarterly Pricing Council → annual sunset / reposition re-test")
    flow(c, 5)
    foot(c, 7)
    c.showPage()

    # 8 Ask
    bg(c)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 7 OF 7 · CLOSE")
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Reposition. Win new revenue. Restore margin.")
    badge(c, "7", "ASK", SUCCESS)
    rr(c, 0.55 * inch, H - 3.25 * inch, W - 1.1 * inch, 1.3 * inch, SURFACE)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, H - 2.15 * inch, "THE CALL")
    wrap(
        c,
        "Reposition AMS as Outcome-Tiered Application Operations. Discount to sustain topline only when traded for scope/term/investment. Fund Deal Desk in Q1; two regional pilots; Pricing Council.",
        0.75 * inch,
        H - 2.5 * inch,
        W - 1.5 * inch,
        size=10,
    )
    rr(c, 0.55 * inch, 1.15 * inch, 5.4 * inch, 2.0 * inch, SURFACE)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, 2.85 * inch, "CASE ASKS — COVERED")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for i, line in enumerate(
        [
            "✓ Packaging · pricing · value proposition",
            "✓ Commercial / Deal Desk guardrails",
            "✓ Invest sequence + what to stop",
            "✓ Sales · Marketing · Finance · regions",
            "✓ Metrics + Pricing Council feedback loop",
        ]
    ):
        c.drawString(0.75 * inch, 2.5 * inch - i * 16, line)
    rr(c, 6.2 * inch, 1.15 * inch, 5.0 * inch, 2.0 * inch, SURFACE)
    c.setFillColor(SUCCESS)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(6.4 * inch, 2.85 * inch, "ASK OF LEADERSHIP")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for i, line in enumerate(
        [
            "Approve reposition + global rate card",
            "Stand up Deal Desk trade rules",
            "Authorize two regional pilots",
            "Name owner + steering on sales + GM",
        ]
    ):
        c.drawString(6.4 * inch, 2.5 * inch - i * 18, "▸  " + line)
    flow(c, 6)
    foot(c, 8)
    c.showPage()

    c.save()
    print(f"Wrote {path}")


if __name__ == "__main__":
    build()
