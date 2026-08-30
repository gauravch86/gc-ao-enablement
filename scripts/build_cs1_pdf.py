#!/usr/bin/env python3
"""Case Study 1 executive PDF — matches debiased PPTX storyline."""

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


def rounded_rect(c, x, y, w, h, fill, radius=8):
    c.setFillColor(fill)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=0)


def footer(c, page, total=6):
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8)
    c.drawString(0.55 * inch, 0.35 * inch, "Case Study 1 · Reposition underperforming offering · Gaurav Chaudhary")
    c.drawRightString(W - 0.55 * inch, 0.35 * inch, f"{page} / {total}")


def label(c, text, y=None):
    if y is None:
        y = H - 0.45 * inch
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, y, text.upper())


def title(c, text, y=None):
    if y is None:
        y = H - 0.85 * inch
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(0.55 * inch, y, text)


def wrap_text(c, text, x, y, max_width, font="Helvetica", size=11, color=TEXT, leading=14):
    c.setFont(font, size)
    c.setFillColor(color)
    words = text.split()
    lines, cur = [], ""
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


def slide_bg(c):
    c.setFillColor(BG)
    c.rect(0, 0, W, H, fill=1, stroke=0)


def build(path="/workspace/case-study-1-reposition-executive.pdf"):
    c = canvas.Canvas(path, pagesize=PAGE)

    # 1
    slide_bg(c)
    label(c, "Strategic Product Manager  ·  Case Study 1  ·  10-minute commercial brief")
    title(c, "Reposition an underperforming portfolio offering")
    wrap_text(c, "Application Managed Services: strong technical delivery, but commercially stuck — flat net sales, declining gross margin, Sales cannot articulate value, inconsistent regional pricing.",
              0.55 * inch, H - 1.25 * inch, W - 1.1 * inch, size=10, color=MUTED, leading=13)

    cards = [
        ("NET SALES (2 YEARS)", "Flat", "Weak conversion · little attach", MUTED),
        ("GROSS MARGIN", "−4–6 points", "Discount · scope creep · transition overrun", DANGER),
        ("DELIVERY SATISFACTION", "Strong", "Execution is not the failure mode", SUCCESS),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (lbl, val, sub, col) in enumerate(cards):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rounded_rect(c, x, H - 3.3 * inch, cw, 1.4 * inch, SURFACE)
        c.setFillColor(MUTED); c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x + cw / 2, H - 2.15 * inch, lbl)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(x + cw / 2, H - 2.55 * inch, val)
        c.setFillColor(MUTED); c.setFont("Helvetica", 8)
        c.drawCentredString(x + cw / 2, H - 2.9 * inch, sub)

    rounded_rect(c, 0.55 * inch, H - 4.75 * inch, W - 1.1 * inch, 1.15 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, H - 3.8 * inch, "EXECUTIVE FRAMING")
    wrap_text(c, "Commercial and go-to-market problem — not a delivery turnaround. Sold as headcount; buyers purchase outcomes, cost predictability, and risk reduction. Decide evolve / reposition / sunset with evidence; fix packaging, pricing discipline, and Sales enablement.",
              0.75 * inch, H - 4.1 * inch, W - 1.5 * inch, size=10, color=TEXT, leading=13)

    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, 1.45 * inch, "AGENDA  ·  10 MINUTES  (+ 5 MINUTES QUESTIONS)")
    c.setFillColor(TEXT); c.setFont("Helvetica", 11)
    c.drawString(0.55 * inch, 1.15 * inch, "1  Diagnose   →   2  Decide   →   3  Package & price   →   4  Prioritize & align   →   5  Measure & ask")
    c.setFillColor(MUTED); c.setFont("Helvetica", 9)
    c.drawString(0.55 * inch, 0.85 * inch, "Leadership question: evolve, reposition, or sunset — with a plan Sales, Deal Desk, Finance, and regions can run.")
    footer(c, 1)
    c.showPage()

    # 2
    slide_bg(c)
    label(c, "Diagnosis framework  ·  Recommendation")
    title(c, "Strong delivery. Broken commercial model. → Reposition.")
    lenses = [
        ("WIN / LOSS", "15–20 deals: price-only vs value unclear"),
        ("MARGIN", "Discount · transition · unpaid scope"),
        ("COMPETITIVE", "Staff augmentation vs outcome peers"),
        ("CUSTOMER VOICE", "High satisfaction + low willingness to pay"),
        ("PORTFOLIO FIT", "18-month margin path? Pull-through?"),
    ]
    lw = (W - 1.3 * inch) / 5
    for i, (h, body) in enumerate(lenses):
        x = 0.55 * inch + i * (lw + 0.08 * inch)
        rounded_rect(c, x, H - 2.9 * inch, lw, 1.3 * inch, SURFACE)
        c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 0.1 * inch, H - 1.8 * inch, h)
        wrap_text(c, body, x + 0.1 * inch, H - 2.1 * inch, lw - 0.2 * inch, size=8, color=TEXT, leading=11)

    causes = [
        ("Commoditized packaging", "Level-1/2/3 staffing towers — procurement benches us vs systems integrators"),
        ("Pricing chaos", "Regional discounts · underpriced knowledge transfer · unpaid scope creep"),
        ("Value story gap", "Buyers want outcomes; Sales pitches headcount replacement"),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (h, body) in enumerate(causes):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rounded_rect(c, x, H - 4.4 * inch, cw, 1.2 * inch, SURFACE)
        c.setFillColor(DANGER); c.rect(x, H - 3.25 * inch, cw, 3, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 10); c.drawString(x + 0.12 * inch, H - 3.5 * inch, h)
        wrap_text(c, body, x + 0.12 * inch, H - 3.8 * inch, cw - 0.25 * inch, size=9, color=TEXT, leading=11)

    decisions = [
        ("SUNSET", "Only if no 18-month margin path.", MUTED, False),
        ("EVOLVE ONLY", "Tooling without new sellable unit.", MUTED, False),
        ("REPOSITION  ✓", "Stabilize → Optimize → Transform", SUCCESS, True),
    ]
    for i, (h, body, col, sel) in enumerate(decisions):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rounded_rect(c, x, 0.55 * inch, cw, 1.45 * inch, SURFACE)
        if sel:
            c.setStrokeColor(SUCCESS); c.setLineWidth(2)
            c.roundRect(x, 0.55 * inch, cw, 1.45 * inch, 8, fill=0, stroke=1)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 11)
        c.drawString(x + 0.15 * inch, 1.65 * inch, h)
        c.setFillColor(TEXT if sel else MUTED); c.setFont("Helvetica", 9)
        c.drawString(x + 0.15 * inch, 1.3 * inch, body)
    footer(c, 2)
    c.showPage()

    # 3
    slide_bg(c)
    label(c, "Revised pricing and packaging  ·  Value proposition")
    title(c, "Sell tiers and outcomes — not staffing towers")
    wrap_text(c, "Outcome-tiered managed applications — restore control, improve total cost of ownership where contracted, globally consistent Sales story with Deal Desk discipline.",
              0.55 * inch, H - 1.2 * inch, W - 1.1 * inch, size=10, color=MUTED, leading=12)

    tiers = [
        ("STABILIZE", GOLD, "Restore control",
         ["Service-level baseline · single accountability", "Top-offender program", "Platform + criticality class", "Floor gross margin ≥ 22%"]),
        ("OPTIMIZE", ACCENT, "Prove efficiency",
         ["Monitoring pack · runbook automation", "Outcome-band vs service levels", "Optional gain-share", "Floor gross margin ≥ 26%"]),
        ("TRANSFORM", SUCCESS, "Operating-model uplift",
         ["Continuous improvement · standard changes", "Milestone elements (Finance sign-off)", "Executive scorecard", "Floor gross margin ≥ 28%"]),
    ]
    tw = (W - 1.3 * inch) / 3
    for i, (name, col, tag, lines) in enumerate(tiers):
        x = 0.55 * inch + i * (tw + 0.1 * inch)
        rounded_rect(c, x, 1.6 * inch, tw, 3.4 * inch, SURFACE)
        c.setFillColor(col); c.rect(x, 1.6 * inch + 3.4 * inch - 4, tw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 12); c.drawString(x + 0.15 * inch, 4.65 * inch, name)
        c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(x + 0.15 * inch, 4.4 * inch, tag)
        c.setFillColor(TEXT); c.setFont("Helvetica", 9)
        for j, line in enumerate(lines):
            c.drawString(x + 0.15 * inch, 4.0 * inch - j * 16, "•  " + line)

    rounded_rect(c, 0.55 * inch, 0.55 * inch, W - 1.1 * inch, 0.85 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 8)
    c.drawString(0.75 * inch, 1.15 * inch, "MODULAR ATTACH  ·  PRICED  ·  MULTI-VENDOR GOVERNANCE ONLY WHEN WIN/LOSS SHOWS DEMAND")
    c.setFillColor(TEXT); c.setFont("Helvetica", 9)
    c.drawString(0.75 * inch, 0.85 * inch, "Mandatory knowledge-transfer on new logos · Optional top-offender / monitoring packs · Consolidation = applicable possibility, not required strategy")
    footer(c, 3)
    c.showPage()

    # 4
    slide_bg(c)
    label(c, "Commercial guardrails  ·  Investment sequencing")
    title(c, "Deal Desk rules that stop the bleed — fund return first")
    rounded_rect(c, 0.55 * inch, 0.55 * inch, 5.6 * inch, H - 1.7 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 11)
    c.drawString(0.75 * inch, H - 1.4 * inch, "PRICING GUARDRAILS")
    rules = [
        ("Floor gross margin", "Stabilize ≥22% · Optimize ≥26% · Transform ≥28%"),
        ("Maximum discount", "≤12% off list (≤8% renewals) — trade scope / term"),
        ("Transition pricing", "Knowledge-transfer mandatory · min 8% Year-1 value"),
        ("Scope change", "Unpaid work cap 2% annual contract value"),
        ("Regional variance", "Like-for-like within ±15% of global rate card"),
        ("Credits / penalties", "Soft-ramp M1–6 · Finance revenue-recognition sign-off"),
    ]
    y = H - 1.75 * inch
    for name, detail in rules:
        c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 10); c.drawString(0.75 * inch, y, name)
        c.setFillColor(TEXT); c.setFont("Helvetica", 9); c.drawString(0.75 * inch, y - 14, detail)
        y -= 40

    rounded_rect(c, 6.4 * inch, H - 4.15 * inch, 4.7 * inch, 2.65 * inch, SURFACE)
    c.setFillColor(SUCCESS); c.setFont("Helvetica-Bold", 11)
    c.drawString(6.6 * inch, H - 1.7 * inch, "FUND  ·  HIGH RETURN")
    fund = ["Q1: Rate card · tiers · Deal Desk toolkit", "Q2: Service-level catalogue · 2 pilots",
            "Q3–4: Packs · renewal step-up · scale", "Year 2: Retire headcount-only contracts"]
    y = H - 2.1 * inch
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for line in fund:
        c.drawString(6.6 * inch, y, "▸  " + line); y -= 22

    rounded_rect(c, 6.4 * inch, 0.55 * inch, 4.7 * inch, 1.85 * inch, SURFACE)
    c.setFillColor(DANGER); c.setFont("Helvetica-Bold", 11)
    c.drawString(6.6 * inch, 2.05 * inch, "STOP / DEPRIORITIZE")
    stops = ["Per-account custom tooling builds", "Below-floor proofs of concept",
             "Regional bespoke rate cards", "Rebrand without Sales / Deal Desk kit"]
    y = 1.7 * inch
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for line in stops:
        c.drawString(6.6 * inch, y, "×  " + line); y -= 18
    footer(c, 4)
    c.showPage()

    # 5
    slide_bg(c)
    label(c, "Cross-functional alignment  ·  Success metrics and feedback loop")
    title(c, "One program — Sales, Marketing, Finance, regions")
    funcs = [
        ("SALES / DEAL DESK", ACCENT, "≥80% bids on tier + rate card by Q2", "Discovery kit · ROI calculator · weekly reviews"),
        ("MARKETING", PURPLE, "One narrative: Application Ops by tier", "One-pagers · proof stories · retire staffing language"),
        ("FINANCE", GOLD, "Revenue recognition · margin floors", "Platform vs milestone memo · bid sign-off"),
        ("REGIONS / DELIVERY", SUCCESS, "Pilots hit service-level catalogue", "Knowledge-transfer playbook · renewal step-ups"),
    ]
    fw = (W - 1.35 * inch) / 4
    for i, (name, col, outcome, arts) in enumerate(funcs):
        x = 0.55 * inch + i * (fw + 0.08 * inch)
        rounded_rect(c, x, H - 3.65 * inch, fw, 2.1 * inch, SURFACE)
        c.setFillColor(col); c.rect(x, H - 3.65 * inch, 4, 2.1 * inch, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.drawString(x + 0.12 * inch, H - 1.8 * inch, name)
        wrap_text(c, outcome, x + 0.12 * inch, H - 2.15 * inch, fw - 0.25 * inch, size=8, color=TEXT, leading=11)
        wrap_text(c, arts, x + 0.12 * inch, H - 2.7 * inch, fw - 0.25 * inch, size=8, color=MUTED, leading=10)

    rounded_rect(c, 0.55 * inch, 0.55 * inch, W - 1.1 * inch, 2.1 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 10)
    c.drawString(0.75 * inch, 2.3 * inch, "12-MONTH SCOREBOARD  ·  Quarterly Pricing Council")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    c.drawString(0.75 * inch, 1.9 * inch, "Net sales  Flat → +8–12%     Gross margin  +3–4 points     Tier adoption  ≥80%")
    c.drawString(0.75 * inch, 1.6 * inch, "Price variance  <15%     Wins on value/total-cost  ≥40%     Renewal step-up  ≥25%")
    c.setFillColor(MUTED); c.setFont("Helvetica", 9)
    wrap_text(c, "Feedback: monthly Deal Desk overrides + change requests → quarterly Pricing Council (±3% rate card) → annual sunset re-test",
              0.75 * inch, 1.2 * inch, W - 1.5 * inch, size=9, color=MUTED, leading=12)
    footer(c, 5)
    c.showPage()

    # 6
    slide_bg(c)
    label(c, "Recommendation  ·  Ask")
    title(c, "Reposition. Land with tiers. Expand when value is proven.")
    rounded_rect(c, 0.55 * inch, H - 3.15 * inch, W - 1.1 * inch, 1.65 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, H - 1.7 * inch, "THE CALL")
    wrap_text(c, "Reposition Application Managed Services as Outcome-Tiered Application Operations (Stabilize → Optimize → Transform). Do not sunset a delivery-capable franchise. Do not evolve tooling without fixing the sellable unit. Fund guardrails in Quarter 1; pilot in two regions; measure with a quarterly Pricing Council.",
              0.75 * inch, H - 2.05 * inch, W - 1.5 * inch, size=10, color=TEXT, leading=13)

    rounded_rect(c, 0.55 * inch, 1.35 * inch, 5.4 * inch, 2.2 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, 3.25 * inch, "THOUGHT PROCESS")
    for i, line in enumerate([
        "Evidence before instinct — five-lens diagnostic",
        "Separate delivery quality from commercial failure",
        "Decide with trade-offs; fund Sales motion first",
        "Optional packs only when win/loss data says so",
    ]):
        c.setFillColor(TEXT); c.setFont("Helvetica", 10)
        c.drawString(0.75 * inch, 2.85 * inch - i * 20, "•  " + line)

    rounded_rect(c, 6.2 * inch, 1.35 * inch, 5.0 * inch, 2.2 * inch, SURFACE)
    c.setFillColor(SUCCESS); c.setFont("Helvetica-Bold", 9)
    c.drawString(6.4 * inch, 3.25 * inch, "ASK OF LEADERSHIP")
    asks = ["Approve reposition + global rate card", "Stand up Deal Desk guardrails this quarter",
            "Authorize two regional pilots", "Name Product owner + monthly steering"]
    y = 2.85 * inch
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    for a in asks:
        c.drawString(6.4 * inch, y, "▸  " + a); y -= 22

    c.setFillColor(MUTED); c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(W / 2, 0.85 * inch, "QUESTIONS  ·  5 MINUTES     ·     Deep dive: HTML case study")
    footer(c, 6)
    c.showPage()
    c.save()
    print(f"Wrote {path}")


if __name__ == "__main__":
    build()
