#!/usr/bin/env python3
"""Case Study 2 executive PDF — 8-slide sequential interview flow."""

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

BG, SURFACE = HexColor("#0F1419"), HexColor("#1A2332")
TEXT, MUTED = HexColor("#E6EDF3"), HexColor("#8B9CB3")
ACCENT, GOLD, SUCCESS = HexColor("#58A6FF"), HexColor("#D4A853"), HexColor("#3FB950")
DANGER, PURPLE, CYAN = HexColor("#F87171"), HexColor("#A371F7"), HexColor("#38BDF8")
W, H = landscape(A4)
TOTAL = 8


def bg(c):
    c.setFillColor(BG); c.rect(0, 0, W, H, fill=1, stroke=0)


def rr(c, x, y, w, h, fill, radius=8):
    c.setFillColor(fill); c.roundRect(x, y, w, h, radius, fill=1, stroke=0)


def foot(c, page):
    c.setFillColor(MUTED); c.setFont("Helvetica", 8)
    c.drawString(0.55 * inch, 0.28 * inch, "Case Study 2 · Reusable Autonomous Operations offering · Gaurav Chaudhary")
    c.drawRightString(W - 0.55 * inch, 0.28 * inch, f"{page} / {TOTAL}")


def wrap(c, text, x, y, mw, size=10, color=TEXT, leading=13):
    c.setFont("Helvetica", size); c.setFillColor(color)
    words, lines, cur = text.split(), [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if c.stringWidth(test, "Helvetica", size) <= mw:
            cur = test
        else:
            if cur: lines.append(cur)
            cur = w
    if cur: lines.append(cur)
    for i, line in enumerate(lines):
        c.drawString(x, y - i * leading, line)


def flow(c, active):
    steps = ["1 Unlock", "2 Autonomy", "3 Measure", "4 Reuse", "5 Enable", "6 P&L", "7 Ask"]
    ww, gap, y = 1.45 * inch, 0.1 * inch, 0.55 * inch
    for i, name in enumerate(steps):
        x = 0.55 * inch + i * (ww + gap)
        rr(c, x, y, ww, 0.28 * inch, SURFACE, 4)
        if i == active:
            c.setFillColor(GOLD); c.rect(x, y + 0.23 * inch, ww, 0.05 * inch, fill=1, stroke=0)
        c.setFillColor(GOLD if i == active else MUTED)
        c.setFont("Helvetica-Bold" if i == active else "Helvetica", 8)
        c.drawCentredString(x + ww / 2, y + 0.08 * inch, name)


def badge(c, num, name, color):
    rr(c, 0.55 * inch, H - 1.55 * inch, 2.7 * inch, 0.35 * inch, SURFACE)
    c.setFillColor(color); c.rect(0.55 * inch, H - 1.55 * inch, 0.07 * inch, 0.35 * inch, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, H - 1.45 * inch, f"STAGE {num}  ·  {name}")


def build(path="/workspace/case-study-2-ao-offering-executive.pdf"):
    c = canvas.Canvas(path, pagesize=landscape(A4))

    # 1
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STRATEGIC PRODUCT MANAGER · CASE STUDY 2 · 10-MINUTE BRIEF · LIVED EXPERIENCE")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 17)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Reusable Autonomous Operations on customer KPIs / OKRs")
    wrap(c, "Large telecom operator · multi-vendor · multi-domain. Propose reusable managed services: reactive → predictive → autonomous.",
         0.55 * inch, H - 1.15 * inch, W - 1.1 * inch, size=10, color=MUTED)
    rr(c, 0.55 * inch, H - 2.9 * inch, W - 1.1 * inch, 1.05 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, H - 2.05 * inch, "NORTH STAR")
    wrap(c, "End-user experience and Net Promoter Score · contractual performance-indicator attainment · market trust. Total cost of ownership funds the outcome — it is not the outcome.",
         0.75 * inch, H - 2.4 * inch, W - 1.5 * inch, size=10)
    for i, (h, b, col) in enumerate([("GIVEN", "Ericsson + third-party under Managed Services", GOLD), ("METHOD", "Design→Deploy→Operate→Assure→Improve · pilots first", ACCENT), ("PROOF", "Telefónica Argentina · Three UK", SUCCESS)]):
        x = 0.55 * inch + i * ((W - 1.3 * inch) / 3 + 0.1 * inch); ww = (W - 1.3 * inch) / 3
        rr(c, x, 1.55 * inch, ww, 1.2 * inch, SURFACE)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 10); c.drawString(x + 0.12 * inch, 2.45 * inch, h)
        wrap(c, b, x + 0.12 * inch, 2.15 * inch, ww - 0.25 * inch, size=9)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(0.55 * inch, 1.2 * inch, "STORYLINE · ONE STAGE PER SLIDE")
    c.setFillColor(TEXT); c.setFont("Helvetica", 10)
    c.drawString(0.55 * inch, 0.95 * inch, "1 Unlock → 2 Autonomy → 3 Measure → 4 Reuse → 5 Enable → 6 P&L → 7 Ask")
    foot(c, 1); c.showPage()

    # 2 Unlock
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 1 OF 7 · ECONOMIC UNLOCK (ENABLER, NOT THE NORTH STAR)")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Vendor consolidation unlocks autonomy — it is not the goal")
    badge(c, "1", "ECONOMIC UNLOCK", GOLD)
    for i, (h, v, s, col) in enumerate([("IT OPS BUDGET", "$50m", "Multi-vendor run cost", GOLD), ("AFTER CONSOLIDATION", "$40m", "~20% TCO reduction", SUCCESS), ("GOVERNANCE", "−20%", "Vendor / SLA overhead", ACCENT)]):
        x = 0.55 * inch + i * ((W - 1.3 * inch) / 3 + 0.1 * inch); ww = (W - 1.3 * inch) / 3
        rr(c, x, H - 3.5 * inch, ww, 1.35 * inch, SURFACE)
        c.setFillColor(MUTED); c.setFont("Helvetica-Bold", 8); c.drawCentredString(x + ww / 2, H - 2.4 * inch, h)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 18); c.drawCentredString(x + ww / 2, H - 2.85 * inch, v)
        c.setFillColor(MUTED); c.setFont("Helvetica", 8); c.drawCentredString(x + ww / 2, H - 3.2 * inch, s)
    for i, (h, body, col) in enumerate([("PAIN TODAY", "Finger-pointing · noisy KPIs · slow cross-domain correlation", DANGER), ("UNLOCKS", "Single accountability · observability baseline · lower governance", SUCCESS), ("CHAIN", "L1 → L2/L3 → infra/cloud → change → SDM → problem/sunset", CYAN)]):
        x = 0.55 * inch + i * ((W - 1.3 * inch) / 3 + 0.1 * inch); ww = (W - 1.3 * inch) / 3
        rr(c, x, 1.05 * inch, ww, 1.7 * inch, SURFACE)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 10); c.drawString(x + 0.12 * inch, 2.45 * inch, h)
        wrap(c, body, x + 0.12 * inch, 2.1 * inch, ww - 0.25 * inch, size=9)
    flow(c, 0); foot(c, 2); c.showPage()

    # 3 Autonomy
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 2 OF 7 · OUTCOME PATH AND PRODUCTIZED LIFECYCLE")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Reactive → predictive → autonomous — with gates")
    badge(c, "2", "AUTONOMY PATH", ACCENT)
    for i, (h, body, col) in enumerate([("REACTIVE", "Ticket storms · multi-vendor blame", DANGER), ("OBSERVE", "One health view across domains", ACCENT), ("PREDICT", "Anomaly detection before impact", PURPLE), ("AUTONOMOUS", "Self-heal · human-in-loop → policy auto", SUCCESS)]):
        x = 0.55 * inch + i * ((W - 1.35 * inch) / 4 + 0.08 * inch); ww = (W - 1.35 * inch) / 4
        rr(c, x, 2.4 * inch, ww, 2.0 * inch, SURFACE)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 9); c.drawString(x + 0.1 * inch, 4.1 * inch, h)
        wrap(c, body, x + 0.1 * inch, 3.7 * inch, ww - 0.2 * inch, size=9)
    rr(c, 0.55 * inch, 1.05 * inch, W - 1.1 * inch, 1.15 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawString(0.75 * inch, 1.95 * inch, "LIFECYCLE")
    wrap(c, "Design → Deploy → Operate → Assure → Improve — one productized journey; intent & policy replace tribal manual action over time.",
         0.75 * inch, 1.6 * inch, W - 1.5 * inch, size=10)
    flow(c, 1); foot(c, 3); c.showPage()

    # 4 Measure
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 3 OF 7 · DATA-DRIVEN SIZING AND CONTRACTUAL PERFORMANCE INDICATORS")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Size from incidents. Contract soft-ramp.")
    badge(c, "3", "MEASURE", CYAN)
    rr(c, 0.55 * inch, 1.1 * inch, 5.5 * inch, 3.7 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 11); c.drawString(0.75 * inch, 4.5 * inch, "DISCOVERY → CONTRACT")
    for i, l in enumerate(["6–12 months incident ingest", "SME + AI → maturity per component", "Top-offender backlog", "CPI catalogue with soft-ramp", "Same scoreboard for roadmap & P&L"]):
        c.setFillColor(TEXT); c.setFont("Helvetica", 10); c.drawString(0.75 * inch, 4.1 * inch - i * 22, "•  " + l)
    rr(c, 6.3 * inch, 1.1 * inch, 5.3 * inch, 3.7 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 11); c.drawString(6.5 * inch, 4.5 * inch, "ILLUSTRATIVE CPI CREDITS")
    for i, l in enumerate(["Availability credits ~20%", "P0/P1 restore ≤4h (~10%)", "Change success ≥98%", "Proactive problems ≥20%", "Repeat P0–P2 ≤10%"]):
        c.setFillColor(TEXT); c.setFont("Helvetica", 10); c.drawString(6.5 * inch, 4.1 * inch - i * 22, "•  " + l)
    flow(c, 2); foot(c, 4); c.showPage()

    # 5 Reuse
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 4 OF 7 · ARCHITECTURE FOR SCALE ACROSS LOGOS")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Productize reuse — next customer inherits the library")
    badge(c, "4", "REUSE LIBRARY", PURPLE)
    items = [("METRICS & RUNBOOKS", "Drag-drop packs — not rebuild", ACCENT), ("SELF-HEAL CLASSES", "Versioned known failure classes", SUCCESS),
             ("ERICSSON + 3PP", "Thin adapters under one MS model", GOLD), ("TRADE-OFF", "Thin adapters > rip-and-replace", CYAN),
             ("IMPROVE → GLOBAL", "Library compounds across regions", PURPLE), ("FUNDING", "Account seeds IP; portfolio R&D absorbs platform", DANGER)]
    for i, (h, b, col) in enumerate(items):
        row, col_i = divmod(i, 3)
        x = 0.55 * inch + col_i * ((W - 1.3 * inch) / 3 + 0.1 * inch); ww = (W - 1.3 * inch) / 3
        y = 3.0 * inch - row * 1.85 * inch
        rr(c, x, y, ww, 1.65 * inch, SURFACE)
        c.setFillColor(col); c.rect(x, y + 1.65 * inch - 3, ww, 3, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 9); c.drawString(x + 0.12 * inch, y + 1.25 * inch, h)
        wrap(c, b, x + 0.12 * inch, y + 0.9 * inch, ww - 0.25 * inch, size=9)
    flow(c, 3); foot(c, 5); c.showPage()

    # 6 Enable
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 5 OF 7 · ADOPTION FOR CUSTOMER OPS AND ERICSSON REGIONS")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 15)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Enablement cadence — shortened, gated, shared")
    badge(c, "5", "ENABLE", SUCCESS)
    for i, (m, t) in enumerate([("M0–M3", "Train · baseline · shadow start"), ("M3–M6", "Forward shadow · soft-ramp · champions"),
                                ("M6–M9", "Reverse shadow · HITL · region kit"), ("M12–M15", "Auto classes · Nth-logo · reuse")]):
        x = 0.55 * inch + i * ((W - 1.35 * inch) / 4 + 0.08 * inch); ww = (W - 1.35 * inch) / 4
        rr(c, x, H - 3.4 * inch, ww, 1.25 * inch, SURFACE)
        c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 11); c.drawString(x + 0.1 * inch, H - 2.35 * inch, m)
        wrap(c, t, x + 0.1 * inch, H - 2.7 * inch, ww - 0.2 * inch, size=8)
    for i, (h, lines, col) in enumerate([("CUSTOMER OPS", "Train · shadow→auto · upskill narrative", ACCENT),
                                         ("ERICSSON REGIONS", "Library-first · Deal Desk tiers · reuse %", PURPLE),
                                         ("PILOTS", "Pass-gates · M9–M12 stabilize · then auto", GOLD)]):
        x = 0.55 * inch + i * ((W - 1.3 * inch) / 3 + 0.1 * inch); ww = (W - 1.3 * inch) / 3
        rr(c, x, 1.05 * inch, ww, 1.55 * inch, SURFACE)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 10); c.drawString(x + 0.12 * inch, 2.3 * inch, h)
        wrap(c, lines, x + 0.12 * inch, 1.95 * inch, ww - 0.25 * inch, size=9)
    flow(c, 4); foot(c, 6); c.showPage()

    # 7 P&L
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 6 OF 7 · FIVE-YEAR MANAGED-SERVICES PROFIT-AND-LOSS")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 14)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Full-time-equivalent ramp · site-mix · tools · penalty lot")
    badge(c, "6", "P&L CONSTRUCT", GOLD)
    wrap(c, "Baseline B. Target X = B−30%. Customer commit ~20%; size at 30% → ~10-point structural EBIT gap.",
         0.55 * inch, H - 1.7 * inch, W - 1.1 * inch, size=10, color=MUTED)
    for i, (y, v) in enumerate([("Y1", "X+20%"), ("Y2", "X+10%"), ("Y3", "X"), ("Y4", "X−10%"), ("Y5", "X−20%")]):
        x = 0.55 * inch + i * ((W - 1.3 * inch) / 5 + 0.08 * inch); ww = (W - 1.3 * inch) / 5
        rr(c, x, H - 3.5 * inch, ww, 1.15 * inch, SURFACE)
        c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 10); c.drawCentredString(x + ww / 2, H - 2.55 * inch, y)
        c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 14); c.drawCentredString(x + ww / 2, H - 2.95 * inch, v)
    for i, (h, v, col) in enumerate([("① Labour", "+10%", SUCCESS), ("② Site-mix", "+10%", SUCCESS), ("③ Tools", "−5%", DANGER), ("④ License", "+5%", SUCCESS), ("⑤ Penalty lot", "+5% if met", SUCCESS)]):
        x = 0.55 * inch + i * ((W - 1.3 * inch) / 5 + 0.08 * inch); ww = (W - 1.3 * inch) / 5
        rr(c, x, 1.05 * inch, ww, 1.55 * inch, SURFACE)
        c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 9); c.drawCentredString(x + ww / 2, 2.3 * inch, h)
        c.setFillColor(col); c.setFont("Helvetica-Bold", 12); c.drawCentredString(x + ww / 2, 1.8 * inch, v)
    flow(c, 5); foot(c, 7); c.showPage()

    # 8 Ask
    bg(c)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, H - 0.4 * inch, "STORYLINE STEP 7 OF 7 · LIVED PROOF AND RECOMMENDATION")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawString(0.55 * inch, H - 0.8 * inch, "Lived proof. One recommendation.")
    badge(c, "7", "PROOF & ASK", SUCCESS)
    rr(c, 0.55 * inch, H - 3.5 * inch, 5.4 * inch, 1.55 * inch, SURFACE)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 10); c.drawString(0.75 * inch, H - 2.15 * inch, "TELEFÓNICA ARGENTINA — ~$100m")
    wrap(c, "Top-offender discipline · KPI transparency · accountability → renewal", 0.75 * inch, H - 2.5 * inch, 5.0 * inch, size=9)
    rr(c, 6.2 * inch, H - 3.5 * inch, 5.0 * inch, 1.55 * inch, SURFACE)
    c.setFillColor(ACCENT); c.setFont("Helvetica-Bold", 10); c.drawString(6.4 * inch, H - 2.15 * inch, "THREE UK — ~$125m")
    wrap(c, "Consolidation · ~75 CPIs · SPOG + self-heal · 474→336 FTE · lowest P1s", 6.4 * inch, H - 2.5 * inch, 4.6 * inch, size=9)
    rr(c, 0.55 * inch, 1.05 * inch, W - 1.1 * inch, 1.9 * inch, SURFACE)
    c.setFillColor(SUCCESS); c.setFont("Helvetica-Bold", 10); c.drawString(0.75 * inch, 2.65 * inch, "RECOMMENDATION")
    wrap(c, "Launch reusable Autonomous Operations managed services (Ericsson + third-party under MS) selling autonomy maturity + observability + outcome KPIs as one journey — with enablement and pilots before scale. Reverse shadow M6–M9 · Auto classes M12–M15.",
         0.75 * inch, 2.3 * inch, W - 1.5 * inch, size=10)
    c.setFillColor(MUTED); c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, 0.95 * inch, "QUESTIONS · 5 MINUTES · Deep dive on any stage: HTML case study")
    flow(c, 6); foot(c, 8); c.showPage()

    c.save()
    print(f"Wrote {path}")


if __name__ == "__main__":
    build()
