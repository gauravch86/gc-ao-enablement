#!/usr/bin/env python3
"""Case Study 1 executive PDF — mirrors the 6-slide PPTX storyline."""

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BG = HexColor("#0F1419")
SURFACE = HexColor("#1A2332")
TEXT = HexColor("#E6EDF3")
MUTED = HexColor("#8B9CB3")
ACCENT = HexColor("#58A6FF")
GOLD = HexColor("#D4A853")
SUCCESS = HexColor("#3FB950")
DANGER = HexColor("#F87171")
PURPLE = HexColor("#A371F7")

PAGE = landscape(A4)  # 841.89 x 595.28
W, H = PAGE


def rounded_rect(c, x, y, w, h, fill, radius=8):
    c.setFillColor(fill)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=0)


def footer(c, page, total=6):
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8)
    c.drawString(0.55 * inch, 0.35 * inch, "Case Study 1 · Reposition underperforming offering · Gaurav Chaudhary")
    c.drawRightString(W - 0.55 * inch, 0.35 * inch, f"{page} / {total}")


def label(c, text, y=H - 0.45 * inch):
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, y, text.upper())


def title(c, text, y=H - 0.85 * inch):
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(0.55 * inch, y, text)


def wrap_text(c, text, x, y, max_width, font="Helvetica", size=11, color=TEXT, leading=14):
    c.setFont(font, size)
    c.setFillColor(color)
    words = text.split()
    lines = []
    cur = ""
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

    # ─── SLIDE 1 ───
    slide_bg(c)
    label(c, "Strategic Product Manager  ·  Case Study 1  ·  10-min brief")
    title(c, "Reposition an underperforming portfolio offering")
    wrap_text(
        c,
        "Application Managed Services (AMS): strong technical delivery, but commercially stuck — "
        "flat net sales, declining margins, Sales cannot articulate value, regional pricing chaos.",
        0.55 * inch, H - 1.25 * inch, W - 1.1 * inch, size=11, color=MUTED, leading=15,
    )

    cards = [
        ("NET SALES (2 YRS)", "Flat", "Weak conversion · no attach", MUTED),
        ("GROSS MARGIN", "−4–6 pts", "Discount · scope creep · KT bleed", DANGER),
        ("DELIVERY NPS / C-SAT", "Strong", "Execution is not the problem", SUCCESS),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (lbl, val, sub, col) in enumerate(cards):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        y = H - 3.35 * inch
        rounded_rect(c, x, y, cw, 1.45 * inch, SURFACE)
        c.setFillColor(MUTED)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x + cw / 2, y + 1.1 * inch, lbl)
        c.setFillColor(col)
        c.setFont("Helvetica-Bold", 22)
        c.drawCentredString(x + cw / 2, y + 0.65 * inch, val)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 9)
        c.drawCentredString(x + cw / 2, y + 0.3 * inch, sub)

    rounded_rect(c, 0.55 * inch, H - 4.85 * inch, W - 1.1 * inch, 1.2 * inch, SURFACE)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, H - 3.9 * inch, "WORKING HYPOTHESIS")
    wrap_text(
        c,
        "Commercially mis-positioned — sold as FTE headcount while the market buys outcomes, "
        "autonomy, and TCO reduction. Delivery excellence exists; the product story, packaging, "
        "and pricing model do not.",
        0.75 * inch, H - 4.2 * inch, W - 1.5 * inch, size=11, color=TEXT, leading=14,
    )

    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.55 * inch, 1.55 * inch, "AGENDA  ·  10 MINUTES")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 11)
    c.drawString(
        0.55 * inch, 1.25 * inch,
        "1  Diagnose   →   2  Decide   →   3  Package & price   →   4  Prioritize & align   →   5  Measure & close",
    )
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9)
    c.drawString(
        0.55 * inch, 0.95 * inch,
        "Proof anchors:  H3G UK  $125m packaging & economics   ·   TEF Argentina  $100m renewal reposition",
    )
    footer(c, 1)
    c.showPage()

    # ─── SLIDE 2 ───
    slide_bg(c)
    label(c, "Diagnosis framework  ·  Recommendation")
    title(c, "Strong delivery. Broken commercial model. → Reposition.")

    lenses = [
        ("WIN / LOSS", "15–20 deals · tag price vs value vs unclear"),
        ("MARGIN", "GM waterfall · discount · KT overrun · CR leak"),
        ("COMPETITIVE", "FTE SI vs outcome / AO product MS shift"),
        ("VOICE", "NPS high + pay willingness low = story gap"),
        ("PORTFOLIO", "18-mo GM path? Attach to AO / products?"),
    ]
    lw = (W - 1.3 * inch) / 5
    for i, (h, body) in enumerate(lenses):
        x = 0.55 * inch + i * (lw + 0.08 * inch)
        rounded_rect(c, x, H - 2.95 * inch, lw, 1.35 * inch, SURFACE)
        c.setFillColor(ACCENT)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 0.12 * inch, H - 1.85 * inch, h)
        wrap_text(c, body, x + 0.12 * inch, H - 2.15 * inch, lw - 0.24 * inch, size=9, color=TEXT, leading=12)

    causes = [
        ("Commoditized packaging", "Sold as L1/L2/L3 FTE towers — procurement benches us vs SIs"),
        ("Pricing chaos", "Regional ad-hoc discounts · underpriced KT · unpaid scope creep"),
        ("Value story gap", "Buyers want TCO & autonomy path; we pitch headcount replacement"),
    ]
    cw = (W - 1.3 * inch) / 3
    for i, (h, body) in enumerate(causes):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rounded_rect(c, x, H - 4.5 * inch, cw, 1.25 * inch, SURFACE)
        c.setFillColor(DANGER)
        c.rect(x, H - 3.3 * inch - 2, cw, 3, fill=1, stroke=0)
        c.setFillColor(DANGER)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x + 0.15 * inch, H - 3.55 * inch, h)
        wrap_text(c, body, x + 0.15 * inch, H - 3.85 * inch, cw - 0.3 * inch, size=9, color=TEXT, leading=12)

    decisions = [
        ("SUNSET", "Rejects delivery DNA & footprint.\nOnly if no 18-mo GM path.", MUTED, False),
        ("EVOLVE ONLY", "Tooling without SKU change\nfixes neither narrative nor GM.", MUTED, False),
        ("REPOSITION  ✓", "Outcome-tiered Application Ops\nStabilize → Optimize → Autonomize", SUCCESS, True),
    ]
    for i, (h, body, col, sel) in enumerate(decisions):
        x = 0.55 * inch + i * (cw + 0.1 * inch)
        rounded_rect(c, x, 0.55 * inch, cw, 1.55 * inch, SURFACE)
        if sel:
            c.setStrokeColor(SUCCESS)
            c.setLineWidth(2)
            c.roundRect(x, 0.55 * inch, cw, 1.55 * inch, 8, fill=0, stroke=1)
        c.setFillColor(col)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x + 0.15 * inch, 1.75 * inch, h)
        c.setFillColor(TEXT if sel else MUTED)
        c.setFont("Helvetica", 9)
        for j, line in enumerate(body.split("\n")):
            c.drawString(x + 0.15 * inch, 1.45 * inch - j * 13, line)

    footer(c, 2)
    c.showPage()

    # ─── SLIDE 3 ───
    slide_bg(c)
    label(c, "Revised pricing & packaging  ·  Value proposition")
    title(c, "Sell tiers & outcomes — not FTE towers")
    wrap_text(
        c,
        "Value prop: Outcome-guaranteed Application Operations with a built-in ramp to autonomy — "
        "reducing customer TCO while protecting margin through standardized tiers & global guardrails.",
        0.55 * inch, H - 1.2 * inch, W - 1.1 * inch, size=10, color=MUTED, leading=13,
    )

    tiers = [
        ("STABILIZE", GOLD, "Restore control",
         ["Backlog burn · SLA baseline", "Single accountability · CPI board", "Top-offender program",
          "Price: platform + app-class", "Floor GM ≥ 22%"]),
        ("OPTIMIZE", ACCENT, "TCO & MTTR proof",
         ["+ Unified observability / SPOG", "+ GenAI SOP · automation factory", "Outcome-band pricing",
          "Gain-share on TCO proof", "Floor GM ≥ 26%"]),
        ("AUTONOMIZE", SUCCESS, "Path to L3–L4 autonomy",
         ["+ Closed-loop · policy packs", "+ Soft-ramp CPI contract", "Milestone + performance",
          "On-ramp to reusable AO SKU", "Floor GM ≥ 30%"]),
    ]
    tw = (W - 1.3 * inch) / 3
    for i, (name, col, tag, lines) in enumerate(tiers):
        x = 0.55 * inch + i * (tw + 0.1 * inch)
        rounded_rect(c, x, 1.55 * inch, tw, 3.5 * inch, SURFACE)
        c.setFillColor(col)
        c.rect(x, 1.55 * inch + 3.5 * inch - 4, tw, 4, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x + 0.18 * inch, 4.7 * inch, name)
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x + 0.18 * inch, 4.4 * inch, tag)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 10)
        for j, line in enumerate(lines):
            c.drawString(x + 0.18 * inch, 4.0 * inch - j * 18, "•  " + line)

    rounded_rect(c, 0.55 * inch, 0.55 * inch, W - 1.1 * inch, 0.8 * inch, SURFACE)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(0.75 * inch, 1.1 * inch, "MODULAR ATTACH  ·  PRICED, NOT FREE")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 9)
    c.drawString(
        0.75 * inch, 0.8 * inch,
        "Observability jump-start · Top-offender sprint · Transition accelerator (mandatory KT) · "
        "Vendor-consolidation bridge · Customer buys service class & CPI — not FTEs",
    )
    footer(c, 3)
    c.showPage()

    # ─── SLIDE 4 ───
    slide_bg(c)
    label(c, "Commercial guardrails  ·  Investment sequencing")
    title(c, "Deal Desk rules that stop the bleed — fund ROI first")

    rounded_rect(c, 0.55 * inch, 0.55 * inch, 5.6 * inch, H - 1.7 * inch, SURFACE)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.75 * inch, H - 1.4 * inch, "PRICING GUARDRAILS")
    rules = [
        ("Floor GM%", "Stabilize ≥22% · Optimize ≥26% · Autonomize ≥30%"),
        ("Max discount", "≤12% off list (≤8% renewals) — trade for scope / term"),
        ("Transition", "KT module mandatory · min 8% of Y1 ACV"),
        ("Scope CR", "Unpaid work cap 2% ACV · 5-day CR cycle"),
        ("Regional variance", "Like-for-like within ±15% of global rate card"),
        ("CPI / penalty", "Soft-ramp M1–6 · credit caps · Finance rev-rec"),
    ]
    y = H - 1.75 * inch
    for name, detail in rules:
        c.setFillColor(ACCENT)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(0.75 * inch, y, name)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 9)
        c.drawString(0.75 * inch, y - 14, detail)
        y -= 42

    rounded_rect(c, 6.4 * inch, H - 4.2 * inch, 4.7 * inch, 2.7 * inch, SURFACE)
    c.setFillColor(SUCCESS)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(6.6 * inch, H - 1.7 * inch, "FUND  ·  HIGH ROI")
    fund = [
        "Q1: Rate card · tier SKUs · ROI · Deal Desk",
        "Q2: CPI catalogue · transition factory · 2 pilots",
        "Q3–4: Observability · SPOG · library · renewals",
        "Y2: Autonomize → reusable AO offering merge",
    ]
    y = H - 2.1 * inch
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for line in fund:
        c.drawString(6.6 * inch, y, "▸  " + line)
        y -= 22

    rounded_rect(c, 6.4 * inch, 0.55 * inch, 4.7 * inch, 1.85 * inch, SURFACE)
    c.setFillColor(DANGER)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(6.6 * inch, 2.05 * inch, "STOP / DEPRIORITIZE")
    stops = [
        "Per-account custom observability builds",
        "Below-floor PoCs to “win logo”",
        "Regional bespoke rate cards",
        "Rebrand without sales / Deal Desk kit",
    ]
    y = 1.7 * inch
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for line in stops:
        c.drawString(6.6 * inch, y, "×  " + line)
        y -= 18

    footer(c, 4)
    c.showPage()

    # ─── SLIDE 5 ───
    slide_bg(c)
    label(c, "Cross-functional alignment  ·  Success metrics & feedback loop")
    title(c, "One program — Sales, Marketing, Finance, Regions")

    funcs = [
        ("SALES / DEAL DESK", ACCENT, "≥80% bids on tier SKU by Q2",
         "Discovery kit · ROI · margin waterfall · weekly pilot review"),
        ("MARKETING", PURPLE, "One narrative: Apps Ops → Autonomy",
         "Tier one-pagers · proof stories · retire FTE language"),
        ("FINANCE", GOLD, "Rev-rec compliant · GM floors in ERP",
         "Platform vs milestone memo · penalty lot · bid sign-off"),
        ("REGIONS / DELIVERY", SUCCESS, "Pilots hit CPI catalogue",
         "KT playbook · SDM board · QBR step-ups · train-the-trainer"),
    ]
    fw = (W - 1.35 * inch) / 4
    for i, (name, col, outcome, arts) in enumerate(funcs):
        x = 0.55 * inch + i * (fw + 0.08 * inch)
        rounded_rect(c, x, H - 3.7 * inch, fw, 2.15 * inch, SURFACE)
        c.setFillColor(col)
        c.rect(x, H - 3.7 * inch, 4, 2.15 * inch, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 0.15 * inch, H - 1.8 * inch, name)
        c.setFillColor(TEXT)
        c.setFont("Helvetica-Bold", 9)
        wrap_text(c, outcome, x + 0.15 * inch, H - 2.15 * inch, fw - 0.3 * inch, size=9, color=TEXT, leading=12)
        wrap_text(c, arts, x + 0.15 * inch, H - 2.7 * inch, fw - 0.3 * inch, size=8, color=MUTED, leading=11)

    rounded_rect(c, 0.55 * inch, 0.55 * inch, W - 1.1 * inch, 2.15 * inch, SURFACE)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(0.75 * inch, 2.35 * inch, "12-MONTH SCOREBOARD  ·  Quarterly Pricing Council")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    c.drawString(0.75 * inch, 1.95 * inch, "Net sales  Flat → +8–12%          Gross margin  recover +3–4 pts          Tier SKU adoption  ≥80%")
    c.drawString(0.75 * inch, 1.65 * inch, "Regional price variance  <15%          Wins citing value/TCO  ≥40%          Renewal step-up  ≥25% Optimize+")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9)
    wrap_text(
        c,
        "Feedback loop: monthly Deal Desk overrides + CR volume → quarterly pricing council (±3% rate card) → "
        "annual sunset re-test & Autonomize → AO SKU merge",
        0.75 * inch, 1.25 * inch, W - 1.5 * inch, size=9, color=MUTED, leading=12,
    )
    footer(c, 5)
    c.showPage()

    # ─── SLIDE 6 ───
    slide_bg(c)
    label(c, "Recommendation  ·  Ask")
    title(c, "Reposition AMS. Land with tiers. Expand to autonomy.")

    rounded_rect(c, 0.55 * inch, H - 3.2 * inch, W - 1.1 * inch, 1.7 * inch, SURFACE)
    c.setFillColor(ACCENT)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, H - 1.7 * inch, "THE CALL")
    wrap_text(
        c,
        "Reposition Application Managed Services as Outcome-Tiered Application Operations "
        "(Stabilize → Optimize → Autonomize). Do not sunset a delivery-capable franchise. "
        "Do not evolve tooling without fixing the commercial SKU. Fund guardrails & enablement in Q1; "
        "pilot in two regions; measure with a quarterly pricing council.",
        0.75 * inch, H - 2.05 * inch, W - 1.5 * inch, size=11, color=TEXT, leading=14,
    )

    rounded_rect(c, 0.55 * inch, 1.35 * inch, 5.4 * inch, 2.25 * inch, SURFACE)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(0.75 * inch, 3.3 * inch, "WHY CREDIBLE")
    c.setFillColor(TEXT)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.75 * inch, 2.95 * inch, "H3G UK — $125m")
    wrap_text(c, "75 CPIs · priced transition · 474→336 economics — sold transformation, not FTEs",
              0.75 * inch, 2.7 * inch, 5.0 * inch, size=9, color=MUTED, leading=12)
    c.setFillColor(TEXT)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.75 * inch, 2.15 * inch, "TEF Argentina — $100m")
    wrap_text(c, "Repositioned accountability & KPI transparency when the story was broken → renewal",
              0.75 * inch, 1.9 * inch, 5.0 * inch, size=9, color=MUTED, leading=12)

    rounded_rect(c, 6.2 * inch, 1.35 * inch, 5.0 * inch, 2.25 * inch, SURFACE)
    c.setFillColor(SUCCESS)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(6.4 * inch, 3.3 * inch, "ASK OF LEADERSHIP")
    asks = [
        "Approve reposition mandate & global rate card",
        "Stand up Deal Desk guardrails this quarter",
        "Authorize 2 regional pilots (UK + LATAM)",
        "Name SPM owner + monthly steering forum",
    ]
    y = 2.9 * inch
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    for a in asks:
        c.drawString(6.4 * inch, y, "▸  " + a)
        y -= 22

    c.setFillColor(MUTED)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(W / 2, 0.85 * inch, "Q&A  ·  5 minutes     ·     Detail: HTML case study for deep-dive")
    footer(c, 6)
    c.showPage()

    c.save()
    print(f"Wrote {path}")


if __name__ == "__main__":
    build()
