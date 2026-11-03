#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Refined CSP asks slide - same format as Ericsson marketing: 4 pillars + outcome bar + answer."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BG = RGBColor(0x12, 0x12, 0x14)
CARD = RGBColor(0xF5, 0xF5, 0xF5)
TEXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED_DARK = RGBColor(0x4A, 0x4A, 0x4A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLUE = RGBColor(0x1E, 0x6B, 0xD6)
BLUE_BAR = RGBColor(0x2A, 0x6F, 0xDB)
GOLD = RGBColor(0xD4, 0xA8, 0x53)

W = Inches(13.333)
H = Inches(7.5)
OUT = "/workspace/csp-asks-refined.pptx"


def fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def set_run(run, size=12, bold=False, color=TEXT_DARK, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def add_text(tf, text, size=12, bold=False, color=TEXT_DARK, align=PP_ALIGN.LEFT, space_after=4):
    if tf.paragraphs[0].text == "" and all(not p.text for p in tf.paragraphs):
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    r = p.add_run()
    r.text = text
    set_run(r, size=size, bold=bold, color=color)
    return p


def blank(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    fill(bg, BG)
    sp = bg._element
    slide.shapes._spTree.remove(sp)
    slide.shapes._spTree.insert(2, sp)
    return slide


def pillar(slide, left, top, width, height, num, title, problem, need):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    fill(card, CARD)
    try:
        card.adjustments[0] = 0.06
    except Exception:
        pass

    tb = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.18), width - Inches(0.36), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    add_text(tf, f"{num}  {title}", size=13, bold=True, color=TEXT_DARK, space_after=8)
    add_text(tf, problem, size=11, bold=False, color=MUTED_DARK, space_after=10)
    add_text(tf, "What they need", size=10, bold=True, color=BLUE, space_after=3)
    add_text(tf, need, size=11, bold=False, color=TEXT_DARK, space_after=0)


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    s = blank(prs)

    # Title
    tb = s.shapes.add_textbox(Inches(0.55), Inches(0.28), Inches(12.2), Inches(0.55))
    add_text(tb.text_frame, "What CSPs are asking us to solve  and how we answer as a portfolio", size=22, bold=True, color=WHITE)

    # Subtitle
    tb = s.shapes.add_textbox(Inches(0.55), Inches(0.78), Inches(12.2), Inches(0.32))
    add_text(
        tb.text_frame,
        "Problem ? sellable offer ? operated outcome  ·  Productized core + configured edge  ·  Growth with margin discipline",
        size=11,
        bold=False,
        color=RGBColor(0xA8, 0xB4, 0xC4),
    )

    pillars = [
        (
            "01",
            "Multi-vendor complexity",
            "Heterogeneous BSS/OSS estates make every integration and change slow, risky, and costly  with unclear accountability across vendors.",
            "A governed multi-vendor integration layer, safe change paths, and single accountability for outcomes  not another point-to-point project.",
        ),
        (
            "02",
            "New digital / B2B2X offers",
            "New IoT and partner services need capabilities the current estate was never designed to deliver  without multi-year rebuilds.",
            "Reusable capability packs and thin adapters to launch offers faster  custom only where it creates durable differentiation.",
        ),
        (
            "03",
            "Faster differentiation",
            "Off-the-shelf products rarely create advantage; CSPs need speed without locking into one-off builds that cannot scale.",
            "Configurable offer ladder (land ? expand ? deepen) with clear time-to-market SLAs and reuse across markets.",
        ),
        (
            "04",
            "AI inside operations",
            "Generic AI pilots do not fit CSP workflows; value appears only when intelligence is embedded in run, change, and commercial processes.",
            "Domain AI & automation delivered as a managed service  with human override, auditability, and model lifecycle governance.",
        ),
    ]

    gap = Inches(0.22)
    left0 = Inches(0.45)
    width = Inches(3.0)
    top = Inches(1.25)
    height = Inches(3.55)
    for i, (num, title, problem, need) in enumerate(pillars):
        pillar(s, left0 + i * (width + gap), top, width, height, num, title, problem, need)

    # Outcome bar
    bar = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.45), Inches(5.0), Inches(12.4), Inches(1.05))
    fill(bar, BLUE_BAR)
    try:
        bar.adjustments[0] = 0.08
    except Exception:
        pass

    tb = s.shapes.add_textbox(Inches(0.65), Inches(5.1), Inches(12.0), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    add_text(tf, "Customer outcome / value", size=11, bold=True, color=WHITE, space_after=4)
    add_text(
        tf,
        "Faster time-to-market for new offers  ·  Lower integration & rework  ·  3040% engineering effort reduction where scoped  ·  "
        "Measurable run outcomes (availability, order/billing cycle, incident load)  ·  Reuse across logos with maintainable cost-to-serve",
        size=11,
        bold=False,
        color=WHITE,
        space_after=0,
    )

    # Answer
    tb = s.shapes.add_textbox(Inches(0.55), Inches(6.25), Inches(12.2), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    add_text(tf, "Our answer", size=11, bold=True, color=GOLD, space_after=3)
    add_text(
        tf,
        "Solve inside the CSPs BSS/OSS with a productized services portfolio: reusable core + configured edge + embedded AI, "
        "sold as clear offers (subscription / managed service), operated under shared KPIs  with Ericsson domain depth and day-2 accountability. "
        "Custom build only where it becomes reusable IP; not as the default.",
        size=12,
        bold=False,
        color=WHITE,
        space_after=0,
    )

    prs.save(OUT)
    print(f"Wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
