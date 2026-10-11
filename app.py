#!/usr/bin/env python3
"""PPT Studio - a local PPT generator web app."""

from __future__ import annotations

import io
import re
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024

ACCENTS = {
    "blue": "2563EB",
    "green": "059669",
    "purple": "7C3AED",
    "red": "DC2626",
    "orange": "EA580C",
    "teal": "0D9488",
}

INK = "0F172A"
MUTED = "475569"
PAPER = "FFFFFF"
NIGHT = "0F172A"
NIGHT_MUTED = "CBD5E1"

SLIDE_W = 13.333
SLIDE_H = 7.5
MAX_BULLETS = 7


def _set_run(run, size, color, bold=False, family="PingFang SC"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Arial"
    run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        if latin is not None:
            latin.addnext(ea)
        else:
            rPr.append(ea)
    ea.set("typeface", family)


def _rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    sp.line.fill.background()
    sp.shadow.inherit = False
    if radius is not None:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    return sp


def _text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP, family="PingFang SC", line_spacing=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    _set_run(run, size, color, bold, family)
    return tb


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _cover(prs, title, subtitle, author, accent, dark):
    s = _blank(prs)
    bg = NIGHT if dark else PAPER
    fg = PAPER if dark else INK
    sub = NIGHT_MUTED if dark else MUTED
    _rect(s, 0, 0, SLIDE_W, SLIDE_H, bg)
    _rect(s, 0, 0, 0.18, SLIDE_H, accent)
    _rect(s, 0.9, 0.95, 0.22, 0.22, accent, MSO_SHAPE.OVAL)
    _text(s, 0.9, 1.7, 11.5, 0.5, "PPT STUDIO · " + datetime.now().strftime("%Y-%m-%d"), 17, accent, True)
    _text(s, 0.9, 2.35, 11.5, 1.9, title, 46, fg, True)
    if subtitle:
        _text(s, 0.9, 3.75, 11.5, 0.8, subtitle, 22, sub)
    _rect(s, 0.9, 4.55, 1.15, 0.09, accent)
    if author:
        _text(s, 0.9, 5.15, 11.5, 0.5, author, 17, sub)


def _content(prs, title, bullets, accent, dark, page_no, total):
    s = _blank(prs)
    bg = NIGHT if dark else PAPER
    fg = PAPER if dark else INK
    sub = NIGHT_MUTED if dark else MUTED
    _rect(s, 0, 0, SLIDE_W, SLIDE_H, bg)
    _rect(s, 0, 0, 0.18, SLIDE_H, accent)
    _rect(s, 0.9, 0.78, 0.42, 0.42, accent, MSO_SHAPE.ROUNDED_RECTANGLE, 0.25)
    _text(s, 0.9, 0.9, 0.42, 0.42, str(page_no), 18, PAPER, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    _text(s, 1.55, 0.82, 10.8, 0.8, title, 30, fg, True)
    _text(s, 1.55, 1.55, 10.8, 0.4, f"{page_no} / {total}", 13, sub)
    y = 2.25
    for b in bullets:
        _rect(s, 0.95, y + 0.09, 0.13, 0.13, accent, MSO_SHAPE.OVAL)
        _text(s, 1.35, y, 10.9, 0.85, b, 19, fg)
        y += 0.72


def _ending(prs, subtitle, accent):
    s = _blank(prs)
    _rect(s, 0, 0, SLIDE_W, SLIDE_H, NIGHT)
    _rect(s, 0, 0, SLIDE_W, 0.16, accent)
    _text(s, 0.9, 2.7, 11.5, 1.2, "谢谢观看", 48, PAPER, True)
    if subtitle:
        _text(s, 0.9, 3.85, 11.5, 0.7, subtitle, 22, NIGHT_MUTED)
    _rect(s, 0.9, 4.45, 1.0, 0.08, accent)
    _text(s, 0.9, 5.15, 11.5, 0.5, "由 PPT Studio 生成 · 原生可编辑 PPTX", 15, "94A3B8")


def parse_blocks(text):
    parts = [b.strip() for b in re.split(r"\n\s*\n", text.strip()) if b.strip()]
    blocks = []
    for b in parts:
        lines = [ln.rstrip() for ln in b.splitlines() if ln.strip()]
        if not lines:
            continue
        title = re.sub(r"^#+\s*", "", lines[0]).strip() or "未命名页面"
        bullets = [re.sub(r"^\s*[-*•]\s*", "", ln).strip() for ln in lines[1:]]
        blocks.append({"title": title, "bullets": bullets})
    return blocks


def build_pptx(title, subtitle, blocks, accent_key, dark, author):
    accent = ACCENTS.get(accent_key, ACCENTS["blue"])
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    _cover(prs, title, subtitle, author, accent, dark)
    if not blocks:
        blocks = [{"title": "从这里开始", "bullets": ["在左侧内容框里输入你的要点", "空行分隔每一页", "每页第一行是标题，其余是要点"]}]
    slides = []
    for blk in blocks:
        bullets = blk["bullets"] or []
        chunks = [bullets[i:i + MAX_BULLETS] for i in range(0, max(1, len(bullets)), MAX_BULLETS)]
        slides.append((blk["title"], chunks[0]))
        for extra in chunks[1:]:
            slides.append((blk["title"] + "（续）", extra))
    total = len(slides)
    for i, (t, b) in enumerate(slides, start=1):
        _content(prs, t, b, accent, dark, i, total)
    _ending(prs, subtitle, accent)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    return render_template("index.html", accents=ACCENTS)


@app.route("/health")
def health():
    return jsonify({"ok": True})


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(force=True, silent=True) or {}
    title = (data.get("title") or "").strip() or "未命名演示"
    subtitle = (data.get("subtitle") or "").strip()
    author = (data.get("author") or "").strip()
    content = (data.get("content") or "").strip()
    accent = data.get("accent") or "blue"
    dark = bool(data.get("dark"))
    blocks = parse_blocks(content)
    buf = build_pptx(title, subtitle, blocks, accent, dark, author)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"presentation_{stamp}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", "5000"))
    host = os.environ.get("HOST", "0.0.0.0")
    app.run(host=host, port=port, debug=False, threaded=True)
