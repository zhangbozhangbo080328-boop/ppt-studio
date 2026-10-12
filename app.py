#!/usr/bin/env python3
"""PPT Studio - a local PPT generator web app."""

from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
import requests
from PIL import Image
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION


def _load_dotenv(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, val = line.split("=", 1)
                os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))
    except FileNotFoundError:
        pass


_load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

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

DEFAULT_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
DEFAULT_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

SYSTEM_PROMPT = (
    "你是专业的演示文稿内容策划。根据用户主题与要求，生成一份结构清晰的 PPT 大纲。"
    "只输出 JSON，不要输出任何解释、注释或 Markdown 代码块标记。"
    "JSON 格式如下："
    '{"title":"演示标题","subtitle":"副标题","slides":[{"title":"页面标题","bullets":["要点1"]}],'
    '"charts":[{"type":"column|bar|pie|doughnut","title":"图表标题","data":[["标签",数值]]}],'
    '"image_queries":["英文图片搜索关键词"]}'
    "要求：共 5-8 页；每页 2-6 个要点；每个要点不超过 24 个汉字；"
    "标题简洁有力，逻辑从背景、要点到结论层层递进。"
    "仅当主题包含明确的数据对比时，才输出 charts（最多 1 个，数据 3-6 组、数值合理）；"
    "仅当配图能明显增强表达时，才输出 image_queries（1-3 个简短英文关键词）；"
    "没有对应需求时，省略 charts 与 image_queries 字段。"
)


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


CHART_TYPES = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}
CHART_PALETTE = ["1F3A5F", "3D5A80", "5B6B8C", "B08D3E", "4F6D5A", "7A2E39", "5A7D6C", "A6762D"]


def add_chart_slide(prs, title, chart_type, data, accent, dark):
    s = _blank(prs)
    bg = NIGHT if dark else PAPER
    fg = PAPER if dark else INK
    sub = NIGHT_MUTED if dark else MUTED
    _rect(s, 0, 0, SLIDE_W, SLIDE_H, bg)
    _rect(s, 0, 0, 0.18, SLIDE_H, accent)
    _text(s, 0.9, 0.82, 11.5, 0.8, title or "数据图表", 28, fg, True)

    labels = [str(d[0]) for d in data]
    values = [float(d[1]) for d in data]
    ct = CHART_TYPES.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    if ct in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        total = sum(values) or 1.0
        values = [v / total for v in values]

    chart_data = CategoryChartData()
    chart_data.categories = labels
    chart_data.add_series("数值", values)
    gf = s.shapes.add_chart(ct, Inches(0.9), Inches(1.8), Inches(11.5), Inches(5.25), chart_data)
    chart = gf.chart
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    series = plot.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = RGBColor.from_string(accent)
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = "0%" if ct in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT) else "0.##"
    dl.number_format_is_linked = False
    dl.font.size = Pt(11)
    dl.font.color.rgb = RGBColor.from_string(fg)
    if ct in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        for i in range(len(labels)):
            pt = series.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = RGBColor.from_string(CHART_PALETTE[i % len(CHART_PALETTE)])
    try:
        ca = chart.category_axis
        ca.tick_labels.font.size = Pt(12)
        ca.tick_labels.font.color.rgb = RGBColor.from_string(sub)
        va = chart.value_axis
        va.tick_labels.font.size = Pt(10)
        va.tick_labels.font.color.rgb = RGBColor.from_string(sub)
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = RGBColor.from_string("334155" if dark else "E7E1D3")
    except Exception:
        pass
    return s


def add_image_slide(prs, caption, image_bytes, accent, dark):
    s = _blank(prs)
    bg = NIGHT if dark else PAPER
    fg = PAPER if dark else INK
    _rect(s, 0, 0, SLIDE_W, SLIDE_H, bg)
    _rect(s, 0, 0, 0.18, SLIDE_H, accent)
    if caption:
        _text(s, 0.9, 0.6, 11.5, 0.7, caption, 22, fg, True)
    try:
        img = Image.open(io.BytesIO(image_bytes))
        iw, ih = img.size
        max_w, max_h, top = 11.5, 5.5, 1.45
        ratio = iw / ih
        w = min(max_w, max_h * ratio)
        h = w / ratio
        x = (SLIDE_W - w) / 2
        y = top + (max_h - h) / 2
        s.shapes.add_picture(io.BytesIO(image_bytes), Inches(x), Inches(y), Inches(w), Inches(h))
    except Exception:
        _text(s, 0.9, 3.2, 11.5, 0.6, "图片无法加载", 20, "DC2626", True)
    return s


def fetch_image(url):
    if not re.match(r"^https?://", url):
        raise ValueError("图片地址需以 http:// 或 https:// 开头")
    resp = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    if not (resp.headers.get("Content-Type") or "").lower().startswith("image/"):
        raise ValueError("该地址不是图片")
    if len(resp.content) > 8 * 1024 * 1024:
        raise ValueError("图片超过 8MB")
    return resp.content


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


def parse_outline(content):
    text = (content or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    data = json.loads(text)
    title = str(data.get("title") or "").strip() or "AI 生成演示"
    subtitle = str(data.get("subtitle") or "").strip()
    slides = []
    for item in (data.get("slides") or [])[:10]:
        t = str(item.get("title") or "").strip() or "未命名页面"
        bullets = [str(b).strip() for b in (item.get("bullets") or []) if str(b).strip()]
        slides.append({"title": t, "bullets": bullets[:MAX_BULLETS]})
    if not slides:
        slides = [{"title": "内容概览", "bullets": ["请补充更多主题细节后重试"]}]
    charts = []
    for c in (data.get("charts") or [])[:1]:
        ctype = str(c.get("type") or "column").lower()
        ctitle = str(c.get("title") or "").strip()
        rows = []
        for r in (c.get("data") or []):
            if isinstance(r, list) and len(r) >= 2:
                try:
                    rows.append([str(r[0]).strip(), float(r[1])])
                except (ValueError, TypeError):
                    continue
        if rows and ctype in CHART_TYPES:
            charts.append({"type": ctype, "title": ctitle, "data": rows})
    queries = [str(q).strip() for q in (data.get("image_queries") or []) if str(q).strip()][:3]
    return title, subtitle, slides, charts, queries


def call_llm(topic, requirements, api_key, base_url, model):
    base = (base_url or "").strip().rstrip("/") or DEFAULT_BASE_URL
    model = (model or "").strip() or DEFAULT_MODEL
    user_content = "主题：" + topic
    if requirements:
        user_content += "\n额外要求：" + requirements
    resp = requests.post(
        base + "/chat/completions",
        headers={"Authorization": "Bearer " + api_key, "Content-Type": "application/json"},
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
            "temperature": 0.7,
        },
        timeout=90,
    )
    if resp.status_code != 200:
        detail = resp.text[:240].replace("\n", " ")
        raise RuntimeError(f"AI 接口返回 {resp.status_code}：{detail}")
    try:
        content = resp.json()["choices"][0]["message"]["content"]
    except Exception:
        raise RuntimeError("AI 返回内容无法解析，请检查模型与接口地址")
    return parse_outline(content)


def fetch_wikimedia_image(query):
    try:
        params = {
            "action": "query",
            "format": "json",
            "generator": "search",
            "gsrsearch": "filetype:bitmap " + query,
            "gsrnamespace": "6",
            "gsrlimit": "1",
            "prop": "imageinfo",
            "iiprop": "url",
            "iiurlwidth": "1600",
        }
        resp = requests.get(
            "https://commons.wikimedia.org/w/api.php",
            params=params,
            timeout=20,
            headers={"User-Agent": "PPT-Studio/1.0"},
        )
        pages = resp.json().get("query", {}).get("pages", {})
        for page in pages.values():
            ii = (page.get("imageinfo") or [{}])[0]
            url = ii.get("thumburl") or ii.get("url")
            if url:
                return url
    except Exception:
        pass
    return None


def build_pptx(title, subtitle, blocks, accent_key, dark, author, charts=None, images=None):
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
    for spec in (charts or []):
        try:
            data = [(str(r[0]), r[1]) for r in (spec.get("data") or []) if len(r) >= 2]
            if data:
                add_chart_slide(prs, spec.get("title") or "", spec.get("type") or "column", data, accent, dark)
        except Exception:
            continue
    for caption, img_bytes in (images or []):
        add_image_slide(prs, caption, img_bytes, accent, dark)
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
    charts = data.get("charts") or []
    image_slides = []
    for url in (data.get("images") or []):
        url = (url or "").strip()
        if not url:
            continue
        try:
            image_slides.append((None, fetch_image(url)))
        except Exception as exc:
            return jsonify({"error": f"图片下载失败：{url}（{exc}）"}), 400
    buf = build_pptx(title, subtitle, blocks, accent, dark, author, charts, image_slides)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"presentation_{stamp}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/generate-ai", methods=["POST"])
def generate_ai():
    data = request.get_json(force=True, silent=True) or {}
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify({"error": "请填写主题"}), 400
    api_key = (data.get("api_key") or "").strip() or os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        return jsonify({"error": "AI 服务未配置（请站点管理员配置 API Key）"}), 400
    try:
        title, subtitle, slides, ai_charts, ai_queries = call_llm(
            topic,
            (data.get("requirements") or "").strip(),
            api_key,
            (data.get("base_url") or "").strip(),
            (data.get("model") or "").strip(),
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 502
    charts = list(ai_charts) + (data.get("charts") or [])
    image_slides = []
    for url in (data.get("images") or []):
        url = (url or "").strip()
        if not url:
            continue
        try:
            image_slides.append((None, fetch_image(url)))
        except Exception as exc:
            return jsonify({"error": f"图片下载失败：{url}（{exc}）"}), 400
    for q in ai_queries:
        image_url = fetch_wikimedia_image(q)
        if not image_url:
            continue
        try:
            image_slides.append((None, fetch_image(image_url)))
        except Exception:
            continue
    buf = build_pptx(
        title,
        subtitle,
        slides,
        data.get("accent") or "blue",
        bool(data.get("dark")),
        (data.get("author") or "").strip(),
        charts,
        image_slides,
    )
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"ai_presentation_{stamp}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", "5000"))
    host = os.environ.get("HOST", "0.0.0.0")
    app.run(host=host, port=port, debug=False, threaded=True)
