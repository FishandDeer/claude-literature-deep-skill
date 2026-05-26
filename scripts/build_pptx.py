#!/usr/bin/env python3
"""
build_pptx.py — 学术风 python-pptx 模板封装(V6 风格)

视觉规范完全遵循 references/ppt-design.md:
  - 16:9, 深蓝主色 (#1F3A68) + 橙色强调 (#C55A11)
  - 5 种 Tint 浅底色用于卡片(TINT_BLUE/ORG/GRN/RED/GOLD)
  - 中文 Microsoft YaHei (Windows) / Noto Sans CJK SC (Linux fallback)
  - 行动式标题、单论点页、低密度
  - 富 Header(eyebrow + 页码 + 双条 + 分隔线)
  - PART 分隔扉页(满版深蓝 + 大号罗马编号 + 章节名 + tagline)

用法(作为库导入):

    from build_pptx import AcademicDeck

    deck = AcademicDeck(total_pages=38, footer_text="组会汇报 · 2026-05")
    deck.add_cover(title="基于 VLM 的驾驶员状态多维感知与接管能力评估",
                   subtitle="—— 模拟器实证权重 × VLM 感知 × 系统验证",
                   eyebrow="DSM × VLM  ·  Driver Take-Over Readiness",
                   author="某某", date="2026-05", org="某某实验室")
    deck.add_toc([
        ("Ⅰ", "研究背景", "时代窗口 · 法规 · 三重张力"),
        ("Ⅱ", "研究现状", "DMS 四时代 · VLM 三范式"),
        ("Ⅲ", "研究方案", "3 RQ · 6 维度 · 系统集成"),
    ])
    deck.add_part_divider("Ⅰ", "研究背景", "从接管底线切入")
    deck.add_content_slide(
        eyebrow="Ⅰ · 研究背景",
        title="L3 准入落地,为什么是 2025 这个时代窗口?",
        subtitle="工信部 L3 试点 / Mercedes Drive Pilot / NHTSA SGO",
        bullets=[
            ("工信部 L3 准入 2025-12-15 启动 [#19]", None),
            ("Mercedes Drive Pilot 全球首款量产 L3", "Nevada/California 已开放"),
        ],
    )
    deck.add_card_grid(
        eyebrow="Ⅱ · 研究现状",
        title="VLM 在 DMS 的三种范式",
        cards=[
            ("C-i 零样本", ["GPT-4V / Gemini 直接 prompt", "无需训练,但精度有限"], "blue"),
            ("C-ii LoRA 微调", ["少样本 + 适配器", "工程友好"], "orange"),
            ("C-iii 端到端", ["LLaVA 风格", "算力要求高"], "green"),
        ],
    )
    deck.save("/path/to/output.pptx")
"""

from __future__ import annotations

import platform
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt


# ============ 视觉常量 V6 ============

PRIMARY    = RGBColor(0x1F, 0x3A, 0x68)   # 深蓝主色
PRIMARY_2  = RGBColor(0x2E, 0x5A, 0x8C)   # 副色
ACCENT     = RGBColor(0xC5, 0x5A, 0x11)   # 橙色强调
GREEN      = RGBColor(0x2E, 0x7D, 0x32)   # 学术绿(用于第三栏对比)
GOLD       = RGBColor(0xC9, 0x9A, 0x2A)   # 学术金

DARK       = RGBColor(0x22, 0x22, 0x22)
MUTED      = RGBColor(0x6B, 0x6B, 0x6B)
GRAY_L     = RGBColor(0xE6, 0xE9, 0xEE)
GRAY_M     = RGBColor(0xC9, 0xCF, 0xD7)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# Tint 浅底色 — 用于卡片背景与对比页底色
TINT_BLUE  = RGBColor(0xEB, 0xF1, 0xF8)
TINT_ORG   = RGBColor(0xFB, 0xEC, 0xDD)
TINT_GRN   = RGBColor(0xE7, 0xF1, 0xE8)
TINT_RED   = RGBColor(0xFB, 0xE6, 0xE6)
TINT_GOLD  = RGBColor(0xFF, 0xF6, 0xE0)

# 兼容历史调色板别名(老脚本仍可 import)
SECONDARY = PRIMARY_2
TEXT_DARK = DARK
TEXT_MUTE = MUTED
PALE      = TINT_BLUE

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# 中文字体:Windows 端 Microsoft YaHei,Linux 端如未安装则需先 `apt install fonts-noto-cjk`
CN_FONT = "Microsoft YaHei" if platform.system() == "Windows" else "Noto Sans CJK SC"
EN_FONT = "Calibri"

# Tint 主题查表(用于 add_card_grid)
THEME = {
    "blue":   (TINT_BLUE, PRIMARY,    PRIMARY),
    "orange": (TINT_ORG,  ACCENT,     ACCENT),
    "green":  (TINT_GRN,  GREEN,      GREEN),
    "red":    (TINT_RED,  RGBColor(0xB0, 0x2A, 0x2A), RGBColor(0xB0, 0x2A, 0x2A)),
    "gold":   (TINT_GOLD, GOLD,       GOLD),
}


# ============ 工具函数 ============

def _set_eastasia(run, font: str = CN_FONT):
    """注入 a:eastAsia 字体节点,避免中文回落到默认西文字体导致方框。"""
    rPr = run._r.get_or_add_rPr()
    e = rPr.find(qn("a:eastAsia"))
    if e is None:
        e = etree.SubElement(rPr, qn("a:eastAsia"))
    e.set("typeface", font)


def _set_run(run, text: str, *, size: int, color=DARK, bold=False,
             italic=False, font=CN_FONT):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    _set_eastasia(run)


def _add_textbox(slide, left, top, width, height, text, *,
                 size=18, color=DARK, bold=False, italic=False,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                 font=CN_FONT, fill=None, line=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
    tf.vertical_anchor = anchor
    if fill is not None:
        tb.fill.solid(); tb.fill.fore_color.rgb = fill
    else:
        tb.fill.background()
    if line is not None:
        tb.line.color.rgb = line; tb.line.width = Pt(0.75)
    else:
        tb.line.fill.background()
    p = tf.paragraphs[0]
    p.alignment = align
    _set_run(p.add_run(), text, size=size, color=color, bold=bold,
             italic=italic, font=font)
    if p.runs[0].text == "":
        p._p.remove(p.runs[0]._r)
    return tb


def _add_filled_rect(slide, left, top, width, height, fill_color,
                     line_color=None, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
    shp.shadow.inherit = False
    return shp


def _add_line(slide, x1, y1, x2, y2, color=GRAY_M, width=0.75):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    x1, y1, x2 - x1, y2 - y1)
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


def _add_paragraphs(slide, x, y, w, h, items, *, size=14, color=DARK,
                    fill=None, line=None, anchor=MSO_ANCHOR.TOP,
                    line_spacing=1.20):
    """items: list[dict] with keys: text, size?, bold?, color?, italic?,
    bullet?, dash?, indent?, align?
    """
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.10); tf.margin_right = Inches(0.10)
    tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
    if fill is not None:
        tb.fill.solid(); tb.fill.fore_color.rgb = fill
    else:
        tb.fill.background()
    if line is not None:
        tb.line.color.rgb = line; tb.line.width = Pt(0.75)
    else:
        tb.line.fill.background()
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = it.get("align", PP_ALIGN.LEFT)
        p.line_spacing = line_spacing
        if it.get("indent"):
            p.level = it["indent"]
        prefix = "・ " if it.get("bullet") else ("—— " if it.get("dash") else "")
        _set_run(p.add_run(),
                 prefix + it["text"],
                 size=it.get("size", size),
                 bold=it.get("bold", False),
                 color=it.get("color", color),
                 italic=it.get("italic", False))
    return tb


# ============ 主类 ============

@dataclass
class AcademicDeck:
    """学术 PPT 构建器(V6 风格)。每个 ``add_*`` 方法对应一种版式。

    Attributes
    ----------
    total_pages : int | None
        若指定,则页码显示为 ``NN / TOTAL``;否则只显示 ``NN``。
    footer_text : str
        每页底部的全局脚注(显示在内容页底部,不在 cover/divider/ending)。
    """

    total_pages: int | None = None
    footer_text: str = ""
    _page_no: int = field(default=0, init=False, repr=False)

    def __post_init__(self):
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self._blank_layout = self.prs.slide_layouts[6]

    # ---------- 内部:页码 / header / footer ----------

    def _next_page(self) -> int:
        self._page_no += 1
        return self._page_no

    def _page_label(self, n: int) -> str:
        if self.total_pages:
            return f"{n:02d} / {self.total_pages:02d}"
        return f"{n:02d}"

    def _add_header(self, slide, eyebrow: str, title: str,
                    subtitle: str | None = None,
                    page_no: int | None = None):
        """V6 富 header:顶部深蓝条 + 橙色细条 + eyebrow + title + 副标题 + 分隔线 + 页码。"""
        _add_filled_rect(slide, 0, 0, SLIDE_W, Inches(0.10), PRIMARY)
        _add_filled_rect(slide, 0, Inches(0.10), SLIDE_W, Inches(0.04), ACCENT)
        # 页码(右上角)
        if page_no is not None:
            _add_textbox(slide, SLIDE_W - Inches(1.8), Inches(0.22),
                         Inches(1.6), Inches(0.32),
                         self._page_label(page_no),
                         size=10, color=MUTED, align=PP_ALIGN.RIGHT,
                         font=EN_FONT)
        # eyebrow(橙色小字)
        if eyebrow:
            _add_textbox(slide, Inches(0.55), Inches(0.22),
                         Inches(8.0), Inches(0.32),
                         eyebrow, size=10, bold=True, color=ACCENT,
                         anchor=MSO_ANCHOR.MIDDLE)
        # 主标题
        _add_textbox(slide, Inches(0.55), Inches(0.55),
                     SLIDE_W - Inches(1.10), Inches(0.55),
                     title, size=22, bold=True, color=PRIMARY,
                     anchor=MSO_ANCHOR.MIDDLE)
        # 副标题(灰色斜体)
        if subtitle:
            _add_textbox(slide, Inches(0.55), Inches(1.10),
                         SLIDE_W - Inches(1.10), Inches(0.32),
                         subtitle, size=12, color=MUTED, italic=True,
                         anchor=MSO_ANCHOR.MIDDLE)
        # 灰色分隔线
        _add_line(slide, Inches(0.55), Inches(1.50),
                  SLIDE_W - Inches(0.55), Inches(1.50), GRAY_M, 0.5)

    def _add_footer(self, slide):
        if not self.footer_text:
            return
        _add_textbox(slide, Inches(0.55), SLIDE_H - Inches(0.35),
                     Inches(10.0), Inches(0.28),
                     self.footer_text,
                     size=9, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)

    # ---------- 1. 封面页 ----------

    def add_cover(self, title: str, subtitle: str = "",
                  eyebrow: str = "",
                  author: str = "", date: str = "", org: str = "",
                  meeting_tag: str = ""):
        """封面页(V6 风格):左侧深蓝竖条 + 橙色细竖条 + eyebrow + 大标题 + 橙色短分隔条 + 副标题 + 元信息。"""
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._page_no = 1  # 封面占第 1 页

        # 左侧深蓝色装饰条 + 橙色细条
        _add_filled_rect(slide, 0, 0, Inches(0.5), SLIDE_H, PRIMARY)
        _add_filled_rect(slide, Inches(0.5), 0, Inches(0.08), SLIDE_H, ACCENT)

        # eyebrow(英文小字、灰色斜体)
        if eyebrow:
            _add_textbox(slide, Inches(1.0), Inches(0.9), Inches(11), Inches(0.4),
                         eyebrow, size=14, color=MUTED, font=EN_FONT, italic=True)

        # 主标题(允许换行,使用一段两行的策略)
        _add_textbox(slide, Inches(1.0), Inches(2.0), Inches(11), Inches(2.2),
                     title, size=40, color=PRIMARY, bold=True,
                     anchor=MSO_ANCHOR.TOP)

        # 橙色短分隔条
        _add_filled_rect(slide, Inches(1.0), Inches(4.4),
                         Inches(0.8), Inches(0.06), ACCENT)

        # 副标题
        if subtitle:
            _add_textbox(slide, Inches(1.0), Inches(4.55),
                         Inches(11), Inches(0.55),
                         subtitle, size=18, color=DARK,
                         anchor=MSO_ANCHOR.MIDDLE)

        # meeting_tag(橙色加粗)
        if meeting_tag:
            _add_textbox(slide, Inches(1.0), Inches(5.20),
                         Inches(11), Inches(0.4),
                         meeting_tag, size=14, bold=True, color=ACCENT)

        # 元信息分隔线
        _add_line(slide, Inches(1.0), Inches(6.4),
                  Inches(7.0), Inches(6.4), GRAY_M, 0.5)

        # 作者 / 单位 / 日期
        if author:
            _add_textbox(slide, Inches(1.0), Inches(6.5),
                         Inches(11), Inches(0.4),
                         f"汇报人:{author}", size=14, bold=True, color=DARK)
        meta = "  ·  ".join(p for p in [org, date] if p)
        if meta:
            _add_textbox(slide, Inches(1.0), Inches(6.85),
                         Inches(11), Inches(0.4),
                         meta, size=11, color=MUTED, italic=True)
        return slide

    # ---------- 2. 目录页 ----------

    def add_toc(self, items: Iterable):
        """目录页。

        items 支持两种格式:
          - 简单字符串列表: ``["研究背景", "相关工作", ...]``
          - 富格式三元组列表: ``[("Ⅰ", "研究背景", "tagline"), ...]``
        """
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide,
                         eyebrow="CONTENTS · 目录",
                         title="本次汇报路线图",
                         subtitle=None,
                         page_no=page)

        items = list(items)
        # 检测富格式
        rich = items and isinstance(items[0], tuple) and len(items[0]) >= 2

        if rich:
            base_y = 1.95
            row_h = 1.10
            left_x = 1.0
            colors = [PRIMARY, ACCENT, GREEN, GOLD]
            for i, item in enumerate(items):
                if len(item) == 3:
                    no, name, desc = item
                else:
                    no, name = item; desc = ""
                color = colors[i % len(colors)]
                y = Inches(base_y + i * row_h)
                _add_filled_rect(slide, Inches(left_x), y,
                                 Inches(0.95), Inches(0.85), color)
                _add_textbox(slide, Inches(left_x), y,
                             Inches(0.95), Inches(0.85),
                             str(no), size=32, bold=True, color=WHITE,
                             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                             font=EN_FONT)
                _add_textbox(slide, Inches(left_x + 1.15), y,
                             Inches(10.5), Inches(0.45),
                             name, size=20, bold=True, color=color,
                             anchor=MSO_ANCHOR.MIDDLE)
                if desc:
                    _add_textbox(slide, Inches(left_x + 1.15),
                                 y + Inches(0.45),
                                 Inches(10.5), Inches(0.40),
                                 desc, size=12, color=MUTED, italic=True)
        else:
            # 兼容旧格式:圆点编号 + 章节名
            top0 = Inches(2.0)
            step = Inches(0.7)
            for i, name in enumerate(items, 1):
                shp = slide.shapes.add_shape(
                    MSO_SHAPE.OVAL, Inches(2.0), top0 + step * (i - 1),
                    Inches(0.55), Inches(0.55))
                shp.fill.solid()
                shp.fill.fore_color.rgb = PRIMARY
                shp.line.fill.background()
                tf = shp.text_frame
                tf.margin_top = tf.margin_bottom = 0
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                _set_run(p.add_run(), f"{i:02d}", size=14, color=WHITE,
                         bold=True, font=EN_FONT)
                if p.runs[0].text == "":
                    p._p.remove(p.runs[0]._r)
                _add_textbox(slide, Inches(2.9), top0 + step * (i - 1),
                             Inches(9.0), Inches(0.55),
                             str(name), size=22, color=DARK,
                             anchor=MSO_ANCHOR.MIDDLE)
        self._add_footer(slide)
        return slide

    # ---------- 3. PART 分隔扉页(V6 新增) ----------

    def add_part_divider(self, part_no: str, part_name: str,
                         tagline: str = "",
                         total_parts: int | None = None):
        """大型 PART 章节分隔页:满版深蓝 + 大号 PART 编号 + 章节名 + tagline。"""
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._next_page()
        _add_filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PRIMARY)

        # 大型 PART 编号(英文 + 罗马数字组合)
        _add_textbox(slide, Inches(0.8), Inches(1.6),
                     Inches(8.0), Inches(2.0),
                     f"PART {part_no}", size=110, bold=True, color=ACCENT,
                     font=EN_FONT)
        # 中文章节名
        _add_textbox(slide, Inches(0.8), Inches(3.6),
                     Inches(11), Inches(1.0),
                     part_name, size=48, bold=True, color=WHITE,
                     anchor=MSO_ANCHOR.MIDDLE)
        # 橙色短线
        _add_filled_rect(slide, Inches(0.85), Inches(4.65),
                         Inches(1.0), Inches(0.06), ACCENT)
        # tagline
        if tagline:
            _add_textbox(slide, Inches(0.8), Inches(4.85),
                         Inches(11), Inches(0.5),
                         tagline, size=18, color=GRAY_L, italic=True,
                         anchor=MSO_ANCHOR.MIDDLE)
        # 右下角进度
        if total_parts:
            _add_textbox(slide, SLIDE_W - Inches(2.0), SLIDE_H - Inches(0.55),
                         Inches(1.6), Inches(0.35),
                         f"{part_no} / {total_parts}",
                         size=11, color=GRAY_M, align=PP_ALIGN.RIGHT,
                         font=EN_FONT)
        return slide

    # ---------- 旧 API 兼容:section_header(老脚本可继续用) ----------

    def add_section_header(self, number: str, name: str, tagline: str = ""):
        """旧 API:与 add_part_divider 等价,保留以避免老脚本崩溃。"""
        return self.add_part_divider(number, name, tagline)

    # ---------- 4. 内容页(标题 + bullets) ----------

    def add_content_slide(self, title: str,
                          bullets: list[tuple[str, str | None]],
                          *, eyebrow: str = "", subtitle: str = ""):
        """通用内容页:富 header + bullets(每项可带二级注释)。

        bullets: ``[(主论点, 次级注释 or None), ...]``
        """
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide, eyebrow=eyebrow, title=title,
                         subtitle=subtitle, page_no=page)

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.7),
                                       Inches(11.7), Inches(5.3))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for primary, secondary in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.25
            _set_run(p.add_run(), "• ", size=20, color=ACCENT, bold=True)
            _set_run(p.add_run(), primary, size=18, color=DARK)
            p.space_after = Pt(6)
            if secondary:
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.LEFT
                p2.line_spacing = 1.18
                _set_run(p2.add_run(), "    " + secondary, size=13,
                         color=MUTED)
                p2.space_after = Pt(10)
        self._add_footer(slide)
        return slide

    # ---------- 5. 对比页(2~3 列) ----------

    def add_compare_slide(self, title: str,
                          columns: list[tuple[str, list[str]]],
                          *, eyebrow: str = "", subtitle: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide, eyebrow=eyebrow, title=title,
                         subtitle=subtitle, page_no=page)

        n = len(columns)
        assert 2 <= n <= 3, "compare_slide supports 2 or 3 columns"
        margin = Inches(0.5)
        gap = Inches(0.3)
        total_w = SLIDE_W - margin * 2 - gap * (n - 1)
        col_w = Emu(int(total_w) // n)
        top = Inches(1.7)
        col_h = Inches(5.3)

        # 三列各分一种主题
        themes = [(PRIMARY, TINT_BLUE), (ACCENT, TINT_ORG), (GREEN, TINT_GRN)]

        for i, (header, items) in enumerate(columns):
            head_color, body_color = themes[i % 3]
            left = margin + (col_w + gap) * i
            # 列头
            _add_filled_rect(slide, left, top, col_w, Inches(0.7), head_color)
            _add_textbox(slide, left, top, col_w, Inches(0.7),
                         header, size=18, color=WHITE, bold=True,
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            # 列体
            _add_filled_rect(slide, left, top + Inches(0.7), col_w,
                             col_h - Inches(0.7), body_color)
            tb = slide.shapes.add_textbox(left + Inches(0.2),
                                          top + Inches(0.85),
                                          col_w - Inches(0.4),
                                          col_h - Inches(1.0))
            tf = tb.text_frame
            tf.word_wrap = True
            first = True
            for item in items:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.alignment = PP_ALIGN.LEFT
                p.line_spacing = 1.22
                _set_run(p.add_run(), "• ", size=14, color=head_color,
                         bold=True)
                _set_run(p.add_run(), item, size=14, color=DARK)
                p.space_after = Pt(6)
        self._add_footer(slide)
        return slide

    # ---------- 6. 卡片网格页(V6 新增) ----------

    def add_card_grid(self, title: str,
                      cards: list[tuple[str, list[str], str]],
                      *, eyebrow: str = "", subtitle: str = ""):
        """卡片网格(2~6 个卡片自动排版)。

        cards: ``[(卡片标题, [bullet1, bullet2, ...], 主题), ...]``
        主题取值: ``"blue" | "orange" | "green" | "red" | "gold"``
        """
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide, eyebrow=eyebrow, title=title,
                         subtitle=subtitle, page_no=page)

        n = len(cards)
        assert 2 <= n <= 6, "card_grid supports 2~6 cards"
        # 1 行 2~3 张 / 2 行 4~6 张
        if n <= 3:
            cols, rows = n, 1
        else:
            cols, rows = (n + 1) // 2, 2

        margin = Inches(0.5)
        gap_x = Inches(0.25)
        gap_y = Inches(0.25)
        area_top = Inches(1.7)
        area_h = Inches(5.3)
        total_w = SLIDE_W - margin * 2 - gap_x * (cols - 1)
        card_w = Emu(int(total_w) // cols)
        card_h = Emu(int(area_h - gap_y * (rows - 1)) // rows)

        for i, (ttl, bullets, theme) in enumerate(cards):
            r, c = divmod(i, cols)
            x = margin + (card_w + gap_x) * c
            y = area_top + (card_h + gap_y) * r
            self._draw_card(slide, x, y, card_w, card_h,
                            ttl, bullets, theme)
        self._add_footer(slide)
        return slide

    def _draw_card(self, slide, x, y, w, h, title, bullets, theme):
        fill, bar, ttl_color = THEME.get(theme, THEME["blue"])
        _add_filled_rect(slide, x, y, w, h, fill)
        _add_filled_rect(slide, x, y, Inches(0.07), h, bar)
        _add_textbox(slide, x + Inches(0.20), y + Inches(0.10),
                     w - Inches(0.30), Inches(0.40),
                     title, size=14, bold=True, color=ttl_color)
        items = [{"text": t, "bullet": True} for t in bullets]
        _add_paragraphs(slide, x + Inches(0.18),
                        y + Inches(0.55),
                        w - Inches(0.30),
                        h - Inches(0.65),
                        items, size=11, color=DARK,
                        line_spacing=1.20)

    # ---------- 7. 表格页(V6 新增) ----------

    def add_table_slide(self, title: str, data: list[list[str]],
                        *, eyebrow: str = "", subtitle: str = "",
                        col_widths: list[float] | None = None,
                        accent_rows: list[int] | None = None,
                        first_col_bold: bool = True,
                        font_size: int = 11,
                        header_size: int = 12,
                        caption: str = ""):
        """学术表格页:深蓝表头 + 隔行底色 + 重点行橙色高亮。

        data 第一行为表头。
        accent_rows 是从 0 起算的行号(注意 0 是表头,通常用 1+)。
        """
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide, eyebrow=eyebrow, title=title,
                         subtitle=subtitle, page_no=page)

        rows, cols = len(data), len(data[0])
        x = Inches(0.6); y = Inches(1.75)
        w = SLIDE_W - Inches(1.2)
        h = Inches(5.0) if caption else Inches(5.3)

        tbl_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
        tbl = tbl_shape.table
        if col_widths:
            total = sum(col_widths)
            for i, cw in enumerate(col_widths):
                tbl.columns[i].width = int(w * cw / total)
        accent_rows = accent_rows or []

        for r in range(rows):
            for c in range(cols):
                cell = tbl.cell(r, c)
                cell.margin_left = Inches(0.06)
                cell.margin_right = Inches(0.06)
                cell.margin_top = Inches(0.03)
                cell.margin_bottom = Inches(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                if r == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = PRIMARY
                    txt_color = WHITE; bold = True; sz = header_size
                elif r in accent_rows:
                    cell.fill.solid(); cell.fill.fore_color.rgb = TINT_ORG
                    txt_color = ACCENT; bold = True; sz = font_size
                elif r % 2 == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = GRAY_L
                    txt_color = DARK
                    bold = (first_col_bold and c == 0)
                    sz = font_size
                else:
                    cell.fill.solid(); cell.fill.fore_color.rgb = WHITE
                    txt_color = DARK
                    bold = (first_col_bold and c == 0)
                    sz = font_size
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
                _set_run(p.add_run(), str(data[r][c]),
                         size=sz, bold=bold, color=txt_color)

        if caption:
            _add_textbox(slide, x, y + h + Inches(0.08), w, Inches(0.3),
                         caption, size=11, color=MUTED, italic=True,
                         align=PP_ALIGN.CENTER)
        self._add_footer(slide)
        return slide

    # ---------- 8. 图表强调页 ----------

    def add_figure_slide(self, title: str, image_path: str,
                         caption: str = "",
                         *, eyebrow: str = "", subtitle: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide, eyebrow=eyebrow, title=title,
                         subtitle=subtitle, page_no=page)

        max_w = Inches(11.0)
        max_h = Inches(5.0)
        try:
            from PIL import Image
            with Image.open(image_path) as im:
                w, h = im.size
            ratio = min(int(max_w) / w, int(max_h) / h)
            pic_w = Emu(int(w * ratio))
            pic_h = Emu(int(h * ratio))
        except Exception:
            pic_w = max_w
            pic_h = max_h

        left = Emu((int(SLIDE_W) - int(pic_w)) // 2)
        top = Inches(1.7)
        slide.shapes.add_picture(image_path, left, top, pic_w, pic_h)

        if caption:
            _add_textbox(slide, Inches(0.5), SLIDE_H - Inches(0.65),
                         Inches(12.3), Inches(0.4),
                         caption, size=13, color=DARK, bold=True,
                         align=PP_ALIGN.CENTER)
        self._add_footer(slide)
        return slide

    # ---------- 9. 参考文献页 ----------

    def add_references_slide(self, refs: list[str],
                             title: str = "参考文献 / References",
                             *, eyebrow: str = "References",
                             subtitle: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        page = self._next_page()
        self._add_header(slide, eyebrow=eyebrow, title=title,
                         subtitle=subtitle, page_no=page)

        tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.65),
                                       Inches(12.1), Inches(5.5))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for i, ref in enumerate(refs, 1):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.18
            _set_run(p.add_run(), f"[{i}] ", size=12, color=ACCENT, bold=True,
                     font=EN_FONT)
            _set_run(p.add_run(), ref, size=12, color=DARK)
            p.space_after = Pt(3)
        self._add_footer(slide)
        return slide

    # ---------- 10. 结尾页 ----------

    def add_ending(self, big_text: str = "谢谢 / Thank You",
                   contact: str = "",
                   *, sub_text: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._next_page()
        _add_filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PRIMARY)
        # 副文(置上)
        if sub_text:
            _add_textbox(slide, 0, Inches(2.0), SLIDE_W, Inches(0.6),
                         sub_text, size=18, color=GRAY_L, italic=True,
                         align=PP_ALIGN.CENTER)
        _add_textbox(slide, 0, Inches(2.8), SLIDE_W, Inches(1.5),
                     big_text, size=72, color=WHITE, bold=True,
                     align=PP_ALIGN.CENTER)
        _add_filled_rect(slide,
                         Emu(int(SLIDE_W) // 2 - int(Inches(1.0)) // 2),
                         Inches(4.5), Inches(1.0), Inches(0.06), ACCENT)
        if contact:
            _add_textbox(slide, 0, Inches(4.9), SLIDE_W, Inches(0.6),
                         contact, size=16, color=GRAY_L,
                         align=PP_ALIGN.CENTER)
        return slide

    # ---------- 11. Speaker notes ----------

    @staticmethod
    def add_notes(slide, text: str):
        """给指定 slide 添加演讲者备注。"""
        notes_tf = slide.notes_slide.notes_text_frame
        notes_tf.text = text
        for p in notes_tf.paragraphs:
            for r in p.runs:
                r.font.name = CN_FONT
                r.font.size = Pt(11)
                _set_eastasia(r)

    # ---------- 保存 ----------

    def save(self, path: str | Path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(path))
        return path


# ============ CLI 自检 ============

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "/tmp/build_pptx_demo.pptx"
    d = AcademicDeck(total_pages=9, footer_text="literature-deep-skill · 自检 demo")
    d.add_cover(
        title="基于 VLM 的驾驶员状态多维感知与接管能力评估",
        subtitle="—— 模拟器实证权重 × VLM 多维感知 × 端到端系统验证",
        eyebrow="DSM × VLM  ·  Driver Take-Over Readiness · 6-Dim",
        author="自检", date="2026-05", org="literature-deep-skill",
        meeting_tag="Meeting Demo  ·  研究背景 + 现状 + 方案",
    )
    d.add_toc([
        ("Ⅰ", "研究背景", "时代窗口 · 法规 · 三重张力"),
        ("Ⅱ", "研究现状", "DMS 四时代 · VLM 三范式"),
        ("Ⅲ", "研究方案", "3 RQ · 6 维度 · 系统集成"),
    ])
    d.add_part_divider("Ⅰ", "研究背景",
                       "从接管底线切入 · 法规 + 人因数据",
                       total_parts=3)
    d.add_content_slide(
        title="L3 准入落地,2025 年是关键时代窗口",
        eyebrow="Ⅰ · 研究背景",
        subtitle="工信部 L3 试点 / Mercedes Drive Pilot / NHTSA SGO",
        bullets=[
            ("工信部 L3 准入试点 2025-12-15 启动", "多家车企已进入名单 [#19]"),
            ("Mercedes Drive Pilot 全球首款量产 L3", "Nevada / California 已开放"),
            ("NHTSA SGO 1500+ 起 L2/L3 事故报告", "接管失败是高发模式"),
        ],
    )
    d.add_card_grid(
        title="VLM 在 DMS 的三种范式横向对比",
        eyebrow="Ⅱ · 研究现状",
        cards=[
            ("C-i 零样本", ["GPT-4V / Gemini 直接 prompt",
                          "无需训练,精度有限",
                          "适合早期验证"], "blue"),
            ("C-ii LoRA 微调", ["少样本 + 适配器",
                              "工程友好",
                              "可端侧部署"], "orange"),
            ("C-iii 端到端", ["LLaVA 风格大模型",
                            "算力要求高",
                            "推理延迟大"], "green"),
        ],
    )
    d.add_compare_slide(
        title="监督学习 vs VLM 三范式 vs 视频 TAL",
        eyebrow="Ⅱ · 研究现状",
        columns=[
            ("监督学习", ["数据依赖大", "天花板已现", "[#1] [#3] [#11]"]),
            ("VLM 三范式", ["语义灵活", "可解释强", "[#6] [#13] [#15]"]),
            ("视频 TAL", ["时序一致", "AMA mAP 0.83", "[#17] [#18]"]),
        ],
    )
    d.add_table_slide(
        title="6 维度状态空间一览",
        eyebrow="Ⅲ · 研究方案",
        data=[
            ["维度", "变量名", "类别数", "可观测性"],
            ["疲劳", "Drowsiness", "4", "高"],
            ["视线", "Gaze_Direction", "4", "高"],
            ["手部", "Hand_Status", "4", "高"],
            ["分心", "Distraction_Type", "5", "中"],
            ["姿态", "Posture", "3", "高"],
            ["情绪", "Emotion", "3", "中"],
        ],
        accent_rows=[4],
        col_widths=[2, 3, 1, 1.5],
        caption="表 1  6 维度变量空间(Distraction 维含复合分心)",
    )
    d.add_references_slide([
        "First Author, et al. Paper Title. Conference/Journal, Year.",
        "Second Author, et al. Another Paper Title. Conference/Journal, Year.",
        "Third Author, et al. Third Paper Title. Conference/Journal, Year.",
    ])
    d.add_ending(
        "Q & A",
        contact="your-email@example.edu",
        sub_text="—— 期待各位老师与同学的指导意见 ——",
    )
    p = d.save(out)
    print(f"Demo PPT saved: {p}")
