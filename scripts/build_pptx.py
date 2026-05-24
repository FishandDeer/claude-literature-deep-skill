#!/usr/bin/env python3
"""
build_pptx.py — 学术风 python-pptx 模板封装

视觉规范完全遵循 references/ppt-design.md:
  - 16:9, 深蓝主色 (#1F3864) + 橙色强调 (#ED7D31)
  - 中文 Microsoft YaHei (Windows) / Noto Sans CJK SC (Linux fallback)
  - 行动式标题、单论点页、低密度

用法(作为库导入):

    from build_pptx import AcademicDeck

    deck = AcademicDeck()
    deck.add_cover(title="研究汇报",
                   subtitle="组会报告",
                   author="张三", date="2026-05",
                   org="某某实验室")
    deck.add_toc(["研究背景", "相关工作", "研究问题", "下一步计划"])
    deck.add_section_header(number="01", name="研究背景",
                            tagline="问题定义 / 现状 / 切入角度")
    deck.add_content_slide(
        title="行动式标题示例:核心论点放在标题里",
        bullets=[
            ("第一个论据,带数据 [#1]", None),
            ("第二个论据,带次级注释 [#2]", "次级注释用更小字号"),
            ("第三个论据 [#3]", None),
        ],
    )
    deck.add_compare_slide(
        title="方法 A vs 方法 B vs 方法 C 的优劣",
        columns=[
            ("方法 A", ["优点 1", "优点 2", "局限 1"]),
            ("方法 B", ["优点 1", "优点 2", "局限 1"]),
            ("方法 C", ["优点 1", "优点 2", "局限 1"]),
        ],
    )
    deck.save("/path/to/output.pptx")
"""

from __future__ import annotations

import platform
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt


# ============ 视觉常量(改这里一处即可全局生效)============

PRIMARY = RGBColor(0x1F, 0x38, 0x64)       # 深蓝主色
SECONDARY = RGBColor(0x2F, 0x55, 0x97)     # 副色
ACCENT = RGBColor(0xED, 0x7D, 0x31)        # 橙色强调
TEXT_DARK = RGBColor(0x33, 0x33, 0x33)
TEXT_MUTE = RGBColor(0x66, 0x66, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PALE = RGBColor(0xD9, 0xE2, 0xF3)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# 中文字体:Windows 端 Microsoft YaHei,Linux 端如未安装则需先 `apt install fonts-noto-cjk`
CN_FONT = "Microsoft YaHei" if platform.system() == "Windows" else "Noto Sans CJK SC"
EN_FONT = "Calibri"


# ============ 工具函数 ============

def _set_run(run, text: str, *, size: int, color=TEXT_DARK, bold=False,
             font=CN_FONT):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    # 同时设置远东字体,避免中文回落到默认西文字体导致方框
    rPr = run._r.get_or_add_rPr()
    from pptx.oxml.ns import qn
    eastAsia = rPr.find(qn("a:ea"))
    if eastAsia is None:
        from lxml import etree
        eastAsia = etree.SubElement(rPr, qn("a:ea"))
    eastAsia.set("typeface", CN_FONT)


def _add_textbox(slide, left, top, width, height, text, *,
                 size=18, color=TEXT_DARK, bold=False,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    _set_run(p.add_run(), text, size=size, color=color, bold=bold)
    # 移除默认空 run
    if p.runs[0].text == "":
        p._p.remove(p.runs[0]._r)
    return tb


def _add_filled_rect(slide, left, top, width, height, fill_color,
                     line_color=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
    shp.shadow.inherit = False
    return shp


# ============ 主类 ============

@dataclass
class AcademicDeck:
    """学术 PPT 构建器。每个 `add_*` 方法对应一种版式。"""

    def __post_init__(self):
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self._blank_layout = self.prs.slide_layouts[6]

    # ---- 通用页脚 ----
    def _add_footer(self, slide, page_num: int | None = None):
        if page_num is not None:
            _add_textbox(slide, Inches(12.5), Inches(7.15), Inches(0.8), Inches(0.3),
                         str(page_num), size=10, color=TEXT_MUTE, align=PP_ALIGN.RIGHT)

    # ---- 1. 封面页 ----
    def add_cover(self, title: str, subtitle: str = "", author: str = "",
                  date: str = "", org: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)

        # 左侧深蓝色装饰条
        _add_filled_rect(slide, 0, 0, Inches(0.6), SLIDE_H, PRIMARY)

        # 主标题
        _add_textbox(slide, Inches(1.2), Inches(2.4), Inches(11.0), Inches(1.5),
                     title, size=40, color=PRIMARY, bold=True)
        # 副标题
        if subtitle:
            _add_textbox(slide, Inches(1.2), Inches(3.9), Inches(11.0), Inches(0.8),
                         subtitle, size=22, color=SECONDARY)
        # 橙色横线
        _add_filled_rect(slide, Inches(1.2), Inches(4.85), Inches(1.5), Emu(38100),
                         ACCENT)

        # 作者 / 日期 / 单位
        meta_parts = [p for p in [author, org, date] if p]
        if meta_parts:
            _add_textbox(slide, Inches(1.2), Inches(5.2), Inches(11.0), Inches(0.5),
                         "  ·  ".join(meta_parts), size=16, color=TEXT_MUTE)
        return slide

    # ---- 2. 目录页 ----
    def add_toc(self, items: Iterable[str]):
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._add_top_title_bar(slide, "目录 / Contents")

        items = list(items)
        top0 = Inches(1.8)
        step = Inches(0.7)
        for i, name in enumerate(items, 1):
            # 圆点编号
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
            _set_run(p.add_run(), f"{i:02d}", size=14, color=WHITE, bold=True)
            if p.runs[0].text == "":
                p._p.remove(p.runs[0]._r)

            _add_textbox(slide, Inches(2.9), top0 + step * (i - 1),
                         Inches(9.0), Inches(0.55),
                         name, size=22, color=TEXT_DARK,
                         anchor=MSO_ANCHOR.MIDDLE)
        return slide

    # ---- 3. 章节扉页 ----
    def add_section_header(self, number: str, name: str, tagline: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        _add_filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PRIMARY)

        # 大号编号
        _add_textbox(slide, Inches(1.2), Inches(2.0), Inches(4.0), Inches(2.5),
                     number, size=120, color=ACCENT, bold=True)
        # 章节名
        _add_textbox(slide, Inches(5.5), Inches(2.6), Inches(7.0), Inches(1.0),
                     name, size=40, color=WHITE, bold=True)
        if tagline:
            _add_textbox(slide, Inches(5.5), Inches(3.7), Inches(7.0), Inches(0.6),
                         tagline, size=18, color=PALE)
        return slide

    # ---- 4. 通用标题条 ----
    def _add_top_title_bar(self, slide, title: str):
        _add_filled_rect(slide, 0, 0, SLIDE_W, Inches(0.9), PRIMARY)
        _add_textbox(slide, Inches(0.5), Inches(0.15), Inches(12.5), Inches(0.6),
                     title, size=24, color=WHITE, bold=True,
                     anchor=MSO_ANCHOR.MIDDLE)
        # 橙色细底线
        _add_filled_rect(slide, 0, Inches(0.9), SLIDE_W, Emu(38100), ACCENT)

    # ---- 5. 内容页(标题 + bullets,每 bullet 可带二级注释)----
    def add_content_slide(self, title: str,
                          bullets: list[tuple[str, str | None]]):
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._add_top_title_bar(slide, title)

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.3),
                                       Inches(11.7), Inches(5.7))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for primary, secondary in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = PP_ALIGN.LEFT
            _set_run(p.add_run(), "• ", size=20, color=PRIMARY, bold=True)
            _set_run(p.add_run(), primary, size=20, color=TEXT_DARK)
            p.space_after = Pt(6)
            if secondary:
                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.LEFT
                _set_run(p2.add_run(), "    " + secondary, size=14,
                         color=TEXT_MUTE)
                p2.space_after = Pt(10)
        return slide

    # ---- 6. 对比页(2~3 列) ----
    def add_compare_slide(self, title: str,
                          columns: list[tuple[str, list[str]]]):
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._add_top_title_bar(slide, title)

        n = len(columns)
        assert 2 <= n <= 3, "compare_slide supports 2 or 3 columns"
        margin = Inches(0.5)
        gap = Inches(0.3)
        total_w = SLIDE_W - margin * 2 - gap * (n - 1)
        col_w = Emu(int(total_w) // n)
        top = Inches(1.3)
        col_h = Inches(5.7)

        for i, (header, items) in enumerate(columns):
            left = margin + (col_w + gap) * i
            # 列头
            _add_filled_rect(slide, left, top, col_w, Inches(0.7),
                             SECONDARY)
            _add_textbox(slide, left, top, col_w, Inches(0.7),
                         header, size=18, color=WHITE, bold=True,
                         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            # 列体
            _add_filled_rect(slide, left, top + Inches(0.7), col_w,
                             col_h - Inches(0.7), PALE)
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
                _set_run(p.add_run(), "• ", size=14, color=PRIMARY, bold=True)
                _set_run(p.add_run(), item, size=14, color=TEXT_DARK)
                p.space_after = Pt(6)
        return slide

    # ---- 7. 图表强调页(标题 + 大图 + 底部一句话结论) ----
    def add_figure_slide(self, title: str, image_path: str,
                         caption: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._add_top_title_bar(slide, title)

        # 计算图片尺寸:留 caption 区
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
        top = Inches(1.3)
        slide.shapes.add_picture(image_path, left, top, pic_w, pic_h)

        if caption:
            _add_textbox(slide, Inches(0.5), Inches(6.6), Inches(12.3),
                         Inches(0.6), caption, size=14, color=TEXT_MUTE,
                         align=PP_ALIGN.CENTER, bold=True)
        return slide

    # ---- 8. 参考文献页 ----
    def add_references_slide(self, refs: list[str], title: str = "参考文献 / References"):
        slide = self.prs.slides.add_slide(self._blank_layout)
        self._add_top_title_bar(slide, title)

        tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.2),
                                       Inches(12.3), Inches(6.0))
        tf = tb.text_frame
        tf.word_wrap = True
        first = True
        for i, ref in enumerate(refs, 1):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = PP_ALIGN.LEFT
            _set_run(p.add_run(), f"[{i}] ", size=12, color=ACCENT, bold=True)
            _set_run(p.add_run(), ref, size=12, color=TEXT_DARK)
            p.space_after = Pt(3)
        return slide

    # ---- 9. 结尾页 ----
    def add_ending(self, big_text: str = "谢谢 / Thank You",
                   contact: str = ""):
        slide = self.prs.slides.add_slide(self._blank_layout)
        _add_filled_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PRIMARY)
        _add_textbox(slide, 0, Inches(2.8), SLIDE_W, Inches(1.5),
                     big_text, size=72, color=WHITE, bold=True,
                     align=PP_ALIGN.CENTER)
        _add_filled_rect(slide, Emu(int(SLIDE_W) // 2 - int(Inches(1.0)) // 2),
                         Inches(4.5), Inches(1.0), Emu(38100), ACCENT)
        if contact:
            _add_textbox(slide, 0, Inches(4.9), SLIDE_W, Inches(0.6),
                         contact, size=16, color=PALE,
                         align=PP_ALIGN.CENTER)
        return slide

    # ---- 保存 ----
    def save(self, path: str | Path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(path))
        return path


# ============ CLI 自检 ============

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "/tmp/build_pptx_demo.pptx"
    d = AcademicDeck()
    d.add_cover("学术 PPT 模板自检", "literature-deep-skill",
                author="自检", date="2026-05", org="本地测试")
    d.add_toc(["研究背景", "相关工作", "方法", "实验", "结论"])
    d.add_section_header("01", "研究背景", "问题定义与现状")
    d.add_content_slide("行动式标题示例:核心论点放在标题里",
        [("第一个论据,带数据 [#1]", None),
         ("第二个论据,带次级注释 [#2]", "次级注释用更小字号、灰色"),
         ("第三个论据 [#3]", None)])
    d.add_compare_slide("方法 A vs 方法 B vs 方法 C",
        [("方法 A", ["优点 1", "优点 2", "局限 1"]),
         ("方法 B", ["优点 1", "优点 2", "局限 1"]),
         ("方法 C", ["优点 1", "优点 2", "局限 1"])])
    d.add_references_slide([
        "First Author, et al. Paper Title. Conference/Journal, Year.",
        "Second Author, et al. Another Paper Title. Conference/Journal, Year.",
    ])
    d.add_ending("Q & A", "your-email@example.edu")
    p = d.save(out)
    print(f"Demo PPT saved: {p}")
