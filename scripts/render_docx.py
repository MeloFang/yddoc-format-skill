#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""按公司版式把 Markdown 渲染成 docx。

格式参数（取自 company-doc-format 技能 + 用户已确认页面设置）：
    页面      A4
    页边距    上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm
    奇偶页不同
    页眉 1.5 / 页脚 2.6 cm
    标题      方正小标宋简体 三号 居中
    一级标题   黑体 四号
    二级标题   楷体_GB2312 小四
    正文      仿宋_GB2312 小四；西文 Times New Roman
    行距      1.5 倍行距（2026-09-09 用户修正，原固定值 28 磅作废）
    落款      正文下空 3 行，日期文本后缩进 4 字符

Markdown 输入约定：
    # 标题                 方正小标宋简体 三号 居中
    > 主送机关              顶格、不缩进（如「各分公司：」）
    ## 一级（一、）         黑体
    ### 二级（（一））       楷体_GB2312
    #### 三级（1.）         仿宋_GB2312（序数后为标题性短语时用）
    ##### 四级（（1））      仿宋_GB2312
    其余行                  正文：仿宋_GB2312 小四，首行缩进 2 字符
    空行分段
    {{SIGN}}            —— 落款单位（下空 3 行后写）
    {{DATE}}            —— 日期（再空 1 行，缩进 4 字符）
    {{ATTACH:文件名}}   —— 附件引用（可选）

判定提醒：序数引出的是**完整叙述句**时按正文写（普通行，仿宋_GB2312）；
只有**标题性短语**才用 ## / ### / #### / ##### 标记。

用法：
    python render_docx.py draft.md out.docx
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# —— 字体 / 字号 ——
FONT_BODY_ZH = "仿宋_GB2312"
FONT_BODY_EN = "Times New Roman"
FONT_TITLE = "方正小标宋简体"
FONT_H1 = "黑体"
FONT_H2 = "楷体_GB2312"

SIZE_TITLE = Pt(16)   # 三号
SIZE_H1 = Pt(14)      # 四号
SIZE_H2 = Pt(12)      # 小四
SIZE_BODY = Pt(12)    # 小四
LINE_SPACING_PT = 28  # 兼容位：外部调用仍按磅传，但默认 _set_para_format 改用 1.5 倍（见下）


def _set_run_font(run, name_zh, name_en=FONT_BODY_EN, size=SIZE_BODY,
                  bold=False):
    run.font.name = name_en
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name_zh)
    rFonts.set(qn("w:ascii"), name_en)
    rFonts.set(qn("w:hAnsi"), name_en)
    run.font.size = size
    run.font.bold = bold


def _set_para_format(p, *, align=None, space_before=None, space_after=None,
                     line_pt=None, first_line_indent=None,
                     line_multiple=1.5):
    """行距默认 1.5 倍（2026-09-09 用户修正，替代固定值 28 磅）。
    line_pt 传值时用固定磅值（EXACTLY，表格单元格等场景）；否则用 line_multiple 倍数。"""
    pf = p.paragraph_format
    if line_pt is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_pt)
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line_multiple
    if align is not None:
        p.alignment = align
    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent


def set_page(doc):
    s = doc.sections[0]
    s.page_height = Cm(29.7)   # A4
    s.page_width = Cm(21.0)
    s.top_margin = Cm(3.7)
    s.bottom_margin = Cm(3.5)
    s.left_margin = Cm(2.8)
    s.right_margin = Cm(2.6)
    s.header_distance = Cm(1.5)
    s.footer_distance = Cm(2.6)
    # 奇偶页不同
    sectPr = s._sectPr
    titlePg = sectPr.find(qn("w:titlePg"))
    if titlePg is None:
        titlePg = OxmlElement("w:titlePg")
        sectPr.append(titlePg)


def add_title(doc, text):
    p = doc.add_paragraph()
    _set_para_format(p, align=WD_ALIGN_PARAGRAPH.CENTER,
                     space_before=Pt(0), space_after=Pt(0))
    r = p.add_run(text)
    _set_run_font(r, FONT_TITLE, FONT_TITLE, SIZE_TITLE)
    return p


def add_h1(doc, text):
    p = doc.add_paragraph()
    _set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT,
                     space_before=Pt(0), space_after=Pt(0),
                     first_line_indent=Cm(0.74))
    r = p.add_run(text)
    _set_run_font(r, FONT_H1, FONT_BODY_EN, SIZE_H1)
    return p


def add_h2(doc, text):
    p = doc.add_paragraph()
    _set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT,
                     space_before=Pt(0), space_after=Pt(0),
                     first_line_indent=Cm(0.74))
    r = p.add_run(text)
    _set_run_font(r, FONT_H2, FONT_BODY_EN, SIZE_BODY)
    return p


def add_h3(doc, text):
    """三级（1.）/四级（（1））标题：仿宋_GB2312 小四。

    序数引出标题性短语时用；若是完整叙述句，调用方应走 add_body。
    """
    p = doc.add_paragraph()
    _set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT,
                     space_before=Pt(0), space_after=Pt(0),
                     first_line_indent=Cm(0.74))
    r = p.add_run(text)
    _set_run_font(r, FONT_BODY_ZH, FONT_BODY_EN, SIZE_BODY)
    return p


def add_flush(doc, text):
    """顶格行：不首行缩进，用于主送机关、附件说明等。"""
    text = text.replace("\u3000", "").replace("\xa0", "").strip()
    p = doc.add_paragraph()
    _set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT,
                     space_before=Pt(0), space_after=Pt(0))
    r = p.add_run(text)
    _set_run_font(r, FONT_BODY_ZH, FONT_BODY_EN, SIZE_BODY)
    return p


def add_body(doc, text):
    # 公司规范：不加粗、无空格空行。空行分段用段落。
    # 正文首行缩进 2 字符 ≈ 0.74cm（小四 * 2 ≈ 24pt ≈ 0.85cm，标准用 0.74）
    text = text.replace("\u3000", "").replace("\xa0", "").strip()
    if not text:
        return None
    p = doc.add_paragraph()
    _set_para_format(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                     space_before=Pt(0), space_after=Pt(0),
                     first_line_indent=Cm(0.74))
    r = p.add_run(text)
    _set_run_font(r, FONT_BODY_ZH, FONT_BODY_EN, SIZE_BODY)
    return p


def add_blank(doc, n=1):
    """加 n 行空段（不写任何字符，仅占位），用来实现『下空 3 行』。"""
    for _ in range(n):
        p = doc.add_paragraph()
        _set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT,
                         space_before=Pt(0), space_after=Pt(0))


def add_sign_line(doc, text, indent_chars=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    # 行距统一 1.5 倍（2026-09-09 用户修正，原固定值 28 磅作废）
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if indent_chars:
        # 首行缩进 N 字符 ≈ N * 0.37cm（小四下粗略估算）
        pf.first_line_indent = Cm(0.37 * indent_chars)
    r = p.add_run(text)
    _set_run_font(r, FONT_BODY_ZH, FONT_BODY_EN, SIZE_BODY)
    return p


def render(md_text, out_path):
    doc = Document()
    set_page(doc)

    # 默认样式改一下，避免 empty 段落默认带 Calibri
    style = doc.styles["Normal"]
    style.font.name = FONT_BODY_EN
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), FONT_BODY_ZH)
    rFonts.set(qn("w:ascii"), FONT_BODY_EN)
    rFonts.set(qn("w:hAnsi"), FONT_BODY_EN)
    style.font.size = SIZE_BODY

    lines = md_text.splitlines()
    i = 0
    sign = None
    date = None
    attach = None
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("# ") and not line.startswith("## "):
            title = line[2:].strip()
            add_title(doc, title)
        elif line.startswith("## "):
            text = line[3:].strip()
            add_h1(doc, text)
        elif line.startswith("### "):
            text = line[4:].strip()
            add_h2(doc, text)
        elif line.startswith("##### "):
            text = line[6:].strip()
            add_h3(doc, text)
        elif line.startswith("#### "):
            text = line[5:].strip()
            add_h3(doc, text)
        elif line.startswith("> "):
            text = line[2:].strip()
            add_flush(doc, text)
        elif line.startswith("{{SIGN}}"):
            sign = line.replace("{{SIGN}}", "").strip() or None
        elif line.startswith("{{DATE}}"):
            date = line.replace("{{DATE}}", "").strip() or None
        elif line.startswith("{{ATTACH:"):
            attach = line[len("{{ATTACH:"):-2].strip() or None
        else:
            add_body(doc, line)
        i += 1

    # 落款：先空 3 行 → 单位（不缩进）→ 空 1 行 → 日期（缩进 4 字符）
    if sign or date:
        add_blank(doc, 3)
        if sign:
            add_sign_line(doc, sign, indent_chars=0)
        if date:
            add_blank(doc, 1)
            add_sign_line(doc, date, indent_chars=4)

    if attach:
        # 附件占一行（用正文样式），前空一行
        add_blank(doc, 1)
        p = doc.add_paragraph()
        _set_para_format(p, align=WD_ALIGN_PARAGRAPH.LEFT,
                         space_before=Pt(0), space_after=Pt(0))
        r = p.add_run(f"附件：{attach}")
        _set_run_font(r, FONT_BODY_ZH, FONT_BODY_EN, SIZE_BODY)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)
    print(f"[render] -> {out_path}")


def main():
    if len(sys.argv) < 3:
        print("usage: python render_docx.py draft.md out.docx")
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        md = f.read()
    render(md, sys.argv[2])


if __name__ == "__main__":
    main()