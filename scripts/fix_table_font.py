# -*- coding: utf-8 -*-
"""
把已有 .docx 中**所有表格单元格**的文字字体统一改成指定中文字体。

用途（公司文档改格式场景）：
    1) 表内文字要统一为仿宋_GB2312（公司规范，2026-09-13 用户明确）；
    2) 表内形似标题的文字（"一、""（一）"等）**不当作标题**，一并按正文处理——
       本脚本的做法是"无差别覆盖表内全部 run"，天然满足该要求。

用法：
    python fix_table_font.py 源文件.docx [-o 输出.docx] [--zh 仿宋_GB2312] [--en "Times New Roman"] [--size 10.5]
    # 不指定 -o 时，输出为「源文件名（表格仿宋修正）.docx」

⚠️ 关键实现要点（踩过两次坑，勿改）：
    - `run.font.name` 只写西文槽位（w:ascii / w:hAnsi）；
      中文字体**必须单独写 w:eastAsia**，且**必须写在最后**，
      否则先设的中文字体被后续的西文设置覆盖 → Word 里中文回退成宋体，
      表现就是"表格中文没有用仿宋"。
    - 除 run 之外，还要同步 **段落标记**（pPr/rPr/rFonts）的字体，
      否则 Word 字体框显示、空单元格新建字符的默认字体不对。
    - 合并单元格：`row.cells` 对纵向合并会返回同一对象，必须用 id(cell._tc) 去重。
"""
import argparse
import os
import sys

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_rfonts(rFonts, zh, en):
    """按正确顺序写字体槽位：西文 → 中文（中文最后，不能被覆盖）"""
    rFonts.set(qn('w:ascii'), en)
    rFonts.set(qn('w:hAnsi'), en)
    rFonts.set(qn('w:cs'), en)
    rFonts.set(qn('w:eastAsia'), zh)   # ← 必须最后
    rFonts.set(qn('w:hint'), 'eastAsia')


def fix_run(run, zh, en, size):
    rPr = run._element.get_or_add_rPr()
    set_rfonts(rPr.get_or_add_rFonts(), zh, en)
    # 清掉字符样式引用，防止样式中的字体把显式设置盖回去
    el = rPr.find(qn('w:rStyle'))
    if el is not None:
        rPr.remove(el)
    if size:
        run.font.size = Pt(size)


def fix_para_mark(para, zh, en, size):
    """同步段落标记（¶）的字体"""
    pPr = para._p.get_or_add_pPr()
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        pPr.append(rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    set_rfonts(rFonts, zh, en)
    if size:
        for tag, val in (('w:sz', str(int(size * 2))), ('w:szCs', str(int(size * 2)))):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rPr.append(el)
            el.set(qn('w:val'), val)


def iter_tables(tables):
    """递归遍历表格（含嵌套表格）"""
    for t in tables:
        yield t
        for row in t.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:      # 去重：合并单元格会重复返回
                    continue
                seen.add(id(cell._tc))
                if cell.tables:
                    yield from iter_tables(cell.tables)


def unique_cells(table):
    for row in table.rows:
        seen = set()
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            yield cell


def main():
    ap = argparse.ArgumentParser(description='统一 docx 表格内文字字体')
    ap.add_argument('src', help='源 .docx 路径')
    ap.add_argument('-o', '--out', help='输出 .docx 路径')
    ap.add_argument('--zh', default='仿宋_GB2312', help='中文字体（默认 仿宋_GB2312）')
    ap.add_argument('--en', default='Times New Roman', help='西文字体（默认 Times New Roman）')
    ap.add_argument('--size', type=float, default=10.5, help='字号磅值（默认 10.5 = 五号）')
    ap.add_argument('--no-size', action='store_true', help='不修改字号')
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.isfile(src):
        sys.exit(f'源文件不存在：{src}')
    out = args.out or os.path.join(
        os.path.dirname(src),
        os.path.splitext(os.path.basename(src))[0] + '（表格仿宋修正）.docx')
    out = os.path.abspath(out)
    if os.path.abspath(src) == out:
        sys.exit('输出文件不能与源文件相同')
    size = None if args.no_size else args.size

    doc = Document(src)
    print(f'读取：{src}')
    print(f'段落 {len(doc.paragraphs)}，表格 {len(doc.tables)}')

    n_cell = n_run = n_mark = 0
    for table in iter_tables(doc.tables):
        for cell in unique_cells(table):
            n_cell += 1
            for para in cell.paragraphs:
                fix_para_mark(para, args.zh, args.en, size)
                n_mark += 1
                if para.runs:
                    for run in para.runs:
                        fix_run(run, args.zh, args.en, size)
                        n_run += 1
                elif para.text.strip():          # 有文字却没 run，补一个
                    fix_run(para.add_run(para.text), args.zh, args.en, size)
                    n_run += 1

    try:
        doc.save(out)
    except PermissionError:
        sys.exit(f'输出文件被占用（可能已在 Word 中打开），请关闭后重试：{out}')

    print(f'单元格 {n_cell}，run {n_run}，段落标记 {n_mark}')
    print(f'已保存：{out}')

    # ---- 校验 ----
    chk = Document(out)
    stat = {}
    for table in iter_tables(chk.tables):
        for cell in unique_cells(table):
            for para in cell.paragraphs:
                for run in para.runs:
                    rPr = run._element.rPr
                    rf = rPr.rFonts if rPr is not None else None
                    key = rf.get(qn('w:eastAsia')) if rf is not None else '<无 rFonts>'
                    stat[key or '<未设>'] = stat.get(key or '<未设>', 0) + 1
    print('校验 · 表内 run 中文字体分布：', stat)
    bad = [k for k in stat if k not in (args.zh,)]
    print('✅ 全部为「%s」' % args.zh if not bad else f'⚠️ 仍存在非目标字体：{bad}')


if __name__ == '__main__':
    main()
