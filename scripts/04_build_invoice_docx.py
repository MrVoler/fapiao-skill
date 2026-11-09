# -*- coding: utf-8 -*-
"""Step 4 排版：把发票图片排成打印用 Word，每页 A4 纵向放 N 张（默认 2 张）。

用法:
    python 04_build_invoice_docx.py --manifest ./work/amounts.json --out ./work/发票.docx

排版规则:
    行高 = (可用高度 - 间隔×(每页张数-1)) / 每页张数，行高精确固定且 cantSplit，
    图片在单元格内等比缩放居中。末尾补 1pt 段落，抵消 Word 自动补的空段撑页。
"""
import argparse
import math
import os

from _common import load_json

PAGE_W, PAGE_H = 210.0, 297.0


def main():
    ap = argparse.ArgumentParser(description='生成每页 2 张发票的打印文档')
    ap.add_argument('--manifest', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--per-page', type=int, default=2, help='每页放几张，默认 2')
    ap.add_argument('--margin', type=float, default=12.0, help='页边距 mm')
    ap.add_argument('--gap', type=float, default=8.0, help='两张之间的间隔 mm')
    ap.add_argument('--box-w', type=float, default=0.0, help='图片区最大宽，默认=页宽-2边距-2')
    args = ap.parse_args()

    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt
    from PIL import Image

    manifest = load_json(args.manifest)
    items = manifest['items']
    if not items:
        raise SystemExit('[错误] 清单里没有发票')

    usable_w = PAGE_W - args.margin * 2
    usable_h = PAGE_H - args.margin * 2
    box_w = args.box_w or (usable_w - 2)
    row_h = (usable_h - args.gap * (args.per_page - 1)) / args.per_page
    box_h = row_h - 2  # 留一点余量，避免图片顶到行线

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Mm(PAGE_W), Mm(PAGE_H)
    sec.left_margin = sec.right_margin = Mm(args.margin)
    sec.top_margin = sec.bottom_margin = Mm(args.margin)
    sec.header_distance = sec.footer_distance = Mm(0)

    normal = doc.styles['Normal']
    normal.font.name = 'Microsoft YaHei'
    normal.font.size = Pt(9)
    pf = normal.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = 1.0

    table = doc.add_table(rows=len(items), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tbl_pr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        el.set(qn('w:space'), '0')
        borders.append(el)
    tbl_pr.append(borders)

    margins = OxmlElement('w:tblCellMar')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), '0')
        el.set(qn('w:type'), 'dxa')
        margins.append(el)
    tbl_pr.append(margins)

    table.columns[0].width = Mm(usable_w)

    rows = []
    for i, it in enumerate(items):
        row = table.rows[i]
        row.height = Mm(row_h)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))

        cell = row.cells[0]
        cell.width = Mm(usable_w)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        path = it['img']
        with Image.open(path) as im:
            iw, ih = im.size
        dpi = im.info.get('dpi', (300, 300))[0] or 300
        w_mm, h_mm = iw / dpi * 25.4, ih / dpi * 25.4
        scale = min(box_w / w_mm, box_h / h_mm)
        fw, fh = w_mm * scale, h_mm * scale

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.add_run().add_picture(path, width=Mm(fw), height=Mm(fh))
        rows.append([it['seq'], it['label'][:14], it['amount'],
                     f'{fw:.1f}x{fh:.1f}', f'{(i // args.per_page) + 1}'])

    # 末尾 1pt 空段：抵消 Word 为表格后补的默认段落，防止最后一页被顶出新页
    tail = doc.add_paragraph()
    tail.paragraph_format.space_before = Pt(0)
    tail.paragraph_format.space_after = Pt(0)
    tail.paragraph_format.line_spacing = Pt(1)
    tail.add_run('').font.size = Pt(1)

    out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)

    pages = math.ceil(len(items) / args.per_page)
    from _common import table as tprint
    tprint(rows, ['#', '平台', '金额', '放置尺寸mm', '页'], ['r', 'l', 'r', 'r', 'r'])
    print()
    print(f'每页 {args.per_page} 张，行高 {row_h:.1f}mm（可用高 {usable_h:.0f}mm），共 {pages} 页')
    print(f'已生成: {out}  ({os.path.getsize(out) / 1024 / 1024:.2f} MB)')


if __name__ == '__main__':
    main()
