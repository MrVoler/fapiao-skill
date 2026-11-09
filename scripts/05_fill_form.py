# -*- coding: utf-8 -*-
"""Step 5 填单：把费用名称、摘要、金额（含分格与大小写）、附单据张数填进报销单模板。

用法:
    python 05_fill_form.py \
        --fee "差旅费" --summary "某项目出差" --amount 1234.56 --bills 8 \
        --out ./work/报销单.docx

    # 日期默认今天，也可指定
    --date "2026 年 11 月 9 日"

设计要点:
    报销人、经理审批故意留空，供手写。
    模板若为两联，只保留第一联（--keep-copies 1）。
"""
import argparse
import copy
import datetime
import os

from _common import die, money, rmb_upper

# 金额分格表头（从右往左）
UNIT_CHARS = {'分', '角', '元', '十', '百', '千', '万', '亿'}


def strip_empty_tail(body, qn):
    while True:
        kids = [c for c in body if c.tag != qn('w:sectPr')]
        if kids and kids[-1].tag == qn('w:p') and not ''.join(kids[-1].itertext()).strip():
            body.remove(kids[-1])
        else:
            break


def in_textbox(el, qn):
    p = el.getparent()
    while p is not None:
        if p.tag == qn('w:txbxContent'):
            return True
        p = p.getparent()
    return False


def rewrite_cell(cell, text, bold=False, color=None, size=12):
    """清空单元格首个段落并写入文本，字体统一宋体。"""
    from docx.oxml.ns import qn
    from docx.shared import Pt
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    if bold:
        run.font.bold = True
    if color is not None:
        run.font.color.rgb = color
    return run


def main():
    ap = argparse.ArgumentParser(description='填写费用报销单')
    ap.add_argument('--template', required=True, help='报销单模板 docx')
    ap.add_argument('--fee', required=True, help='费用名称（必须先和用户确认）')
    ap.add_argument('--summary', required=True, help='摘要（必须先和用户确认）')
    ap.add_argument('--amount', required=True, help='金额，如 1234.56')
    ap.add_argument('--bills', default=None, help='附单据张数，默认留空')
    ap.add_argument('--date', default=None, help='日期，默认今天')
    ap.add_argument('--keep-copies', type=int, default=1, help='保留几联，默认 1')
    ap.add_argument('--out', required=True)
    args = ap.parse_args()

    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    amount = money(args.amount)
    cap = rmb_upper(amount)
    today = args.date or datetime.date.today().strftime('%Y 年 %m 月 %d 日')

    doc = Document(args.template)
    body = doc.element.body

    # ---- 1. 只保留前 N 联 ----
    # 一联 = 标题段 + 日期段 + 表格 + 签字段；按标题段切分边界整体删除，
    # 只按表格位置删会残留第二联的标题和日期。
    tables_el = [c for c in body if c.tag == qn('w:tbl')]
    if len(tables_el) > args.keep_copies:
        kids = [c for c in body if c.tag != qn('w:sectPr')]
        starts = [i for i, c in enumerate(kids)
                  if c.tag == qn('w:p') and '报销单' in ''.join(c.itertext())]
        if len(starts) > args.keep_copies:
            cut = kids[starts[args.keep_copies]]
        else:
            cut = tables_el[args.keep_copies]  # 兜底：按表格切
        removing = False
        for ch in list(body):
            if ch.tag == qn('w:sectPr'):
                break
            if ch is cut:
                removing = True
            if removing:
                body.remove(ch)
    strip_empty_tail(body, qn)

    # ---- 2. 日期 ----
    date_done = False
    for p in body.iter(qn('w:p')):
        for t in p.iter(qn('w:t')):
            if in_textbox(t, qn) or not t.text:
                continue
            if '年' in t.text and '日' in t.text and '月' in t.text:
                t.text = today
                t.set(qn('xml:space'), 'preserve')
                date_done = True
                pPr = p.find(qn('w:pPr'))
                if pPr is not None:
                    ind = pPr.find(qn('w:ind'))
                    if ind is not None:
                        pPr.remove(ind)
                    jc = pPr.find(qn('w:jc'))
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.insert(0, jc)
                    jc.set(qn('w:val'), 'center')
                break
        if date_done:
            break
    if not date_done:
        print('[注意] 没找到日期段，跳过')

    # ---- 3. 附单据 N 张（WPS 文本框） ----
    if args.bills:
        for t in body.iter(qn('w:t')):
            if t.text and '附单据' in t.text:
                t.text = f' 附单据   {args.bills}  张 '
                t.set(qn('xml:space'), 'preserve')

    # ---- 4. 定位表格、表头、明细行 ----
    tbl = doc.tables[0]
    fee_col = summary_col = None
    units_row = None
    for ri, row in enumerate(tbl.rows):
        texts = [c.text.strip() for c in row.cells]
        if fee_col is None and '费用名称' in texts:
            fee_col = texts.index('费用名称')
        if summary_col is None and '摘要' in texts:
            summary_col = texts.index('摘要')
        if '分' in texts and '角' in texts and '元' in texts:
            units_row = ri
            break
    if fee_col is None or summary_col is None:
        die('模板里没找到「费用名称 / 摘要」表头，请检查模板')
    if units_row is None:
        units_row = 0
        print('[注意] 未识别金额分格表头，按最后 10 列处理')

    detail_row = units_row + 1
    n_cols = len(tbl.rows[units_row].cells)
    amount_cols = list(range(summary_col + 1, n_cols))
    if len(amount_cols) < 2:
        die('金额分格列数异常：' + str(amount_cols))

    # ---- 5. 填费用名称 / 摘要 ----
    rewrite_cell(tbl.cell(detail_row, fee_col), args.fee)
    rewrite_cell(tbl.cell(detail_row, summary_col), args.summary)

    # ---- 6. 金额分格：末位对齐，¥ 放在首位数字左侧 ----
    cents = str(int(amount * 100))
    if len(cents) > len(amount_cols):
        die(f'金额 {amount} 超出模板位数（最多 {len(amount_cols)} 位）')
    start = len(amount_cols) - len(cents)
    for i, col in enumerate(amount_cols):
        txt = '' if i < start else cents[i - start]
        if txt:
            rewrite_cell(tbl.cell(detail_row, col), txt, bold=True)
    yen_col = amount_cols[start - 1] if start > 0 else None
    if yen_col is not None:
        rewrite_cell(tbl.cell(detail_row, yen_col), '¥', bold=True)

    # ---- 7. 总计金额（小写 / 大写） ----
    red = RGBColor(0xC0, 0x00, 0x00)
    hit_small = hit_cap = False
    for row in tbl.rows:
        first = row.cells[0].text
        if '小写' in first and not hit_small:
            p = row.cells[0].paragraphs[0]
            run = p.add_run(f'¥{amount}')
            run.font.size = Pt(12)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.bold = True
            run.font.color.rgb = red
            hit_small = True
        elif '大写' in first and not hit_cap:
            p = row.cells[0].paragraphs[0]
            run = p.add_run(cap)
            run.font.size = Pt(12)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.bold = True
            run.font.color.rgb = red
            hit_cap = True
    if not (hit_small and hit_cap):
        print('[注意] 没找到「总计金额」行，请人工补填')

    # ---- 8. 末尾分页，让发票从第 2 页开始 ----
    sect = body[-1]
    pb = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '2')
    rPr.append(sz)
    pPr.append(rPr)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:line'), '20')
    spacing.set(qn('w:lineRule'), 'exact')
    pPr.append(spacing)
    pb.append(pPr)
    r = OxmlElement('w:r')
    r.append(copy.deepcopy(rPr))
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    pb.append(r)
    sect.addprevious(pb)

    out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)

    print(f'费用名称: {args.fee}')
    print(f'摘要    : {args.summary}')
    print(f'金额    : ¥{amount}  →  分格 {cents}  大写 {cap}')
    if args.bills:
        print(f'附单据  : {args.bills} 张')
    print(f'日期    : {today}')
    print(f'已生成  : {out}')


if __name__ == '__main__':
    main()
