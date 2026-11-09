# -*- coding: utf-8 -*-
"""Step 6 合并：把填好的报销单页与发票排版表拼成一个 Word。

用法:
    python 06_merge.py --form ./work/报销单.docx --invoice ./work/发票.docx --out ./work/报销材料.docx

为什么不用 Word COM 的 InsertFile:
    当文档第一个块是表格时，InsertFile 会把插入点落在表格首个单元格里，
    结果整张表单被塞进单元格。改用「表格深拷贝 + 图片关系重绑」，与路径、顺序完全解耦。
"""
import argparse
import copy
import io
import os
import tempfile

from _common import die


def rebind_images(dst_doc, src_doc, tbl_el, qn):
    """深拷贝后图片的 r:embed 仍指向源文档，需要在本包里重新登记并改写。"""
    blips = list(tbl_el.iter(qn('a:blip')))
    tmp_dir = tempfile.mkdtemp(prefix='fapiao_img_')
    for i, blip in enumerate(blips):
        old_rid = blip.get(qn('r:embed'))
        if not old_rid:
            continue
        try:
            image_part = src_doc.part.related_parts[old_rid]
            blob = image_part.blob
            ext = os.path.splitext(str(image_part.partname))[1] or '.png'
        except KeyError:
            die(f'源文档缺少图片关系 {old_rid}')

        new_rid = None
        # 优先用流式登记，避免落盘
        if hasattr(dst_doc.part, 'get_or_add_image_part'):
            try:
                _, new_rid = dst_doc.part.get_or_add_image_part(io.BytesIO(blob))
            except Exception:
                new_rid = None
        if new_rid is None:
            tmp = os.path.join(tmp_dir, f'{i:03d}{ext}')
            with open(tmp, 'wb') as f:
                f.write(blob)
            new_rid, _ = dst_doc.part.get_or_add_image(tmp)
        blip.set(qn('r:embed'), new_rid)
    return len(blips)


def main():
    ap = argparse.ArgumentParser(description='合并报销单与发票页')
    ap.add_argument('--form', required=True, help='填好的报销单（单联）')
    ap.add_argument('--invoice', required=True, help='发票排版文档')
    ap.add_argument('--out', required=True)
    args = ap.parse_args()

    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Mm

    for p in (args.form, args.invoice):
        if not os.path.exists(p):
            die(f'文件不存在: {p}')

    dst = Document(args.form)
    src = Document(args.invoice)

    # 统一页边距：发票表格按 184mm 设计，模板自带边距通常装不下
    mis = src.sections[0]
    for doc_sec in dst.sections:
        doc_sec.left_margin = mis.left_margin
        doc_sec.right_margin = mis.right_margin
        doc_sec.top_margin = mis.top_margin
        doc_sec.bottom_margin = mis.bottom_margin
        doc_sec.page_width = mis.page_width
        doc_sec.page_height = mis.page_height

    if not src.tables:
        die('发票文档里没有表格，请先运行 04_build_invoice_docx.py')

    tbl = copy.deepcopy(src.tables[0]._tbl)
    n = rebind_images(dst, src, tbl, qn)

    body = dst.element.body
    sect = body[-1]
    sect.addprevious(tbl)

    out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    dst.save(out)

    d = Document(out)
    print(f'发票图片 {n} 张，表格 {len(d.tables)} 个')
    print(f'已生成: {out}  ({os.path.getsize(out) / 1024 / 1024:.2f} MB)')


if __name__ == '__main__':
    main()
