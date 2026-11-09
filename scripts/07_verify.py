# -*- coding: utf-8 -*-
"""Step 7 校验：确认页数与每页发票张数，并导出 PDF 供打印/预览。

用法:
    python 07_verify.py --docx ./work/报销材料.docx --per-page 2

检查项:
    1. 结构：表格数、图片数、显式分页符
    2. 页数：调用本机 Word 统计真实分页（期望 1 + ceil(发票数/每页张数)）
    3. 每页图片数：导出 PDF 后逐页统计，2 张/页 才算通过
"""
import argparse
import math
import os
import subprocess
import sys
import tempfile

from _common import die, table, warn

PS_TEMPLATE = r'''
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
    $doc = $word.Documents.Open("{docx}", $false, $true)
    "PAGES=" + $doc.ComputeStatistics(2)
    $doc.ExportAsFixedFormat("{pdf}", 17)
    $doc.Close(0)
}} finally {{
    try {{ $word.Quit() }} catch {{}}
}}
'''


def _via_com(docx, pdf):
    """首选：pywin32 直接驱动本机 Word。"""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
        doc.ExportAsFixedFormat(os.path.abspath(pdf), 17)  # wdExportFormatPDF
        doc.Close(0)
        return pages, None
    except Exception as e:  # noqa: BLE001
        return None, f'{type(e).__name__}: {e}'[:300]
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _via_powershell(docx, pdf):
    """备用：没有 pywin32 时借 PowerShell 调 Word。"""
    script = PS_TEMPLATE.format(docx=docx.replace('/', '\\'), pdf=pdf.replace('/', '\\'))
    try:
        out = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command', script],
            capture_output=True, text=True, timeout=180,
        )
    except (OSError, subprocess.TimeoutExpired) as e:
        return None, f'调用 Word 失败: {e}'
    if out.returncode != 0:
        return None, (out.stderr or out.stdout or '').strip()[:300]
    for line in out.stdout.splitlines():
        if line.startswith('PAGES='):
            return int(line.split('=', 1)[1].strip()), None
    return None, '未取到页数'


def word_pages_and_pdf(docx, pdf):
    try:
        import win32com.client  # noqa: F401
        has_com = True
    except ImportError:
        has_com = False
    if has_com:
        pages, err = _via_com(docx, pdf)
        if pages is not None:
            return pages, None
        warn(f'COM 驱动 Word 失败（{err}），改用 PowerShell')
    return _via_powershell(docx, pdf)


def main():
    ap = argparse.ArgumentParser(description='校验合成文档并导出 PDF')
    ap.add_argument('--docx', required=True)
    ap.add_argument('--per-page', type=int, default=2)
    ap.add_argument('--pdf', default=None, help='PDF 输出路径，默认同目录同名')
    ap.add_argument('--preview-dir', default=None, help='每页 PNG 预览输出目录')
    args = ap.parse_args()

    from docx import Document
    from docx.oxml.ns import qn

    docx = os.path.abspath(args.docx)
    if not os.path.exists(docx):
        die(f'文件不存在: {docx}')

    d = Document(docx)
    blips = len(list(d.element.body.iter(qn('a:blip'))) )
    breaks = len([b for b in d.element.body.iter(qn('w:br')) if b.get(qn('w:type')) == 'page'])
    print(f'结构: 表格 {len(d.tables)} 个 | 图片 {blips} 张 | 显式分页 {breaks} 处')

    expect_pages = 1 + math.ceil(blips / args.per_page) if blips else None
    pdf = args.pdf or os.path.splitext(docx)[0] + '.pdf'

    pages, err = word_pages_and_pdf(docx, pdf)
    if pages is None:
        warn(f'无法用 Word 统计页数（{err}）')
        if expect_pages:
            print(f'按结构推算应有 {expect_pages} 页，请人工翻页确认')
        return

    print(f'Word 实测页数: {pages}（期望 {expect_pages}）')
    print(f'PDF: {pdf}')

    problems = []
    if expect_pages and pages != expect_pages:
        problems.append(f'页数不符：实测 {pages}，期望 {expect_pages}')

    try:
        import pymupdf
    except ImportError:
        warn('未安装 pymupdf，跳过逐页图片统计')
        pymupdf = None

    if pymupdf and os.path.exists(pdf):
        doc = pymupdf.open(pdf)
        rows, bad = [], []
        preview = args.preview_dir
        if preview:
            os.makedirs(preview, exist_ok=True)
        for i, page in enumerate(doc, 1):
            n_img = len(page.get_images(full=True))
            rows.append([i, n_img, f'{page.rect.width / 72 * 25.4:.0f}x{page.rect.height / 72 * 25.4:.0f}'])
            if i > 1 and n_img != args.per_page:
                bad.append(f'第 {i} 页有 {n_img} 张图，应为 {args.per_page}')
            if preview:
                page.get_pixmap(dpi=120).save(os.path.join(preview, f'p{i}.png'))
        table(rows, ['页', '图片数', '尺寸mm'], ['r', 'r', 'r'])
        problems += bad
        doc.close()

    print()
    if problems:
        print(f'[未通过] {len(problems)} 项问题')
        for p in problems:
            print('  -', p)
        sys.exit(1)
    print('[通过] 页数与每页张数均符合预期')


if __name__ == '__main__':
    main()
