# -*- coding: utf-8 -*-
"""一条龙：汇集 → 核价 → 转图 → 排版 → 填单 → 合并 → 校验。

用法:
    python pipeline.py --work ./work --src 发票1.zip 发票2.zip \
        --fee "交通费" --summary "9 月市内打车" --out ./work/报销材料.docx

费用名称与摘要必须由用户确认后显式传入，脚本不设默认值。
"""
import argparse
import math
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TPL = os.path.join(HERE, '..', 'assets', '费用报销单模板.docx')


def run(script, *args):
    cmd = [sys.executable, os.path.join(HERE, script)] + list(args)
    print(f'\n=== {script} ===', flush=True)
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f'[中断] {script} 执行失败', file=sys.stderr)
        sys.exit(r.returncode)


def main():
    ap = argparse.ArgumentParser(description='发票 → 报销材料 全流程')
    ap.add_argument('--work', required=True, help='工作目录')
    ap.add_argument('--src', nargs='*', help='发票压缩包或目录（已有 collect.json 时可省略）')
    ap.add_argument('--fee', required=True, help='费用名称')
    ap.add_argument('--summary', required=True, help='摘要')
    ap.add_argument('--amount', default=None, help='金额，默认取发票合计')
    ap.add_argument('--bills', default=None, help='附单据张数，默认取发票张数')
    ap.add_argument('--date', default=None, help='日期，默认今天')
    ap.add_argument('--template', default=DEFAULT_TPL)
    ap.add_argument('--per-page', type=int, default=2)
    ap.add_argument('--dpi', type=int, default=300)
    ap.add_argument('--out', required=True, help='最终 docx 路径')
    args = ap.parse_args()

    work = os.path.abspath(args.work)
    os.makedirs(work, exist_ok=True)

    if args.src:
        run('01_collect.py', '--src', *args.src, '--work', work)
    elif not os.path.exists(os.path.join(work, 'collect.json')):
        sys.exit('[错误] 没有 collect.json，也没有给 --src')

    if not os.path.exists(os.path.join(work, 'amounts.json')):
        run('02_amounts.py', '--work', work)

    run('03_render.py', '--work', work, '--dpi', str(args.dpi))

    manifest = os.path.join(work, 'amounts.json')
    invoice_docx = os.path.join(work, '发票排版.docx')
    run('04_build_invoice_docx.py', '--manifest', manifest,
        '--out', invoice_docx, '--per-page', str(args.per_page))

    import json
    with open(manifest, encoding='utf-8') as f:
        data = json.load(f)
    total = args.amount or data['total']
    bills = args.bills or str(data['count'])

    form_docx = os.path.join(work, '报销单.docx')
    form_args = ['--template', args.template, '--fee', args.fee, '--summary', args.summary,
                 '--amount', str(total), '--bills', str(bills), '--out', form_docx]
    if args.date:
        form_args += ['--date', args.date]
    run('05_fill_form.py', *form_args)

    run('06_merge.py', '--form', form_docx, '--invoice', invoice_docx, '--out', args.out)
    run('07_verify.py', '--docx', args.out, '--per-page', str(args.per_page),
        '--preview-dir', os.path.join(work, '预览'))

    expect = 1 + math.ceil(int(bills) / args.per_page)
    print(f'\n完成：{args.out}')
    print(f'发票 {bills} 张，合计 ¥{total}，预计 {expect} 页（含报销单 1 页）')


if __name__ == '__main__':
    main()
