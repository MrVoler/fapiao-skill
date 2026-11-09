# -*- coding: utf-8 -*-
"""一条龙：发票 → 报销材料.docx

对普通人只需要两个信息：费用名称、摘要。其余全部自动。

用法:
    python pipeline.py --src 发票.zip --fee "交通费" --summary "10 月打车"

    # 也可以给多个包、给一个装满 PDF 的文件夹
    python pipeline.py --src 发票1.zip 发票2.zip --fee "交通费" --summary "10 月打车"
    python pipeline.py --src ./我的发票 --fee "差旅费" --summary "上海出差"

默认产出:
    <发票所在目录>/报销单-<费用名称>-<日期>.docx
    中间文件在 <发票所在目录>/报销中间文件/
"""
import argparse
import datetime
import json
import math
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TPL = os.path.abspath(os.path.join(HERE, '..', 'assets', '费用报销单模板.docx'))

DEPS = [
    ('docx', 'python-docx'),
    ('pymupdf', 'pymupdf'),
    ('PIL', 'pillow'),
]


def check_deps():
    missing = []
    for mod, pkg in DEPS:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if missing:
        print('[缺少依赖] 请先安装：', file=sys.stderr)
        print(f'  {sys.executable} -m pip install ' + ' '.join(missing), file=sys.stderr)
        sys.exit(2)


def run(script, *args):
    cmd = [sys.executable, os.path.join(HERE, script)] + list(args)
    print(f'\n=== {script} ===', flush=True)
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f'\n[中断] {script} 没跑通，上面是具体原因', file=sys.stderr)
        sys.exit(r.returncode)


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|\s]+', '', s)[:20] or '报销'


def main():
    ap = argparse.ArgumentParser(description='发票 → 报销材料，一条命令跑完')
    ap.add_argument('--src', nargs='*', help='发票压缩包或文件夹')
    ap.add_argument('--fee', required=True, help='费用名称（问用户）')
    ap.add_argument('--summary', required=True, help='摘要（问用户）')
    ap.add_argument('--work', help='中间文件目录，默认放在发票旁边')
    ap.add_argument('--out', help='最终 Word 路径，默认放在发票旁边')
    ap.add_argument('--template', default=DEFAULT_TPL, help='报销单模板')
    ap.add_argument('--date', help='默认今天')
    ap.add_argument('--amount', help='默认取发票合计')
    ap.add_argument('--bills', help='默认取发票张数')
    ap.add_argument('--per-page', type=int, default=2)
    ap.add_argument('--dpi', type=int, default=300)
    args = ap.parse_args()

    check_deps()

    if not os.path.exists(args.template):
        sys.exit(f'[错误] 模板不存在: {args.template}')

    srcs = [os.path.abspath(s) for s in (args.src or [])]
    base = os.path.dirname(srcs[0]) if srcs else os.getcwd()
    if srcs and os.path.isdir(srcs[0]):
        base = srcs[0]

    work = os.path.abspath(args.work or os.path.join(base, '报销中间文件'))
    os.makedirs(work, exist_ok=True)

    stamp = datetime.date.today().strftime('%Y%m%d')
    out = os.path.abspath(args.out or os.path.join(base, f'报销单-{safe_name(args.fee)}-{stamp}.docx'))

    if srcs:
        run('01_collect.py', '--src', *srcs, '--work', work)
    elif not os.path.exists(os.path.join(work, 'collect.json')):
        sys.exit('[错误] 没给 --src，且中间目录里没有已汇集的发票')

    if not os.path.exists(os.path.join(work, 'amounts.json')):
        run('02_amounts.py', '--work', work)

    run('03_render.py', '--work', work, '--dpi', str(args.dpi))

    manifest = os.path.join(work, 'amounts.json')
    with open(manifest, encoding='utf-8') as f:
        data = json.load(f)

    invoice_docx = os.path.join(work, '发票排版.docx')
    run('04_build_invoice_docx.py', '--manifest', manifest,
        '--out', invoice_docx, '--per-page', str(args.per_page))

    total = args.amount or data['total']
    bills = args.bills or str(data['count'])

    form_docx = os.path.join(work, '报销单.docx')
    form_args = ['--template', args.template, '--fee', args.fee, '--summary', args.summary,
                 '--amount', str(total), '--bills', str(bills), '--out', form_docx]
    if args.date:
        form_args += ['--date', args.date]
    run('05_fill_form.py', *form_args)

    run('06_merge.py', '--form', form_docx, '--invoice', invoice_docx, '--out', out)
    run('07_verify.py', '--docx', out, '--per-page', str(args.per_page),
        '--preview-dir', os.path.join(work, '预览'))

    expect = 1 + math.ceil(int(bills) / args.per_page)
    print('\n' + '=' * 52)
    print(f'做好了：{out}')
    print(f'{bills} 张发票，合计 ¥{total}，共 {expect} 页（第 1 页报销单，之后每页 2 张）')
    print(f'中间文件和预览图在：{work}')
    print('=' * 52)


if __name__ == '__main__':
    main()
