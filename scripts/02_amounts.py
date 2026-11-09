# -*- coding: utf-8 -*-
"""Step 2 核价：从数电票 XML 提取含税金额（权威来源），并与文件名标注逐张核对。

用法:
    python 02_amounts.py --work ./work

产物:
    work/amounts.json   按金额降序的明细 + 合计
    work/发票金额汇总.csv
"""
import argparse
import csv
import os
import re
import xml.etree.ElementTree as ET
from decimal import Decimal

from _common import ensure_dir, load_json, money, save_json, table, warn

# 数电票 XML 字段候选名（不同开票渠道命名略有差异）
FIELDS = {
    'amount': ['TotalTax-includedAmount', 'TotalTaxIncludedAmount', 'TotalAmountWithTax'],
    'net': ['TotalAmWithoutTax', 'TotalAmountWithoutTax', 'TotalAmWithoutTaxAmount'],
    'tax': ['TotalTaxAm', 'TotalTaxAmount', 'TotalTax'],
    'seller': ['SellerName', 'SellerName', 'SalesName'],
    'no': ['EIid', 'InvoiceNum', 'InvoiceNumber'],
    'date': ['RequestTime', 'IssueTime', 'InvoiceDate'],
}

NAME_RE = re.compile(r'【(?P<label>[^-】]+)-(?P<amt>\d+(?:\.\d+)?)元(?:-(?P<trips>\d+)个行程)?】')


def parse_xml(path):
    """返回 {字段: 值}，标签忽略命名空间只比对本地名。"""
    out = {}
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError:
        return out
    flat = {}
    for el in root.iter():
        tag = el.tag.split('}')[-1]
        if el.text and el.text.strip() and tag not in flat:
            flat[tag] = el.text.strip()
    for key, names in FIELDS.items():
        for n in names:
            if n in flat:
                out[key] = flat[n]
                break
    return out


def parse_pdf(path):
    """兜底：从 PDF 文本里找价税合计。多数电子发票内嵌字体，可能取不到。"""
    try:
        from pypdf import PdfReader
    except ImportError:
        return {}
    try:
        reader = PdfReader(path)
        text = '\n'.join((p.extract_text() or '') for p in reader.pages)
    except Exception:
        return {}
    flat = re.sub(r'\s+', ' ', text)
    got = {}
    m = re.search(r'(?:价税合计|小写)[^0-9]{0,12}(\d+\.\d{2})', flat)
    if m:
        got['amount'] = m.group(1)
    m = re.search(r'发票号码[:：]?\s*(\d{8,25})', flat)
    if m:
        got['no'] = m.group(1)
    return got


def main():
    ap = argparse.ArgumentParser(description='提取并核对发票金额')
    ap.add_argument('--work', required=True)
    args = ap.parse_args()

    work = os.path.abspath(args.work)
    collect = load_json(os.path.join(work, 'collect.json'))

    items, problems, seen_no = [], [], {}
    for inv in collect['invoices']:
        name = inv['name']
        meta = {}
        if inv.get('xml') and os.path.exists(inv['xml']):
            meta = parse_xml(inv['xml'])
        source = 'XML' if meta.get('amount') else None

        if not meta.get('amount'):
            meta.update({k: v for k, v in parse_pdf(inv['file']).items() if v})
            source = 'PDF文本' if meta.get('amount') else None

        m = NAME_RE.search(name)
        label = m.group('label') if m else os.path.splitext(name)[0][:16]
        hint = money(m.group('amt')) if m else None

        if not meta.get('amount'):
            problems.append(f'无法取得金额: {name}')
            continue
        amount = money(meta['amount'])

        if hint is not None and abs(hint - amount) > Decimal('0.01'):
            problems.append(f'金额不一致: {name} 文件名 {hint} / 票据 {amount}')

        no = meta.get('no', '')
        if no:
            if no in seen_no:
                warn(f'重复发票号 {no}，跳过: {name}')
                continue
            seen_no[no] = name

        items.append({
            'label': label,
            'name': name,
            'file': inv['file'],
            'amount': str(amount),
            'net': str(money(meta['net'])) if meta.get('net') else None,
            'tax': str(money(meta['tax'])) if meta.get('tax') else None,
            'seller': meta.get('seller') or '',
            'no': no,
            'date': (meta.get('date') or '')[:19],
            'source': source,
            'hint': str(hint) if hint is not None else None,
        })

    if not items:
        raise SystemExit('[错误] 没有任何可用金额，检查发票文件')

    items.sort(key=lambda x: Decimal(x['amount']), reverse=True)
    total = sum((Decimal(x['amount']) for x in items), Decimal('0'))
    net_total = sum((Decimal(x['net']) for x in items if x['net']), Decimal('0'))
    tax_total = sum((Decimal(x['tax']) for x in items if x['tax']), Decimal('0'))

    for i, it in enumerate(items, 1):
        it['seq'] = i

    manifest = {
        'work': work,
        'count': len(items),
        'total': str(money(total)),
        'net_total': str(money(net_total)) if net_total else None,
        'tax_total': str(money(tax_total)) if tax_total else None,
        'items': items,
        'problems': problems,
    }
    save_json(manifest, os.path.join(work, 'amounts.json'))

    csv_path = os.path.join(work, '发票金额汇总.csv')
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['序号', '平台', '销售方', '发票号码', '价税合计(元)',
                    '不含税(元)', '税额(元)', '开票时间', '取数来源'])
        for it in items:
            w.writerow([it['seq'], it['label'], it['seller'] or '-', it['no'],
                        it['amount'], it['net'] or '-', it['tax'] or '-',
                        it['date'], it['source']])
        w.writerow(['合计', '', '', '', str(money(total)),
                    str(money(net_total)) if net_total else '-',
                    str(money(tax_total)) if tax_total else '-', '', ''])

    table(
        [[it['seq'], it['label'][:16], it['amount'], it['source'], it['no'][:20]]
         for it in items],
        ['#', '平台', '价税合计', '取数', '发票号'],
        ['r', 'l', 'r', 'c', 'l'],
    )
    print()
    print(f'共 {len(items)} 张，价税合计 {money(total)} 元')
    if net_total:
        print(f'不含税 {money(net_total)} 元，税额 {money(tax_total)} 元')
    print(f'明细: {os.path.join(work, "amounts.json")}')
    print(f'汇总: {csv_path}')

    if problems:
        print()
        print(f'[需人工确认] {len(problems)} 项')
        for p in problems:
            print('  -', p)


if __name__ == '__main__':
    main()
