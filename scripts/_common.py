# -*- coding: utf-8 -*-
"""公共工具：金额大写、单位换算、清单读写。"""
import json
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

MM_PER_INCH = 25.4

_DIGITS = '零壹贰叁肆伍陆柒捌玖'
_UNITS = ['', '拾', '佰', '仟']
_BIG = ['', '万', '亿', '万亿']


def money(v):
    """转成两位小数的 Decimal。"""
    return Decimal(str(v)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


def rmb_upper(amount):
    """人民币金额转大写。123.45 -> 壹佰贰拾叁元肆角伍分"""
    cents = int(money(amount) * 100)
    if cents == 0:
        return '零元整'

    def sec4(n):
        s, zero = '', False
        for i in range(3, -1, -1):
            d = n // (10 ** i) % 10
            if d == 0:
                if s:
                    zero = True
            else:
                if zero:
                    s += '零'
                    zero = False
                s += _DIGITS[d] + _UNITS[i]
        return s

    yuan, rest = divmod(cents, 100)
    jiao, fen = divmod(rest, 10)

    if yuan == 0:
        head = '零'
    else:
        groups, n = [], yuan
        while n > 0:
            groups.append(n % 10000)
            n //= 10000
        parts = []
        for idx in range(len(groups) - 1, -1, -1):
            g = groups[idx]
            if g == 0:
                continue
            if idx < len(groups) - 1 and g < 1000 and parts:
                parts.append('零')
            parts.append(sec4(g) + _BIG[idx])
        head = ''.join(parts)

    if jiao == 0 and fen == 0:
        tail = '整'
    else:
        tail = ''
        if jiao:
            tail += _DIGITS[jiao] + '角'
        elif yuan:
            tail += '零'
        if fen:
            tail += _DIGITS[fen] + '分'

    return f'{head}元{tail}'


def load_json(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def save_json(obj, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def die(msg, code=1):
    print(f'[错误] {msg}', file=sys.stderr)
    sys.exit(code)


def warn(msg):
    print(f'[注意] {msg}')


def table(rows, headers, aligns=None):
    """终端等宽表格输出；中文按 2 字符宽度对齐。"""
    def width(s):
        return sum(2 if ord(c) > 127 else 1 for c in str(s))

    def pad(s, w, al='l'):
        s = str(s)
        gap = w - width(s)
        if al == 'r':
            return ' ' * gap + s
        if al == 'c':
            left = gap // 2
            return ' ' * left + s + ' ' * (gap - left)
        return s + ' ' * gap

    aligns = aligns or ['l'] * len(headers)
    data = [headers] + [[str(c) for c in r] for r in rows]
    widths = [max(width(r[i]) for r in data) for i in range(len(headers))]
    lines = []
    for k, r in enumerate(data):
        lines.append('  '.join(pad(r[i], widths[i], aligns[i]) for i in range(len(headers))))
        if k == 0:
            lines.append('  '.join('-' * w for w in widths))
    print('\n'.join(lines))
