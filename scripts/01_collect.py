# -*- coding: utf-8 -*-
"""Step 1 汇集：从压缩包/目录收集电子发票，按「发票」与「行程单」分类存放。

用法:
    python 01_collect.py --src a.zip b.zip --work ./work
    python 01_collect.py --src ./下载目录 --work ./work

产物:
    work/发票PDF/*.pdf        真正的发票
    work/行程单PDF/*.pdf      行程明细（不是发票，单独放）
    work/结构化数据/*.xml     数电票 XML（金额的权威来源）
    work/其他/*              无法判定的 PDF
    work/collect.json         清单
"""
import argparse
import os
import re
import shutil
import zipfile

from _common import ensure_dir, save_json, table, warn

FIX_RE = re.compile(r'[\\/:*?"<>|]')


def fix_name(name):
    """ZIP 内中文文件名常被当作 cp437 存储，还原成 GBK。"""
    if any('\u4e00' <= c <= '\u9fff' for c in name):
        return name
    try:
        return name.encode('cp437').decode('gbk')
    except (UnicodeEncodeError, UnicodeDecodeError):
        try:
            return name.encode('cp437').decode('utf-8')
        except (UnicodeEncodeError, UnicodeDecodeError):
            return name


def safe_name(name):
    return FIX_RE.sub('_', os.path.basename(name))


def classify(name):
    if '行程单' in name:
        return 'trip'
    if '发票' in name:
        return 'invoice'
    return 'unknown'


class Collector:
    def __init__(self, work):
        self.work = work
        self.dirs = {
            'invoice': ensure_dir(os.path.join(work, '发票PDF')),
            'trip': ensure_dir(os.path.join(work, '行程单PDF')),
            'unknown': ensure_dir(os.path.join(work, '其他')),
            'xml': ensure_dir(os.path.join(work, '结构化数据')),
        }
        self.seen = {}          # (basename, size) -> 目标路径，用于跨包去重
        self.items = {'invoice': [], 'trip': [], 'unknown': []}
        self.xmls = {}

    def _target(self, kind, name, size):
        key = (name, size)
        if key in self.seen:
            return None, True
        dst = os.path.join(self.dirs[kind], name)
        stem, ext = os.path.splitext(name)
        k = 2
        while os.path.exists(dst):
            dst = os.path.join(self.dirs[kind], f'{stem}_{k}{ext}')
            k += 1
        self.seen[key] = dst
        return dst, False

    def add_stream(self, kind, name, size, opener):
        dst, dup = self._target(kind, safe_name(name), size)
        if dup:
            return None
        with opener() as src, open(dst, 'wb') as out:
            shutil.copyfileobj(src, out)
        self.items[kind].append(dst)
        return dst

    def add_file(self, kind, path):
        name = os.path.basename(path)
        size = os.path.getsize(path)
        dst, dup = self._target(kind, name, size)
        if dup:
            return None
        shutil.copy2(path, dst)
        self.items[kind].append(dst)
        return dst

    def add_xml(self, name, size, opener):
        dst = os.path.join(self.dirs['xml'], safe_name(name))
        stem, ext = os.path.splitext(dst)
        k = 2
        while os.path.exists(dst):
            dst = f'{stem}_{k}{ext}'
            k += 1
        with opener() as src, open(dst, 'wb') as out:
            shutil.copyfileobj(src, out)
        self.xmls[os.path.splitext(safe_name(name))[0]] = dst
        return dst

    def scan_zip(self, path):
        z = zipfile.ZipFile(path)
        for info in z.infolist():
            if info.is_dir():
                continue
            name = fix_name(info.filename)
            low = name.lower()
            opener = (lambda i=info: z.open(i))
            if low.endswith('.pdf'):
                self.add_stream(classify(name), name, info.file_size, opener)
            elif low.endswith('.xml'):
                self.add_xml(name, info.file_size, opener)
        z.close()

    def scan_dir(self, root):
        # 目录里没有原始压缩包，XML 可能以同名文件躺在 pdf 旁边
        for dirpath, _, files in os.walk(root):
            for fn in files:
                p = os.path.join(dirpath, fn)
                low = fn.lower()
                if low.endswith('.pdf'):
                    self.add_file(classify(fn), p)
                elif low.endswith('.xml'):
                    self.add_xml(fn, os.path.getsize(p), lambda p=p: open(p, 'rb'))


def main():
    ap = argparse.ArgumentParser(description='汇集电子发票，分类存放')
    ap.add_argument('--src', nargs='+', required=True, help='压缩包或目录，可多个')
    ap.add_argument('--work', required=True, help='工作目录')
    args = ap.parse_args()

    work = ensure_dir(os.path.abspath(args.work))
    c = Collector(work)

    for src in args.src:
        if not os.path.exists(src):
            warn(f'跳过不存在的路径: {src}')
            continue
        if os.path.isdir(src):
            c.scan_dir(src)
        elif zipfile.is_zipfile(src):
            c.scan_zip(src)
        else:
            warn(f'既不是压缩包也不是目录: {src}')

    # 文件名不带「发票/行程单」的，若有同名 XML 则按发票处理
    for p in list(c.items['unknown']):
        stem = os.path.splitext(os.path.basename(p))[0]
        if stem in c.xmls:
            dst = os.path.join(c.dirs['invoice'], os.path.basename(p))
            shutil.move(p, dst)
            c.items['unknown'].remove(p)
            c.items['invoice'].append(dst)

    invoices = []
    for p in c.items['invoice']:
        stem = os.path.splitext(os.path.basename(p))[0]
        invoices.append({
            'file': p,
            'name': os.path.basename(p),
            'xml': c.xmls.get(stem),
        })
    trips = [{'file': p, 'name': os.path.basename(p)} for p in c.items['trip']]

    manifest = {
        'work': work,
        'sources': [os.path.abspath(s) for s in args.src],
        'invoices': invoices,
        'trips': trips,
        'unknown': c.items['unknown'],
    }
    save_json(manifest, os.path.join(work, 'collect.json'))

    table(
        [[i + 1, os.path.basename(x['file'])[:44], '有' if x['xml'] else '—']
         for i, x in enumerate(invoices)],
        ['#', '发票文件', 'XML'],
        ['r', 'l', 'c'],
    )
    print()
    print(f'发票 {len(invoices)} 张 | 行程单 {len(trips)} 张 | 其他 {len(c.items["unknown"])} 个')
    print(f'清单: {os.path.join(work, "collect.json")}')
    if not invoices:
        raise SystemExit('[错误] 没有找到任何发票 PDF')


if __name__ == '__main__':
    main()
