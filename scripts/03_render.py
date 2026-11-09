# -*- coding: utf-8 -*-
"""Step 3 转图：把每张发票 PDF 渲染成 300DPI PNG，并裁掉四周白边。

用法:
    python 03_render.py --work ./work [--dpi 300] [--no-trim]

产物:
    work/图片/NN.png   与 amounts.json 的 seq 对应
    amounts.json 里补写 img / img_w_mm / img_h_mm
"""
import argparse
import os

from _common import ensure_dir, load_json, save_json, table, warn
from deps import ensure


def trim(img, pad=8, threshold=12):
    from PIL import Image, ImageChops
    bg = Image.new('RGB', img.size, (255, 255, 255))
    mask = ImageChops.difference(img, bg).convert('L').point(lambda x: 255 if x > threshold else 0)
    box = mask.getbbox()
    if not box:
        return img
    box = (max(0, box[0] - pad), max(0, box[1] - pad),
           min(img.width, box[2] + pad), min(img.height, box[3] + pad))
    # 裁得太狠说明整页就有大片底色，宁可不裁
    if (box[2] - box[0]) * (box[3] - box[1]) < img.width * img.height * 0.2:
        return img
    return img.crop(box)


def main():
    ensure()
    ap = argparse.ArgumentParser(description='发票 PDF 转高清图并裁白边')
    ap.add_argument('--work', required=True)
    ap.add_argument('--dpi', type=int, default=300)
    ap.add_argument('--no-trim', action='store_true', help='不裁白边')
    args = ap.parse_args()

    try:
        import pymupdf
    except ImportError:
        raise SystemExit('[错误] 缺少 pymupdf，请先 pip install pymupdf pillow')

    from PIL import Image
    import io

    work = os.path.abspath(args.work)
    out_dir = ensure_dir(os.path.join(work, '图片'))
    manifest = load_json(os.path.join(work, 'amounts.json'))

    rows = []
    for it in manifest['items']:
        doc = pymupdf.open(it['file'])
        if len(doc) > 1:
            warn(f'{it["name"]} 有 {len(doc)} 页，只取第 1 页')
        pix = doc[0].get_pixmap(dpi=args.dpi, colorspace=pymupdf.csRGB)
        img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
        doc.close()
        if not args.no_trim:
            img = trim(img)

        path = os.path.join(out_dir, f'{it["seq"]:02d}.png')
        img.save(path, dpi=(args.dpi, args.dpi), optimize=True)

        w_mm = img.width / args.dpi * 25.4
        h_mm = img.height / args.dpi * 25.4
        it['img'] = path
        it['img_w_mm'] = round(w_mm, 1)
        it['img_h_mm'] = round(h_mm, 1)
        it['img_ratio'] = round(w_mm / h_mm, 3)
        rows.append([it['seq'], it['label'][:16], f'{w_mm:.1f}x{h_mm:.1f}',
                     f'{w_mm / h_mm:.2f}', f'{img.width}x{img.height}'])

    save_json(manifest, os.path.join(work, 'amounts.json'))

    table(rows, ['#', '平台', '尺寸mm', '宽高比', '像素'], ['r', 'l', 'r', 'r', 'r'])
    print()
    print(f'已生成 {len(rows)} 张图，目录: {out_dir}')
    widest = max(manifest['items'], key=lambda x: x['img_w_mm'])
    print(f'最宽 {widest["img_w_mm"]}mm、最矮 {min(i["img_h_mm"] for i in manifest["items"])}mm')


if __name__ == '__main__':
    main()
