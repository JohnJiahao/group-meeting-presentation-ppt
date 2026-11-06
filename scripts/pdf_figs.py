# -*- coding: utf-8 -*-
"""从论文 PDF 里按“图注/表注”自动裁切图片与表格，输出 400dpi PNG。

用法：
    python pdf_figs.py paper.pdf -o figs fig1="Fig. 1" fig2="Figure 2" \
        tab2="Table 2" tab5="Table 5"

规则（Elsevier/Springer 常见排版）：
  * 图注在图下方  -> 图片区域取“图注上方”最近的一块图像/绘图区域；
  * 表注在表上方  -> 表格区域取“表注下方”直到出现大段空白（>14pt）。
  * 图注/表注本身不裁进结果里（幻灯片里用中文图注说明即可）。

注意：PDF 里图上的文字常是矢量文本，直接抽取内嵌位图会得到“空框”，
所以本脚本统一按页面区域渲染（get_pixmap），不要用 extract_image。
裁完务必肉眼检查一遍（尤其是跨栏大表、双图并排的 Figure）。
"""
import argparse
import os
import re

import fitz

CAP_RE = re.compile(r'^(Fig(?:ure)?\.?|Table)\s*([0-9]+[A-Za-z]?)\b', re.I)


def find_caption(page, label):
    kind = 'table' if label.strip().lower().startswith('tab') else 'fig'
    m = re.search(r'([0-9]+[A-Za-z]?)', label)
    if not m:
        return None
    num = m.group(1)
    for block in page.get_text('dict')['blocks']:
        if block['type'] != 0:
            continue
        btxt = ''.join(''.join(sp['text'] for sp in ln['spans']) for ln in block['lines'])
        if len(btxt) > 160:            # 正文引用句整段更长，直接跳过
            continue
        for i, line in enumerate(block['lines']):
            if i > 0:
                break
            txt = ''.join(sp['text'] for sp in line['spans']).strip()
            mm = CAP_RE.match(txt)
            if not mm:
                continue
            k = 'table' if mm.group(1).lower().startswith('tab') else 'fig'
            if k != kind or mm.group(2) != num:
                continue
            # 排掉正文里的引用句 “Table 3 shows the benchmarking results …”
            if len(txt) > 90:
                continue
            return fitz.Rect(block['bbox'])   # 图注常分两行，取整块
    return None


def caption_block(page, cap):
    return fitz.Rect(cap)


def figure_rect(page, cap):
    capb = caption_block(page, cap)
    best = None
    for info in page.get_image_info():
        b = fitz.Rect(info['bbox'])
        if b.y1 > capb.y0 + 3:                 # 必须在图注上方
            continue
        ov = min(b.x1, capb.x1) - max(b.x0, capb.x0)
        if ov < 0.3 * min(b.width, capb.width):
            continue
        if best is None or b.y1 > best.y1:
            best = b
    if best is not None:
        return fitz.Rect(best.x0 - 2, best.y0 - 2, best.x1 + 2, best.y1 + 2)
    # 矢量图：取图注上方 420pt 内、同栏的文字/绘图块并集
    rects = []
    for block in page.get_text('dict')['blocks']:
        b = fitz.Rect(block['bbox'])
        if b.y1 <= capb.y0 + 2 and b.y0 > capb.y0 - 420 and b.x0 < capb.x1 and b.x1 > capb.x0:
            rects.append(b)
    if not rects:
        return None
    u = rects[0]
    for r in rects[1:]:
        u |= r
    return fitz.Rect(u.x0 - 4, u.y0 - 4, u.x1 + 4, min(u.y1, capb.y0) + 2)


def table_rect(page, cap):
    """表注在表上方。优先用表格的横线（drawings）定边界——跨栏表也能取全；
    没有画线时退回“按行距扫文字”的兜底逻辑。"""
    capb = caption_block(page, cap)

    rules = []
    for dr in page.get_drawings():
        r = fitz.Rect(dr['rect'])
        if r.width > 80 and r.height < 4 and capb.y1 - 3 <= r.y0 <= capb.y1 + 320:
            if abs(r.x0 - capb.x0) < 60:          # 与表注左端对齐的同一条表
                rules.append(r)
    if len(rules) >= 2:
        u = fitz.Rect(rules[0])
        for r in rules[1:]:
            u |= r
        return fitz.Rect(u.x0 - 2, max(u.y0 - 2, capb.y1 + 1), u.x1 + 2, u.y1 + 3)

    rows = []
    for block in page.get_text('dict')['blocks']:
        if block['type'] != 0:
            continue
        for line in block['lines']:
            b = fitz.Rect(line['bbox'])
            if b.y0 < capb.y1 - 1 or b.y0 > capb.y1 + 320:
                continue
            rows.append(b)
    if not rows:
        return None
    rows.sort(key=lambda r: (r.y0, r.x0))
    band = [rows[0]]
    for r in rows[1:]:
        if r.y0 - band[-1].y1 > 14:        # 出现段落空隙 -> 表格结束
            break
        band.append(r)
    y0, y1 = min(r.y0 for r in band), max(r.y1 for r in band)
    inband = [r for r in rows if r.y0 >= y0 - 1 and r.y1 <= y1 + 1]
    x0 = min(r.x0 for r in inband) - 6
    x1 = max(r.x1 for r in inband) + 6
    return fitz.Rect(x0, capb.y1 + 1, x1, y1 + 4)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('specs', nargs='+', help='name="Fig. 2" 这样的键值对')
    ap.add_argument('-o', '--out', default='figs')
    ap.add_argument('--dpi', type=int, default=400)
    ap.add_argument('--pad', type=float, default=2.0)
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    doc = fitz.open(args.pdf)
    for spec in args.specs:
        name, _, label = spec.partition('=')
        label = label.strip('"\' ')
        hit = None
        for page in doc:
            cap = find_caption(page, label)
            if cap is None:
                continue
            rect = table_rect(page, cap) if label.lower().startswith('tab') else figure_rect(page, cap)
            if rect is None:
                continue
            hit = (page.number + 1, rect)
            break
        if hit is None:
            print('[MISS] %s (%s)' % (name, label))
            continue
        pno, r = hit
        r = fitz.Rect(r.x0 - args.pad, r.y0 - args.pad, r.x1 + args.pad, r.y1 + args.pad)
        r &= doc[pno - 1].rect
        pix = doc[pno - 1].get_pixmap(clip=r, dpi=args.dpi)
        fn = os.path.join(args.out, name + '.png')
        pix.save(fn)
        print('[OK] %-10s p%-3d %s  %dx%d' % (name, pno, tuple(round(v, 1) for v in r), pix.width, pix.height))


if __name__ == '__main__':
    main()
