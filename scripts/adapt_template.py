# -*- coding: utf-8 -*-
"""自动识别任意模板 deck 的“角色”，输出 JSON / Markdown 映射表。

解决的是这个问题：`inspect_template.py` 只把形状全列出来，换了一套别人的模板，
形状名可能是 `Title 1` / `TextBox 3` / `Rectangle 12`，靠名字写 build 脚本很痛苦。
本脚本按**位置 + 字号 + 占位符类型 + 面积**猜出每页的：标题 / 正文 / 图注 / 图片位 / 页码，
生成 `tpl_map.json`，build 脚本按角色取名字即可。

用法：
    python adapt_template.py 别人模板.pptx                 # 打印概览
    python adapt_template.py 别人模板.pptx -o tpl_map.json  # 写 JSON（推荐）
    python adapt_template.py 别人模板.pptx -m tpl_map.md    # 同时写人看的表

识别规则（按优先级）：
  title    标题占位符(idx 0) > 名字含“标题/Title” > 顶部 40% 区域内字号最大且 >=20pt > 最靠上的文本
  body     其余文本框中面积最大的一个
  caption  下半页、字号 <=16 的次大文本（通常是“图 X …”说明）
  picture  所有图片形状，按位置排序
  page_no  页码占位符 / 纯数字且面积很小的文本
误判时改 JSON 里的 role 即可，不必改脚本。
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU = 914400.0
ROLE_ORDER = ['title', 'subtitle', 'body', 'caption', 'picture', 'page_no', 'header', 'other']


def font_size(shape):
    """形状内最大字号（pt）；取不到时按占位符类型给个估值。"""
    sizes = []
    if not shape.has_text_frame:
        return None
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                sizes.append(r.font.size.pt)
    if sizes:
        return max(sizes)
    try:
        if shape.is_placeholder:
            pt = shape.placeholder_format.type
            name = str(pt)
            if 'TITLE' in name:
                return 32
            if 'SUBTITLE' in name:
                return 20
    except Exception:
        pass
    return 18


def box(shape):
    return [round((shape.left or 0) / EMU, 2), round((shape.top or 0) / EMU, 2),
            round((shape.width or 0) / EMU, 2), round((shape.height or 0) / EMU, 2)]


def text_of(shape):
    if not shape.has_text_frame:
        return ''
    return shape.text_frame.text.replace('\v', ' ').replace('\n', ' | ').strip()


def is_page_number(shape, txt):
    try:
        if shape.is_placeholder and 'SLIDE_NUMBER' in str(shape.placeholder_format.type):
            return True
    except Exception:
        pass
    return txt.isdigit() and shape.height and shape.height / EMU < 0.6


def classify(slide, prs, idx=0):
    W, H = prs.slide_width / EMU, prs.slide_height / EMU
    items = []
    for sh in slide.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            items.append({'name': sh.name, 'type': 'picture', 'role': None,
                          'box': box(sh), 'font': None, 'text': ''})
            continue
        if not sh.has_text_frame or not text_of(sh):
            continue
        items.append({'name': sh.name, 'type': str(sh.shape_type), 'role': None,
                      'box': box(sh), 'font': font_size(sh), 'text': text_of(sh)[:60]})

    texts = [i for i in items if i['type'] != 'picture']

    for it in texts:
        if is_page_number_by_item(it):
            it['role'] = 'page_no'
    pool = [i for i in texts if i['role'] is None]

    title = None
    for i in pool:
        try:
            sh = next(s for s in slide.shapes if s.name == i['name'])
            if sh.is_placeholder and sh.placeholder_format.idx == 0:
                title = i
                break
        except Exception:
            pass
    if not title:
        for i in pool:
            if '标题' in i['name'] or 'Title' in i['name'] or 'title' in i['name']:
                title = i
                break
    if not title:
        lo = H * (0.15 if idx == 0 else 0.05)
        band = [i for i in pool
                if lo <= i['box'][1] <= H * 0.75 and (i['font'] or 0) >= 20]
        if band:
            if idx == 0:
                title = max(band, key=lambda i: (i['font'] or 0) * i['box'][2])
            else:
                title = max(band, key=lambda i: (i['font'] or 0, i['box'][2]))
        else:
            top_area = [i for i in pool if i['box'][1] < H * 0.4 and (i['font'] or 0) >= 20]
            if top_area:
                title = max(top_area, key=lambda i: (i['font'] or 0, -i['box'][1]))
    if not title and pool:
        title = min(pool, key=lambda i: i['box'][1])
    if title:
        title['role'] = 'title'
        pool.remove(title)

    if idx == 0:                      # 封面：除标题外的大字是副标题/汇报人/日期
        for i in pool:
            i['role'] = 'subtitle' if (i['font'] or 0) >= 16 else (
                'header' if i['box'][1] < H * 0.12 else 'other')
        pool = []
    elif pool:
        body = max(pool, key=lambda i: i['box'][2] * i['box'][3])
        body['role'] = 'body'
        pool.remove(body)

    capt = [i for i in pool if i['box'][1] >= H * 0.5 and (i['font'] or 99) <= 16]
    if capt:
        cap = max(capt, key=lambda i: i['box'][2] * i['box'][3])
        cap['role'] = 'caption'
        pool.remove(cap)

    for i in pool:
        i['role'] = 'header' if i['box'][1] < H * 0.12 else 'other'

    for i in items:
        if i['type'] == 'picture':
            i['role'] = 'picture'

    items.sort(key=lambda i: ROLE_ORDER.index(i['role']))
    return items


def is_page_number_by_item(it):
    t = it['text'].strip()
    return t.isdigit() and it['box'][3] < 0.6


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('-o', '--out', default=None, help='输出 JSON 映射表')
    ap.add_argument('-m', '--md', default=None, help='另存一份 Markdown 概览')
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    data = {
        'deck': args.pptx,
        'size': [round(prs.slide_width / EMU, 2), round(prs.slide_height / EMU, 2)],
        'slides': [],
    }
    rows = []
    for idx, s in enumerate(prs.slides):
        items = classify(s, prs, idx)
        data['slides'].append({
            'index': idx,
            'layout': s.slide_layout.name,
            'shapes': items,
        })
        title = next((i['text'][:24] for i in items if i['role'] == 'title'), '')
        roles = ', '.join('%s=%s' % (i['role'], i['name']) for i in items
                          if i['role'] in ('title', 'body', 'caption'))
        rows.append('| %d | %s | %s | %s |' % (idx, s.slide_layout.name, title, roles))

    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print('written', args.out)

    md = ['# 模板角色映射', '', '页面尺寸 %s x %s in' % tuple(data['size']), '',
          '| # | layout | title | roles |', '|---|---|---|---|'] + rows
    md_text = '\n'.join(md)
    if args.md:
        with open(args.md, 'w', encoding='utf-8') as f:
            f.write(md_text)
        print('written', args.md)
    if not args.out and not args.md:
        sys.stdout.buffer.write(md_text.encode('utf-8'))


if __name__ == '__main__':
    main()
