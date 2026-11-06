# -*- coding: utf-8 -*-
"""盘点模板 deck：列出每一页的版式、形状名、位置尺寸和文字，用于“套模板”时定位。

用法：
    python inspect_template.py 模板.pptx            # 概览
    python inspect_template.py 模板.pptx -o tpl.txt # 输出到文件（中文路径/编码更稳）

输出里重点看：
  * 页序（0 基）与版式名 -> 决定复用哪一页当“文字页/图片页/封面”
  * 形状名（如 文本框 2 / 文本框 4 / 图片 7）-> deckkit 里用名字定位
  * L/T/W/H（英寸）-> 照抄坐标，保证风格一致
"""
import argparse
import sys

from pptx import Presentation


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('-o', '--out', default=None)
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    out = []
    out.append('slide size %.2f x %.2f in' % (prs.slide_width / 914400, prs.slide_height / 914400))
    out.append('== layouts ==')
    for i, l in enumerate(prs.slide_layouts):
        out.append('%d %s' % (i, l.name))
    out.append('== slides ==')
    for i, s in enumerate(prs.slides):
        out.append('--- slide %d  layout=%s' % (i, s.slide_layout.name))
        for sh in s.shapes:
            txt = ''
            if sh.has_text_frame:
                txt = sh.text_frame.text.replace('\n', ' | ')[:100]
            out.append('  [%s] %s  L=%.2f T=%.2f W=%.2f H=%.2f  %s' % (
                sh.shape_type, sh.name,
                (sh.left or 0) / 914400, (sh.top or 0) / 914400,
                (sh.width or 0) / 914400, (sh.height or 0) / 914400, txt))
    text = '\n'.join(out)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(text)
        print('written', args.out)
    else:
        sys.stdout.buffer.write(text.encode('utf-8'))


if __name__ == '__main__':
    main()
