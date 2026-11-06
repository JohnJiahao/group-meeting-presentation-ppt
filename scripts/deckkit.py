# -*- coding: utf-8 -*-
"""deckkit —— 以用户已有 pptx 为模板，按页填充文本/图片的最小工具库。

典型用法（在你的 build_deck.py 里）：

    import sys, os
    sys.path.insert(0, r"<skill>/scripts")
    from deckkit import Deck

    d = Deck(template_pptx, out_pptx)
    S = d.slides                      # 模板原有页
    new = d.dup(S[2])                 # 复制一页（保留所有形状与样式）
    d.drop(new, "文本框 2")            # 按形状名删除
    d.delete(S[9]); d.reorder([...])  # 删除 / 重排
    tf = d.setup(S[0], "文本框 4", 0.68, 3.9, 12.2, 0.6)
    d.para(tf, [("中文标题", {"bold": True})], first=True, size=24, color=d.BLUE)
    d.pic(S[0], "fig2.png", 7.4, 2.8, w=5.4)
    d.save()

设计原则：模板的形状名与坐标一律由调用方给出（先用 scripts/inspect_template.py 查看），
本库只负责“复制页 / 找形状 / 写文字 / 插图 / 排序”这些机械动作。
"""
import copy
import os

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

DEFAULT_BLUE = RGBColor(0x1F, 0x5F, 0xA8)
DEFAULT_DARK = RGBColor(0x1C, 0x1F, 0x23)
DEFAULT_RED = RGBColor(0xFF, 0x00, 0x00)
DEFAULT_GRAY = RGBColor(0x59, 0x59, 0x59)


def _highlight(run, color='FFFF00'):
    rPr = run._r.get_or_add_rPr()
    hl = rPr.makeelement(A + 'highlight', {})
    clr = hl.makeelement(A + 'srgbClr', {'val': color})
    hl.append(clr)
    latin = rPr.find(A + 'latin')
    if latin is not None:
        latin.addprevious(hl)
    else:
        rPr.append(hl)


class Deck(object):
    def __init__(self, template, out=None):
        self.prs = Presentation(template)
        self.out = out or template
        self.slides = list(self.prs.slides)
        self.BLUE = DEFAULT_BLUE
        self.DARK = DEFAULT_DARK
        self.RED = DEFAULT_RED
        self.GRAY = DEFAULT_GRAY

    # ---------------------------------------------------------------- 页面操作
    def dup(self, src):
        """复制一页（含全部形状与格式），追加到末尾并返回新页。"""
        new = self.prs.slides.add_slide(src.slide_layout)
        for shp in list(new.shapes):
            shp._element.getparent().remove(shp._element)
        for shp in src.shapes:
            new.shapes._spTree.append(copy.deepcopy(shp._element))
        return new

    def delete(self, slide):
        sid = slide.slide_id
        lst = self.prs.slides._sldIdLst
        for sldId in list(lst):
            if sldId.get('id') == str(sid):
                self.prs.part.drop_rel(sldId.get(R + 'id'))
                lst.remove(sldId)

    def reorder(self, order):
        """按 order 里的 Slide 对象顺序重排。注意：id 必须在清空 sldIdLst 之前取。"""
        lst = self.prs.slides._sldIdLst
        ids = [s.slide_id for s in order]
        elems = {int(e.get('id')): e for e in lst}
        for e in list(lst):
            lst.remove(e)
        for i in ids:
            lst.append(elems[i])

    def set_order_by_index(self, order_idx):
        """用“当前页序索引”重排，方便先算好顺序。"""
        cur = list(self.prs.slides)
        self.reorder([cur[i] for i in order_idx])

    # ---------------------------------------------------------------- 形状操作
    @staticmethod
    def shape(slide, name):
        for s in slide.shapes:
            if s.name == name:
                return s
        raise KeyError('%s 中找不到形状：%s' % (slide.name if hasattr(slide, 'name') else '', name))

    def drop(self, slide, name):
        s = self.shape(slide, name)
        s._element.getparent().remove(s._element)

    def setup(self, slide, name, l, t, w, h, anchor=MSO_ANCHOR.TOP):
        """定位并清空某个已有形状，返回可写的 text_frame。"""
        s = self.shape(slide, name)
        s.left, s.top, s.width, s.height = Inches(l), Inches(t), Inches(w), Inches(h)
        tf = s.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = 0
        tf.clear()
        return tf

    def textbox(self, slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
        s = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
        tf = s.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = 0
        return tf

    def pic(self, slide, path, l, t, w=None, h=None):
        return slide.shapes.add_picture(path, Inches(l), Inches(t),
                                        Inches(w) if w else None, Inches(h) if h else None)

    # ---------------------------------------------------------------- 文字
    def para(self, tf, runs, first=False, size=16, color=DEFAULT_DARK, font='微软雅黑',
             bold=False, align=PP_ALIGN.LEFT, sb=0, sa=6, line=1.2, indent=None):
        """runs: 字符串 或 [(文本, {bold/color/size/hl/font}), ...]

        中文段落一律左对齐（不要用 JUSTIFY，否则会出现标点跑到行首、
        中英文之间被拉开大空隙），并写入 lang="zh-CN" 以启用中文换行规则。
        """
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(sb)
        p.space_after = Pt(sa)
        p.line_spacing = line
        if indent is not None:                      # 首行缩进（英寸）
            pPr = p._pPr if p._pPr is not None else p.get_or_add_pPr()
            pPr.set('marL', '0')
            pPr.set('indent', str(int(indent * 914400)))
        if isinstance(runs, str):
            runs = [(runs, {})]
        for text, o in runs:
            r = p.add_run()
            r.text = text
            f = r.font
            f.name = o.get('font', font)
            f.size = Pt(o.get('size', size))
            f.bold = o.get('bold', bold)
            f.color.rgb = o.get('color', color)
            rPr = r._r.get_or_add_rPr()
            rPr.set('lang', 'zh-CN')
            rPr.set('altLang', 'en-US')
            if o.get('hl'):
                _highlight(r)
        return p

    # ---------------------------------------------------------------- 保存
    def save(self, path=None):
        p = path or self.out
        self.prs.save(p)
        print('saved', p, 'slides=', len(self.prs.slides._sldIdLst))
        return p
