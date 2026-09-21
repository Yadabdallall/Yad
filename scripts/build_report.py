# -*- coding: utf-8 -*-
"""
دروستکردنی فایلی Word بۆ خشتەی چالاکییە وەرزشییەکانی
لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی (2026)
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------- ڕەنگەکان
NAVY        = "1F3864"   # ڕەنگی سەردێڕی خشتە
RED_BORDER  = "C00000"   # چوارچێوەی سوور
ROW_ALT     = "EDF2F8"   # ڕەنگی ڕیزە جیاکەرەوەکان
WHITE       = "FFFFFF"

FONT_LATIN = "Calibri"
FONT_CS    = "Arial"     # فۆنتی سکریپتی عەرەبی/کوردی


# ---------------------------------------------------------------- ڕیزبەندی XML
# OOXML داوای ڕیزبەندییەکی دیاریکراو دەکات بۆ منداڵە ئێلێمێنتەکان،
# بۆیە هەر ئێلێمێنتێک لە شوێنی دروستی خۆیدا دادەنرێت.
_ORDER = {
    "w:rPr": ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
              "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
              "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
              "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
              "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
              "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang"],
    "w:pPr": ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore",
              "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers",
              "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
              "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
              "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
              "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
              "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
              "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
              "w:rPr", "w:sectPr", "w:pPrChange"],
    "w:tcPr": ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
               "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
               "w:tcFitText", "w:vAlign", "w:hideMark"],
    "w:trPr": ["w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
               "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader",
               "w:tblCellSpacing", "w:jc", "w:hidden"],
    "w:tblPr": ["w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
                "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
                "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
                "w:tblLayout", "w:tblCellMar", "w:tblLook"],
}


def _local(tag):
    """گۆڕینی تاگی {namespace}name بۆ w:name"""
    if tag.startswith("{"):
        return "w:" + tag.split("}", 1)[1]
    return tag


def insert_ordered(parent, child):
    """دانانی ئێلێمێنتێک لە شوێنی دروستی خۆیدا بەپێی ڕیزبەندی OOXML"""
    order = _ORDER.get(_local(parent.tag))
    if order is None:
        parent.append(child)
        return child
    try:
        idx = order.index(_local(child.tag))
    except ValueError:
        parent.append(child)
        return child
    for existing in parent:
        name = _local(existing.tag)
        if name in order and order.index(name) > idx:
            existing.addprevious(child)
            return child
    parent.append(child)
    return child



# ---------------------------------------------------------------- یاریدەدەر
def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    insert_ordered(tcPr, shd)


def set_cell_borders(cell, color_hex=RED_BORDER, sz=8):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color_hex)
        borders.append(el)
    insert_ordered(tcPr, borders)


def set_cell_margins(cell, top=60, bottom=60, start=100, end=100):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    insert_ordered(tcPr, mar)


def set_vertical_center(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    insert_ordered(tcPr, va)


def make_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    insert_ordered(pPr, bidi)


def style_run(run, size=10, bold=False, color=None, font_cs=FONT_CS):
    run.font.name = FONT_LATIN
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), FONT_LATIN)
    rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    rFonts.set(qn("w:cs"), font_cs)
    # قەبارەی فۆنت بۆ سکریپتی کۆمپلێکس
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    insert_ordered(rPr, szCs)
    rtl = OxmlElement("w:rtl")
    rtl.set(qn("w:val"), "1")
    insert_ordered(rPr, rtl)
    if bold:
        bcs = OxmlElement("w:bCs")
        bcs.set(qn("w:val"), "1")
        insert_ordered(rPr, bcs)


def fill_cell(cell, text, size=10, bold=False, color=None,
              align=WD_ALIGN_PARAGRAPH.RIGHT, bg=None, rtl=True):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    if rtl:
        make_rtl(p)
    pf = p.paragraph_format
    pf.space_before = Pt(2)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.15
    run = p.add_run(text)
    style_run(run, size=size, bold=bold, color=color)
    set_cell_borders(cell)
    set_vertical_center(cell)
    set_cell_margins(cell)
    if bg:
        set_cell_bg(cell, bg)


def table_rtl(table):
    """ڕیزکردنی ستوونەکان لە ڕاست بۆ چەپ"""
    tblPr = table._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    insert_ordered(tblPr, bidi)


def set_table_fixed_widths(table, widths):
    """پانی جێگیر بۆ ستوونەکان (بەبێ ئەمە Word پانییەکان پشتگوێ دەخات)"""
    tbl = table._tbl
    tblPr = tbl.tblPr

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    insert_ordered(tblPr, layout)

    total = sum(w.twips for w in widths)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    insert_ordered(tblPr, tblW)

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(w.twips))
        grid.append(col)
    tblPr.addnext(grid)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    insert_ordered(trPr, hdr)


def no_split(row):
    """ڕێگری لە دابەشبوونی ڕیز بەسەر دوو لاپەڕەدا"""
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    insert_ordered(trPr, el)


def set_row_height(row, cm):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(cm * 567)))
    h.set(qn("w:hRule"), "atLeast")
    insert_ordered(trPr, h)


def add_title(doc, text, size=16, space_after=10, color=NAVY):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    make_rtl(p)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(text)
    style_run(run, size=size, bold=True, color=color)
    return p


# ---------------------------------------------------------------- داتا
# (ژمارە، ناوی چالاکی، بەروار، ژمارەی بەشداربووان، وردەکاری)
ACTIVITIES = [
    (1, "یاری کۆتایی خولی تۆپی باڵەی تیپە میللییەکانی کوڕان (پاڵەوان: پەروەردەی وەرزشی زانکۆی سلێمانی)",
     "28/4/2026", "18 تیپ (252 یاریزانی کوڕ)",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، یاری کۆتایی پاڵەوانێتی تۆپی باڵە بۆ تیپە میللییەکان لە سلێمانی ئەنجامدرا، کە تیپی پەروەردەی وەرزشی زانکۆی سلێمانی پاڵەوانی پاڵەوانێتییەکە بوو. بە بەشداری 18 تیپ کە دەکاتە 252 یاریزانی کوڕ."),

    (2, "دابەشکردنی بڕوانامەی پشتێنی ڕەش و خولی دادوەری کیۆکۆشین لە پارکی هەواری شار",
     "1/5/2026", "44 یاریزان",
     "بە ئامادەبوونی سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، بڕوانامەی پشتێنی ڕەش بەسەر 44 یاریزانی پێشکەوتووی کیۆکۆشیندا دابەشکرا لە پارکی «هەواری شار»، هەروەها دەورەیەکی دادوەری بۆ کیۆکۆشین کرایەوە بە بەشداری 44 یاریزان."),

    (3, "ڕێوڕەسمی ڕێزلێنانی تیپەکانی یانەی پێشمەرگە (خوار 17 و 21 ساڵ) بۆ بڕینەیان بۆ خولی ئەستێرەکانی عێراق",
     "5/5/2026", "ئامار بەردەست نییە",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، ڕێوڕەسمی ڕێزلێنان لە تیپەکانی یانەی پێشمەرگە (خوار 17 و 21 ساڵ) ئەنجامدرا بە بۆنەی بڕینەیان بۆ خولی ئەستێرەکانی عێراق."),

    (4, "پاڵەوانێتی تێنسی زەوی بۆ یانەکانی سلێمانی (پلەی یەکەم: یانەی ئاشتی)",
     "8/5/2026", "7 یانە (70 بۆ 80 کوڕ، 15 بۆ 20 کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی تێنسی زەوی بۆ یانەکانی سلێمانی ئەنجامدرا و یانەی «ئاشتی» پلەی یەکەمی بەدەستهێنا. بە بەشداری 7 یانە، لە 70 بۆ 80 کوڕ و لە 15 بۆ 20 کچ."),

    (5, "پاڵەوانێتی کراوەی سنووکەر لە سلێمانی",
     "10/5/2026", "64 یاریزان",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی کراوەی سنووکەر ئەنجامدرا بە بەشداری 64 یاریزان، و یاریزان محەمەد جەناب پلەی یەکەمی بەدەستهێنا."),

    (6, "چالاکی وەرزشی تایبەت بە کێشە ئۆڵۆمپییەکانی تایکواندۆ لە فامیلی مۆڵ",
     "14/5/2026", "29 یاریزان (18 کوڕ، 11 کچ)",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، چالاکییەکی وەرزشی تایبەت بە کێشە ئۆڵۆمپییەکانی تایکواندۆ لە «فامیلی مۆڵ» ئەنجامدرا بە بەشداری 29 یاریزان (18 کوڕ و 11 کچ)."),

    (7, "پاڵەوانێتی بەرزکردنەوەی قورسایی بۆ خوێندکارانی کۆلێژی پەروەردەی وەرزشی سلێمانی",
     "16/5/2026", "11 خوێندکار (4 کوڕ، 7 کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی بەرزکردنەوەی قورسایی بۆ خوێندکارانی کۆلێژی پەروەردەی وەرزشی زانکۆی سلێمانی ئەنجامدرا بە بەشداری 11 خوێندکار (4 کوڕ و 7 کچ)."),

    (8, "سیمیناری زانستی و هونەری (تاولۆ و ساندا) بۆ ڕاهێنەرانی ووشوو کۆنگ فو",
     "20/5/2026", "ئامار بەردەست نییە",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، سیمینارێکی زانستی و هونەری لە بوارەکانی (تاولۆ و ساندا) بۆ ڕاهێنەرانی ووشوو کۆنگ فو ئەنجامدرا."),

    (9, "پاڵەوانێتی کراوەی بادمینتۆن بۆ کوڕان و کچان",
     "23/5/2026", "ئامار بەردەست نییە",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی کراوەی بادمینتۆن بۆ کوڕان و کچان کۆتایی هات."),

    (10, "دەستپێکردنی خولی فێرگەکانی تۆپی پێی کچان و پێشوازیکردن لە هەڵبژاردەی ناشئاتی عێراق",
     "15/6/2026", "9 تیپ (140 یاریزانی کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی و بە ئامادەبوونی خاتوو پەیمان جەمال، خولی فێرگە تۆپی پێیەکانی کچان لە سلێمانی دەستیپێکرد، لەگەڵ پێشوازیکردن لە شاندی هەڵبژاردەی ناشئاتی عێراق. 9 تیپ و 140 یاریزانی کچ."),

    (11, "پاڵەوانێتی تیر و کەوانی یانەکانی سنووری پارێزگای سلێمانی (کوڕان و کچان)",
     "19/6/2026", "8 یانە (سلێمانی، سیروانی نوێ، ئاشتی، پێشمەرگە، هەڵەبجە، ئافرۆدێت، خورماڵ، نەورۆز)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی تیر و کەوان بۆ یانەکانی سنووری پارێزگای سلێمانی بۆ هەردوو ڕەگەزی کوڕان و کچان ئەنجامدرا بە بەشداری 8 یانە: سلێمانی، سیروانی نوێ، ئاشتی، پێشمەرگە، هەڵەبجە، ئافرۆدێت، خورماڵ و نەورۆز."),

    (12, "کەرنەڤاڵی وەرزشی هەمەجۆر بەبۆنەی ڕۆژی جیهانی ئۆڵۆمپی لە سلێمانی",
     "23/6/2026", "300 بۆ 350 یاریزان",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، میهرەجانێکی وەرزشی هەمەجۆر بە بۆنەی ڕۆژی جیهانی ئۆڵۆمپی لە سلێمانی ئەنجامدرا، بە نزیکەی 300 بۆ 350 یاریزان لە یارییە جۆراوجۆرەکاندا."),

    (13, "منافەساتی پاڵەوانێتی «پادڵ بۆرد»ی عێراق لە سلێمانی",
     "25/6/2026", "نزیکەی 28 یاریزانی کوڕ",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، منافەساتی پاڵەوانێتی «پادڵ بۆرد»ی عێراق لە سلێمانی ئەنجامدرا بە بەشداری نزیکەی 28 یاریزانی کوڕ."),

    (14, "کۆتاییهاتنی پاڵەوانێتی کراوەی تۆپی مێز بۆ سەرجەم تەمەنەکان",
     "28/6/2026", "نزیکەی 80 یاریزانی کوڕ و کچ",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی کراوەی تۆپی مێز لە سلێمانی بۆ جۆراوجۆر تەمەن کۆتایی هات بە بەشداری نزیکەی 80 یاریزانی کوڕ و کچ."),

    (15, "یارییەکانی قۆناغی دووەمی خولی تۆپی دەستی عێراق لە هۆڵی زانکۆی سلێمانی",
     "5/7/2026", "7 یانە (نزیکەی 100 یاریزان)",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، یارییەکانی قۆناغی دووەمی خولی تۆپی دەستی عێراق لە هۆڵی زانکۆی سلێمانی ئەنجامدران بە بەشداری 7 یانە و نزیکەی 100 یاریزان."),

    (16, "خولی پەرەپێدانی هونەری «تاولۆ» و خولی دادوەری «ساندا» بۆ ووشوو کۆنگ فو",
     "5/7/2026", "79 کەس (67 تاولۆ + 12 دادوەری ساندا)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، بۆ یەکەم جار لە شاردا خولێکی پەرەپێدان لە هونەری «تاولۆ» بۆ وەرزشوانانی ووشوو کۆنگ فو ئەنجامدرا بە بەشداری نزیکەی 67 کەس، هەروەها 12 کەس لە خولی دادوەری «ساندا» بەشدار بوون."),

    (17, "پاڵەوانێتی «سوپەری عێراق» بۆ کیک بۆکسینگ لە هۆڵەکانی کۆلێژی پەروەردەی وەرزش",
     "22/7/2026", "96 یاریزان (81 کوڕ، 15 کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی «سوپەری عێراق» بۆ کیک بۆکسینگ لە هۆڵەکانی کۆلێژی پەروەردەی وەرزشی ئەنجامدرا بە بەشداری نزیکەی 96 یاریزان کە 15 یان کچ بوون."),

    (18, "کۆتاییهاتنی پاڵەوانێتی یانەکانی عێراق بۆ تایکواندۆ لە سلێمانی",
     "26/7/2026", "83 بۆ 84 یانە",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی یانەکانی عێراق بۆ تایکواندۆ لە سلێمانی کۆتایی هات بە بەشداری 83 یانە."),

    (19, "میهرەجانی «نوێکردنەوەی تۆپی سەبەتە» بۆ یانە و فێرگەکان (خوار 12 ساڵ)",
     "30/7/2026", "نزیکەی 50 یاریزان",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، میهرەجانی «نوێکردنەوەی تۆپی سەبەتە» بۆ یانە و فێرگە تۆپی سەبەتەکان بۆ تەمەنی خوار 12 ساڵ ئەنجامدرا بە بەشداری نزیکەی 50 یاریزان."),

    (20, "خولی پەرەپێدانی وەرزشی «پیلاتس و ئایرۆبیک» بۆ ڕاهێنەران لە سلێمانی",
     "31/7/2026", "نزیکەی 35 ڕاهێنەری کچ",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، بۆ یەکەم جار لە سلێمانی خولێکی پەرەپێدان لە وەرزشەکانی «پیلاتس و ئایرۆبیک» کۆتایی هات بە بەشداری نزیکەی 35 ڕاهێنەری هۆڵەکان (کچان)."),

    (21, "میهرەجانی «مینی باسکێت» بۆ تیپەکانی کچان و کوڕان",
     "31/7/2026", "200 یاریزانی کچ و کوڕ",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، میهرەجانی «مینی باسکێت» بۆ تیپەکانی کچان لە هۆڵی کۆلێژی پەروەردەی وەرزشی ئەنجامدرا بە بەشداری 200 یاریزانی کچ و کوڕ."),

    (22, "یاری دۆستانەی تۆپی سەبەتە (هەڵبژاردەی سنە بەرامبەر هەڵبژاردەی سلێمانی)",
     "6/8/2026", "نزیکەی 28 یاریزان",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، یارییەکی دۆستانەی تۆپی سەبەتە لەنێوان هەڵبژاردەی شاری سنە و هەڵبژاردەی شاری سلێمانی ئەنجامدرا بە بەشداری نزیکەی 28 یاریزان."),

    (23, "پاڵەوانێتی وەرزشیی ڕێکخراوی جیو جیتسۆ لە زانکۆی سلێمانی",
     "13/8/2026", "90 بەشداربوو",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، ڕێکخراوی جیو جیتسۆ پاڵەوانێتییەکی وەرزشی لە گەڕەکی زانکۆی سلێمانی ڕێکخست بە بەشداری 90 کەس."),

    (24, "پاڵەوانێتی مەلەوانگەکانی سلێمانی بۆ کچان و کوڕان",
     "14/8/2026", "50 یاریزانی کوڕ و کچ",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی مەلەوانگەکان لە سلێمانی بە بەشداری بەرفراوانی کوڕان و کچان ئەنجامدرا، بە بەشداری 50 یاریزانی کوڕ و کچ."),

    (25, "کۆبوونەوەی هونەری بۆ پاڵەوانێتی تۆپی باڵەی فێرگە وەرزشییەکان",
     "18/8/2026", "14 تیپ، 7 فێرگە (196 یاریزانی کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، کۆبوونەوەی هونەری بۆ دەستپێکردنی پاڵەوانێتی تۆپی باڵەی فێرگە وەرزشییەکان لە سلێمانی ئەنجامدرا: 14 تیپ لە 7 فێرگە و 196 یاریزانی کچ."),

    (26, "خولی ڕاهێنان و ناوبژیوانی بۆ وەرزشی ووشوو کۆنگ فو",
     "27/8/2026", "90 یاریزان (کوڕ و کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، خولێکی ڕاهێنان و ناوبژیوانی بۆ وەرزشی ووشوو کۆنگ فو لە سلێمانی ئەنجامدرا بە بەشداری 90 یاریزان لە هەردوو ڕەگەز."),

    (27, "پاڵەوانێتی کراوەی شەترەنجی سلێمانی",
     "27/8/2026", "49 یاریزان (کوڕ و کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی کراوەی شەترەنجی سلێمانی ئەنجامدرا بە بەشداری 49 یاریزانی کوڕ و کچ."),

    (28, "میهرەجانی چوارەمی کاراتێی یانەی پێشمەرگە",
     "28/8/2026", "350 یاریزان",
     "بە ئامادەبوونی د. محەمەد ئیبراهیم کەنعان، سەرۆکی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، میهرەجانی چوارەمی کاراتێی یانەی پێشمەرگە ئەنجامدرا بە بەشداری 350 یاریزان."),

    (29, "خولی دادوەری تایکواندۆ",
     "28/8/2026", "30 یاریزانی کوڕ و کچ",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، خولێکی دادوەری تایکواندۆ ئەنجامدرا بە بەشداری 30 یاریزانی کچ و کوڕ."),

    (30, "پاڵەوانێتی ووشوو کۆنگ فو (تاولۆ و ساندا)",
     "28/8/2026", "110 یاریزان (50 تاولۆ، 60 ساندا)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی ووشوو کۆنگ فو ئەنجامدرا بە بەشداری 110 یاریزان، کە 50 یان لە تاولۆ و 60 یان لە ساندا بەشدار بوون."),

    (31, "دەستپێکردنی پاڵەوانێتی تۆپی باڵەی فێرگە وەرزشییەکان",
     "2/9/2026", "14 تیپ",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی تۆپی باڵەی فێرگە وەرزشییەکان دەستیپێکرد بە بەشداری 14 تیپ."),

    (32, "پاڵەوانێتی گۆڕەپان و مەیدانی یانەکانی کوردستان و سلێمانی (ساحە و مەیدان)",
     "5/9/2026", "18 یانە (8 کچان، 10 کوڕان – 200 بۆ 250 کچ و کوڕ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی گۆڕەپان و مەیدانی یانەکانی کوردستان و سلێمانی ئەنجامدرا بە بەشداری 18 یانە (8 یانەی کچان و 10 یانەی کوڕان) و نزیکەی 200 بۆ 250 کچ و کوڕ."),

    (33, "پاڵەوانێتی تێنسی زەوی",
     "9/9/2026", "7 یانە (70 بۆ 80 کوڕ، 13 کچ)",
     "بە سەرپەرشتی لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی، پاڵەوانێتی تێنسی زەوی ئەنجامدرا بە بەشداری 7 یانە، لە 70 بۆ 80 کوڕ و 13 کچ."),
]


# ---------------------------------------------------------------- بنیاتنان
def build_document(out_path):
    doc = Document()

    # ڕێکخستنی لاپەڕە
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.3)
    sec.bottom_margin = Cm(1.3)
    sec.left_margin = Cm(1.3)
    sec.right_margin = Cm(1.3)

    normal = doc.styles["Normal"]
    normal.font.name = FONT_LATIN
    normal.font.size = Pt(10)
    normal.element.rPr.rFonts.set(qn("w:cs"), FONT_CS)

    # ---------------------------------------- خشتەی یەکەم: کورتەی چالاکییەکان
    add_title(doc, "خشتەی چالاکییە وەرزشییەکان بەپێی ژمارەی بەشداربووان", size=16, space_after=4)
    add_title(doc, "لیژنەی نیشتمانی ئۆڵۆمپی عێراق – لقی سلێمانی | ساڵی 2026",
              size=11, space_after=10, color="595959")

    widths = [Cm(1.2), Cm(8.6), Cm(2.5), Cm(6.1)]
    headers = ["ژ", "ناوی چالاکی", "بەروار", "ژمارەی بەشداربووان"]

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table_rtl(table)

    hdr = table.rows[0]
    repeat_header(hdr)
    set_row_height(hdr, 0.95)
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = widths[i]
        fill_cell(cell, text, size=11, bold=True, color=WHITE,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=NAVY)

    for idx, (num, name, date, count, _detail) in enumerate(ACTIVITIES):
        bg = WHITE if idx % 2 == 0 else ROW_ALT
        row = table.add_row()
        set_row_height(row, 0.7)
        no_split(row)
        cells = row.cells
        for i in range(4):
            cells[i].width = widths[i]
        fill_cell(cells[0], str(num), size=10, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg)
        fill_cell(cells[1], name, size=10, align=WD_ALIGN_PARAGRAPH.RIGHT, bg=bg)
        fill_cell(cells[2], date, size=10,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg, rtl=False)
        fill_cell(cells[3], count, size=10,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg)

    set_table_fixed_widths(table, widths)

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    make_rtl(note)
    note.paragraph_format.space_before = Pt(8)
    style_run(note.add_run("تێبینی: چالاکییەکان بەپێی ڕێکەوت ڕیزکراون. "
                           "کۆی گشتی چالاکییەکان: %d چالاکی." % len(ACTIVITIES)),
              size=9, color="595959")

    # ---------------------------------------- خشتەی دووەم: وردەکاری
    doc.add_page_break()
    add_title(doc, "وردەکاری چالاکییە وەرزشییەکان", size=16, space_after=10)

    widths2 = [Cm(1.2), Cm(2.5), Cm(14.7)]
    headers2 = ["ژ", "بەروار", "وردەکاری چالاکی"]

    t2 = doc.add_table(rows=1, cols=3)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.autofit = False
    table_rtl(t2)

    hdr2 = t2.rows[0]
    repeat_header(hdr2)
    set_row_height(hdr2, 0.95)
    for i, text in enumerate(headers2):
        cell = hdr2.cells[i]
        cell.width = widths2[i]
        fill_cell(cell, text, size=11, bold=True, color=WHITE,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=NAVY)

    for idx, (num, _name, date, _count, detail) in enumerate(ACTIVITIES):
        bg = WHITE if idx % 2 == 0 else ROW_ALT
        row = t2.add_row()
        set_row_height(row, 0.7)
        no_split(row)
        cells = row.cells
        for i in range(3):
            cells[i].width = widths2[i]
        fill_cell(cells[0], str(num), size=10, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg)
        fill_cell(cells[1], date, size=10,
                  align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg, rtl=False)
        fill_cell(cells[2], detail, size=10,
                  align=WD_ALIGN_PARAGRAPH.JUSTIFY, bg=bg)

    set_table_fixed_widths(t2, widths2)

    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(here, "output", "چالاکییە-وەرزشییەکان-سلێمانی-2026.docx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    build_document(out)
    print("دروستکرا:", out)
