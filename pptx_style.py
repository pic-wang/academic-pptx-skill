"""House style helpers for academic .pptx decks (python-pptx).

Every visual rule in STYLE.md is encoded here. Build decks by calling these
helpers instead of setting fonts, colours and borders by hand.

Inline markup accepted by every text helper:
    **text**  -> red (C00000) bold run   (key phrase)
    *text*    -> red (C00000) regular run
    plain     -> 1A1A1A regular run

Canvas: 10 x 5.625 in (16:9). The reference decks were authored on
13.33 x 7.5 in; all sizes below are the reference sizes x 0.75.
"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- constants
FONT = "Cambria"
TEXT = RGBColor(0x1A, 0x1A, 0x1A)
BLACK = RGBColor(0x00, 0x00, 0x00)
RED = RGBColor(0xC0, 0x00, 0x00)
GRAY = RGBColor(0xF2, 0xF2, 0xF2)
MUTED = RGBColor(0x6B, 0x6B, 0x6B)
AGENDA_INACTIVE = RGBColor(0xBF, 0xBF, 0xBF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W, SLIDE_H = 10.0, 5.625
MARGIN = 0.5
CONTENT_W = SLIDE_W - 2 * MARGIN

SIZE = {
    "title": 22,        # page title
    "section": 16,      # red-bullet section header
    "body": 15,         # body text (= 20 pt on 13.33 in canvas)
    "body_small": 13,   # dense body / table cells
    "takehome": 15,
    "caption": 12,
    "cite": 8,
    "page": 10,
    "deck_title": 22,
    "agenda": 18,
}

TAKEHOME_Y = 4.72
CITE_Y = 5.30

_MARKUP = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*)")


# ---------------------------------------------------------------- deck setup
def new_presentation(template: str | Path | None = None) -> Presentation:
    """Blank 16:9 deck, or `template` with its existing slides removed."""
    if template and Path(template).exists():
        prs = Presentation(str(template))
        sldIdLst = prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            prs.part.drop_rel(sldId.rId)
            sldIdLst.remove(sldId)
    else:
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_W)
        prs.slide_height = Inches(SLIDE_H)
    return prs


def blank_slide(prs):
    """Add a slide on the layout with the fewest placeholders."""
    layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    slide = prs.slides.add_slide(layout)
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    return slide


# ---------------------------------------------------------------- text core
def _style_run(run, size, bold=False, color=TEXT, italic=False):
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):  # keep CJK / complex-script runs in Cambria too
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)


def add_runs(paragraph, text, size, bold=False, color=TEXT):
    """Write `text` into `paragraph`, honouring **red bold** / *red* markup."""
    for part in _MARKUP.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            r = paragraph.add_run()
            r.text = part[2:-2]
            _style_run(r, size, bold=True, color=RED)
        elif part.startswith("*") and part.endswith("*"):
            r = paragraph.add_run()
            r.text = part[1:-1]
            _style_run(r, size, bold=bold, color=RED)
        else:
            r = paragraph.add_run()
            r.text = part
            _style_run(r, size, bold=bold, color=color)


def _textbox(slide, x, y, w, h, wrap=True, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    return tb, tf


def _set_bullet(paragraph, char="•", font="Arial", marL=0.25, color=None):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(marL)))
    pPr.set("indent", str(-Inches(marL)))
    for tag in ("a:buNone", "a:buChar", "a:buFont", "a:buClr"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    if color is not None:
        buClr = etree.SubElement(pPr, qn("a:buClr"))
        etree.SubElement(buClr, qn("a:srgbClr")).set("val", str(color))
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", font)
    etree.SubElement(pPr, qn("a:buChar")).set("char", char)


def _line_spacing(paragraph, pct):
    pPr = paragraph._p.get_or_add_pPr()
    ln = pPr.find(qn("a:lnSpc"))
    if ln is None:
        ln = etree.SubElement(pPr, qn("a:lnSpc"))
        pPr.insert(0, ln)
    for c in list(ln):
        ln.remove(c)
    etree.SubElement(ln, qn("a:spcPct")).set("val", str(int(pct * 1000)))


# ---------------------------------------------------------------- page parts
def add_page_title(slide, text, x=MARGIN, y=0.38, w=CONTENT_W - 1.4):
    """22 pt Cambria bold black, top-left. No accent line underneath.

    Must fit on one line: about 52 characters at the default width.
    """
    plain = _MARKUP.sub(lambda m: m.group(0).strip("*"), text)
    if len(plain) * SIZE["title"] * 0.0066 > w:
        import warnings
        warnings.warn(f"page title may wrap ({len(plain)} chars): {plain!r}")
    tb, tf = _textbox(slide, x, y, w, 0.5, anchor=MSO_ANCHOR.TOP)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    add_runs(p, text, SIZE["title"], bold=True, color=BLACK)
    return tb


def add_section_header(slide, text, x=MARGIN, y=1.05, w=CONTENT_W):
    """Red bullet + 16 pt Cambria bold red text."""
    tb, tf = _textbox(slide, x, y, w, 0.35)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _set_bullet(p, marL=0.22, color=RED)
    add_runs(p, text, SIZE["section"], bold=True, color=RED)
    return tb


def add_body(slide, items, x=MARGIN + 0.22, y=1.45, w=CONTENT_W - 0.22, h=2.5,
             size=None, bullets=True, space_after=6):
    """Body paragraphs, left-aligned.

    `items`: list of str, or (level, str). Level 0 = black bullet, level 1 =
    en-dash sub-bullet. Pass bullets=False for plain lead-in sentences.
    """
    size = size or SIZE["body"]
    tb, tf = _textbox(slide, x, y, w, h)
    first = True
    for it in items:
        level, text = it if isinstance(it, tuple) else (0, it)
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(space_after)
        if bullets:
            _set_bullet(p, char="•" if level == 0 else "–",
                        marL=0.22 + 0.25 * level, color=TEXT)
        add_runs(p, text, size - (1 if level else 0))
    return tb


def add_takehome_box(slide, text, y=TAKEHOME_Y, w=None, x=None, size=None):
    """Red dashed-border box at the bottom. Mandatory on every content slide.

    Width defaults to fit the text (single line) up to CONTENT_W; the box is
    centred horizontally. Put the key phrase in **red bold** markup.
    """
    size = size or SIZE["takehome"]
    plain = _MARKUP.sub(lambda m: m.group(0).strip("*"), text)
    if w is None:
        w = min(CONTENT_W, 0.30 + len(plain) * size * 0.0062)
    if x is None:
        x = (SLIDE_W - w) / 2
    lines = max(1, round(len(plain) * size * 0.0062 / (w - 0.2) + 0.49))
    h = 0.12 + lines * size * 1.25 / 72
    tb, tf = _textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.MIDDLE)
    tf.margin_left = tf.margin_right = Inches(0.1)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs(p, text, size)
    spPr = tb._element.spPr
    for el in spPr.findall(qn("a:ln")):
        spPr.remove(el)
    ln = etree.SubElement(spPr, qn("a:ln"))
    ln.set("w", "19050")
    sf = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", "C00000")
    etree.SubElement(ln, qn("a:prstDash")).set("val", "dash")
    return tb


def add_caption(slide, label, text, x, y, w, align=PP_ALIGN.CENTER):
    """12 pt caption: bold 'Fig. N.' / 'Tab. N.' prefix, single line."""
    tb, tf = _textbox(slide, x, y, w, 0.28)
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = label.rstrip() + " "
    _style_run(r, SIZE["caption"], bold=True)
    add_runs(p, text, SIZE["caption"])
    return tb


def add_citation(slide, text, x=0.1, y=CITE_Y, w=8.6):
    """Full reference in small muted text at the bottom-left."""
    tb, tf = _textbox(slide, x, y, w, 0.25)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    add_runs(p, text, SIZE["cite"], color=MUTED)
    return tb


def add_page_number(slide, n):
    tb, tf = _textbox(slide, SLIDE_W - 0.6, SLIDE_H - 0.32, 0.5, 0.25)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    add_runs(p, str(n), SIZE["page"])
    return tb


def add_logo(slide, path, w=1.45):
    """Institution logo in the top-right corner (optional)."""
    if path and Path(path).exists():
        return slide.shapes.add_picture(str(path), Inches(SLIDE_W - w - 0.1),
                                        Inches(0.08), width=Inches(w))
    return None


def add_figure(slide, path, x, y, w=None, h=None, caption=None):
    """Picture plus optional (label, text) caption directly beneath it.

    With both `w` and `h`, the image is scaled to fit inside that box with its
    aspect ratio kept, and centred horizontally in it. The caption height
    (0.3 in) is reserved inside the box.
    """
    if w and h:
        from PIL import Image
        with Image.open(path) as im:
            ar = im.width / im.height
        box_h = h - (0.3 if caption else 0)
        fw, fh = (w, w / ar) if w / ar <= box_h else (box_h * ar, box_h)
        x, w, h = x + (w - fw) / 2, fw, fh
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y),
                                   Inches(w) if w else None,
                                   Inches(h) if h else None)
    if caption:
        label, text = caption
        add_caption(slide, label, text, x, y + Emu(pic.height).inches + 0.03,
                    Emu(pic.width).inches)
    return pic


def add_concept_boxes(slide, boxes, y=3.05, h=1.45, x=MARGIN, w=CONTENT_W, gap=0.25):
    """Row of rounded red-bordered boxes: [(header, [bullet, ...]), ...]."""
    n = len(boxes)
    bw = (w - gap * (n - 1)) / n
    shapes = []
    for i, (head, items) in enumerate(boxes):
        bx = x + i * (bw + gap)
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx),
                                    Inches(y), Inches(bw), Inches(h))
        sh.adjustments[0] = 0.12
        sh.fill.background()
        sh.line.color.rgb = RED
        sh.line.width = Pt(1.5)
        sh.shadow.inherit = False
        tf = sh.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = Inches(0.12)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(4)
        add_runs(p, head, SIZE["body_small"], bold=True)
        for it in items:
            q = tf.add_paragraph()
            q.alignment = PP_ALIGN.LEFT
            _set_bullet(q, marL=0.18, color=TEXT)
            add_runs(q, it, 11)
        shapes.append(sh)
    return shapes


# ---------------------------------------------------------------- tables
def _cell_border(tag, w_emu, color="000000"):
    ln = etree.Element(qn(tag))
    ln.set("w", str(w_emu))
    ln.set("cap", "flat")
    ln.set("cmpd", "sng")
    ln.set("algn", "ctr")
    if w_emu:
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
    else:
        etree.SubElement(ln, qn("a:noFill"))
    return ln


def _apply_three_line_borders(table, rule_w=12700):
    """Black top rule, black header underline, black bottom rule. No grid.

    OOXML order trap: inside <a:tcPr>, lnL/lnR/lnT/lnB must precede the cell
    fill, or PowerPoint ignores the borders. This rebuilds tcPr in order.
    """
    n_rows = len(table.rows)
    for r, row in enumerate(table.rows):
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            fills = [copy.deepcopy(el) for el in tcPr
                     if el.tag in (qn("a:solidFill"), qn("a:noFill"))]
            for el in list(tcPr):
                tcPr.remove(el)
            top = rule_w if r in (0, 1) else 0
            bottom = rule_w if r in (0, n_rows - 1) else 0
            tcPr.append(_cell_border("a:lnL", 0))
            tcPr.append(_cell_border("a:lnR", 0))
            tcPr.append(_cell_border("a:lnT", top))
            tcPr.append(_cell_border("a:lnB", bottom))
            for el in fills:
                tcPr.append(el)


def add_table(slide, header, rows, x, y, w, col_widths=None, highlight=(),
              size=None, row_h=0.32):
    """Three-line academic table with zebra stripes.

    `highlight`: iterable of (row_idx, col_idx) data-cell positions (0-based,
    excluding the header) rendered red bold. Values may also use markup.
    """
    size = size or SIZE["body_small"]
    n_rows, n_cols = len(rows) + 1, len(header)
    gf = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y),
                                Inches(w), Inches(row_h * n_rows))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)
    for attr in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol", "bandCol"):
        if attr in tblPr.attrib:
            del tblPr.attrib[attr]
    if col_widths:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = Inches(cw)
    for r in range(n_rows):
        tbl.rows[r].height = Inches(row_h)
    hl = set(highlight)
    for r in range(n_rows):
        values = header if r == 0 else rows[r - 1]
        for c in range(n_cols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.text = ""
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            val = str(values[c])
            if r > 0 and (r - 1, c) in hl:
                val = f"**{val.strip('*')}**"
            add_runs(p, val, size, bold=(r == 0), color=BLACK if r == 0 else TEXT)
            if r == 0 or r % 2 == 0:
                cell.fill.background()
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = GRAY
    _apply_three_line_borders(tbl)
    return gf


# ---------------------------------------------------------------- whole slides
def title_slide(prs, title, authors, presenter=None, affiliation=None,
                venue=None, logo=None):
    """Centred title (key phrase via markup), authors, presenter block."""
    s = blank_slide(prs)
    add_logo(s, logo, w=1.8)
    tb, tf = _textbox(s, 1.0, 1.55, 8.0, 1.1, anchor=MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs(p, title, SIZE["deck_title"], bold=True, color=BLACK)
    tb, tf = _textbox(s, 1.0, 2.95, 8.0, 0.35)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs(p, authors, SIZE["body"])
    if venue:
        tb, tf = _textbox(s, 1.0, 3.32, 8.0, 0.3)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, venue, SIZE["caption"], color=MUTED)
    lines = [l for l in (presenter and f"Presenter: {presenter}", affiliation) if l]
    for i, line in enumerate(lines):
        yy = 4.55 + i * 0.36
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.55), Inches(yy + 0.05),
                                 Inches(0.05), Inches(0.24))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RED
        bar.line.fill.background()
        bar.shadow.inherit = False
        tb, tf = _textbox(s, 5.68, yy, 4.2, 0.34)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        if line.startswith("Presenter: "):
            r = p.add_run()
            r.text = "Presenter: "
            _style_run(r, SIZE["body"], bold=True, color=BLACK)
            add_runs(p, line[len("Presenter: "):], SIZE["body"])
        else:
            add_runs(p, line, SIZE["body"], bold=True, color=BLACK)
    return s


def agenda_slide(prs, sections, current, page=None, logo=None, title="Content"):
    """Section progress slide: current item black, others gray hollow bullets."""
    s = blank_slide(prs)
    add_logo(s, logo)
    add_page_title(s, title)
    tb, tf = _textbox(s, 2.0, 1.25, 7.0, 3.2, anchor=MSO_ANCHOR.MIDDLE)
    for i, sec in enumerate(sections):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _line_spacing(p, 200)
        active = (i == current) or (sec == current)
        if active:
            _set_bullet(p, "•", "Arial", marL=0.3, color=BLACK)
        else:
            _set_bullet(p, "o", "Courier New", marL=0.3, color=BLACK)
        add_runs(p, sec, SIZE["agenda"], bold=True,
                 color=BLACK if active else AGENDA_INACTIVE)
    if page is not None:
        add_page_number(s, page)
    return s


def content_slide(prs, title, page=None, logo=None):
    """Blank content slide with page title, optional logo and page number."""
    s = blank_slide(prs)
    add_logo(s, logo)
    add_page_title(s, title)
    if page is not None:
        add_page_number(s, page)
    return s


def check_takehome(prs, exempt=(1,)):
    """Return 1-based indices of slides missing a dashed take-home box."""
    missing = []
    for i, s in enumerate(prs.slides, 1):
        if i in exempt:
            continue
        if not any(b'prstDash val="dash"' in etree.tostring(sh._element)
                   and sh.has_text_frame for sh in s.shapes):
            missing.append(i)
    return missing
