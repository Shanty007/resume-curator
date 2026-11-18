"""Render resume JSON -> DOCX (native Word structures only) -> PDF (LibreOffice, tagged PDF/A-2b).

ATS rules enforced by construction:
  * single text flow in reading order: no tables, text boxes, shapes with text, or columns
  * nothing in headers/footers
  * real Unicode text: bullets are a real U+2022 in the body font, not Symbol/Wingdings
  * standard section headings (EN/DE)
  * fonts embedded in the PDF; tagged PDF for structure
"""
from __future__ import annotations

import copy
import datetime as _dt
import shutil
import subprocess
import tempfile
from pathlib import Path

import docx
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Emu

from .templates import TEMPLATES, HEADINGS

A4_W_CM, A4_H_CM = 21.0, 29.7


# ---------------------------------------------------------------- low-level helpers
def _set_run_font(run, name):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), name)


def _shade(p, fill):
    ppr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    ppr.append(shd)


def _border(p, sides, color, sz=8, space=1):
    """sides: dict side->size (eighths of a point)."""
    ppr = p._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr"); ppr.append(pbdr)
    for side in ("top", "left", "bottom", "right"):
        if side in sides:
            b = OxmlElement(f"w:{side}")
            b.set(qn("w:val"), "single"); b.set(qn("w:sz"), str(sides[side]))
            b.set(qn("w:space"), str(space if side != "left" else 6)); b.set(qn("w:color"), color)
            pbdr.append(b)


def _char_spacing(run, pt):
    if not pt:
        return
    rpr = run._element.get_or_add_rPr()
    sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), str(int(pt * 20))); rpr.append(sp)


def _no_ligatures(run):
    # w14:ligatures none, so parsers never see U+FB01 'ﬁ' glyphs
    rpr = run._element.get_or_add_rPr()
    lig = OxmlElement("w14:ligatures")
    lig.set("{http://schemas.microsoft.com/office/word/2010/wordml}val", "none")
    rpr.append(lig)


class Doc:
    def __init__(self, spec, pdf_mode, lang):
        self.s = spec
        self.lang = lang
        self.h = HEADINGS[lang]
        self.font = spec["font_pdf" if pdf_mode else "font_docx"]
        self.hfont = spec["head_font_pdf" if pdf_mode else "head_font_docx"]
        self.d = docx.Document()
        self._page()
        self._styles()
        self._numbering()
        # python-docx starts with one empty paragraph; drop it
        body = self.d.element.body
        for p in list(body.iterchildren(qn("w:p"))):
            body.remove(p)

    # -------- page/styles
    def _page(self):
        sec = self.d.sections[0]
        sec.page_width, sec.page_height = Cm(A4_W_CM), Cm(A4_H_CM)
        t, r, b, l = self.s["margins_cm"]
        sec.top_margin, sec.right_margin, sec.bottom_margin, sec.left_margin = Cm(t), Cm(r), Cm(b), Cm(l)
        sec.header_distance = Cm(0.5); sec.footer_distance = Cm(0.5)
        self.text_w_cm = A4_W_CM - r - l

    def _styles(self):
        st = self.d.styles["Normal"]
        st.font.name = self.font; st.font.size = Pt(self.s["body_pt"])
        st.font.color.rgb = RGBColor.from_string(self.s["text"])
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), self.font)
        pf = st.paragraph_format
        pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = 1.12
        # document language (helps hyphenation / spell-check, harmless for ATS)
        lang = OxmlElement("w:lang"); lang.set(qn("w:val"), "de-DE" if self.lang == "de" else "en-GB")
        rpr.append(lang)

    def _numbering(self):
        """Create a bullet list whose glyph is a real U+2022 in the body font."""
        numbering = self.d.part.numbering_part.element
        abs_id, num_id = 77, 77
        indent = Cm(self.bullet_left_cm()).twips if hasattr(Cm(1), "twips") else int(self.bullet_left_cm() * 567)
        hang = int(0.38 * 567)
        a = OxmlElement("w:abstractNum"); a.set(qn("w:abstractNumId"), str(abs_id))
        mlt = OxmlElement("w:multiLevelType"); mlt.set(qn("w:val"), "singleLevel"); a.append(mlt)
        lvl = OxmlElement("w:lvl"); lvl.set(qn("w:ilvl"), "0")
        for tag, val in (("w:start", "1"), ("w:numFmt", "bullet"), ("w:lvlText", "•"), ("w:lvlJc", "left")):
            e = OxmlElement(tag); e.set(qn("w:val"), val); lvl.append(e)
        ppr = OxmlElement("w:pPr"); ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(indent)); ind.set(qn("w:hanging"), str(hang)); ppr.append(ind); lvl.append(ppr)
        rpr = OxmlElement("w:rPr"); rf = OxmlElement("w:rFonts")
        for at in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(at), self.font)
        rpr.append(rf)
        col = OxmlElement("w:color"); col.set(qn("w:val"), self.s["accent"]); rpr.append(col)
        lvl.append(rpr); a.append(lvl)
        # abstractNum must precede num elements
        first_num = numbering.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(a)
        else:
            numbering.append(a)
        n = OxmlElement("w:num"); n.set(qn("w:numId"), str(num_id))
        an = OxmlElement("w:abstractNumId"); an.set(qn("w:val"), str(abs_id)); n.append(an)
        numbering.append(n)
        self.num_id = num_id

    def bullet_left_cm(self):
        base = self.s["date_col_cm"] if self.s["layout"] == "tabular" else 0
        return base + 0.42

    def content_left_cm(self):
        return self.s["date_col_cm"] if self.s["layout"] == "tabular" else 0

    # -------- paragraph builders
    def p(self, runs=(), align=None, before=0, after=0, left_cm=None, first_cm=None,
          right_cm=None, tabs=(), keep_next=False, line=None):
        para = self.d.add_paragraph()
        pf = para.paragraph_format
        pf.space_before = Pt(before); pf.space_after = Pt(after)
        if line:
            pf.line_spacing = line
        if left_cm is not None:
            pf.left_indent = Cm(left_cm)
        if right_cm is not None:
            pf.right_indent = Cm(right_cm)
        if first_cm is not None:
            pf.first_line_indent = Cm(first_cm)
        if align == "center":
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for pos, kind in tabs:
            pf.tab_stops.add_tab_stop(Cm(pos), WD_TAB_ALIGNMENT.RIGHT if kind == "right" else WD_TAB_ALIGNMENT.LEFT)
        pf.keep_with_next = keep_next
        pf.widow_control = True
        for r in runs:
            self.run(para, r[0], **r[1]) if isinstance(r, tuple) else self.run(para, r)
        return para

    def run(self, para, text, **f):
        r = para.add_run(text)
        _set_run_font(r, f.get("font", self.font))
        if f.get("size"):
            r.font.size = Pt(f["size"])
        r.bold = f.get("bold")
        r.italic = f.get("italic")
        col = f.get("color")
        if col:
            r.font.color.rgb = RGBColor.from_string(self.s.get(col, col))
        if f.get("caps"):
            r.font.all_caps = True
        if f.get("smallcaps"):
            r.font.small_caps = True
        _char_spacing(r, f.get("spacing"))
        if f.get("shade"):
            rpr = r._element.get_or_add_rPr()
            shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), f["shade"]); rpr.append(shd)
        _no_ligatures(r)
        return r

    def bullet(self, text, after=1.2):
        para = self.p([(text, {})], after=after, line=1.1)
        ppr = para._p.get_or_add_pPr()
        numpr = OxmlElement("w:numPr")
        il = OxmlElement("w:ilvl"); il.set(qn("w:val"), "0")
        ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(self.num_id))
        numpr.append(il); numpr.append(ni)
        ppr.insert(0, numpr)
        return para

    def heading(self, key_or_text, first=False):
        h = self.s["heading"]
        text = self.h.get(key_or_text, key_or_text)
        f = {"bold": True, "size": h["size"], "color": h["color"], "font": self.hfont,
             "spacing": h.get("spacing_pt")}
        if h["case"] == "upper":
            f["caps"] = True
        elif h["case"] == "smallcaps":
            f["smallcaps"] = True
        before = 8 if first else 13
        if h.get("pill"):
            f["color"] = "FFFFFF"
            f["shade"] = self.s["accent"]
            text = "\u2002" + text + "\u2002"
        left = self.content_left_cm() if self.s.get("heading_in_col") else None
        para = self.p([(text, f)], before=before, after=6 if h.get("pill") else 5, keep_next=True,
                      align=h.get("align"), left_cm=left)
        col = self.s["accent"] if h["color"] == "accent" else self.s["text"]
        if h["rule"] == "bottom":
            _border(para, {"bottom": 8}, self.s["accent"], space=2)
        elif h["rule"] == "bottom-thin":
            _border(para, {"bottom": 4}, self.s["accent"], space=2)
        elif h["rule"] == "both":
            _border(para, {"top": 4, "bottom": 4}, col, space=3)
        if h.get("bar"):
            _border(para, {"left": 24}, self.s["accent"])
            para.paragraph_format.left_indent = Cm(0.22)
        return para


# ---------------------------------------------------------------- schema ordering
# Word rejects files whose pPr/rPr children are out of schema order ("unreadable
# content"), even though LibreOffice tolerates it. Normalise before saving.
_PPR = ("pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
        "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct "
        "topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind "
        "contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment "
        "textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange").split()
_RPR = ("rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint "
        "noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u "
        "effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath").split()
_BDR = "top left bottom right between bar".split()
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _reorder(el, order):
    rank = {(_W + n): i for i, n in enumerate(order)}
    kids = list(el)
    kids_sorted = sorted(kids, key=lambda k: rank.get(k.tag, 1000))  # w14:* etc. last
    if kids_sorted != kids:
        for k in kids:
            el.remove(k)
        for k in kids_sorted:
            el.append(k)


def _normalise(root):
    for el in root.iter(_W + "pPr"):
        _reorder(el, _PPR)
    for el in root.iter(_W + "rPr"):
        _reorder(el, _RPR)
    for el in root.iter(_W + "pBdr"):
        _reorder(el, _BDR)


# ---------------------------------------------------------------- content blocks
def _date_range(start, end, h):
    end = h["present"] if (end or "").lower() in ("present", "heute", "current", "now") else end
    if start and end:
        return f"{start} – {end}"
    return start or end or ""


def _contact_items(cv):
    c = cv.get("contact", {})
    order = ["location", "phone", "email", "linkedin", "github", "website"]
    return [c[k] for k in order if c.get(k)]


def _header(D: Doc, cv, photo):
    s = D.s
    items = _contact_items(cv)
    name, headline = cv.get("name", ""), cv.get("headline", "")
    w = D.text_w_cm
    if s["header"] == "band":
        # Deterministic geometry: exact line heights so the background band
        # (a text-free picture behind the text) always fits the header block.
        pad_top, pad_bot = 22, 16                      # pt
        name_lh = s["name_pt"] * 1.2
        head_lh = s["headline_pt"] * 1.35
        c_size = s["body_pt"] - 0.5
        c_lh = c_size * 1.45
        contact = "  |  ".join(items)
        ind_cm = 3.4 if (photo or s.get("monogram")) and s.get("band_badge") else 0
        chars_per_line = int((D.text_w_cm - ind_cm) * 11.5 * 10 / c_size / 1.0 / 2.2)
        head_lines = 1
        if headline:
            est_w = len(headline) * s["headline_pt"] * 0.52 / 72 * 2.54
            head_lines = max(1, -(-est_w // (D.text_w_cm - ind_cm)))
        c_lines = max(1, -(-len(contact) // max(chars_per_line, 40))) if items else 0
        total = pad_top + name_lh + (head_lh * head_lines + 2 if headline else 0) + (c_lines * c_lh + 6 if items else 0) + pad_bot
        bt, bs_, bsep = s.get("band_text", "FFFFFF"), s.get("band_sub", "E6F2F2"), s.get("band_sep", "9FD3D3")
        band_color = s.get("band_color", s["accent"])
        ind = 3.4 if (photo or s.get("monogram")) and s.get("band_badge") else None
        para = D.p([(name, {"size": s["name_pt"], "bold": True, "color": bt, "font": D.hfont})],
                   before=pad_top, left_cm=ind)
        para.paragraph_format.line_spacing = Pt(name_lh)
        if headline:
            para = D.p([(headline, {"size": s["headline_pt"], "color": bs_})], before=2, left_cm=ind)
            para.paragraph_format.line_spacing = Pt(head_lh)
        if items:
            runs = []
            for i, it in enumerate(items):
                if i:
                    runs.append(("  |  ", {"color": bsep, "size": c_size}))
                runs.append((it, {"color": bt, "size": c_size}))
            para = D.p(runs, before=6, left_cm=ind)
            para.paragraph_format.line_spacing = Pt(c_lh)
        band_h = total / 72 * 2.54
        if ind:
            t, r_, b_, l_ = s["margins_cm"]
            badge = _badge_png(D, cv, photo, 2.6)
            _anchor_picture(D, badge, 2.6, 2.6, x_cm=l_, y_cm=max(0.3, (band_h - 2.6) / 2),
                            rel="page")
        _anchor_picture(D, _solid_png(band_color), A4_W_CM, band_h, x_cm=0, y_cm=0,
                        rel="page", behind=True)
        D.p([], before=pad_bot)
        return

    align = "center" if s["header"] == "center" else None
    name_f = {"size": s["name_pt"], "bold": s.get("name_bold", True),
              "color": s.get("name_color", "text"), "font": D.hfont}
    right_cm = None  # photo floats with square wrap; text flows around it
    hl = None
    if s.get("heading_in_col"):
        hl = D.content_left_cm()
    elif s.get("monogram") == "left":
        hl = 2.45
    D.p([(name, name_f)], align=align, after=1, right_cm=right_cm, line=1.0, left_cm=hl)
    if headline:
        D.p([(headline, {"size": s["headline_pt"], "color": "accent" if s["header"] != "center" else "muted",
                         "italic": s["header"] == "center" or s.get("headline_italic", False),
                         "font": s.get("headline_font_" + ("pdf" if D.font == s["font_pdf"] else "docx"), D.font)})],
            align=align, after=5, right_cm=right_cm, left_cm=hl)
    if items:
        runs = []
        sep = "  ·  " if s["header"] == "center" else "  |  "
        for i, it in enumerate(items):
            if i:
                runs.append((sep, {"color": "accent"}))
            runs.append((it, {"color": "muted", "size": s["body_pt"] - (1 if align else 0.5)}))
        para = D.p(runs, align=align, after=2, right_cm=right_cm, left_cm=hl)
        if s["header"] == "split":
            _border(para, {"bottom": 6}, s["accent"], space=6)
    if s.get("monogram") == "left":
        _anchor_picture(D, _badge_png(D, cv, None, 1.9, square=True), 1.9, 1.9, x_cm=0, y_cm=0.05, rel="margin")
    if s.get("panel"):
        # photo or monogram sits in the tinted sidebar, not in the text flow
        pw = s["panel"]["w_cm"]
        _anchor_picture(D, _badge_png(D, cv, photo, 3.0), 3.0, 3.0, x_cm=(pw - 3.0) / 2,
                        y_cm=1.3, rel="page")
        photo = None
    if photo:
        _add_floating_photo(D, photo)
        # make sure the first section rule starts below the 4 cm photo
        c_size = s["body_pt"] - 0.5
        chars = int((D.text_w_cm - 3.5) * 11.5 * 10 / c_size / 2.2)
        c_lines = max(1, -(-len("  |  ".join(items)) // max(chars, 30))) if items else 0
        est = s["name_pt"] * 1.25 + 1 + (s["headline_pt"] * 1.3 + 5 if headline else 0) \
            + c_lines * s["body_pt"] * 1.3 + 2
        from PIL import Image as _Im
        with _Im.open(photo) as _im:
            ratio = 1.0 if s.get("photo_shape") == "circle" else _im.height / _im.width
        gap = 3.0 * ratio / 2.54 * 72 - est - 10  # real photo height (pt) minus text height
        if gap > 0:
            D.p([], before=gap)


def _anchor_picture(D: Doc, img_path, width_cm, height_cm=None, x_cm=None, y_cm=0.0,
                    rel="margin", align_right=False, behind=False, wrap=False):
    """Insert a picture as a floating anchor at the very top of the body.
    Pictures carry no text, so ATS parsers skip them cleanly; the real text stays
    in ordinary paragraphs."""
    para = D.p([])
    run = para.add_run()
    pic = run.add_picture(str(img_path), width=Cm(width_cm),
                          height=Cm(height_cm) if height_cm else None)
    inline = run._element.find(".//" + qn("wp:inline"))
    anchor = OxmlElement("wp:anchor")
    for k, v in dict(distT="0", distB="0", distL="0", distR="0", simplePos="0",
                     relativeHeight="251659264" if not behind else "251650000",
                     behindDoc="1" if behind else "0", locked="0", layoutInCell="1",
                     allowOverlap="1").items():
        anchor.set(k, v)
    sp = OxmlElement("wp:simplePos"); sp.set("x", "0"); sp.set("y", "0"); anchor.append(sp)
    ph = OxmlElement("wp:positionH"); ph.set("relativeFrom", rel)
    if align_right:
        al = OxmlElement("wp:align"); al.text = "right"; ph.append(al)
    else:
        off = OxmlElement("wp:posOffset"); off.text = str(int((x_cm or 0) * 360000)); ph.append(off)
    anchor.append(ph)
    pv = OxmlElement("wp:positionV"); pv.set("relativeFrom", rel)
    off = OxmlElement("wp:posOffset"); off.text = str(int(y_cm * 360000)); pv.append(off); anchor.append(pv)
    for child in list(inline):
        if child.tag == qn("wp:extent"):
            anchor.append(copy.deepcopy(child))
            ee = OxmlElement("wp:effectExtent")
            for a in ("l", "t", "r", "b"):
                ee.set(a, "0")
            anchor.append(ee)
            if wrap:
                ws = OxmlElement("wp:wrapSquare"); ws.set("wrapText", "bothSides")
                anchor.append(ws)
            else:
                anchor.append(OxmlElement("wp:wrapNone"))
        elif child.tag in (qn("wp:docPr"), qn("wp:cNvGraphicFramePr"), qn("a:graphic")):
            anchor.append(copy.deepcopy(child))
    inline.getparent().replace(inline, anchor)
    body = D.d.element.body
    body.remove(para._p)
    body.insert(0, para._p)
    para.paragraph_format.line_spacing = Pt(1)
    return para


def _solid_png(hex_color):
    from PIL import Image
    path = Path(tempfile.gettempdir()) / f"curator_{hex_color}.png"
    if not path.exists():
        Image.new("RGB", (8, 8), "#" + hex_color).save(path)
    return path


def _initials(name):
    parts = [p for p in name.replace("-", " ").split() if p[:1].isalpha()]
    return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper() if parts else "CV"


def _badge_png(D: Doc, cv, photo, size_cm, square=False):
    """Circular photo, or a monogram badge when there is no photo. Pure picture:
    the name is still real text in the body, so ATS loses nothing."""
    from PIL import Image, ImageDraw, ImageFont, ImageOps
    s = D.s
    px = 600
    mask = Image.new("L", (px, px), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, px - 1, px - 1), radius=60 if square else px // 2, fill=255)
    if photo:
        im = ImageOps.fit(Image.open(photo).convert("RGB"), (px, px), centering=(0.5, 0.35))
    else:
        im = Image.new("RGB", (px, px), "#" + s.get("badge_bg", s["accent"]))
        d = ImageDraw.Draw(im)
        fpath = str(FONT_DIR / s.get("badge_font", "Poppins-Medium.ttf"))
        f = None
        for cand in (fpath, "arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"):   # Windows/Linux fallbacks
            try:
                f = ImageFont.truetype(cand, int(px * 0.38)); break
            except OSError:
                continue
        f = f or ImageFont.load_default(size=int(px * 0.38))
        txt = _initials(cv.get("name", ""))
        bb = d.textbbox((0, 0), txt, font=f)
        d.text(((px - (bb[2] - bb[0])) / 2 - bb[0], (px - (bb[3] - bb[1])) / 2 - bb[1]), txt,
               font=f, fill="#" + s.get("badge_fg", "FFFFFF"))
    out = Image.new("RGBA", (px, px), (0, 0, 0, 0))
    out.paste(im, (0, 0), mask)
    if s.get("badge_ring") and not square:
        ImageDraw.Draw(out).ellipse((6, 6, px - 7, px - 7), outline="#" + s["badge_ring"], width=14)
    path = Path(tempfile.gettempdir()) / f"curator_badge_{abs(hash((str(photo), cv.get('name'), s['label'], square)))}.png"
    out.save(path)
    return path


def _header_picture(D: Doc, img_path, w_cm, h_cm, x_cm, y_cm):
    """Background graphic placed in the page header part so it repeats on every page.
    It is a picture with no text; header text stays empty, so parsers see nothing."""
    hdr = D.d.sections[0].header
    para = hdr.paragraphs[0]
    run = para.add_run()
    run.add_picture(str(img_path), width=Cm(w_cm), height=Cm(h_cm))
    _inline_to_anchor(run._element, x_cm, y_cm, behind=True)


def _inline_to_anchor(run_el, x_cm, y_cm, behind=True):
    inline = run_el.find(".//" + qn("wp:inline"))
    anchor = OxmlElement("wp:anchor")
    for k, v in dict(distT="0", distB="0", distL="0", distR="0", simplePos="0",
                     relativeHeight="251640000", behindDoc="1" if behind else "0", locked="0",
                     layoutInCell="1", allowOverlap="1").items():
        anchor.set(k, v)
    sp = OxmlElement("wp:simplePos"); sp.set("x", "0"); sp.set("y", "0"); anchor.append(sp)
    for tag, val in (("wp:positionH", x_cm), ("wp:positionV", y_cm)):
        ph = OxmlElement(tag); ph.set("relativeFrom", "page")
        off = OxmlElement("wp:posOffset"); off.text = str(int(val * 360000)); ph.append(off)
        anchor.append(ph)
    for child in list(inline):
        if child.tag == qn("wp:extent"):
            anchor.append(copy.deepcopy(child))
            ee = OxmlElement("wp:effectExtent")
            for a in ("l", "t", "r", "b"):
                ee.set(a, "0")
            anchor.append(ee)
            anchor.append(OxmlElement("wp:wrapNone"))
        elif child.tag in (qn("wp:docPr"), qn("wp:cNvGraphicFramePr"), qn("a:graphic")):
            anchor.append(copy.deepcopy(child))
    inline.getparent().replace(inline, anchor)


def _add_floating_photo(D: Doc, photo_path):
    if D.s.get("photo_shape") == "circle":
        photo_path = _badge_png(D, {}, photo_path, 3.0)
    para = _anchor_picture(D, photo_path, 3.0, rel="margin", align_right=True, wrap=True)
    a = para._p.find(".//" + qn("wp:anchor"))
    a.set("distL", str(int(0.5 * 360000))); a.set("distB", str(int(0.3 * 360000)))


def _personal(D: Doc, cv):
    per = cv.get("personal") or {}
    if not per:
        return
    D.heading("personal")
    col = 3.6 if D.s["layout"] != "tabular" else D.s["date_col_cm"]
    for k, v in per.items():
        label = D.h.get(k, k)
        D.p([(f"{label}", {"color": "muted"}), ("\t", {}), (str(v), {})],
            tabs=[(col, "left")], left_cm=col, first_cm=-col, after=1.5)


def _entry(D: Doc, title, org, loc, start, end, bullets, desc=None):
    s = D.s
    n0 = len(D.d.element.body.findall(qn("w:p")))
    _entry_inner(D, title, org, loc, start, end, bullets, desc)
    if s.get("timeline"):
        # continuous rule down the left of each entry; every paragraph starts at the same x
        # (bullets use a hanging indent from that x) so the rule stays straight
        tl = s["timeline"]
        body = D.d.element.body.findall(qn("w:p"))[n0:]
        for p_el in body:
            if p_el.tag != qn("w:p"):
                continue
            from docx.text.paragraph import Paragraph
            para = Paragraph(p_el, D.d._body)
            is_b = p_el.find(".//" + qn("w:numPr")) is not None
            pf = para.paragraph_format
            if is_b:
                pf.left_indent = Cm(tl["indent_cm"] + 0.42)
                pf.first_line_indent = Cm(-0.42)
            else:
                pf.left_indent = Cm(tl["indent_cm"])
            _border(para, {"left": tl.get("sz", 12)}, tl["color"])
            pBdr = p_el.find(qn("w:pPr") + "/" + qn("w:pBdr"))
            pBdr.find(qn("w:left")).set(qn("w:space"), str(tl.get("space_pt", 8)))
        # marker: first (title) line gets the accent colour, thicker
        first = next((e for e in body if e.tag == qn("w:p")), None)
        if first is not None:
            lb = first.find(qn("w:pPr") + "/" + qn("w:pBdr") + "/" + qn("w:left"))
            lb.set(qn("w:color"), s["accent"]); lb.set(qn("w:sz"), str(tl.get("marker_sz", 24)))


def _entry_inner(D: Doc, title, org, loc, start, end, bullets, desc=None):
    s = D.s
    dr = _date_range(start, end, D.h)
    orgline = ", ".join(x for x in (org, loc) if x)
    w = D.text_w_cm
    if s["layout"] == "tabular":
        col = s["date_col_cm"]
        D.p([(dr, {"color": "muted", "size": s["body_pt"] - 0.5}), ("\t", {}),
             (title, {"bold": True})], tabs=[(col, "left")], left_cm=col, first_cm=-col,
            before=6, after=0.5, keep_next=True)
        if s.get("panel") and (org or loc):
            D.p([(loc or "", {"color": "muted", "size": s["body_pt"] - 0.5}), ("\t", {}),
                 (org or "", {"color": "accent"})], tabs=[(col, "left")], left_cm=col,
                first_cm=-col, after=2.5, keep_next=bool(bullets))
        elif orgline:
            D.p([(orgline, {"color": "accent", "italic": False})], left_cm=col, after=2.5, keep_next=bool(bullets))
    elif s["layout"] == "classic":
        long = len(title) + len(org or "") + len(dr) > D.text_w_cm * 5.2 * 10 / s["body_pt"]
        if long:   # company moves to the second line so the date never wraps
            D.p([(title, {"bold": True}), ("\t", {}),
                 (dr, {"color": "muted", "size": s["body_pt"] - 0.5})],
                tabs=[(w, "right")], before=6, after=0.5, keep_next=True)
            D.p([(org or "", {"color": "accent"}), (", " + loc if loc else "", {"italic": True, "color": "muted"})],
                after=2.5, keep_next=True)
            loc = ""
        else:
            D.p([(title, {"bold": True}), (", " if org else "", {}),
                 (org or "", {"bold": False, "color": "accent"}), ("\t", {}),
                 (dr, {"color": "muted", "size": s["body_pt"] - 0.5})],
                tabs=[(w, "right")], before=6, after=0.5, keep_next=True)
        if loc:
            D.p([(loc, {"italic": True, "color": "muted", "size": s["body_pt"] - 0.5})], after=2.5, keep_next=True)
    else:  # stacked (resume.io-like): Title, Company / dates · location in muted small
        D.p([(title, {"bold": True, "size": s["body_pt"] + 0.5}), (", " if org else "", {}),
             (org or "", {"bold": False})], before=7, after=0.5, keep_next=True)
        meta = "  ·  ".join(x for x in (dr, loc) if x)
        if meta:
            D.p([(meta, {"color": "muted", "size": s["body_pt"] - 0.8, "caps": False})], after=2.5, keep_next=True)
    if desc:
        D.p([(desc, {})], left_cm=D.content_left_cm(), after=2)
    for b in bullets or []:
        D.bullet(b)


def _section_summary(D, cv):
    if cv.get("summary"):
        D.heading("summary")
        left = None if D.s.get("full_width_intro") else D.content_left_cm()
        D.p([(cv["summary"], {})], left_cm=left, after=2, line=1.15)


def _section_competencies(D, cv):
    comps = cv.get("competencies") or []
    if not comps:
        return
    D.heading("competencies")
    col = 0 if D.s.get("full_width_intro") else D.content_left_cm()
    if D.s.get("chips"):
        ch = D.s["chips"]
        runs = []
        for i, it in enumerate(comps):
            if i:
                runs.append(("  ", {}))
            runs.append(("\u2009" + it + "\u2009", {"shade": ch["bg"], "color": ch.get("fg", "text"),
                                                   "size": D.s["body_pt"] - 0.6}))
        para = D.p(runs, left_cm=col or None, after=3)
        para.paragraph_format.line_spacing = 1.55
        return
    runs = []
    for i, it in enumerate(comps):
        if i:
            runs.append(("  |  ", {"color": "accent"}))
        runs.append((it, {}))
    D.p(runs, left_cm=col or None, after=2, line=1.25)


def _section_experience(D, cv):
    if cv.get("experience"):
        D.heading("experience")
        for e in cv["experience"]:
            _entry(D, e.get("title", ""), e.get("company", ""), e.get("location", ""),
                   e.get("start", ""), e.get("end", ""), e.get("bullets", []), e.get("description"))


def _section_education(D, cv):
    if cv.get("education"):
        D.heading("education")
        for e in cv["education"]:
            _entry(D, e.get("degree", ""), e.get("institution", ""), e.get("location", ""),
                   e.get("start", ""), e.get("end", ""), e.get("details", []))


def _section_skills(D, cv):
    groups = cv.get("skills") or []
    if not groups:
        return
    D.heading("skills")
    col = D.content_left_cm()
    if D.s.get("chips"):
        ch = D.s["chips"]
        for g in groups:
            if g.get("category"):
                D.p([(g["category"], {"bold": True, "size": D.s["body_pt"] - 0.3,
                                      "color": ch.get("label", "text")})],
                    left_cm=col or None, before=3, after=2, keep_next=True)
            runs = []
            for i, it in enumerate(g.get("items", [])):
                if i:
                    runs.append(("  ", {}))
                runs.append(("\u2009" + it + "\u2009", {"shade": ch["bg"], "color": ch.get("fg", "text"),
                                                       "size": D.s["body_pt"] - 0.6}))
            para = D.p(runs, left_cm=col or None, after=3)
            para.paragraph_format.line_spacing = 1.55
        return
    for g in groups:
        items = ", ".join(g.get("items", []))
        if g.get("category"):
            if D.s["layout"] == "tabular":
                D.p([(g["category"], {"color": "muted"}), ("\t", {}), (items, {})],
                    tabs=[(col, "left")], left_cm=col, first_cm=-col, after=2)
            else:
                D.p([(g["category"] + ": ", {"bold": True}), (items, {})], after=2)
        else:
            D.p([(items, {})], left_cm=col, after=2)


def _section_languages(D, cv):
    langs = cv.get("languages") or []
    if not langs:
        return
    D.heading("languages")
    col = D.content_left_cm()
    if D.s["layout"] == "tabular":
        for l in langs:
            D.p([(l.get("language", ""), {"color": "muted"}), ("\t", {}), (l.get("level", ""), {})],
                tabs=[(col, "left")], left_cm=col, first_cm=-col, after=1.5)
    else:
        runs = []
        for i, l in enumerate(langs):
            if i:
                runs.append(("   |   ", {"color": "accent"}))
            runs.append((l.get("language", ""), {"bold": True}))
            if l.get("level"):
                runs.append((f" ({l['level']})", {}))
        D.p(runs, after=2)


def _section_certs(D, cv):
    certs = cv.get("certifications") or []
    if not certs:
        return
    D.heading("certifications")
    for c in certs:
        if isinstance(c, str):
            c = {"name": c}
        txt = c.get("name", "")
        extra = ", ".join(x for x in (c.get("issuer"), c.get("date")) if x)
        if D.s["layout"] == "tabular":
            col = D.s["date_col_cm"]
            D.p([(c.get("date", ""), {"color": "muted"}), ("\t", {}), (txt, {}),
                 ((", " + c["issuer"]) if c.get("issuer") else "", {"color": "muted"})],
                tabs=[(col, "left")], left_cm=col, first_cm=-col, after=1.5)
        else:
            D.bullet(txt + (f" ({extra})" if extra else ""))


def _section_projects(D, cv):
    pr = cv.get("projects") or []
    if not pr:
        return
    D.heading("projects")
    for p in pr:
        _entry(D, p.get("name", ""), p.get("subtitle", ""), "", p.get("start", ""),
               p.get("end", ""), p.get("bullets", []), p.get("description"))


def _section_extra(D, cv):
    for sec in cv.get("extra_sections") or []:
        if not sec.get("items"):
            continue
        D.heading(sec["title"])
        for it in sec["items"]:
            D.bullet(it)


SECTION_FUNCS = {"summary": _section_summary, "competencies": _section_competencies,
                 "experience": _section_experience,
                 "education": _section_education, "skills": _section_skills,
                 "languages": _section_languages, "certifications": _section_certs,
                 "projects": _section_projects, "extra": _section_extra}


def build_docx(cv: dict, template: str, out_path: Path, pdf_mode=False, lang="en",
               photo=None, signature_place=None, keep_source_order=False):
    spec = TEMPLATES[template]
    D = Doc(spec, pdf_mode, lang)
    use_photo = photo if (photo and spec.get("photo")) else None
    _header(D, cv, use_photo)
    if spec.get("stripe"):
        _header_picture(D, _solid_png(spec["accent"]), 0.55, A4_H_CM, 0, 0)
    if spec.get("panel"):
        _header_picture(D, _solid_png(spec["panel"]["color"]), spec["panel"]["w_cm"], A4_H_CM, 0, 0)
        if spec["panel"].get("edge"):
            _header_picture(D, _solid_png(spec["panel"]["edge"]), 0.12, A4_H_CM, spec["panel"]["w_cm"], 0)
    if spec.get("top_rule"):
        _header_picture(D, _solid_png(spec["accent"]), A4_W_CM, spec["top_rule"], 0, 0)
    _personal(D, cv)
    order = spec["order"]
    if keep_source_order and cv.get("section_order"):
        src = [k for k in cv["section_order"] if k in SECTION_FUNCS]
        order = src + [k for k in order if k not in src]
    first = True
    for key in order:
        SECTION_FUNCS[key](D, cv)
    if signature_place:
        date = _dt.date.today().strftime("%d.%m.%Y")
        D.p([(f"{signature_place}, {date}", {"color": "muted"})], before=18, keep_next=True)
        D.p([(cv.get("name", ""), {"italic": True})], before=14)
    # metadata: helps some parsers and looks professional in file properties
    cp = D.d.core_properties
    cp.title = f"{cv.get('name', '')} – {'Lebenslauf' if lang == 'de' else 'CV'}"
    cp.author = cv.get("name", "")
    cp.subject = cv.get("headline", "")
    cp.keywords = ", ".join(i for g in cv.get("skills", []) for i in g.get("items", []))[:250]
    cp.comments = ""
    _normalise(D.d.element)
    _normalise(D.d.styles.element)
    D.d.save(str(out_path))
    return out_path


FONT_DIR = Path(__file__).resolve().parent.parent / "fonts"


def _soffice() -> str:
    """Find LibreOffice on Linux, macOS or Windows."""
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(cand).exists():
            return cand
    raise RuntimeError("LibreOffice not found: install it and make sure 'soffice' is on PATH.")


def docx_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    """LibreOffice export: tagged PDF/A-2b (fonts embedded, Unicode text map).
    The bundled fonts are loaded through a private LibreOffice profile, so every
    install renders identically and nothing is installed system-wide."""
    filt = ('pdf:writer_pdf_Export:{"UseTaggedPDF":{"type":"boolean","value":"true"},'
            '"SelectPdfVersion":{"type":"long","value":"2"},'
            '"ExportBookmarks":{"type":"boolean","value":"false"}}')
    profile = Path(tempfile.mkdtemp(prefix="lo_"))
    if FONT_DIR.is_dir():
        dest = profile / "user" / "fonts"
        dest.mkdir(parents=True, exist_ok=True)
        for f in FONT_DIR.iterdir():
            if f.suffix.lower() in (".ttf", ".otf"):
                shutil.copy2(f, dest / f.name)
    subprocess.run([_soffice(), f"-env:UserInstallation={profile.as_uri()}", "--headless",
                    "--convert-to", filt, "--outdir", str(out_dir), str(docx_path)],
                   capture_output=True, timeout=180)
    shutil.rmtree(profile, ignore_errors=True)
    return out_dir / (docx_path.stem + ".pdf")


def render(cv: dict, template: str, out_dir: Path, basename: str, lang="en", photo=None,
           signature_place=None, keep_source_order=False):
    out_dir.mkdir(parents=True, exist_ok=True)
    docx_out = out_dir / f"{basename}_{template}.docx"
    build_docx(cv, template, docx_out, pdf_mode=False, lang=lang, photo=photo,
               signature_place=signature_place, keep_source_order=keep_source_order)
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / f"{basename}_{template}.docx"
        build_docx(cv, template, tmp, pdf_mode=True, lang=lang, photo=photo,
                   signature_place=signature_place, keep_source_order=keep_source_order)
        pdf = docx_to_pdf(tmp, out_dir)
    return docx_out, pdf
