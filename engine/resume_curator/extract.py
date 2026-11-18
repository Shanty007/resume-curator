"""Text extraction from DOCX / PDF / TXT — text layer only, never OCR.

If a PDF has no usable text layer (scanned / exported-as-image, e.g. some Canva
exports), we refuse instead of OCR-ing: OCR output is exactly the kind of noisy
text that loses ATS ranking. The user is told to supply a text-based file.
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class NoTextLayerError(Exception):
    pass


def _docx_text(path: Path) -> str:
    """Read DOCX body in document order: paragraphs, table cells, and text
    boxes (txbxContent) — text boxes are where Canva/Word templates hide stuff."""
    import docx
    from docx.oxml.ns import qn

    d = docx.Document(str(path))
    lines: list[str] = []
    seen_txbx = set()

    def para_text(p_el) -> str:
        parts = []
        for node in p_el.iter():
            if node.tag == qn("w:t") and node.text:
                # skip text that lives inside a nested text box; handled separately
                parts.append(node.text)
            elif node.tag in (qn("w:tab"),):
                parts.append("\t")
            elif node.tag in (qn("w:br"), qn("w:cr")):
                parts.append("\n")
        return "".join(parts)

    # styles that carry numbering (e.g. Word's built-in "List Bullet")
    list_styles = set()
    for st in d.styles:
        try:
            el = st.element
        except Exception:
            continue
        if el.find(".//" + qn("w:numPr")) is not None:
            list_styles.add(el.get(qn("w:styleId")))

    def is_list_para(p_el):
        if p_el.find(".//" + qn("w:numPr")) is not None:
            return True
        ps = p_el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        return ps is not None and ps.get(qn("w:val")) in list_styles

    body = d.element.body
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            # numbered/bulleted paragraphs -> prefix a bullet so structure survives
            is_list = is_list_para(el)
            t = para_text(el)
            for sub in t.split("\n"):
                if sub.strip():
                    lines.append(("• " if is_list else "") + sub.strip())
        elif el.tag == qn("w:tbl"):
            for row in el.iter(qn("w:tr")):
                for cell in row.iter(qn("w:tc")):
                    for p in cell.iter(qn("w:p")):
                        is_list = is_list_para(p)
                        t = para_text(p).strip()
                        if t:
                            lines.append(("• " if is_list else "") + t)
    # header text (some people put contact info there)
    hdr = []
    for s in d.sections:
        for p in s.header.paragraphs:
            if p.text.strip():
                hdr.append(p.text.strip())
    text = "\n".join(hdr + lines)
    return _clean(text)


def _lines_from_words(words, tol=3.0):
    words = sorted(words, key=lambda w: (round(w["top"] / tol), w["x0"]))
    lines, cur, cur_top = [], [], None
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if cur_top is None or abs(w["top"] - cur_top) <= tol:
            cur.append(w); cur_top = w["top"] if cur_top is None else cur_top
        else:
            lines.append(cur); cur, cur_top = [w], w["top"]
    if cur:
        lines.append(cur)
    out = []
    for ln in lines:
        ln.sort(key=lambda w: w["x0"])
        seg = [ln[0]]
        for a, b in zip(ln, ln[1:]):
            if b["x0"] - a["x1"] > 40:          # big horizontal gap = separate line
                out.append(" ".join(x["text"] for x in seg)); seg = []
            seg.append(b)
        out.append(" ".join(x["text"] for x in seg))
    return out


def _find_gutter(words, width):
    """x position of a vertical gutter that splits the page into two text columns."""
    if len(words) < 30:
        return None
    best = None
    for x in range(int(width * 0.18), int(width * 0.78), 2):
        crossing = sum(1 for w in words if w["x0"] < x < w["x1"])
        left = sum(1 for w in words if w["x1"] <= x)
        right = len(words) - left - crossing
        if left < 0.12 * len(words) or right < 0.12 * len(words):
            continue
        # the gutter must be empty for most of the page height
        if crossing <= max(1, 0.02 * len(words)):
            gap_ok = not any(x - 6 < w["x1"] <= x and False for w in words)
            score = crossing
            if best is None or score < best[0]:
                best = (score, x)
    return best[1] if best else None


def _pdf_text_columns(path: Path) -> str:
    """Column-aware reading order: sidebars (Canva / builder exports) are read as a
    block instead of being interleaved line by line with the main column."""
    import pdfplumber
    out = []
    with pdfplumber.open(str(path)) as pdf:
        for pg in pdf.pages:
            words = pg.extract_words(extra_attrs=["size"], keep_blank_chars=False, use_text_flow=False)
            if not words:
                continue
            g = _find_gutter(words, pg.width)
            if g is None:
                out += _lines_from_words(words)
                continue
            left = [w for w in words if w["x1"] <= g]
            right = [w for w in words if w["x0"] >= g]
            full = [w for w in words if w["x0"] < g < w["x1"]]
            cols = [left, right]
            # the column holding the biggest type (the name) goes first
            cols.sort(key=lambda c: -max((w["size"] for w in c), default=0))
            out += _lines_from_words(full) + _lines_from_words(cols[0]) + _lines_from_words(cols[1])
    return "\n".join(out)


def _pdf_text(path: Path) -> str:
    text = ""
    try:
        text = _pdf_text_columns(path)
    except Exception:
        text = ""
    if len(re.sub(r"\s", "", text)) < 80:
        try:
            text = subprocess.run(["pdftotext", "-layout", "-enc", "UTF-8", str(path), "-"],
                                  capture_output=True, text=True, timeout=60).stdout
            text = "\n".join(p for ln in text.splitlines() for p in re.split(r"\s{6,}", ln.strip()) if p)
        except Exception:
            pass
    if len(re.sub(r"\s", "", text)) < 80:
        raise NoTextLayerError(
            f"{path.name} has no text layer (scanned or image-only PDF). "
            "OCR is deliberately not used because it degrades ATS parsing. "
            "Please supply the original DOCX or a text-based PDF."
        )
    if text.count("(cid:") > 20:
        raise NoTextLayerError(
            f"{path.name} uses fonts without a Unicode map (text extracts as '(cid:..)'). "
            "ATS systems will read garbage from it too. Please supply the DOCX."
        )
    return _clean(text)


LIGATURES = {"ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi", "ﬄ": "ffl"}


def _clean(text: str) -> str:
    for k, v in LIGATURES.items():
        text = text.replace(k, v)
    # unmapped glyph at line start is almost always a bullet symbol
    text = re.sub(r"(?m)^\s*\(cid:\d+\)\s*", "• ", text)
    text = re.sub(r"\(cid:\d+\)", "", text)
    # private-use bullets from Symbol/Wingdings fonts -> real bullet
    text = re.sub(r"[-]", "•", text)
    text = text.replace(" ", " ").replace("​", "")
    lines = [re.sub(r"[ \t]+", " ", l).strip() for l in text.splitlines()]
    out, blank = [], 0
    for l in lines:
        if not l:
            blank += 1
            if blank > 1:
                continue
        else:
            blank = 0
        out.append(l)
    return "\n".join(out).strip()


def extract_text(path: str | Path) -> str:
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".docx":
        return _docx_text(path)
    if ext == ".pdf":
        return _pdf_text(path)
    if ext in (".txt", ".md"):
        return _clean(path.read_text(encoding="utf-8", errors="replace"))
    if ext in (".doc", ".odt", ".rtf"):
        # convert with LibreOffice to docx first (still a text conversion, no OCR)
        outdir = path.parent
        subprocess.run(["soffice", "--headless", "--convert-to", "docx", "--outdir",
                        str(outdir), str(path)], capture_output=True, timeout=120)
        return _docx_text(outdir / (path.stem + ".docx"))
    raise ValueError(f"Unsupported input type: {ext}")
