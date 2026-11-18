"""ATS self-check: read the generated files back the way a parser does, then verify
that every fact survived, in the right order, as clean Unicode.

PDF: `pdftotext -raw` gives the content-stream order, which is what naive parsers
see. pdfplumber gives the layout order, which is what smarter parsers see. Both
must pass.
DOCX: structural lint for tables, text boxes, header/footer text and multi-column
sections, plus a coverage check on the body text.
"""
from __future__ import annotations

import re
import subprocess
import zipfile
from pathlib import Path


def _norm(t: str) -> str:
    t = t.replace("–", "-").replace("—", "-").replace("‑", "-")
    t = re.sub(r"-\s*\n\s*", "-", t)          # line-broken hyphenated tokens
    t = re.sub(r"\s+", " ", t)
    return t.lower().strip()


def _tokens(cv):
    """Ordered list of (label, text) that must be findable."""
    out = [("name", cv.get("name", ""))]
    c = cv.get("contact", {})
    for k in ("email", "phone", "linkedin"):
        if c.get(k):
            out.append((k, c[k]))
    for e in cv.get("experience", []):
        out += [("job title", e.get("title", "")), ("company", e.get("company", ""))]
        out += [("bullet", b) for b in e.get("bullets", [])]
    for e in cv.get("education", []):
        out += [("degree", e.get("degree", "")), ("institution", e.get("institution", ""))]
    for g in cv.get("skills", []):
        out += [("skill", i) for i in g.get("items", [])]
    for l in cv.get("languages", []):
        out.append(("language", l.get("language", "")))
    return [(a, b) for a, b in out if b and b.strip()]


def _coverage(text, toks):
    nt = _norm(text)
    missing = []
    for label, t in toks:
        probe = _norm(t)
        if label == "bullet":
            probe = " ".join(probe.split()[:8])
        if probe not in nt:
            missing.append(f"{label}: {t[:60]}")
    return missing


def _order_ok(text, cv):
    """Name near the top, then jobs and degrees findable in sequence."""
    nt = _norm(text)
    name = _norm(cv.get("name", ""))
    if name and nt.find(name) > 300:
        return False
    pos = 0
    for e in cv.get("experience", []) + cv.get("education", []):
        t = _norm(e.get("title") or e.get("degree") or "")
        if not t:
            continue
        i = nt.find(t, pos)
        if i < 0:
            return False
        pos = i + len(t)
    return True


def _bad_chars(text):
    issues = []
    if re.search(r"[-]", text):
        issues.append("private-use glyphs (icon/Symbol font bullets)")
    if re.search(r"[ﬀ-ﬆ]", text):
        issues.append("ligature glyphs (ﬁ/ﬂ) that break keyword matching")
    if "(cid:" in text or "�" in text:
        issues.append("unmapped glyphs (cid/replacement chars)")
    if re.search(r"(?:\b[A-Za-z] ){5,}", text):
        issues.append("letter-spaced words split into single letters")
    return issues


def check_pdf(pdf: Path, cv: dict) -> dict:
    raw = subprocess.run(["pdftotext", "-raw", "-enc", "UTF-8", str(pdf), "-"],
                         capture_output=True, text=True).stdout
    import pdfplumber
    with pdfplumber.open(str(pdf)) as d:
        layout = "\n".join((p.extract_text() or "") for p in d.pages)
        pages = len(d.pages)
        # ruled tables only: filled, unstroked rectangles (chip / pill backgrounds,
        # text highlights) are paint behind linear text, not layout tables
        def _ruled(pg):
            return pg.filter(lambda o: not (o.get("object_type") == "rect"
                                            and o.get("fill") and not o.get("stroke")))
        has_tables = any(_ruled(p).find_tables() for p in d.pages)
    fonts = subprocess.run(["pdffonts", str(pdf)], capture_output=True, text=True).stdout
    font_lines = [l for l in fonts.splitlines()[2:] if l.strip()]
    not_embedded = [l.split()[0] for l in font_lines if l.split()[-5] == "no"]
    type3 = [l.split()[0] for l in font_lines if "Type 3" in l]
    toks = _tokens(cv)
    res = {
        "file": pdf.name, "pages": pages,
        "text_layer": len(raw.strip()) > 200,
        "missing_stream_order": _coverage(raw, toks),
        "missing_layout_order": _coverage(layout, toks),
        "order_ok_stream": _order_ok(raw, cv),
        "order_ok_layout": _order_ok(layout, cv),
        "bad_chars": _bad_chars(raw + layout),
        "fonts_not_embedded": not_embedded,
        "type3_fonts": type3,
        "table_regions": has_tables,
        "tokens_checked": len(toks),
    }
    res["score"] = _score(res)
    return res


def check_docx(path: Path, cv: dict) -> dict:
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf8")
    issues = []
    if "<w:tbl>" in doc or "<w:tbl " in doc:
        issues.append("tables in body")
    if "txbxContent" in doc:
        issues.append("text boxes")
    if re.search(r'<w:cols [^>]*w:num="[2-9]"', doc):
        issues.append("multi-column section")
    for n in z.namelist():
        if re.match(r"word/(header|footer)\d*\.xml", n):
            x = z.read(n).decode("utf8")
            if re.search(r"<w:t[ >][^<]*\S", x):
                issues.append(f"text in {n.split('/')[-1]}")
    import docx
    d = docx.Document(str(path))
    text = "\n".join(p.text for p in d.paragraphs)
    missing = _coverage(text, _tokens(cv))
    res = {"file": path.name, "structure_issues": issues, "missing": missing,
           "order_ok": _order_ok(text, cv), "bad_chars": _bad_chars(text)}
    res["score"] = 100 - 25 * len(issues) - 3 * len(missing) - (15 if not res["order_ok"] else 0) \
        - 10 * len(res["bad_chars"])
    res["score"] = max(0, res["score"])
    return res


def _score(r):
    s = 100
    if not r["text_layer"]:
        return 0
    s -= 3 * len(r["missing_stream_order"]) + 3 * len(r["missing_layout_order"])
    s -= 0 if r["order_ok_stream"] else 15
    s -= 0 if r["order_ok_layout"] else 15
    s -= 10 * len(r["bad_chars"])
    s -= 10 if r["fonts_not_embedded"] else 0
    s -= 5 if r.get("type3_fonts") else 0
    s -= 10 if r["table_regions"] else 0
    s -= 5 if r["pages"] > 2 else 0
    return max(0, s)


def report(results) -> str:
    lines = []
    for r in results:
        flag = "PASS" if r["score"] >= 90 else "WARN" if r["score"] >= 75 else "FAIL"
        lines.append(f"[{flag}] {r['file']}: ATS score {r['score']}/100")
        for k, v in r.items():
            if k in ("file", "score", "tokens_checked"):
                continue
            if isinstance(v, list) and v:
                lines.append(f"      {k}: {', '.join(map(str, v[:6]))}")
            elif v is False and k.startswith("order"):
                lines.append(f"      {k}: FAILED")
            elif k == "table_regions" and v:
                lines.append("      table-like regions detected")
    return "\n".join(lines)
