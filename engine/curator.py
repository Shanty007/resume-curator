#!/usr/bin/env python3
"""Resume Curator: turn a plain DOCX/PDF/TXT resume into several ATS-safe European templates.

  # 1) extract + parse (review/edit the JSON afterwards)
  python3 curator.py parse my_resume.pdf -o cv.json

  # 2) render every template (DOCX + PDF) and run the ATS self-check
  python3 curator.py render cv.json -o out/ --templates all

  # one-shot
  python3 curator.py convert my_resume.docx -o out/ --templates spree,rhine,danube

  # German Lebenslauf: German headings, photo, personal data, place/date signature line
  python3 curator.py render cv.json -o out/ --templates spree --lang de \
      --photo me.jpg --signature Berlin

  # tailor to a job ad: job.json (from the ad), tailored copy, then check both ways
  python3 curator.py match master.json job.json          # before: what's missing / buried
  python3 curator.py match tailored.json job.json        # after
  python3 curator.py guard master.json tailored.json --job job.json   # nothing invented

  python3 curator.py list      # show templates
  python3 curator.py check cv.json out/*.pdf out/*.docx
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from resume_curator.extract import extract_text, NoTextLayerError  # noqa: E402
from resume_curator.parse import parse_text  # noqa: E402
from resume_curator.render import render  # noqa: E402
from resume_curator.templates import TEMPLATES  # noqa: E402
from resume_curator.atscheck import check_pdf, check_docx, report  # noqa: E402
from resume_curator.tailor import match, match_report, guard, guard_report  # noqa: E402


def _slug(name):
    import re
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_") or "resume"


def cmd_list(_):
    for k, t in TEMPLATES.items():
        print(f"{k:8s} {t['label']}\n         best for: {t['best_for']}\n         pattern:  {t['inspired_by']}")


def cmd_parse(a):
    try:
        text = extract_text(a.input)
    except NoTextLayerError as e:
        sys.exit(f"ERROR: {e}")
    cv = parse_text(text)
    out = Path(a.output or Path(a.input).with_suffix(".json"))
    out.write_text(json.dumps(cv, indent=2, ensure_ascii=False), encoding="utf-8")
    if a.text:
        Path(a.text).write_text(text, encoding="utf-8")
    print(f"Parsed -> {out}  (review it: job titles, companies, dates)")
    return cv, out


def _render_all(cv, a):
    names = list(TEMPLATES) if a.templates in ("all", None) else a.templates.split(",")
    out = Path(a.output)
    base = _slug(cv.get("name", "resume"))
    results = []
    for t in names:
        if t not in TEMPLATES:
            print(f"skip unknown template {t}")
            continue
        d, p = render(cv, t, out, base, lang=a.lang, photo=a.photo,
                      signature_place=a.signature, keep_source_order=a.keep_order)
        results += [check_pdf(p, cv), check_docx(d, cv)]
        print(f"  {t:8s} -> {d.name}, {p.name}")
    rep = report(results)
    (out / "ats_report.txt").write_text(rep, encoding="utf-8")
    print("\nATS self-check\n" + rep)


def cmd_render(a):
    cv = json.loads(Path(a.json).read_text(encoding="utf-8"))
    _render_all(cv, a)


def cmd_convert(a):
    a.output_json = None
    try:
        text = extract_text(a.input)
    except NoTextLayerError as e:
        sys.exit(f"ERROR: {e}")
    cv = parse_text(text)
    Path(a.output).mkdir(parents=True, exist_ok=True)
    (Path(a.output) / "resume.json").write_text(json.dumps(cv, indent=2, ensure_ascii=False))
    _render_all(cv, a)


def cmd_check(a):
    cv = json.loads(Path(a.json).read_text(encoding="utf-8"))
    res = []
    for f in a.files:
        f = Path(f)
        res.append(check_pdf(f, cv) if f.suffix == ".pdf" else check_docx(f, cv))
    print(report(res))


def cmd_match(a):
    cv = json.loads(Path(a.json).read_text(encoding="utf-8"))
    job = json.loads(Path(a.job).read_text(encoding="utf-8"))
    print(match_report(match(cv, job)))


def cmd_guard(a):
    m = json.loads(Path(a.master).read_text(encoding="utf-8"))
    t = json.loads(Path(a.tailored).read_text(encoding="utf-8"))
    job = json.loads(Path(a.job).read_text(encoding="utf-8")) if a.job else None
    print(guard_report(guard(m, t, job)))


def cmd_selftest(_):
    """Render the bundled sample in every template and verify ATS + guard."""
    import tempfile
    root = Path(__file__).resolve().parent.parent / "samples"
    cv = json.loads((root / "arjun.json").read_text(encoding="utf-8"))
    ok = True
    with tempfile.TemporaryDirectory() as td:
        for t in TEMPLATES:
            d, p = render(cv, t, Path(td), "selftest")
            for r in (check_pdf(p, cv), check_docx(d, cv)):
                good = r["score"] >= 90
                ok &= good
                print(f"  {'ok ' if good else 'BAD'} {r['file']:<28} {r['score']}/100")
    g = guard(cv, json.loads((root / "sample_tailored.json").read_text(encoding="utf-8")))
    print("  guard on sample tailoring:", "clean" if not g["issues"] else g["issues"])
    ok &= not g["issues"]
    print("SELFTEST", "PASSED" if ok else "FAILED")
    sys.exit(0 if ok else 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    sp.add_parser("list").set_defaults(f=cmd_list)

    p = sp.add_parser("parse"); p.add_argument("input"); p.add_argument("-o", "--output")
    p.add_argument("--text", help="also save the raw extracted text here"); p.set_defaults(f=cmd_parse)

    for name, fn, first in (("render", cmd_render, "json"), ("convert", cmd_convert, "input")):
        r = sp.add_parser(name); r.add_argument(first); r.add_argument("-o", "--output", default="out")
        r.add_argument("--templates", default="all", help="comma list or 'all'")
        r.add_argument("--lang", default="en", choices=["en", "de"])
        r.add_argument("--photo", help="photo for templates that support one (spree, neckar)")
        r.add_argument("--signature", metavar="CITY", help="add 'City, DD.MM.YYYY' + name (DE style)")
        r.add_argument("--keep-order", action="store_true", help="keep the source section order")
        r.set_defaults(f=fn)

    m = sp.add_parser("match", help="Ctrl+F test: job-ad terms vs resume")
    m.add_argument("json"); m.add_argument("job"); m.set_defaults(f=cmd_match)
    g = sp.add_parser("guard", help="honesty check: tailored vs master resume")
    g.add_argument("master"); g.add_argument("tailored"); g.add_argument("--job")
    g.set_defaults(f=cmd_guard)

    sp.add_parser("selftest", help="render bundled sample in all templates and verify").set_defaults(f=cmd_selftest)

    c = sp.add_parser("check"); c.add_argument("json"); c.add_argument("files", nargs="+")
    c.set_defaults(f=cmd_check)
    a = ap.parse_args()
    a.f(a)


if __name__ == "__main__":
    main()
