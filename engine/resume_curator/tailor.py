"""Job-ad tailoring helpers.

match : the recruiter "Ctrl+F test". It checks where each job-ad term shows up in a
        resume (headline, summary, competencies, skills, bullets), so missing or
        buried requirements are visible.
guard : the honesty check. It compares a tailored resume with the master and flags
        anything that could sound off: new numbers, new tools/terms that appear
        nowhere in the master, changed employers/dates, changed job titles.

job.json (written by Claude from the job ad):
{"title": "Test Engineer Infotainment", "company": "", "years": 3,
 "domain": ["automotive"], "must_have": ["CANoe", "test automation"],
 "nice_to_have": [], "tools": [], "soft": []}
"""
from __future__ import annotations

import re


def _n(t: str) -> str:
    t = (t or "").lower().replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", t).strip()


def _zones(cv: dict) -> dict[str, str]:
    z = {
        "headline": cv.get("headline", ""),
        "summary": cv.get("summary", ""),
        "competencies": " | ".join(cv.get("competencies", [])),
        "skills": " | ".join(i for g in cv.get("skills", []) for i in g.get("items", []) + [g.get("category", "")]),
        "titles": " | ".join(e.get("title", "") for e in cv.get("experience", [])),
        "bullets": " ".join(b for e in cv.get("experience", []) for b in e.get("bullets", []))
                   + " " + " ".join(b for p in cv.get("projects", []) for b in p.get("bullets", [])),
        "education": " ".join(e.get("degree", "") + " " + " ".join(e.get("details", []))
                              for e in cv.get("education", [])),
        "certs": " ".join((c.get("name", "") if isinstance(c, dict) else c) for c in cv.get("certifications", [])),
    }
    return {k: _n(v) for k, v in z.items()}


def _count(term: str, text: str) -> int:
    t = re.escape(_n(term))
    return len(re.findall(rf"(?<![a-z0-9]){t}(?![a-z0-9])", text))


def match(cv: dict, job: dict) -> dict:
    z = _zones(cv)
    rows = []
    for kind in ("must_have", "tools", "domain", "nice_to_have", "soft"):
        for term in job.get(kind, []):
            hits = {k: _count(term, v) for k, v in z.items() if _count(term, v)}
            rows.append({"term": term, "kind": kind, "hits": hits})
    title = job.get("title", "")
    title_hit = _count(title, z["headline"]) + _count(title, z["titles"]) if title else 0
    top = z["headline"] + " " + z["summary"] + " " + z["competencies"]
    # "above the fold": what a recruiter sees in the first seconds
    must = [r for r in rows if r["kind"] in ("must_have", "tools")]
    fold = sum(1 for r in must if _count(r["term"], top))
    years = job.get("years")
    years_ok = bool(years) and bool(re.search(rf"\b{int(years)}\+?\s*(?:years|jahre)", z["summary"])) \
        or bool(years) and any(int(m) >= int(years) for m in re.findall(r"\b(\d{1,2})\+?\s*(?:years|jahre)", z["summary"]))
    return {"rows": rows, "title": title, "title_in_resume": title_hit > 0,
            "must_total": len(must), "must_found": sum(1 for r in must if r["hits"]),
            "must_above_fold": fold, "years": years, "years_in_summary": years_ok}


def match_report(m: dict) -> str:
    L = [f"Target title '{m['title']}': {'FOUND' if m['title_in_resume'] else 'MISSING (headline)'}"]
    if m["years"]:
        L.append(f"Years ({m['years']}+) stated in summary: {'yes' if m['years_in_summary'] else 'NO'}")
    L.append(f"Must-have/tools found anywhere: {m['must_found']}/{m['must_total']}   "
             f"in headline/summary/competencies: {m['must_above_fold']}/{m['must_total']}")
    for r in m["rows"]:
        where = ", ".join(f"{k}×{v}" for k, v in r["hits"].items()) or "—  MISSING"
        L.append(f"  [{r['kind']:<12}] {r['term']:<32} {where}")
    return "\n".join(L)


# ---------------------------------------------------------------- honesty guard
_NUM = re.compile(r"(?<![\w.])\d+(?:[.,]\d+)?\s*(?:%|\+|x\b)?")
_TERM = re.compile(r"\b(?:[A-Z][A-Za-z0-9+#./-]{1,}|[a-z]+[A-Z][A-Za-z0-9]*)\b")  # Capitalised / camelCase tokens
_CLICHES = ["results-driven", "results driven", "proven track record", "dynamic", "passionate",
            "team player", "synergy", "go-getter", "detail-oriented", "hard-working", "hardworking",
            "think outside the box", "self-starter", "rockstar", "ninja", "spearheaded", "leveraged",
            "highly motivated", "best-in-class", "world-class"]


def _all_text(cv: dict) -> str:
    parts = [cv.get("headline", ""), cv.get("summary", ""), " ".join(cv.get("competencies", []))]
    for g in cv.get("skills", []):
        parts += [g.get("category", "")] + g.get("items", [])
    for e in cv.get("experience", []) + cv.get("projects", []):
        parts += [e.get("title", e.get("name", "")), e.get("company", ""), e.get("subtitle", "")] + e.get("bullets", [])
    for e in cv.get("education", []):
        parts += [e.get("degree", ""), e.get("institution", "")] + e.get("details", [])
    for c in cv.get("certifications", []):
        parts.append(c.get("name", "") if isinstance(c, dict) else c)
    return "\n".join(p for p in parts if p)


def guard(master: dict, tailored: dict, job: dict | None = None) -> dict:
    mt, tt = _all_text(master), _all_text(tailored)
    mtn = _n(mt)
    issues = []
    # 1. numbers that don't exist in the master = invented metrics
    m_nums = {re.sub(r"\s", "", x) for x in _NUM.findall(mt)}
    new_nums = sorted({re.sub(r"\s", "", x) for x in _NUM.findall(tt)} - m_nums)
    if new_nums:
        issues.append(("NEW NUMBERS", new_nums))
    # 2. capitalised terms (tools, standards, products) not in the master
    new_terms = set()
    for m_ in _TERM.finditer(tt):
        tok = m_.group(0)
        if _n(tok).replace("-", " ") in mtn.replace("-", " "):
            continue
        before = tt[max(0, m_.start() - 2):m_.start()]
        sentence_start = m_.start() == 0 or "\n" in before or before.strip() in (".", ":", ";")
        techy = sum(ch.isupper() for ch in tok) >= 2 or any(ch.isdigit() or ch in "+#./" for ch in tok)
        if techy or not sentence_start:      # ordinary capitalised sentence openers are not "terms"
            new_terms.add(tok)
    new_terms = sorted(new_terms)
    if new_terms:
        issues.append(("NEW TERMS (check you really have this)", new_terms))
    # 3. facts that must never change
    def facts(cv):
        return [(e.get("company"), e.get("start"), e.get("end")) for e in cv.get("experience", [])]
    if facts(master) != facts(tailored):
        issues.append(("EMPLOYERS/DATES CHANGED", [str(facts(tailored))]))
    tm = [e.get("title", "") for e in master.get("experience", [])]
    tt_ = [e.get("title", "") for e in tailored.get("experience", [])]
    changed = [f"{a} -> {b}" for a, b in zip(tm, tt_) if a != b]
    if changed:
        issues.append(("JOB TITLES CHANGED (only same role in other words; never a higher level)", changed))
    if [e.get("degree") for e in master.get("education", [])] != [e.get("degree") for e in tailored.get("education", [])]:
        issues.append(("DEGREES CHANGED", ["education must stay identical"]))
    # 4. cliché / AI-fluff words
    low = _n(tt)
    fluff = [c for c in _CLICHES if c in low and c not in mtn]
    if fluff:
        issues.append(("CLICHÉS ADDED", fluff))
    # 5. keyword stuffing: a term repeated too often
    if job:
        for term in job.get("must_have", []) + job.get("tools", []):
            c = _count(term, low)
            if c > 6:   # headline+summary+competencies+skills+2 bullets is normal; more reads as stuffing
                issues.append(("POSSIBLE STUFFING", [f"{term} ×{c}"]))
    # 6. competencies must be backed by the master
    unbacked = [c for c in tailored.get("competencies", []) if _n(c) not in mtn and not any(
        w in mtn for w in _n(c).split() if len(w) > 3)]
    if unbacked:
        issues.append(("COMPETENCIES NOT BACKED BY MASTER", unbacked))
    return {"issues": issues}


def guard_report(g: dict) -> str:
    if not g["issues"]:
        return "Honesty guard: clean. Nothing in the tailored resume goes beyond the master."
    L = ["Honesty guard: review each item. Keep it only if it is true and you can defend it in an interview."]
    for name, items in g["issues"]:
        L.append(f"  {name}: {', '.join(map(str, items[:15]))}")
    return "\n".join(L)
