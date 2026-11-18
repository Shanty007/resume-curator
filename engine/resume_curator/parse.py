"""Heuristic resume text -> structured JSON (EN + DE headings).

Good enough for clean resumes; always review the JSON. When run inside Claude
(the skill), Claude writes the JSON itself from the extracted text, which is
more reliable than any heuristic.
"""
from __future__ import annotations

import re

SECTION_SYNONYMS = {
    "summary": ["summary", "profile", "professional summary", "about me", "objective",
                "career objective", "personal statement", "profil", "kurzprofil",
                "über mich", "zusammenfassung", "professional profile", "executive summary"],
    "experience": ["experience", "work experience", "professional experience",
                   "employment", "employment history", "work history", "career history",
                   "relevant experience", "berufserfahrung", "berufliche erfahrung",
                   "beruflicher werdegang", "praktische erfahrung", "internships",
                   "professional background"],
    "education": ["education", "academic background", "education and training",
                  "academic qualifications", "qualifications", "ausbildung",
                  "bildungsweg", "studium", "schulbildung", "akademischer werdegang"],
    "skills": ["skills", "technical skills", "key skills", "core competencies",
               "competencies", "expertise", "tools", "it skills", "kenntnisse",
               "fähigkeiten", "kompetenzen", "edv-kenntnisse", "it-kenntnisse",
               "skills & tools", "technologies", "tech stack", "areas of expertise"],
    "languages": ["languages", "language skills", "sprachen", "sprachkenntnisse"],
    "certifications": ["certifications", "certificates", "licenses",
                       "licenses & certifications", "courses", "training",
                       "zertifikate", "zertifizierungen", "weiterbildung", "fortbildungen"],
    "projects": ["projects", "key projects", "selected projects", "projekte"],
    "contact": ["contact", "contact details", "contact information", "contact info",
                "kontakt", "kontaktdaten"],
    "personal": ["personal details", "personal information", "personal data",
                 "persönliche daten", "angaben zur person"],
}
_HEAD = {}
for k, vs in SECTION_SYNONYMS.items():
    for v in vs:
        _HEAD[v] = k

MONTHS = (r"jan(?:uary|uar)?|feb(?:ruary|ruar)?|mar(?:ch)?|märz|apr(?:il)?|may|mai|jun(?:e|i)?|"
          r"jul(?:y|i)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|okt(?:ober)?|nov(?:ember)?|"
          r"dec(?:ember)?|dez(?:ember)?")
DATE = rf"(?:(?:{MONTHS})\.?\s+\d{{4}}|\d{{1,2}}[./]\d{{4}}|\d{{4}}[-/.]\d{{1,2}}|\d{{4}})"
PRESENT = r"present|current|now|today|ongoing|heute|aktuell|bis heute|jetzt|dato"
RANGE_RE = re.compile(
    rf"(?P<start>{DATE})\s*(?:[-–—~]|to|bis)\s*(?P<end>{DATE}|{PRESENT})|(?:seit|since)\s+(?P<since>{DATE})",
    re.I)
SINGLE_DATE_RE = re.compile(rf"(?<![\w/]){DATE}(?![\w/])", re.I)
BULLET_RE = re.compile(r"^\s*(?:[•●○▪■◦‣⁃∙·\-\*–»>✓✔]|\d+[.)])\s+")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PHONE_RE = re.compile(r"(?:\+|00)?\d[\d\s()/.-]{7,}\d")
LINKEDIN_RE = re.compile(r"(?:https?://)?(?:[a-z]{2,3}\.)?linkedin\.com/[^\s|,;]+", re.I)
GITHUB_RE = re.compile(r"(?:https?://)?github\.com/[^\s|,;]+", re.I)
URL_RE = re.compile(r"(?:https?://|www\.)[^\s|,;]+", re.I)

MONTH_NUM = {"jan": 1, "feb": 2, "mar": 3, "mär": 3, "apr": 4, "may": 5, "mai": 5, "jun": 6,
             "jul": 7, "aug": 8, "sep": 9, "oct": 10, "okt": 10, "nov": 11, "dec": 12, "dez": 12}


def norm_date(s: str | None) -> str:
    """-> 'MM/YYYY', 'YYYY' or 'Present'."""
    if not s:
        return ""
    s = s.strip().rstrip(".")
    if re.fullmatch(PRESENT, s, re.I):
        return "Present"
    m = re.fullmatch(rf"({MONTHS})\.?\s+(\d{{4}})", s, re.I)
    if m:
        return f"{MONTH_NUM[m.group(1)[:3].lower()]:02d}/{m.group(2)}"
    m = re.fullmatch(r"(\d{1,2})[./](\d{4})", s)
    if m:
        return f"{int(m.group(1)):02d}/{m.group(2)}"
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})", s)
    if m:
        return f"{int(m.group(2)):02d}/{m.group(1)}"
    return s


def heading_key(line: str) -> str | None:
    t = re.sub(r"[:|#*_]+$", "", line.strip()).strip().lower()
    t = re.sub(r"\s+", " ", t)
    if len(t) > 40 or not t:
        return None
    return _HEAD.get(t)


def _split_sections(lines):
    sections = {"_header": []}
    order = ["_header"]
    cur = "_header"
    for ln in lines:
        k = heading_key(ln)
        if k:
            cur = k if k not in sections else k  # merge duplicates
            if cur not in sections:
                sections[cur] = []
                order.append(cur)
            continue
        # unknown ALL-CAPS short line => custom section (not inside list-like sections,
        # where all-caps items such as "SAP" or "SQL" are common)
        if (cur not in ("skills", "languages", "certifications")
                and re.fullmatch(r"[A-ZÄÖÜ][A-ZÄÖÜ &\-]{3,34}", ln)
                and len(ln.split()) <= 4
                and (cur != "_header" or len(sections["_header"]) >= 3)):
            cur = "x:" + ln.title()
            sections.setdefault(cur, [])
            if cur not in order:
                order.append(cur)
            continue
        sections[cur].append(ln)
    return sections, order


def _parse_header(lines):
    contact = {}
    name = headline = ""
    rest = []
    blob = " | ".join(lines)
    if m := EMAIL_RE.search(blob):
        contact["email"] = m.group(0)
    if m := LINKEDIN_RE.search(blob):
        contact["linkedin"] = m.group(0).replace("https://", "").replace("http://", "")
    if m := GITHUB_RE.search(blob):
        contact["github"] = m.group(0).replace("https://", "")
    for m in PHONE_RE.finditer(blob):
        digits = re.sub(r"\D", "", m.group(0))
        if 8 <= len(digits) <= 15 and not re.fullmatch(r"(19|20)\d{2}\s*[-–]\s*(19|20)\d{2}", m.group(0).strip()):
            contact["phone"] = m.group(0).strip()
            break
    for ln in lines:
        stripped = ln
        for rx in (EMAIL_RE, LINKEDIN_RE, GITHUB_RE, URL_RE):
            stripped = rx.sub("", stripped)
        if contact.get("phone"):
            stripped = stripped.replace(contact["phone"], "")
        stripped = re.sub(r"^[\s|•·,;:/-]+|[\s|•·,;:/-]+$", "", stripped)
        stripped = re.sub(r"\s*[|•·]\s*[|•·]\s*", " | ", stripped)
        if not stripped:
            continue
        if not name and len(stripped.split()) <= 5 and not any(ch.isdigit() for ch in stripped):
            name = stripped.title() if stripped.isupper() else stripped
        elif not headline and len(stripped) < 90 and "," not in stripped[:3]:
            # a location-looking line goes to location
            if re.fullmatch(r"[A-ZÄÖÜ][\w .äöüß-]+,\s*[A-ZÄÖÜ][\w .äöüß-]+", stripped) and len(stripped) < 40:
                contact.setdefault("location", stripped)
            else:
                headline = stripped
        else:
            rest.append(stripped)
    for r in rest:
        if "location" not in contact and re.search(r"^[A-ZÄÖÜ][\w .äöüß-]+,\s*[A-ZÄÖÜ][\w .äöüß-]+$", r) and len(r) < 40:
            contact["location"] = r
    # location embedded in a pipe-separated contact line
    if "location" not in contact:
        for ln in lines:
            for seg in re.split(r"[|•·]", ln):
                seg = seg.strip()
                if re.fullmatch(r"[A-ZÄÖÜ][a-zäöüß.-]+(?: [A-ZÄÖÜ][a-zäöüß.-]+)?,\s*[A-ZÄÖÜ][a-zäöüß]+", seg):
                    contact["location"] = seg
                    break
    return name, headline, contact


def _split_header_text(txt):
    """'Title | Company | City' / 'Title at Company, City' / 'Title, Company'."""
    txt = txt.strip(" |,–-")
    if re.search(r"\s[|•·]\s", txt):          # explicit separators win
        parts = [p.strip() for p in re.split(r"\s+[|•·]\s+", txt) if p.strip()]
    else:
        parts = [p.strip() for p in re.split(r"\s+(?:at|bei|@)\s+|\s+[–—]\s+", txt) if p.strip()]
    if len(parts) == 1 and "," in txt:
        parts = [p.strip() for p in txt.split(",") if p.strip()]
    return parts


def _entries(lines, kind):
    """Group lines into dated entries with header lines + bullets."""
    entries = []
    cur = None

    def new():
        return {"_head": [], "start": "", "end": "", "bullets": []}

    for i, ln in enumerate(lines):
        is_bullet = bool(BULLET_RE.match(ln))
        rng = RANGE_RE.search(ln)
        single = None if rng else SINGLE_DATE_RE.search(ln) if kind == "education" else None
        if is_bullet:
            if cur is None:
                cur = new(); entries.append(cur)
            cur["bullets"].append(BULLET_RE.sub("", ln).strip())
            continue
        if rng or single:
            if cur is None or cur["start"] or cur["bullets"]:
                # header lines just before the date line belong to the new entry
                carry = []
                if cur is not None and cur["bullets"] == [] and not cur["start"]:
                    carry = cur["_head"]; entries.remove(cur)
                cur = new(); cur["_head"] = carry; entries.append(cur)
            if rng:
                cur["start"] = norm_date(rng.group("start") or rng.group("since"))
                cur["end"] = norm_date(rng.group("end")) if rng.group("end") else "Present"
                rest = (ln[:rng.start()] + " " + ln[rng.end():]).strip(" |,–-()")
            else:
                cur["end"] = norm_date(single.group(0))
                rest = (ln[:single.start()] + " " + ln[single.end():]).strip(" |,–-()")
            if rest:
                cur["_head"].append(rest)
            continue
        # plain line
        if cur is None:
            cur = new(); entries.append(cur); cur["_head"].append(ln); continue
        if cur["bullets"]:
            prev = cur["bullets"][-1]
            nxt = lines[i + 1] if i + 1 < len(lines) else ""
            looks_new = bool(RANGE_RE.search(nxt)) or (len(ln) < 60 and ln[:1].isupper()
                                                       and prev.endswith((".", "%", ")")))
            if not looks_new and (ln[:1].islower() or not prev.endswith((".", "!", "?"))):
                cur["bullets"][-1] = prev + " " + ln
            else:
                cur = new(); entries.append(cur); cur["_head"].append(ln)
        elif len(cur["_head"]) >= 3 and len(ln) > 60:
            cur["bullets"].append(ln)  # prose description
        else:
            cur["_head"].append(ln)

    out = []
    for e in entries:
        head = []
        for h in e["_head"]:
            head.extend(_split_header_text(h))
        loc = ""
        for j, h in enumerate(head):
            if re.fullmatch(r"[A-ZÄÖÜ][\w.äöüß-]+(?: [A-ZÄÖÜ][\w.äöüß-]+)?(?:,\s*[A-ZÄÖÜ][\w.äöüß -]+)?", h) \
                    and j >= 2 or re.search(r"\b(remote|hybrid)\b", h, re.I):
                loc = h; head.pop(j); break
        a = head[0] if head else ""
        b = head[1] if len(head) > 1 else ""
        extra = head[2:]
        if not loc and extra and len(extra[0]) < 35:
            loc = extra.pop(0)
        if kind == "experience":
            item = {"title": a, "company": b, "location": loc, "start": e["start"],
                    "end": e["end"], "bullets": extra + e["bullets"]}
        elif kind == "education":
            item = {"degree": a, "institution": b, "location": loc, "start": e["start"],
                    "end": e["end"], "details": extra + e["bullets"]}
        else:
            item = {"name": a, "subtitle": b, "start": e["start"], "end": e["end"],
                    "bullets": extra + e["bullets"]}
        out.append(item)
    return out


def _skills(lines):
    groups = []
    loose = []
    for ln in lines:
        ln = BULLET_RE.sub("", ln).strip()
        m = re.match(r"^([^:]{2,40}):\s*(.+)$", ln)
        if m:
            items = [x.strip() for x in re.split(r"[,;|•·]", m.group(2)) if x.strip()]
            groups.append({"category": m.group(1).strip(), "items": items})
        else:
            loose += [x.strip() for x in re.split(r"[,;|•·]", ln) if x.strip()]
    if loose:
        groups.append({"category": "" if not groups else "Other", "items": loose})
    return groups


LEVEL_RE = re.compile(r"\b(A1|A2|B1|B2|C1|C2|native|mother tongue|muttersprache|fluent|"
                      r"fließend|verhandlungssicher|business fluent|professional|basic|"
                      r"grundkenntnisse|intermediate|advanced|gut|sehr gut|conversational)\b.*",
                      re.I)


def _languages(lines):
    out = []
    for ln in lines:
        for part in re.split(r"[,;|•·]\s*(?=[A-ZÄÖÜ])", BULLET_RE.sub("", ln)):
            part = part.strip()
            if not part:
                continue
            m = re.match(r"^([A-Za-zÄÖÜäöüß ]+?)\s*(?:[-–:(]|\s)\s*(.*?)\)?$", part)
            if m and m.group(2):
                out.append({"language": m.group(1).strip(), "level": m.group(2).strip(" ()-–")})
            else:
                out.append({"language": part, "level": ""})
    return out


def parse_text(text: str) -> dict:
    lines = [l for l in text.splitlines() if l.strip()]
    sections, order = _split_sections(lines)
    name, headline, contact = _parse_header(sections.get("_header", []) + sections.get("contact", []))
    cv = {"name": name, "headline": headline, "contact": contact, "personal": {},
          "summary": "", "experience": [], "education": [], "skills": [], "languages": [],
          "certifications": [], "projects": [], "extra_sections": [],
          "section_order": [k for k in order if k != "_header"]}
    for k in order:
        body = sections[k]
        if k == "summary":
            cv["summary"] = " ".join(BULLET_RE.sub("", l) for l in body)
        elif k == "experience":
            cv["experience"] = _entries(body, "experience")
        elif k == "education":
            cv["education"] = _entries(body, "education")
        elif k == "skills":
            cv["skills"] = _skills(body)
        elif k == "languages":
            cv["languages"] = _languages(body)
        elif k == "certifications":
            cv["certifications"] = [{"name": BULLET_RE.sub("", l).strip()} for l in body]
        elif k == "projects":
            cv["projects"] = _entries(body, "projects")
        elif k == "personal":
            for l in body:
                if ":" in l:
                    a, b = l.split(":", 1)
                    cv["personal"][a.strip()] = b.strip()
        elif k.startswith("x:"):
            cv["extra_sections"].append({"title": k[2:], "items": [BULLET_RE.sub("", l).strip() for l in body]})
    # if no summary heading, a long prose line in the header is the summary
    return cv
