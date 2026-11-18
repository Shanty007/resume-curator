---
name: resume-curator
description: Turn a plain DOCX/PDF/TXT resume into 10 ATS-safe resume templates (DOCX + PDF), and tailor it to a job ad without inventing anything. No OCR. Use for resume/CV/Lebenslauf formatting, redesigns or job-specific tailoring.
---

# Resume Curator

Converts a text-based resume into ten designed templates. Each comes out as a DOCX and a PDF, and each is verified by an automatic ATS read-back check. It can also tailor a resume to a job ad without inventing anything.

- **Classic set:** spree (German tabular Lebenslauf), isar (executive serif), rhine (colour header band), elbe (Swiss minimal), danube (compact tech, skills first), neckar (accent stripe).
- **Studio set:** havel (tinted sidebar + photo/monogram circle), mosel (timeline + skill tags), saale (black & white serif monogram), lahn (navy band + pill headings).

## 0. Engine
The engine ships with this skill in `engine/`, next to this SKILL.md: `curator.py` plus the `resume_curator/` package.
1. Copy the `engine/` folder into the working directory. Don't edit it in place.
2. Check the dependencies: python-docx, pdfplumber and Pillow for Python, plus LibreOffice (`soffice`) and poppler (`pdftotext`, `pdffonts`). Run `which soffice pdftotext` and try the imports.
3. Run `python3 engine/curator.py selftest` once. It must print `SELFTEST PASSED`.

## 1. Extract text (never OCR)
```bash
python3 engine/curator.py parse <resume> -o heuristic.json --text raw.txt
```
- If the tool reports "no text layer" (a scanned or image-only PDF), stop. Tell the user that OCR is deliberately not used because it degrades ATS parsing, and ask for the original DOCX or a text-based PDF.
- The PDF reader handles columns, so a Canva/builder sidebar is read as one block.

## 2. Structure the resume yourself
Read `raw.txt` and write `master.json` yourself, using `heuristic.json` only as a starting point. You are more accurate than the heuristic parser. Schema:
```json
{"name":"","headline":"","contact":{"email":"","phone":"","location":"","linkedin":"","github":"","website":""},
 "personal":{"dob":"","nationality":""},
 "summary":"",
 "competencies":[],
 "experience":[{"title":"","company":"","location":"","start":"MM/YYYY","end":"MM/YYYY|Present","bullets":[]}],
 "education":[{"degree":"","institution":"","location":"","start":"","end":"","details":[]}],
 "skills":[{"category":"","items":[]}],
 "languages":[{"language":"","level":"C1"}],
 "certifications":[{"name":"","issuer":"","date":""}],
 "projects":[{"name":"","subtitle":"","start":"","end":"","bullets":[]}],
 "extra_sections":[{"title":"","items":[]}]}
```
Rules:
- Keep the user's wording. Do not invent achievements, numbers, dates or skills. Fix only obvious typos and splitting errors.
- Dates go in MM/YYYY format. Language levels use CEFR (A1–C2 or Native) when the source gives them.
- `personal` holds things like work permit, date of birth and nationality. Fill it only if the source has them or the user asks. Omit empty fields rather than writing placeholders.

## 2b. Tailor to a job ad (only when the user gives one)
`master.json` is never edited.

1. Write `job.json` from the ad, using its exact wording: `{"title","company","years","domain":[],"must_have":[],"tools":[],"nice_to_have":[],"soft":[]}`.
2. Run `python3 engine/curator.py match master.json job.json`. This is the recruiter's Ctrl+F test: it shows which terms are missing, buried, or already visible at the top.
3. Write `tailored.json` as a copy of the master and change only these five things:
   1. **Headline:** mirror the ad's job title when it is the same kind of role in other words. Never claim a higher level (Senior/Lead/Manager) than the user holds. Keep the real titles under experience.
   2. **Summary:** three sentences, no clichés.
      - Sentence 1: identity + years + domain, matching the must-haves.
      - Sentence 2: the strongest relevant proof from the master, using only numbers that already exist there.
      - Sentence 3: what the user brings to this team, using the ad's tools.
      - Check: a stranger reading only the summary should see that every hard requirement is met.
   3. **competencies:** 9 to 12 terms copied verbatim from the ad. Include a term only if the master proves it.
   4. **Skills:** put the categories the ad cares about first, and items within them. Add nothing new.
   5. **Bullets:** within each role, move the most relevant bullets up and swap in the ad's terminology for the same fact (Action + Context + Result). One sentence each. Never add a metric, tool, scope or responsibility the master doesn't contain. If a strong bullet lacks a number, ask the user once for real numbers, listing those bullets; otherwise leave it without one.

   Never touch employers, dates, degrees, grades or certifications. Requirements the user doesn't meet stay off the resume: list them as gaps for the cover letter or interview.
4. Run `match tailored.json job.json`. Aim for the target title FOUND, years in the summary, and every must-have the user has visible in headline/summary/competencies.
5. Run `guard master.json tailored.json --job job.json`. Fix every flag (new numbers, new terms, changed titles/facts, clichés, stuffing) unless the user confirms it's true.
6. Render `tailored.json` in step 4, into `out_<company>/`.
7. In the reply give: what changed, the match score before -> after, the guard result, and the gaps.

## 3. Ask only what changes the output (one AskUserQuestion, skip if already known)
- **Templates:** all ten (default), the classic six, the studio four, or a subset.
- **Language:** `en`, or `de` for German headings and "heute" for present roles.
- **Photo:** spree, neckar, havel, mosel and lahn can take one, and the user must attach it. Recommend one for German-speaking employers, none for UK/IE/NL/international tech. For German Lebenslauf targets, also offer a signature line with a city.

## 4. Render and check
```bash
python3 engine/curator.py render resume.json -o out/ --templates all [--lang de] [--photo photo.jpg] [--signature Berlin]
```
This writes `<Name>_<template>.docx/.pdf` and `out/ats_report.txt`. Every file must score PASS (90 or above). If one doesn't:
- Missing tokens usually mean a JSON mistake. Fix the JSON and re-render.
- More than 2 pages: tighten bullets with the user's agreement, or recommend danube.
- Never "fix" a failure by shrinking type below the template defaults.

## 5. Visual check, then deliver
- Rasterize page 1 of each PDF (`pdftoppm -r 75 -png -f 1 -l 1`) and look at the images. Check for overlaps, an orphaned heading, or the photo colliding with text.
- Build one labeled gallery PNG of all chosen templates.
- Send the gallery plus the DOCX and PDF files.
- In the reply: one line per template saying who it suits, the ATS result, and one recommendation.

## ATS principles (explain if asked)
- Single text flow: no tables, text boxes or columns, and nothing but text-free pictures in headers.
- Visual blocks (bands, sidebars, stripes, photos, monograms) are text-free pictures behind or beside real text. Timelines are paragraph borders; tags and pills are character shading on real text.
- Real Unicode bullets, ligatures off, static fonts embedded, tagged PDF/A-2b.
- Standard section headings in EN/DE, A4 paper, MM/YYYY dates.
