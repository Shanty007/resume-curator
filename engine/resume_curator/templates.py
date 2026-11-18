"""Template design specs.

Each design takes a visual pattern popular with Resume.io / Novorésumé / Canva and
rebuilds it from ATS-safe parts: one text flow, no tables, no text boxes, nothing in
headers or footers. The visual effects come from paragraph shading, paragraph borders,
tab stops, hanging indents and page borders. These sit on real text or have no text
at all, so parsers read the content in visual order.

font_docx = a font every Word install has (used in the .docx you hand out)
font_pdf  = a metric-compatible or nicer font embedded in the PDF (static, non-variable
            fonts only: variable fonts export as Type 3, which some parsers choke on)
"""

COMMON_ORDER = ["summary", "competencies", "experience", "education", "skills", "languages",
                "certifications", "projects", "extra"]
TECH_ORDER = ["summary", "competencies", "skills", "experience", "projects", "education",
              "certifications", "languages", "extra"]

TEMPLATES = {
    "spree": {
        "label": "Spree: German tabular (tabellarischer Lebenslauf)",
        "inspired_by": "classic DE Lebenslauf / Resume.io 'Berlin'-style date gutter",
        "best_for": "German Mittelstand, automotive, engineering, public sector",
        "layout": "tabular", "date_col_cm": 3.6,
        "font_docx": "Calibri", "font_pdf": "Carlito",
        "head_font_docx": "Calibri", "head_font_pdf": "Carlito",
        "body_pt": 10.5, "name_pt": 24, "headline_pt": 12,
        "accent": "1F4E79", "text": "222222", "muted": "5A6270",
        "margins_cm": (1.9, 2.0, 1.8, 2.0),  # top, right, bottom, left
        "header": "left", "band": False,
        "heading": {"case": "upper", "rule": "bottom", "size": 11.5, "color": "accent",
                    "spacing_pt": 0.5, "bar": False},
        "order": COMMON_ORDER, "photo": True,
    },
    "isar": {
        "label": "Isar: executive serif",
        "inspired_by": "Resume.io / Novorésumé 'Traditional' & 'Professional' serif layouts",
        "best_for": "finance, law, consulting, senior and management roles",
        "layout": "classic",
        "font_docx": "Cambria", "font_pdf": "Caladea",
        "head_font_docx": "Cambria", "head_font_pdf": "Caladea",
        "body_pt": 10.5, "name_pt": 25, "headline_pt": 11.5,
        "accent": "7A1F2B", "text": "1E1E1E", "muted": "5F5F5F",
        "margins_cm": (1.9, 2.1, 1.8, 2.1),
        "header": "center", "band": False,
        "heading": {"case": "smallcaps", "rule": "both", "size": 12, "color": "text",
                    "spacing_pt": 0.6, "bar": False, "align": "center"},
        "order": COMMON_ORDER, "photo": False,
    },
    "rhine": {
        "label": "Rhine: modern colour header band",
        "inspired_by": "Novorésumé / Canva colour-block header look",
        "best_for": "tech, product, sales, marketing, international firms",
        "layout": "classic",
        "font_docx": "Arial", "font_pdf": "Inter",
        "head_font_docx": "Arial", "head_font_pdf": "Inter",
        "body_pt": 10, "name_pt": 26, "headline_pt": 12,
        "accent": "0F6E6E", "text": "1F2933", "muted": "52606D",
        "margins_cm": (0.0, 1.9, 1.7, 1.9),  # band is a background picture
        "header": "band", "band": True,
        "heading": {"case": "upper", "rule": "none", "size": 11, "color": "accent",
                    "spacing_pt": 0.8, "bar": False},
        "order": COMMON_ORDER, "photo": False,
    },
    "elbe": {
        "label": "Elbe: Swiss minimal",
        "inspired_by": "Resume.io minimal single-column layouts with date lines above the role",
        "best_for": "design-adjacent, research, startups, any role that values a clean look",
        "layout": "stacked",
        "font_docx": "Arial", "font_pdf": "Liberation Sans",
        "head_font_docx": "Arial", "head_font_pdf": "Liberation Sans",
        "body_pt": 10, "name_pt": 30, "headline_pt": 12,
        "accent": "D7263D", "text": "111111", "muted": "6B6B6B",
        "margins_cm": (2.0, 2.2, 1.8, 2.2),
        "header": "left", "band": False, "name_bold": False,
        "heading": {"case": "upper", "rule": "none", "size": 9.5, "color": "accent",
                    "spacing_pt": 1.0, "bar": False},
        "order": COMMON_ORDER, "photo": False,
    },
    "danube": {
        "label": "Danube: compact tech one-pager",
        "inspired_by": "Novorésumé 'Tech' / Resume.io compact layouts (skills first)",
        "best_for": "software, IT, ADAS/embedded, test and validation engineering",
        "layout": "classic",
        "font_docx": "Calibri", "font_pdf": "Fira Sans",
        "head_font_docx": "Calibri", "head_font_pdf": "Fira Sans",
        "body_pt": 9.5, "name_pt": 22, "headline_pt": 11,
        "accent": "2B59C3", "text": "1B1F24", "muted": "57606A",
        "margins_cm": (1.5, 1.6, 1.4, 1.6),
        "header": "split", "band": False,
        "heading": {"case": "upper", "rule": "none", "size": 10.5, "color": "accent",
                    "spacing_pt": 0.5, "bar": True},
        "order": TECH_ORDER, "photo": False,
    },
    "neckar": {
        "label": "Neckar: accent stripe (Canva-style, but ATS-safe)",
        "inspired_by": "Canva sidebar/stripe designs, rebuilt without text boxes",
        "best_for": "marketing, HR, hospitality, creative-adjacent corporate roles",
        "layout": "stacked",
        "font_docx": "Calibri", "font_pdf": "Carlito",
        "head_font_docx": "Calibri Light", "head_font_pdf": "Poppins",
        "body_pt": 9.8, "name_pt": 28, "headline_pt": 12,
        "accent": "6D28D9", "text": "1F1B2D", "muted": "5B5670",
        "margins_cm": (1.8, 1.9, 1.7, 2.4),
        "header": "left", "band": False, "stripe": True, "name_color": "accent",
        "heading": {"case": "upper", "rule": "bottom-thin", "size": 11, "color": "accent",
                    "spacing_pt": 0.6, "bar": False},
        "order": COMMON_ORDER, "photo": True,
    },
}

# ------------------------------------------------------------------ "Studio" set
# Canva-style looks rebuilt from ATS-safe parts:
#   tinted sidebar  -> text-free picture in the page header (repeats on every page);
#                      the "sidebar" text is a tab-stop gutter in the same text line
#   circle photo    -> picture, or a monogram badge when there is no photo
#   timeline        -> paragraph left borders (no shapes, no text boxes)
#   skill chips     -> character shading on real text
#   pill headings   -> character shading on a real heading
TEMPLATES.update({
    "havel": {
        "label": "Havel: tinted sidebar (Canva sidebar look, single text flow)",
        "inspired_by": "Canva olive/sage sidebar resumes with photo circle",
        "best_for": "corporate, engineering, consulting; looks two-column, parses as one",
        "layout": "tabular", "date_col_cm": 4.7, "heading_in_col": True,
        "font_docx": "Calibri", "font_pdf": "Inter",
        "head_font_docx": "Calibri", "head_font_pdf": "Poppins",
        "body_pt": 9.6, "name_pt": 27, "headline_pt": 10.8,
        "accent": "4A5A2C", "text": "1E2216", "muted": "5C6352",
        "margins_cm": (1.7, 1.7, 1.6, 1.25),
        "header": "left", "band": False, "name_color": "accent",
        "panel": {"w_cm": 5.55, "color": "ECEFE3", "edge": "C9D1B4"},
        "badge_bg": "4A5A2C", "badge_fg": "ECEFE3", "badge_ring": "FFFFFF",
        "heading": {"case": "upper", "rule": "bottom-thin", "size": 10.5, "color": "accent",
                    "spacing_pt": 0.8, "bar": False},
        "order": COMMON_ORDER, "photo": True,
    },
    "mosel": {
        "label": "Mosel: timeline + skill chips",
        "inspired_by": "Canva timeline resumes and tag-style skill clouds",
        "best_for": "tech, product, automotive software, startups",
        "layout": "stacked",
        "timeline": {"indent_cm": 0.38, "color": "E4E1EA", "sz": 12, "marker_sz": 30, "space_pt": 9},
        "chips": {"bg": "FBE3D9", "fg": "7A2A12", "label": "muted"},
        "font_docx": "Calibri", "font_pdf": "Inter",
        "head_font_docx": "Calibri", "head_font_pdf": "Poppins",
        "body_pt": 9.6, "name_pt": 30, "headline_pt": 12,
        "accent": "E0542C", "text": "1D1B26", "muted": "625E70",
        "margins_cm": (1.8, 1.9, 1.6, 1.9),
        "header": "left", "band": False, "top_rule": 0.35,
        "heading": {"case": "upper", "rule": "none", "size": 10.5, "color": "accent",
                    "spacing_pt": 1.0, "bar": False},
        "order": TECH_ORDER, "photo": True, "photo_shape": "circle", "badge_ring": "E0542C",
    },
    "saale": {
        "label": "Saale: monogram editorial (black & white serif)",
        "inspired_by": "Canva black-and-white minimalist & monogram CVs",
        "best_for": "management, consulting, finance, senior engineering",
        "layout": "classic", "monogram": "left",
        "font_docx": "Calibri", "font_pdf": "Inter",
        "head_font_docx": "Georgia", "head_font_pdf": "Crimson Text",
        "headline_font_pdf": "Crimson Text", "headline_font_docx": "Georgia", "headline_italic": True,
        "body_pt": 9.5, "name_pt": 31, "headline_pt": 13.5,
        "accent": "8A6A3B", "text": "161616", "muted": "5E5E5E",
        "margins_cm": (1.8, 2.0, 1.6, 2.0),
        "header": "left", "band": False,
        "badge_bg": "161616", "badge_fg": "F3EEE4", "badge_font": "CrimsonText-SemiBold.ttf",
        "heading": {"case": "smallcaps", "rule": "bottom-thin", "size": 13.5, "color": "text",
                    "spacing_pt": 0.5, "bar": False},
        "order": COMMON_ORDER, "photo": False,
    },
    "lahn": {
        "label": "Lahn: bold navy header, badge, pill headings",
        "inspired_by": "Canva bold-header photo resumes with label-style headings",
        "best_for": "sales, project/programme management, customer-facing tech",
        "layout": "classic",
        "chips": {"bg": "E7EBFB", "fg": "1E2A78", "label": "muted"},
        "font_docx": "Arial", "font_pdf": "Inter",
        "head_font_docx": "Arial", "head_font_pdf": "Poppins",
        "body_pt": 9.6, "name_pt": 27, "headline_pt": 11.5,
        "accent": "2B44C7", "text": "151A2D", "muted": "5A6178",
        "margins_cm": (0.0, 1.8, 1.6, 1.8),
        "header": "band", "band": True, "band_badge": True, "monogram": True,
        "band_color": "141B34", "band_text": "FFFFFF", "band_sub": "AFC0FF", "band_sep": "4D5FA8",
        "badge_bg": "2B44C7", "badge_fg": "FFFFFF", "badge_ring": "FFFFFF",
        "heading": {"case": "upper", "rule": "none", "size": 9.5, "color": "accent",
                    "spacing_pt": 0.8, "bar": False, "pill": True},
        "order": COMMON_ORDER, "photo": True,
    },
})

HEADINGS = {
    "en": {"competencies": "Core Competencies", "summary": "Profile", "experience": "Work Experience", "education": "Education",
           "skills": "Skills", "languages": "Languages", "certifications": "Certifications",
           "projects": "Projects", "personal": "Personal Details", "present": "Present",
           "dob": "Date of birth", "nationality": "Nationality"},
    "de": {"competencies": "Kernkompetenzen", "summary": "Profil", "experience": "Berufserfahrung", "education": "Ausbildung",
           "skills": "Kenntnisse", "languages": "Sprachen", "certifications": "Zertifikate",
           "projects": "Projekte", "personal": "Persönliche Daten", "present": "heute",
           "dob": "Geburtsdatum", "nationality": "Staatsangehörigkeit"},
}
