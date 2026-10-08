"""
parse_resumes_simple.py  --  read every file the classifier sorted and write ONE JSON FILE PER FILE.

PARSING ONLY: it does not decide whether a file is a resume (that is classify_resumes.py) and it never touches
MeiliSearch. It reads both folders made by the classifier:

    classification/Valid/            ->  Output/Valid/      one JSON per valid resume
    classification/Invalid/<reason>/ ->  Output/Invalid/    one JSON per invalid file, with the REASON inside

The JSON file has the same name as the input file:  "Rahul Sharma CV.pdf"  ->  "Rahul Sharma CV.json".
(If two inputs would give the same JSON name, the second one becomes "Rahul Sharma CV (docx).json".)

Every JSON starts with
    "status"   "valid" or "invalid"          "reason"    why it is invalid (null for valid)
    "category" the classifier's reason folder   "source_file", "parsed" (true if the file could be parsed)
An invalid file that is a readable document (a duplicate, an invoice, a job description ...) is parsed too, so you can
see what it contains. A file that cannot be a resume (a video, a program, a company-list spreadsheet, an empty file)
gets a JSON with the reason only ("parsed": false). A valid CSV / Excel candidate list is parsed row by row into
"candidates".

FIELDS IT EXTRACTS
    name, email, phone, linkedin
    skills, titles, industries
    education   (a degree and an education are the same thing: one entry = degree + institution + year)
    experience  (title, company, start, end) + total years of experience
    projects    (name + the skills used in it)
    internships, certifications, salary (current / expected)

HOW IT WORKS (the numbers match the numbered blocks in this file)
    1. read the file into lines (text + font size + bold)      PDF / DOCX / DOC / RTF / ODT / HTML / images (OCR)
    2. cut the lines into sections using the section headings
    3. contact data      regex finds candidates, libraries check them (email-validator, phonenumbers)
    4. name              scoring of the first lines + spaCy PERSON, combined (see compare_methods.py for the test)
    5. skills            dictionary match (n-grams) + fuzzy match (rapidfuzz) + a confidence score
    6. titles, industries
    7. education, experience, internships, projects, certifications, salary
    8. validation layer  every field is checked before it is written; problems are listed in "issues"
    9. candidate lists   CSV / Excel files with one person per row

HOW TO RUN (PowerShell, from this folder)
    python parse_resumes_simple.py                        # Valid AND Invalid  ->  ../Output/Valid and ../Output/Invalid
    python parse_resumes_simple.py --only valid           # only the valid resumes
    python parse_resumes_simple.py --limit 50             # only the first 50 files of each folder, to try it
    python parse_resumes_simple.py --workers 3
"""

import os
import re
import sys
import json
import time
import shutil
import hashlib
import tempfile
import argparse
import subprocess
import collections
from concurrent.futures import ProcessPoolExecutor, as_completed

import pymupdf                       # PDF reading
import docx                          # DOCX reading
import phonenumbers                  # phone numbers: find + validate + normalise
from email_validator import validate_email, EmailNotValidError   # email check
from rapidfuzz import fuzz, process  # fuzzy matching of near-identical words
import dateparser                    # dates such as "Jan 2020", "03/2021"

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
LOOKUPS = os.path.join(HERE, "lookups")
DEFAULT_CLASSIFICATION = os.path.join(ROOT, "classification")
DEFAULT_OUTPUT = os.path.join(ROOT, "Output")
DEFAULT_INPUT = os.path.join(DEFAULT_CLASSIFICATION, "Valid")          # (compare_methods.py samples from here)


# =====================================================================================================
# 0. SMALL HELPERS
# =====================================================================================================

class Line:
    """One line of the resume: its text, font size (0 if unknown) and whether it is bold."""
    __slots__ = ("text", "size", "bold")

    def __init__(self, text, size=0.0, bold=False):
        self.text, self.size, self.bold = text, size, bold


def long_path(path):
    """Windows needs the \\\\?\\ prefix for very long paths."""
    path = os.path.abspath(path)
    return "\\\\?\\" + path if os.name == "nt" and not path.startswith("\\\\?\\") else path


def clean(text):
    """Collapse whitespace and drop private-use / bullet symbols that PDFs and Word leave in the text."""
    text = re.sub(r"[\ue000-\uf8ff\u2022\u25aa\u25cf\u25a0\u2023\u2043\u27a2\u2756\u2714\u2713\uf0b7\uf0a7\uf0d8]", " ", text)
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def norm_key(text):
    """Key used to compare words: lower case, only letters/digits/+/#, so 'MS-Excel' == 'ms excel' == 'MSExcel'."""
    key = re.sub(r"[^a-z0-9+#]", "", text.lower())
    return ("." + key) if text.strip().startswith(".") else key


# =====================================================================================================
# 1. READING THE FILE  ->  list of Line
# =====================================================================================================

def find_tesseract():
    for path in (r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(path):
            return path
    return shutil.which("tesseract")


def ocr_image(img):
    """Text of a picture with Tesseract OCR."""
    import pytesseract
    path = find_tesseract()
    if path:
        pytesseract.pytesseract.tesseract_cmd = path
    return [Line(clean(t)) for t in pytesseract.image_to_string(img).splitlines() if clean(t)]


def read_pdf(path):
    """Text layer of a PDF (with font size and bold). If there is hardly any text it is a scan: use OCR.
    Hyperlinks are added as lines too, because a 'LinkedIn' word is often a link without visible URL."""
    import io
    from PIL import Image
    pdf = pymupdf.open(path)
    lines, links = [], []
    try:
        for page in pdf:
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    spans = line["spans"]
                    text = clean("".join(s["text"] for s in spans))
                    if text:
                        lines.append(Line(text, round(spans[0]["size"], 1), "bold" in spans[0]["font"].lower()))
            links += [l["uri"] for l in page.get_links() if l.get("uri")]
        if sum(len(l.text) for l in lines) < 100:           # a scanned PDF: read the pages as pictures
            Image.MAX_IMAGE_PIXELS = None
            lines = []
            for number, page in enumerate(pdf):
                if number >= 8:
                    break
                lines += ocr_image(Image.open(io.BytesIO(page.get_pixmap(dpi=200).tobytes("png"))))
            if not lines:                                    # some PDFs hold the scan as a picture inside
                for page in list(pdf)[:4]:
                    for info in page.get_images(full=True):
                        try:
                            picture = Image.open(io.BytesIO(pdf.extract_image(info[0])["image"])).convert("L")
                        except Exception:
                            continue
                        if min(picture.size) > 200:
                            lines += ocr_image(picture)
    finally:
        pdf.close()
    lines += [Line(u) for u in links if "linkedin.com" in u.lower()]
    return lines


def read_docx(path):
    """Paragraphs AND table cells (many resumes are tables), plus header text and hyperlinks."""
    document = docx.Document(path)
    lines = []

    def add_paragraph(p):
        text = clean(p.text)
        if not text:
            return
        runs = [r for r in p.runs if r.text.strip()]
        size = 0.0
        for r in runs:
            try:
                if r.font.size:
                    size = r.font.size.pt
                    break
            except (ValueError, TypeError):               # a malformed size in the file: ignore it
                continue
        bold = bool(runs) and all(r.bold or (p.style is not None and p.style.font.bold) for r in runs)
        lines.append(Line(text, size, bold))

    for section in document.sections:                       # the header often holds name and contact data
        for p in section.header.paragraphs:
            add_paragraph(p)
    for item in document.iter_inner_content():              # paragraphs and tables in document order
        if hasattr(item, "rows"):
            seen = set()
            for row in item.rows:
                for cell in row.cells:
                    if id(cell._tc) in seen:
                        continue
                    seen.add(id(cell._tc))
                    for p in cell.paragraphs:
                        add_paragraph(p)
        else:
            add_paragraph(item)
    for rel in document.part.rels.values():                 # hyperlinks (LinkedIn)
        if rel.is_external and "linkedin.com" in rel.target_ref.lower():
            lines.append(Line(rel.target_ref))
    xml = document.element.xml
    existing = " ".join(l.text for l in lines)
    for box in re.findall(r"<w:txbxContent>(.*?)</w:txbxContent>", xml, flags=re.DOTALL):      # text boxes / frames
        for paragraph in re.findall(r"<w:p[ >].*?</w:p>", box, flags=re.DOTALL):
            text = clean("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", paragraph)))
            if text and text not in existing:
                lines.insert(0, Line(text))
    if sum(len(l.text) for l in lines) < 60:                # nothing read at all: take every piece of text in the XML
        lines += [Line(clean(t)) for t in re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml) if clean(t)]
    return lines


def read_text_file(path):
    raw = open(path, "rb").read()
    for encoding in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            return [Line(clean(t)) for t in raw.decode(encoding).splitlines() if clean(t)]
        except (UnicodeDecodeError, UnicodeError):
            continue
    return []


def sniff_type(path):
    """The real type of a file from its first bytes (the extension can lie)."""
    with open(path, "rb") as f:
        head = f.read(16)
    if head.startswith(b"%PDF"):
        return "pdf"
    if head.startswith(b"PK"):
        return "docx"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "old"                        # .doc
    if head.startswith(b"{\\rtf"):
        return "old"
    if head[:3] == b"\xff\xd8\xff" or head.startswith(b"\x89PNG") or head[:2] == b"BM" or head[:4] in (b"II*\x00", b"MM\x00*"):
        return "image"
    low = open(path, "rb").read(300).lstrip().lower()
    if low.startswith((b"<!doctype html", b"<html", b"<head", b"<meta")):
        return "old"                        # a web page saved as a resume
    return "text"


def find_libreoffice():
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return shutil.which("soffice")


def convert_old_files(paths, temp_folder, target="docx"):
    """DOC / RTF / ODT / HTML -> DOCX with LibreOffice (target='csv': old spreadsheets -> CSV). LibreOffice is slow to
    start, so files are converted in batches. Returns {old path: converted path}."""
    office = find_libreoffice()
    done = {}
    if not office or not paths:
        return done
    os.makedirs(temp_folder, exist_ok=True)
    copies = {}
    for number, path in enumerate(paths):
        is_web = target == "docx" and open(path, "rb").read(300).lstrip().lower().startswith((b"<!doctype", b"<html", b"<head", b"<meta"))
        copy = os.path.join(temp_folder, f"{number:06d}" + (".html" if is_web else (os.path.splitext(path)[1].lower() or ".doc")))
        shutil.copy2(path, copy)
        copies[path] = copy
    for web_pages in (False, True):
        group = [c for c in copies.values() if c.endswith(".html") == web_pages]
        for start in range(0, len(group), 40):
            args = [office, "--headless", "--norestore"]
            if target == "csv":
                args += ["--convert-to", "csv"]
            else:
                args += ["--infilter=HTML (StarWriter)", "--convert-to", "docx:MS Word 2007 XML"] if web_pages else ["--convert-to", "docx"]
            try:
                subprocess.run(args + ["--outdir", temp_folder] + group[start:start + 40], capture_output=True,
                               timeout=60 * len(group[start:start + 40]), check=False)
            except Exception:
                pass
    for path, copy in copies.items():
        out = os.path.splitext(copy)[0] + "." + target
        if os.path.exists(out):
            done[path] = out
    return done


def read_lines(path):
    """Any file -> list of Line. Raises an error with a clear message if nothing can be read."""
    kind = sniff_type(path)
    if kind == "pdf":
        return read_pdf(path)
    if kind == "docx":
        return read_docx(path)
    if kind == "image":
        from PIL import Image
        Image.MAX_IMAGE_PIXELS = None
        return ocr_image(Image.open(path))
    if kind == "old":
        raise ValueError("old format (doc/rtf/odt/html) must be converted first")
    return read_text_file(path)


# =====================================================================================================
# LOOKUPS (loaded once per process)
# =====================================================================================================

class Lookups:
    """The word lists: skills.json, titles.json, industries.json, industry clues, section headings."""

    def __init__(self):
        load = lambda name: json.load(open(os.path.join(LOOKUPS, name), encoding="utf-8"))
        extra = load("skills_extra.json")
        # skills: key -> nice name.  'add' = extra known skills, 'block' = words that are never skills,
        # 'whole_item_only' = ambiguous words (C, R, Go, Excel, Word) accepted only when they are a whole list item
        self.skills = {}
        for name in load("skills.json") + extra.get("add", []):
            key = norm_key(name)
            if len(key) >= 2:
                self.skills.setdefault(key, name)
        for name in extra.get("block", []):
            self.skills.pop(norm_key(name), None)
        self.whole_item_only = {norm_key(n) for n in extra.get("whole_item_only", [])}
        for key in self.whole_item_only:
            self.skills.setdefault(key, key.upper() if len(key) <= 2 else key.title())
        self.known_tech = {norm_key(n) for n in extra.get("add", [])}
        self.skill_buckets = collections.defaultdict(list)       # first letter -> keys (makes fuzzy matching fast)
        for key in self.skills:
            self.skill_buckets[key[:1]].append(key)
        self.titles = {norm_key(t): t for t in load("titles.json") if len(norm_key(t)) >= 4}
        self.industries = load("industries.json")
        clues = load("industry_keywords.json")
        self.industry_of = {}                                      # clue (lower) -> industry
        for industry, words in clues.items():
            if not industry.startswith("_"):
                for w in words:
                    self.industry_of.setdefault(w.lower(), industry)
        self.industry_regex = re.compile(r"\b(?:" + "|".join(re.escape(w) for w in sorted(self.industry_of, key=len, reverse=True)) + r")\b",
                                         re.IGNORECASE)
        # section headings: our own list + the skill-heading variants from 'extra skill.json'
        skill_headings = [h for group in load("extra skill.json").values() for h in group]
        self.headings = {
            "summary": ["summary", "professional summary", "profile", "profile summary", "career objective", "objective", "career objectives",
                        "about me", "career summary", "executive summary", "professional profile", "personal summary"],
            "skills": ["skills", "key skills", "core skills", "skill set", "areas of expertise", "core competencies", "expertise"] + skill_headings,
            "experience": ["experience", "work experience", "professional experience", "employment history", "career history", "work history",
                           "employment", "professional background", "relevant experience", "experience summary", "employment details",
                           "organizational experience", "organisational experience"],
            "education": ["education", "academic qualification", "academic qualifications", "educational qualification", "educational qualifications",
                          "qualification", "qualifications", "academics", "academic background", "education and training", "educational background",
                          "academic details", "education details", "scholastic record", "academic profile"],
            "projects": ["projects", "academic projects", "personal projects", "key projects", "project details", "project experience",
                         "major projects", "projects undertaken", "project summary", "projects handled", "relevant projects", "project",
                         "selected projects", "software projects", "project highlights", "work samples", "portfolio"],
            "internships": ["internship", "internships", "internship experience", "industrial training", "summer training", "internship details",
                            "apprenticeship", "training and internship", "vocational training"],
            "certifications": ["certifications", "certificates", "certification", "licenses and certifications", "professional certifications",
                               "courses", "training and certifications", "certifications and training", "courses and certifications",
                               "courses completed", "online courses", "licenses", "accreditations", "credentials", "professional courses"],
            "personal": ["personal details", "personal information", "personal profile", "declaration", "personal data", "basic details"],
            "other": ["languages", "languages known", "hobbies", "interests", "references", "achievements", "awards", "extra curricular activities",
                      "strengths", "hobbies and interests", "awards and achievements", "extracurricular activities", "activities", "publications"],
        }
        self.heading_of = {}
        for section, words in self.headings.items():
            for w in words:
                self.heading_of[norm_key(w)] = section
        self.heading_keys = list(self.heading_of)
        self._nlp = None

    def nlp(self):
        """spaCy small English model, loaded the first time it is needed (it finds PERSON / ORG / DATE)."""
        if self._nlp is None:
            import spacy
            self._nlp = spacy.load("en_core_web_sm", exclude=["parser", "lemmatizer", "attribute_ruler", "senter"])
        return self._nlp


_LOOKUPS = None


def lookups():
    global _LOOKUPS
    if _LOOKUPS is None:
        _LOOKUPS = Lookups()
    return _LOOKUPS


# =====================================================================================================
# 2. SECTIONS
# =====================================================================================================

def heading_section(text, bold=False, size=0.0, body_size=0.0):
    """If this line is a section heading, which section? Exact match first, then fuzzy for typos and
    lost spaces ('CAREEROBJECTIVE'). Only short lines can be headings."""
    L = lookups()
    text = clean(text).rstrip(":-–— ").strip()
    if not text or len(text) > 45 or len(text.split()) > 6 or re.search(r"[@\d]", text):
        return None
    key = norm_key(text.replace("&", " and "))
    if key in L.heading_of:
        return L.heading_of[key]
    styled = bold or text.isupper() or text.istitle() or (body_size and size > body_size + 0.5)
    if styled and len(text.split()) <= 4 and len(key) >= 6:
        match = process.extractOne(key, L.heading_keys, scorer=fuzz.ratio, score_cutoff=88)
        if match:
            return L.heading_of[match[0]]
    return None


def split_sections(lines):
    """{'top': [...], 'skills': [...], 'experience': [...], ...}. The lines before the first heading are 'top'.
    'Technical Skills: Java, SQL' (a heading with text after the colon) starts a section with that text."""
    sizes = collections.Counter(l.size for l in lines if l.size)
    body_size = sizes.most_common(1)[0][0] if sizes else 0.0
    sections = collections.defaultdict(list)
    current = "top"
    for line in lines:
        found = heading_section(line.text, line.bold, line.size, body_size)
        if found:
            current = found
            continue
        if ":" in line.text:                                  # heading with the first item on the same line
            head, rest = line.text.split(":", 1)
            inline = heading_section(head, True)
            if inline and rest.strip():
                current = inline
                sections[current].append(Line(rest.strip(), line.size, line.bold))
                continue
        sections[current].append(line)
    return sections


# =====================================================================================================
# 3. CONTACT DATA  (regex finds candidates, a library validates them)
# =====================================================================================================

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
LINKEDIN_RE = re.compile(r"(?:https?://)?(?:[a-z]{2,3}\.)?(?:www\.)?linkedin\.com/in/([A-Za-z0-9_%\-]+)", re.IGNORECASE)


def find_email(lines):
    """First VALID email. An address that a PDF wrapped over two lines ('name@gmail.' + 'com') is joined first."""
    texts = [l.text for l in lines]
    joined = [t + texts[i + 1] if i + 1 < len(texts) and ("@" in t) and not EMAIL_RE.search(t) else t for i, t in enumerate(texts)]
    for text in texts + joined:
        for candidate in EMAIL_RE.findall(text.replace(" @", "@").replace("@ ", "@")):
            candidate = re.sub(r"^(?:e-?mail(?:-?id)?|mail(?:-?id)?)[:\-_]*", "", candidate.strip(".,;:-"), flags=re.IGNORECASE)
            try:
                return validate_email(candidate, check_deliverability=False).normalized
            except EmailNotValidError:
                continue
    return None


def find_phone(lines):
    """First valid phone number. phonenumbers reads many formats (+91, spaces, hyphens, brackets) and checks
    that the number can really exist. Default country: India. The number is returned in international format."""
    top, rest = lines[:40], lines[40:]
    for group in (top, rest):
        text = "\n".join(l.text for l in group)
        for match in phonenumbers.PhoneNumberMatcher(text, "IN"):
            number = match.number
            if phonenumbers.is_valid_number(number):
                return phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.E164)
    return None


def find_linkedin(lines):
    """LinkedIn profile URL, normalised to https://www.linkedin.com/in/username. A URL cut over two lines is joined."""
    texts = [l.text for l in lines]
    for i, text in enumerate(texts):
        m = LINKEDIN_RE.search(text)
        if not m:
            continue
        slug = m.group(1)
        nxt = texts[i + 1].strip() if i + 1 < len(texts) else ""
        if (not slug or slug.endswith("-")) and re.fullmatch(r"[A-Za-z0-9_%\-]{3,}", nxt):
            slug += nxt
        if len(slug) >= 3:
            return "https://www.linkedin.com/in/" + slug.strip("-/")
    return None


# =====================================================================================================
# 4. NAME   (two methods compared in compare_methods.py, then combined)
# =====================================================================================================

NAME_BAD_WORDS = {"resume", "curriculum", "vitae", "cv", "profile", "summary", "objective", "contact", "address", "email", "mobile", "phone",
                  "experience", "education", "skills", "projects", "declaration", "personal", "details", "career", "page", "date", "name",
                  "thanking", "yours", "sincerely", "regards", "references", "certificate", "professional", "technical", "information",
                  "software", "engineer", "developer", "manager", "analyst", "consultant", "executive", "fresher", "linkedin", "github",
                  "knowledge", "road", "street", "nagar", "colony", "creating", "families", "market", "building", "apartment", "flat",
                  "plot", "village", "district", "state", "country", "nationality", "gender", "religion", "marital", "status", "father",
                  "mother", "hobbies", "strengths", "languages", "known", "technology", "technologies", "business", "development",
                  "management", "marketing", "finance", "sales", "operations", "general", "senior", "junior", "team", "leader", "seeking",
                  "looking", "dedicated", "highly", "motivated", "passionate", "result", "oriented", "overview", "key", "core", "areas",
                  "work", "working", "career", "job", "position", "role", "responsible", "responsibilities", "training", "internship",
                  "projects", "project", "academic", "qualification", "qualifications", "achievements", "awards", "interests", "other",
                  "reference", "available", "immediate", "joiner", "notice", "period", "current", "expected", "salary", "ctc", "total",
                  "relevant", "digital", "cloud", "data", "science", "system", "systems", "network", "security", "design", "quality"}
COMPANY_WORDS = re.compile(r"\b(pvt|ltd|limited|inc|llc|llp|corp|corporation|company|technologies|solutions|services|systems|industries|"
                           r"consultancy|enterprises|bank|university|college|institute|school|hospital)\b", re.IGNORECASE)
TITLE_HEAD = re.compile(r"\b(engineer|developer|manager|executive|officer|consultant|analyst|lead|head|specialist|associate|assistant|architect|"
                        r"designer|supervisor|coordinator|administrator|programmer|tester|technician|accountant|director|scientist|"
                        r"recruiter|representative|intern|trainee|president|vp|owner|founder|partner|auditor|advisor|planner|buyer|"
                        r"operator|clerk|cashier|agent|teacher|lecturer|professor|trainer|writer|editor|nurse|doctor|pharmacist)s?\b", re.IGNORECASE)
INDIAN_PLACES = {"bangalore", "bengaluru", "mumbai", "delhi", "hyderabad", "chennai", "pune", "kolkata", "noida", "gurgaon", "gurugram", "ahmedabad",
                 "jaipur", "lucknow", "kochi", "indore", "bhopal", "nagpur", "india", "karnataka", "maharashtra", "tamilnadu", "telangana", "kerala",
                 "gujarat", "rajasthan", "uttar", "pradesh", "dubai", "riyadh", "saudi", "arabia", "uae", "usa", "london", "singapore"}


def name_candidates(lines, email, filename):
    """METHOD A  - scoring.  Look at the first meaningful lines, drop what is clearly not a name (email, phone, URL,
    headings, job titles, companies, places) and score the rest by position, font size, capitalisation, number of
    words, and whether the words also appear in the email address or file name."""
    L = lookups()
    local = re.sub(r"[^a-z]", "", (email or "").split("@")[0].lower())
    base = re.sub(r"[^a-z]", "", os.path.splitext(os.path.basename(filename))[0].lower())
    sizes = [l.size for l in lines[:25] if l.size]
    biggest = max(sizes) if sizes else 0
    typical = sorted(sizes)[len(sizes) // 2] if sizes else 0
    out = []
    for i, line in enumerate(lines[:25]):
        text = clean(line.text)
        if re.fullmatch(r"(?:[A-Za-z]\s){3,}[A-Za-z]", text):             # letter-spaced: 'M U R A L I' -> 'MURALI'
            text = text.replace(" ", "")
        # a line that is only the label: 'NAME' / 'Name :' -> the name is on the next line (': BATTULA NARENDER KUMAR')
        labelled = bool(re.match(r"^name\s*[:\-]", text, re.IGNORECASE)) or (i > 0 and re.fullmatch(r"name\s*:?", clean(lines[i - 1].text), re.IGNORECASE) is not None)
        text = re.sub(r"^[:\-\s]+", "", text)
        text = re.sub(r"^(?:name\s*[:\-]|resume of|cv of|curriculum vitae of)\s*", "", text, flags=re.IGNORECASE)
        # 'Rakesh V Email: x@y.com'  /  'Sasi .K Email-id: ...'  -> keep what is before the contact label
        text = re.split(r"\s+(?:e-?mail|email-?id|mobile|mob\.?|phone|ph\.?|contact|tel|cell|address)\b", text, flags=re.IGNORECASE)[0]
        text = re.sub(r"^(?:mr|mrs|ms|miss|dr|shri|smt)\.?\s+", "", text, flags=re.IGNORECASE)
        text = re.split(r"\s*[|,]\s*|\s+[-–—]\s+|\s{3,}", text)[0].strip(" .-")      # "Name | Software Engineer" -> "Name"
        if len(text) < 3 or re.search(r"[@\d:/\\()]|www|http", text) or heading_section(text) or PROSE_START_RE.match(text):
            continue
        words = text.split()
        if not 1 <= len(words) <= 4 or not all(re.fullmatch(r"\.?[A-Za-z][A-Za-z.'\-]*", w) for w in words):
            continue
        # every word starts with a capital letter (except name particles): 'A versatile', 'Second Class sti' are not names
        if not all(w.lstrip(".")[:1].isupper() or w.lower() in {"van", "de", "der", "bin", "al", "la", "di", "da", "von", "bhai", "ben"} for w in words):
            continue
        if {w.lower().strip(".") for w in re.split(r"[\s\-]+", text)} & (NAME_BAD_WORDS | {"add", "winning", "dob"}):
            continue
        if len(words) == 1 and norm_key(text) in L.skills:           # 'JavaScript', 'Python' on its own line is a skill
            continue
        score = 1.0 - i * 0.06                                        # earlier is better
        score += 1.2 if labelled else 0                               # 'Name: X'  is as sure as it gets
        score += 0.5 if biggest and line.size >= biggest - 0.5 and line.size > typical else 0
        score += 0.1 if line.bold else 0
        score += 0.2 if text.isupper() or text.istitle() else -0.2
        score += {1: -0.3, 2: 0.4, 3: 0.4, 4: 0.1}[len(words)]
        letters = [re.sub(r"[^a-z]", "", w.lower()) for w in words]
        if any(len(w) >= 3 and (w in local or w in base) for w in letters):
            score += 0.7                                             # the email / file name agrees
        if COMPANY_WORDS.search(text) or TITLE_HEAD.search(text) or norm_key(text) in L.titles:
            score -= 1.0
        if norm_key(text) in L.skills:
            score -= 0.8
        if {w.lower() for w in words} & INDIAN_PLACES:
            score -= 0.8
        out.append((round(score, 2), text))
    return sorted(out, reverse=True)


def name_by_scoring(lines, email, filename):
    candidates = name_candidates(lines, email, filename)
    return (candidates[0][1], candidates[0][0]) if candidates else (None, 0.0)


def name_by_spacy(lines):
    """METHOD B - spaCy named-entity recognition: the first PERSON in the top lines."""
    text = "\n".join(clean(l.text) for l in lines[:12] if not EMAIL_RE.search(l.text))
    for ent in lookups().nlp()(text).ents:
        words = ent.text.split()
        if ent.label_ == "PERSON" and 1 <= len(words) <= 4 and not re.search(r"[\d@]", ent.text) \
                and not ({w.lower() for w in words} & NAME_BAD_WORDS) and not COMPANY_WORDS.search(ent.text):
            return clean(ent.text), 0.6
    return None, 0.0


def name_from_email_or_file(email, filename):
    """Last resort: 'rahul.sharma92@gmail.com' -> 'Rahul Sharma'."""
    for source in ((email or "").split("@")[0], re.sub(r"(?i)resume|cv|naukri|\d+|\[.*?\]|\(.*?\)", " ", os.path.splitext(os.path.basename(filename))[0])):
        parts = [p for p in re.split(r"[._\-\s]+", re.sub(r"\d+", "", source)) if len(p) >= 2 and p.lower() not in ("email", "mail", "id", "gmail", "yahoo")]
        if 1 <= len(parts) <= 3 and all(p.isalpha() for p in parts):
            return " ".join(p.capitalize() for p in parts), 0.3
    return None, 0.0


FILE_NOISE = re.compile(r"(?i)\b(?:resume|resumes|cv|curriculum|vitae|updated|final|new|latest|copy|profile|download|docx?|pdf|"
                        r"jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b")


def name_from_file(filename):
    """A fuller name from the FILE NAME, for resumes whose own text gives only a first name or a glued word.
    'Rutvik CV - Rutwik Dave.docx' -> 'Rutwik Dave';  'Naukri_RashmiJori[11y_0m].pdf' -> 'Rashmi Jori';
    'PalavesKumar_3.2yrs_iPhone_Bangalore.pdf' -> 'Palaves Kumar'.  Returns None when no 2-4 word name can be read."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = re.sub(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-", "", stem)       # upload id in front
    stem = re.sub(r"^\d{10,}", "", stem)
    if " - " in stem:                                                  # 'Resume1 - Shivu Kumar h m' : the part after the dash is the name
        tail = stem.rsplit(" - ", 1)[1].split()
        if 2 <= len(tail) <= 4 and all(re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*", w) for w in tail):
            return " ".join(w.capitalize() if (w.islower() or w.isupper()) else w for w in tail)
    stem = re.sub(r"\[.*?\]|\(.*?\)", " ", stem)
    stem = re.sub(r"(?i)^(?:naukri|updazz|iimjobs|hirist|resume|cv)[_\- .]*", "", stem)
    stem = re.sub(r"\d+(?:\.\d+)?\s*(?:yrs?|years?)\b.*$|\d+y[_ ]?\d+m.*$", " ", stem, flags=re.IGNORECASE)
    stem = FILE_NOISE.sub(" ", re.sub(r"[_.+\-]+", " ", stem))
    stem = re.sub(r"\d+", " ", stem)
    stem = re.sub(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", stem)                       # CamelCase -> words
    words = [w for w in stem.split() if re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*", w) and len(w.strip(".")) >= 2]
    if 2 <= len(words) <= 4 and sum(len(w.strip(".")) < 3 for w in words) <= 1 \
            and not ({w.lower().strip(".") for w in words} & (NAME_BAD_WORDS | INDIAN_PLACES | {"iphone", "android", "java", "net"})):
        return " ".join(w.capitalize() if (w.islower() or w.isupper()) else w for w in words)
    return None


def looks_like_person(name):
    """A 2-4 word name made only of ordinary name words: no company / job / place / skill / heading word. Such a name from the
    top of the resume is never replaced by the file name."""
    L = lookups()
    words = name.split()
    low = {w.lower().strip(".,") for w in words}
    return (2 <= len(words) <= 4 and not (low & (NAME_BAD_WORDS | INDIAN_PLACES)) and not COMPANY_WORDS.search(name) and not TITLE_HEAD.search(name)
            and norm_key(name) not in L.skills and norm_key(name) not in L.titles and not PROSE_START_RE.match(name))


def better_with_file(name, confidence, filename):
    """A one-word or e-mail-derived name is replaced by the file name's fuller name when the two share a word."""
    full = name_from_file(filename)
    if not full:
        return name, confidence
    if not name:
        return full, 0.3
    joined = re.sub(r"[^a-z]", "", name.lower())
    parts = [re.sub(r"[^a-z]", "", w.lower()) for w in full.split()]
    if len(name.split()) == 1 and (joined == "".join(parts) or any(len(w) >= 3 and w in joined for w in parts)):
        return full, max(confidence, 0.5)
    return name, confidence


def agrees_with_contact(name, email, filename):
    """Do the words of this name also appear in the email address or the file name? (Real names usually do.)"""
    local = re.sub(r"[^a-z]", "", (email or "").split("@")[0].lower())
    base = re.sub(r"[^a-z]", "", os.path.splitext(os.path.basename(filename))[0].lower())
    words = [re.sub(r"[^a-z]", "", w.lower()) for w in (name or "").split()]
    return any(len(w) >= 3 and (w in local or w in base or (len(w) >= 5 and (w[:4] in local or w[:4] in base))) for w in words)


def find_name(lines, email, filename):
    """COMBINED:  1. the scoring method, if it is sure (score >= 1.0);
                  2. otherwise spaCy's PERSON, but only if the email / file name agrees with it (spaCy alone calls
                     'Creating Families' a person);
                  3. otherwise a name built from the email or file name ('rahul.sharma92@gmail.com' -> 'Rahul Sharma');
                  4. otherwise whatever the scoring method found, with low confidence."""
    name, score = name_by_scoring(lines, email, filename)
    if name and email and not agrees_with_contact(name, email, filename) and not looks_like_person(name):    # the text gives no person, the e-mail AND the file name agree on one
        from_file = name_from_file(filename)
        if from_file and agrees_with_contact(from_file, email, ""):
            return from_file, 0.6
    if name and score >= 1.0:
        return better_with_file(name, min(1.0, score / 2), filename)
    spacy_name, spacy_score = name_by_spacy(lines)
    if spacy_name and agrees_with_contact(spacy_name, email, filename):
        return better_with_file(spacy_name, spacy_score, filename)
    if name and agrees_with_contact(name, email, filename):
        return better_with_file(name, min(1.0, max(score, 0.1) / 2), filename)
    from_file = name_from_file(filename)
    if from_file and not (name and looks_like_person(name)) and (not email or agrees_with_contact(from_file, email, "")):
        return from_file, 0.5                                       # e-mail and file name agree: better than a glued e-mail word
    from_contact = name_from_email_or_file(email, filename)
    if from_contact[0]:
        return better_with_file(from_contact[0], from_contact[1], filename)
    if name:
        return name, min(1.0, max(score, 0.1) / 2) * 0.5
    return from_file, (0.3 if from_file else 0.0)


# =====================================================================================================
# 5. SKILLS   (dictionary n-grams + fuzzy + confidence score)
# =====================================================================================================

TOKEN_RE = re.compile(r"\.?[A-Za-z0-9][A-Za-z0-9+#./&'\-]*")


def find_terms(text, table, max_words=4):
    """All dictionary terms in a piece of text: try 4-word, 3-word, 2-word, 1-word phrases at every position
    and keep the longest match. Returns [(key, original words)]. 'table' maps norm_key -> nice name."""
    tokens = TOKEN_RE.findall(text)
    found, i = [], 0
    while i < len(tokens):
        hit = None
        for n in range(min(max_words, len(tokens) - i), 0, -1):
            words = tokens[i:i + n]
            key = norm_key(" ".join(words))
            if key in table:
                hit = (key, " ".join(words), n)
                break
        if hit:
            found.append(hit[:2])
            i += hit[2]
        else:
            i += 1
    return found


SKILL_ITEM_BAD = re.compile(r"\b(skills?|environment|responsibilit\w*|experience|objective|seeking|looking|years?|hobbies|strengths?|"
                            r"declaration|date of birth|passionate|motivated|dedicated|ability|willing)\b", re.IGNORECASE)


def skill_items(lines):
    """The items of a skills list: split on commas, bullets, pipes, semicolons."""
    items = []
    for line in lines:
        for part in re.split(r"[,;|•·●▪]|\s{2,}|\s-\s", line.text):
            part = re.sub(r"^[\w &/+#.\-]{2,40}:\s*", "", part.strip())      # drop a label such as 'Languages:'
            part = clean(part).strip(" .:-")
            if 2 <= len(part) <= 45 and re.search(r"[A-Za-z]", part) and not SKILL_ITEM_BAD.search(part) and len(line.text) <= 160:
                items.append(part)
    return items


def find_skills(sections, use_fuzzy=True):
    """Skill -> confidence.  Where a skill was found decides how sure we are:
         in the Skills section            0.9      (the list is meant to hold skills)
         in experience / projects text    0.4  (+0.3 if it is a multi-word skill or well-known tech, +0.2 if repeated)
         fuzzy match of a skills-list item 0.8 x similarity
       Only skills with confidence >= 0.5 are kept. Ambiguous one-letter / common words (C, R, Go, Excel ...)
       count only when they are a whole item of the skills list."""
    L = lookups()
    scores = {}

    def add(name, value):
        scores[name] = max(scores.get(name, 0), value) if value < 1.0 else value

    section_lines = sections.get("skills", [])
    items = skill_items(section_lines)
    exact_keys = set()
    for item in items:                                               # 1. skills section, item by item
        key = norm_key(item)
        if key in L.skills:
            add(L.skills[key], 0.9)
            exact_keys.add(key)
        else:
            for k, words in find_terms(item, L.skills):
                nice = L.skills[k]
                if k in L.whole_item_only or not (" " in nice.strip() or k in L.known_tech or re.search(r"[A-Z].*[A-Z]|\d|[#+.]", nice[1:])):
                    continue                                        # a generic single word cut out of a longer item ('Cyber' of 'Cyber Crime')
                add(nice, 0.9)
                exact_keys.add(k)
    if use_fuzzy:                                                    # 2. near matches: 'MS-Excell' -> 'MS Excel'
        for item in items[:40]:
            key = norm_key(item)
            if key in L.skills or len(key) < 8 or not 2 <= len(item.split()) <= 4:     # multi-word items only: 'Pivot Table' -> 'Pivot Tables'
                continue
            match = process.extractOne(key, L.skill_buckets.get(key[:1], []), scorer=fuzz.ratio, score_cutoff=92)
            if match and match[0] not in L.whole_item_only:
                add(L.skills[match[0]], round(0.8 * match[1] / 100, 2))
    body = []                                                        # 3. skills mentioned in the rest of the resume
    for name in ("experience", "projects", "internships", "summary", "top"):
        body += [l.text for l in sections.get(name, [])]
    counts = collections.Counter()
    for k, words in find_terms("\n".join(body), L.skills):
        if k in L.whole_item_only:
            continue
        counts[k] += 1
    for k, n in counts.items():
        nice = L.skills[k]
        score = 0.4
        if k in L.known_tech or re.search(r"[A-Z].*[A-Z]|\d|[#+.]", nice[1:]):     # well-known tech, or written like one (AWS, C++, Node.js)
            score += 0.3
        if n >= 2:
            score += 0.2
        if k in exact_keys:
            score = max(score, 0.9)
        if nice not in scores or scores[nice] < score:
            scores[nice] = round(score, 2)
    return {name: s for name, s in sorted(scores.items(), key=lambda kv: -kv[1]) if s >= 0.5}


# =====================================================================================================
# 6. TITLES and INDUSTRIES
# =====================================================================================================

def find_titles(lines, sections, use_fuzzy=False):
    """Job titles: match the titles list in the headline (top lines), the summary and experience lines.
    A title must end with an occupation word (Engineer, Manager, Executive ...). With use_fuzzy a near match
    ('Sr. Softwere Engineer') is accepted too."""
    L = lookups()
    pool = [l.text for l in sections.get("top", [])[:10]] + [l.text for l in sections.get("summary", [])[:6]]
    pool += [l.text for l in sections.get("experience", [])[:200]]
    found = {}
    for text in pool:
        if len(text.split()) > 14 or EMAIL_RE.search(text):
            continue
        for key, words in find_terms(text, L.titles, max_words=6):
            name = L.titles[key]
            if TITLE_HEAD.search(name.split()[-1] + " ") and (len(name.split()) >= 2 or text in [l.text for l in sections.get("top", [])[:6]]):
                found.setdefault(key, name)
        # pattern rule: 'Senior Account Executive' is a title even when titles.json lacks it
        words = text.split()
        if 2 <= len(words) <= 5 and TITLE_HEAD.search(words[-1] + " ") and (text.istitle() or text.isupper()) \
                and not COMPANY_WORDS.search(text) and not re.search(r"[\d@:]", text) and not PROSE_START_RE.match(text) \
                and not ({w.lower().strip(".,") for w in words} & NAME_BAD_WORDS - {"senior", "junior", "general", "digital", "software", "business"}):
            found.setdefault(norm_key(text), clean(text).title() if text.isupper() else clean(text))
        if use_fuzzy and len(text.split()) <= 8 and TITLE_HEAD.search(text):
            key = norm_key(text)
            match = process.extractOne(key, [k for k in L.titles if k[:1] == key[:1]], scorer=fuzz.ratio, score_cutoff=93)
            if match:
                found.setdefault(match[0], L.titles[match[0]])
    # drop a title that is only a part of a longer one found in the same resume
    names = list(found.values())
    return [n for n in names if not any(n != m and n.lower() in m.lower() for m in names)][:8]


def find_industries(lines, sections, use_fuzzy=False):
    """Industries from clue words ('banking', 'thermal power plant', 'Infosys', 'Tata Motors' ...) in the experience /
    summary / project text -- never in education, where a college name would read as 'Education'. The industry that
    is mentioned most comes first. (A fuzzy match against industries.json was tested and added nothing.)"""
    L = lookups()
    text = "\n".join(l.text for name in ("experience", "projects", "internships", "summary", "top") for l in sections.get(name, []))
    counts = collections.Counter(L.industry_of[m.group(0).lower()] for m in L.industry_regex.finditer(text))
    if use_fuzzy:                                           # kept only so compare_methods.py can show it adds nothing
        for line in lines:
            m = re.match(r"\s*(?:industry|domain|sector)\s*[:\-]\s*(.+)", line.text, re.IGNORECASE)
            if m:
                for part in re.split(r"[,;/|]", m.group(1)):
                    match = process.extractOne(part.strip(), L.industries, scorer=fuzz.WRatio, score_cutoff=88)
                    if match:
                        counts[match[0]] += 2
    return [name for name, _ in counts.most_common(5)]


# =====================================================================================================
# 7. EDUCATION, EXPERIENCE, INTERNSHIPS, PROJECTS, CERTIFICATIONS, SALARY
# =====================================================================================================

# Degrees. The short forms BE / ME / BA / MA / CA must be capitals (otherwise they are ordinary words).
DEGREE_RE = re.compile(
    r"\b(?:(?i:b\.?\s?tech|m\.?\s?tech|b\.?\s?sc|m\.?\s?sc|b\.?\s?com|m\.?\s?com|b\.?\s?pharm|m\.?\s?pharm|b\.?\s?arch|b\.?\s?b\.?\s?a|m\.?\s?b\.?\s?a|b\.?\s?c\.?\s?a|m\.?\s?c\.?\s?a|"
    r"llb|llm|mbbs|pgdm|pgdca|ph\.?\s?d|bachelors?(?=\s+(?:of|in|degree|\())|bachelor['’]s|masters?(?=\s+(?:of|in|degree|\())|master['’]s|doctorate|diploma|post[- ]?graduat\w*|graduat(?:e|ion)|"
    r"intermediate|hsc|ssc|sslc|puc|10th|12th|x{1,2}th|iti|b\.?\s?e\.?\s?d|m\.?\s?s\.?\s?w|b\.?\s?s\.?\s?w|cma|"
    r"matriculation|secondary school|senior secondary|higher secondary)|B\.?E\.?|M\.?E\.?|B\.?A\.?|M\.?A\.?|C\.?A\.?)(?![A-Za-z\-])")
INSTITUTION_RE = re.compile(r"\b(university|college|collage|institute|institue|school|polytechnic|academy|iit|nit|iiit|iim|vidyalaya|vidyapeeth|mahavidyalaya|"
                            r"campus|board|institution|govt|government|centre|center|cbse|icse|junior|"
                            r"high school|sr\.? sec|public school|deemed|autonomous|symbiosis|amity|manipal|bits)\b", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
SCORE_RE = re.compile(r"(\d{2}(?:\.\d+)?\s?%|(?:cgpa|gpa|sgpa)\s*[:\-]?\s*\d(?:\.\d+)?(?:\s*/\s*10)?|\d(?:\.\d+)?\s*/\s*10|"
                      r"first class(?: with distinction)?|distinction|second class|merit)", re.IGNORECASE)
MONTHS = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
MONTHS_STRICT = (r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
                 r"nov(?:ember)?|dec(?:ember)?)\b\.?")
_DATE_PART = (r"(?:" + MONTHS + r"[\s,'’\-]*)?(?:19|20)\d{2}(?:[\s,]*" + MONTHS_STRICT + r")?|\d{1,2}[/.\-](?:19|20)\d{2}|" + MONTHS + r"\s*['’]\s*\d{2}\b")
_END_WORDS = r"present|current|till date|till now|till|now|ongoing|date|today"
DATE_RANGE_RE = re.compile(
    r"(?P<start>" + _DATE_PART + r")\s*(?:-|–|—|−|~|to|till|until|through|thru)\s*(?P<end>" + _END_WORDS + r"|" + _DATE_PART + r")",
    re.IGNORECASE)
SINCE_RE = re.compile(r"\b(?:since|from|w\.?e\.?f\.?|joined(?:\s+in)?)\s+((?:\d{1,2}(?:st|nd|rd|th)?\s+)?(?:" + MONTHS_STRICT + r"\s*,?\s*)?(?:19|20)\d{2})\b", re.IGNORECASE)
ONE_DATE_RE = re.compile(r"^\W*(?:" + _DATE_PART + r"|" + _END_WORDS + r")\W*$", re.IGNORECASE)
LABEL_COMPANY_RE = re.compile(r"^\W*(?:name of (?:the )?)?(?:employer|company|organi[sz]ation)(?:\s+name)?\s*[:\-–]\s*(.+)$", re.IGNORECASE)
LABEL_TITLE_RE = re.compile(r"^\W*(?:designation|position|role|job title|title)\s*[:\-–]\s*(.+)$", re.IGNORECASE)
COMPANY_SUFFIX_RE = re.compile(r"\b(pvt\.?|private|ltd\.?|limited|llp|inc\.?|llc|corp\.?|corporation|co\.|company|technologies|technology|solutions|"
                               r"systems|services|industries|labs?|infotech|software|consultancy|consulting|enterprises|group|bank|motors|"
                               r"global|international|associates|partners|studio|networks|digital)\b", re.IGNORECASE)
PROSE_START_RE = re.compile(r"^(developed|designed|created|implemented|worked|working|used|using|built|responsible|involved|handled|performed|"
                            r"managed|prepared|analy[sz]ed|conducted|led|leading|maintained|tested|wrote|integrated|deployed|configured|"
                            r"automated|identified|improved|reduced|achieved|process|provide|provided|ensure|ensuring|support|supported|"
                            r"monitor|monitoring|assist|assisted|coordinate|coordinated|the|this|we|i|my|our|a|an|to|for|with|as|by|"
                            r"managing|handling|preparing|coordinating|maintaining|developing|designing|creating|providing|performing|conducting|"
                            r"analyzing|analysing|reviewing|resolving|liaising|implementing|testing|documenting|collaborating|understanding|"
                            r"gathering|delivering|supporting|training|mentoring|responsible|key|major|actively|participated|ensured|experience|"
                            r"studied|improved|increased|utili[sz]ed|leveraged|contributed|resolved|optimi[sz]ed|migrated|troubleshooting|patching|used)\b", re.IGNORECASE)


GENERIC_INSTITUTION = re.compile(r"^(?:name of (?:the )?)?(?:college|school|university|institute|institution|board)(?:\s*[/&,]\s*(?:college|school|university|institute|institution|board))*"
                                 r"(?:\s+name)?$|\b(?:year of|passing|percentage|completion)\b", re.IGNORECASE)
AMBIGUOUS_DEGREE = {"BE", "ME", "BA", "MA", "CA", "B.E", "M.E", "B.A", "M.A", "C.A"}
AMBIGUOUS_OK = re.compile(r"^\s*(?:\(|[-–:,|.]|$|in\b|of\b|\d|(?:computer|electronics|electrical|mechanical|civil|information|chemical|instrumentation|"
                          r"production|industrial|automobile|biotech\w*|aeronautical|telecom\w*|it|cse|ece|eee|mech|economics|english|history|arts|"
                          r"commerce|political|sociology|psychology|accountancy)\b)", re.IGNORECASE)
ACRONYM_SKIP = {"CGPA", "SGPA", "GPA", "CSE", "ECE", "EEE", "MECH", "CIVIL", "MBA", "BCA", "MCA", "PGDM", "HSC", "SSC", "SSLC", "PUC", "CBSE", "ICSE"}


def real_institution(text):
    """An institution line: has a school word and is more than the generic table header ('College/University', 'Board')."""
    text = clean(text).strip(" .,-:")
    return bool(text) and 3 < len(text) <= 90 and bool(INSTITUTION_RE.search(text)) and not GENERIC_INSTITUTION.search(text) \
        and not DEGREE_RE.fullmatch(text) and not re.search(r"@|https?://", text)


def find_education(sections, all_lines, use_spacy=False):
    """Education = degree + institution + year. Look in the Education section; if the resume has none, scan every
    line for degree words. For each line with a degree:
        degree       the degree words and up to 6 words after them ('B.Tech in Computer Science')
        institution  the part of the line after 'from' / ',' / '-' that has a word like University / College / School,
                     else the nearest such line (the one before, or the two after)
        year         the last 4-digit year in the same small window"""
    lines = sections.get("education") or all_lines
    texts = [l.text for l in lines]
    entries, seen = [], set()
    for i, text in enumerate(texts):
        m = DEGREE_RE.search(text)
        if not m or (PROSE_START_RE.match(text) and len(text.split()) > 12):
            continue
        if m.group(0).rstrip(".") in AMBIGUOUS_DEGREE and (not AMBIGUOUS_OK.match(text[m.end():]) or (
                len(text.split()) > 9 and not INSTITUTION_RE.search(text) and not YEAR_RE.search(text))):
            continue                                                # 'ME Thailand and Power Line Connection between ...'
        rest = text[m.start():]
        degree_text = re.split(r"[,|;(:\[]|\s/\s|\s{2,}|\s[-–—]\s|\b(?:from|at|with|in the year|in\s+(?:19|20)\d{2})\b|\b(?:" + MONTHS + r"\s*)?(?:19|20)\d{2}\b",
                               rest, flags=re.IGNORECASE)[0]
        institution = None
        words = degree_text.split()                                 # 'B.Sc. ACHARYA NAGARJUNA UNIVERSITY' / 'B.Tech in CSE in Andhra University'
        key_at = next((k for k, w in enumerate(words) if k > 0 and INSTITUTION_RE.search(w)), None)
        if key_at is not None:
            begin = key_at
            while begin > 1 and (key_at - begin) < 5 and words[begin - 1][:1].isupper() and words[begin - 1].lower() not in ("in", "of", "and", "from", "at") \
                    and not DEGREE_RE.match(words[begin - 1]):
                begin -= 1
            end = key_at + 1
            if end + 1 < len(words) and words[end].lower() in ("of", "for"):
                end += 2
            candidate = clean(" ".join(words[begin:end])).strip(" .,-:")
            if begin >= 1 and real_institution(candidate):
                institution = candidate
                words = words[:begin]
                while words and words[-1].lower() in ("in", "from", "at", "of"):
                    words.pop()
                degree_text = " ".join(words)
        degree = clean(" ".join(degree_text.split()[:7])).strip(" .-:)")
        if len(degree) < 2 or INSTITUTION_RE.fullmatch(degree):
            continue
        # 1. the institution may be on the same line: 'B.Tech from Avanthi College of Engineering, Vizag in 2014'
        tail = re.sub(DATE_RANGE_RE, " ", rest[len(" ".join(degree_text.split())):] if institution else rest[len(degree_text):])
        tail = re.sub(r"\([^)]*\)", " ", tail)
        tail = re.sub(r"\b(?:" + MONTHS + r"\s*)?(?:19|20)\d{2}\b", " ", tail, flags=re.IGNORECASE)
        segments = [re.sub(r"^(?:in|at|from|of)\s+|\s+(?:in|of|at)$", "", clean(re.split(r"\bwith\b|\(|\d{2}(?:\.\d+)?\s?%|\bcgpa\b", seg, flags=re.IGNORECASE)[0]).strip(" .,-:[]/"), flags=re.IGNORECASE)
                    for seg in re.split(r"\bfrom\b|\s[-–—|:]\s|[:|]|\s/\s|\]|\s{2,}|,\s*(?=[A-Z])", tail, flags=re.IGNORECASE)]
        for segment in ([] if institution else segments):
            if real_institution(segment):
                institution = segment
                break
        if not institution:                                         # '10th from UP Board in 2013' / 'BCA from MMHAPU PATNA in 2018' / '[Nashik Board]'
            fm = re.search(r"\bfrom\s+(?:M/s\s+)?([A-Z][^,()\[\]]*?)(?=\s+in\s+(?:19|20)\d{2}|\s+with\b|\s*[,(\[]|\s+(?:19|20)\d{2}|\s*$)", rest)
            if fm:
                cand = re.sub(r"\s+(?:in|of|at)$", "", clean(fm.group(1)).strip(" .,-:"))
                if 2 <= len(cand) <= 60 and len(cand.split()) <= 6 and not DEGREE_RE.match(cand) and not YEAR_RE.search(cand) \
                        and not (set(w.lower().strip(".") for w in cand.split()) <= INDIAN_PLACES) and not SCORE_RE.search(cand):
                    institution = cand
        if not institution:                                         # 'B.E. (CSE) | RGPV, Bhopal | 2016': an acronym is the institution
            for segment in segments:
                if re.fullmatch(r"[A-Z]{3,10}", segment) and segment not in ACRONYM_SKIP and not DEGREE_RE.fullmatch(segment):
                    institution = segment
                    break
        # 2. otherwise the nearest line that has an institution word and no degree
        if not institution:
            for j in (i - 1, i + 1, i + 2):
                if 0 <= j < len(texts) and real_institution(texts[j]) and not DEGREE_RE.search(texts[j]):
                    institution = clean(re.split(r"\s[-–—|]\s|\b(?:19|20)\d{2}\b|\(", texts[j])[0]).strip(" .,-:")[:90] or None
                    break
        if not institution and use_spacy:                           # optional spaCy layer (tested: mostly noise, so it is off by default)
            for ent in lookups().nlp()(" . ".join(texts[max(0, i - 1):i + 3])).ents:
                if ent.label_ in ("ORG", "FAC") and 2 < len(ent.text) <= 80 and not DEGREE_RE.fullmatch(ent.text.strip()):
                    institution = clean(ent.text)
                    break
        if institution and (norm_key(institution) in norm_key(degree) or norm_key(degree) in norm_key(institution)):
            institution = None                                      # the 'institution' is just the degree text again
        window = " ".join(texts[max(0, i - 1):i + 3])
        years = YEAR_RE.findall(" ".join(texts[i:i + 2])) or YEAR_RE.findall(texts[i - 1] if i else "")    # this degree's own years
        years = sorted(dict.fromkeys(years))
        year = years[-1] if years else None
        if not sections.get("education") and not institution and not year:
            continue                                                # no Education section: a bare degree word alone is not enough
        key = (norm_key(degree), norm_key(institution or ""))
        if key in seen:
            continue
        seen.add(key)
        score = SCORE_RE.search(" ".join(texts[i:i + 3]))
        entries.append({"degree": degree, "institution": institution, "year": year,
                        "start_year": years[0] if len(years) >= 2 else None,
                        "end_year": year, "score": clean(score.group(0)) if score else None})
    # tables read column by column leave the institutions far from their degrees: give the unused ones, in order, to the entries without
    if any(not e["institution"] for e in entries):
        taken = {norm_key(e["institution"]) for e in entries if e["institution"]}
        pool = []
        for t in texts:
            c = clean(re.split(r"\s[-–—|]\s|\(|\b(?:19|20)\d{2}\b", t)[0]).strip(" .,-:")
            if real_institution(c) and not DEGREE_RE.search(c) and norm_key(c) not in taken and norm_key(c) not in {norm_key(x) for x in pool}:
                pool.append(c)
        for e in entries:
            if not e["institution"] and pool and sections.get("education"):
                e["institution"] = pool.pop(0)
    return entries[:8]


def parse_date(text):
    """'Jan 2020' / '03/2021' / '2019' -> (year, month) or None; 'Present' -> today."""
    text = text.strip().lower()
    text = re.sub(r"([a-z]{3,9})\.?\s*['’]\s*(\d{2})$", r"\1 20\2", text)                     # Jan'20 -> Jan 2020
    text = re.sub(r"^((?:19|20)\d{2})[\s,]*([a-z]{3,9})\.?$", r"\2 \1", text)                   # 2018 December -> December 2018
    if re.match(r"present|current|till|now|ongoing|date|today", text):
        now = time.localtime()
        return now.tm_year, now.tm_mon
    m = re.match(r"(\d{1,2})[/.\-]((?:19|20)\d{2})$", text)
    if m:
        return int(m.group(2)), min(12, int(m.group(1)))
    d = dateparser.parse(text, settings={"PREFER_DAY_OF_MONTH": "first", "REQUIRE_PARTS": ["year"]})
    return (d.year, d.month) if d else None


def months_between(start, end):
    return max(0, (end[0] - start[0]) * 12 + end[1] - start[1] + 1)


def shorten(text, limit=450):
    """A job / project description: whitespace cleaned, cut at a sentence end under the limit."""
    text = clean(text)
    if len(text) <= limit:
        return text
    cut = text[:limit]
    end = max(cut.rfind(". "), cut.rfind("; "))
    return (cut[:end + 1] if end > limit // 3 else cut.rsplit(" ", 1)[0]) .strip()


def is_strong_skill(key):
    """A dictionary skill that is safe to report from running text: multi-word, well-known tech, or written like tech
    (AWS, C++, Node.js). One generic word ('load', 'output', 'cluster') is not."""
    L = lookups()
    nice = L.skills[key]
    return key not in L.whole_item_only and (" " in nice.strip() or key in L.known_tech or re.search(r"[A-Z].*[A-Z]|\d|[#+.]", nice[1:]) is not None)


def prepare_date_lines(texts):
    """Text of the lines with the PDF's lost dash (U+FFFD) put back, brackets removed, and two lines that hold only one date each
    ('October 2018' / 'January 2020') joined into one range."""
    out = [re.sub(r"[\[\]]", " ", t.replace("\ufffd", "-")) for t in texts]
    for i in range(len(out) - 1):                                   # 'Company' / '- X Pvt Ltd'  ->  'Company: X Pvt Ltd'
        label = re.fullmatch(r"\W*(company|employer|organi[sz]ation|designation|position|role|job title)(?:\s+name)?\W*", out[i], re.IGNORECASE)
        if label and out[i + 1].strip(" -–:"):
            out[i] = label.group(1) + ": " + out[i + 1].strip(" -–:")
            out[i + 1] = ""
    for i in range(len(out) - 1):
        if ONE_DATE_RE.match(out[i]) and ONE_DATE_RE.match(out[i + 1]) and not DATE_RANGE_RE.search(out[i]):
            out[i] = out[i].strip(" ,-–") + " - " + out[i + 1].strip(" ,-–")
            out[i + 1] = ""
    return out


def find_undated_entries(lines):
    """A resume that gives no dates: a title line next to a company line (either order), or 'Title at Company' / 'Title - Company'."""
    L = lookups()
    texts = [clean(l.text) for l in lines]

    def head(t):
        t = t.strip(" .-:|*>")
        if not t or len(t.split()) > 9 or PROSE_START_RE.match(t) or re.search(r"@|https?|\d{4}", t) or heading_section(t) \
                or t[:1].islower() or t.endswith((",", ";")) or LABEL_COMPANY_RE.match(t) or LABEL_TITLE_RE.match(t) \
                or re.fullmatch(r"(?:company|employer|designation|position|role|fresher|experience|total experience)\W*", t, re.IGNORECASE):
            return ""
        return t

    def is_title(t):
        return norm_key(t) in L.titles or (TITLE_HEAD.search(t) is not None and not COMPANY_SUFFIX_RE.search(t))

    def is_company(t):
        return COMPANY_SUFFIX_RE.search(t) is not None

    def company_like(t):
        return bool(t) and not is_title(t) and (is_company(t) or (len(t.split()) <= 5 and t[:1].isupper() and norm_key(t) not in L.skills
                                                                  and not ({w.lower().strip(",") for w in t.split()} & INDIAN_PLACES)))
    entries, used = [], set()
    for i, raw in enumerate(texts):                                 # 'Company: X' / 'Designation: Y' written as labels
        mc = LABEL_COMPANY_RE.match(raw)
        if mc and clean(mc.group(1)):
            title = None
            for j in range(i + 1, min(len(texts), i + 5)):
                mt = LABEL_TITLE_RE.match(texts[j])
                if mt:
                    title = clean(mt.group(1))
                    used.add(j)
                    break
            entries.append({"title": title, "company": clean(mc.group(1))[:80], "start": "", "end": "", "months": 0, "current": False,
                            "description": "", "skills": []})
            used.add(i)
    for i, raw in enumerate(texts):
        t = head(raw)
        if not t or i in used:
            continue
        parts = [p.strip() for p in re.split(r"\s*[|@,]\s*|\s+[-–—]\s+|\s+at\s+", t) if p.strip()]
        title = company = None
        if is_title(t) or any(is_title(p) for p in parts):
            title = next((p for p in parts if is_title(p)), t)
            company = next((p for p in parts if p != title and is_company(p)), None)
            if company is None and len(parts) > 1:
                company = next((p for p in parts if p != title and company_like(p)), None)
            if company is None:
                for j in (i + 1, i - 1):
                    if 0 <= j < len(texts) and j not in used and company_like(head(texts[j])):
                        company = head(texts[j])
                        used.add(j)
                        break
        elif is_company(t):
            company = t
            if i + 1 < len(texts) and is_title(head(texts[i + 1])):
                title = head(texts[i + 1])
                used.add(i + 1)
        else:
            continue
        if title and company and norm_key(title) == norm_key(company):
            company = None
        stop = next((k for k in range(i + 1, len(texts)) if k not in used and head(texts[k]) and (is_title(head(texts[k])) or is_company(head(texts[k])))), len(texts))
        body = " ".join(texts[i + 1:stop])
        entries.append({"title": title, "company": company, "start": "", "end": "", "months": 0, "current": False,
                        "description": shorten(body) if not body.startswith(company or "\0") else "",
                        "skills": list(dict.fromkeys(L.skills[k].lower() for k, _ in find_terms(body, L.skills) if is_strong_skill(k)))[:15]})
        used.add(i)
    seen, unique = set(), []
    for e in entries:
        k = (norm_key(e["company"] or ""), norm_key(e["title"] or ""))
        if k not in seen and (e["title"] or e["company"]):
            seen.add(k)
            unique.append(e)
    return unique[:12]


def find_entries(lines, undated=False):
    """Experience / internship entries. Every date range ('Jan 2020 - Present', '2018-2021') starts an entry. The title
    and company come from that line and the line before it: a company has a suffix (Pvt Ltd, Technologies ...),
    a title ends with an occupation word. When no suffix is found spaCy's ORG label is tried."""
    L = lookups()
    entries = []
    texts = prepare_date_lines([l.text for l in lines])
    for i, text in enumerate(texts):
        m = DATE_RANGE_RE.search(text)
        if m:
            start, end, end_word, span = parse_date(m.group("start")), parse_date(m.group("end")), m.group("end"), m.span()
        else:                                                       # 'Since Jan 2020' / 'Joined in 2019' : still working there
            sm = SINCE_RE.search(text) if (len(text.split()) <= 10 or re.match(r"\W*(?:currently\s+)?(?:working|employed|serving)\b", text, re.IGNORECASE)) else None
            if not sm:
                continue
            start, end, end_word, span = parse_date(sm.group(1)), parse_date("present"), "present", sm.span()
        if not start or not end:
            continue
        no_days = re.sub(r"\b\d{1,2}(?:st|nd|rd|th),?\s+(?=" + MONTHS + r")", "", text, flags=re.IGNORECASE)     # '12th Aug 2019' is a date, not the 12th class
        if DEGREE_RE.search(no_days) or INSTITUTION_RE.search(text) or (i and INSTITUTION_RE.search(texts[i - 1]) and DEGREE_RE.search(texts[i - 1])):
            continue                                                # a study period, not a job
        header = [text[:span[0]] + " " + text[span[1]:]]
        header_rows = []
        for j in (i - 1, i + 1):
            if 0 <= j < len(texts) and texts[j].strip() and len(texts[j].split()) <= 12 and not PROSE_START_RE.match(texts[j]) and not DATE_RANGE_RE.search(texts[j]):
                header.append(texts[j])
                header_rows.append(j)
        title = company = None
        around = (texts[i - 1] + " " + text) if i else text           # the sentence may start on the line before the date
        sentence = re.search(r"(?:working|worked|serving|employed|currently)\s+as\s+(?:an?\s+)?(.+?)\s+(?:in|at|with|for)\s+(.+?)(?:\s+(?:from|since|for|till)\b|\.|,|\[|$)", around, re.IGNORECASE)
        if sentence:                                                # 'Working as Associate Consultant in Cognizant from Oct 2019'
            title, company = clean(sentence.group(1)), clean(sentence.group(2))
        else:
            sentence = re.search(r"(?:working|worked|employed|serving)\s+(?:in|at|with|for)\s+(.+?)\s+since\b.*?\bas\s+(?:an?\s+)?(.+?)(?:\.|,|$)", around, re.IGNORECASE)
            if sentence:                                            # 'Working in Infosys Ltd. since 14th February, 2011 as Senior Systems Engineer.'
                company, title = clean(sentence.group(1)).strip(" ."), clean(sentence.group(2))
        for h in header:
            for part in re.split(r"\s*[|,@]\s*|\s+[-–—]\s*|\s+at\s+|\s{3,}", h):
                part = re.sub(r"^(?:\d+[.)]\s*|with\s+|at\s+|in\s+)", "", clean(part), flags=re.IGNORECASE).strip(" .-:()")
                if not part or len(part.split()) > 8 or re.search(r"[@\d]{4,}", part) or heading_section(part) \
                        or re.match(r"(?:role|roles|responsibilit|project|client|environment|description|duties|key|technolog)\w*\b", part, re.IGNORECASE):
                    continue
                if title is None and (norm_key(part) in L.titles or TITLE_HEAD.search(part)) and not COMPANY_SUFFIX_RE.search(part):
                    title = part
                elif company is None and COMPANY_SUFFIX_RE.search(part):
                    company = part
        if title is None:                                           # a title in the lines around the date
            for j in range(max(0, i - 3), min(len(texts), i + 4)):
                found = [L.titles[k] for k, _ in find_terms(texts[j], L.titles, max_words=6)
                         if TITLE_HEAD.search(L.titles[k].split()[-1] + " ") and len(L.titles[k].split()) >= 2]
                if found:
                    title = found[0]
                    break
        if company is None:                                         # second chance: spaCy organisation
            for ent in L.nlp()(" . ".join(header)).ents:
                if ent.label_ == "ORG" and ent.text.strip() != title and len(ent.text.split()) <= 6 \
                        and norm_key(ent.text) not in L.skills and norm_key(ent.text) not in L.titles and not re.search(r"compliance|management|analysis|testing|support", ent.text, re.I):
                    company = clean(ent.text)
                    break
        if company is None or title is None:                        # 'Employer: X' / 'Designation: Y' written as labels near the dates
            near = list(dict.fromkeys(texts[max(0, i - 6):i + 7]))
            if company is None:
                company = next((clean(LABEL_COMPANY_RE.match(t).group(1)) for t in near if LABEL_COMPANY_RE.match(t)), None)
            if title is None:
                title = next((clean(LABEL_TITLE_RE.match(t).group(1)) for t in near if LABEL_TITLE_RE.match(t)), None)
        if company is None and title is not None:                   # 'COGNIZANT' on the line next to the title: a short capitalised line is the company
            for h in header[1:]:
                h = clean(h).strip(" .-:|,")
                if h and h != title and 1 <= len(h.split()) <= 4 and (h.isupper() or h.istitle()) and not re.search(r"\d", h) \
                        and norm_key(h) not in L.skills and norm_key(h) not in L.titles and not TITLE_HEAD.search(h) and not heading_section(h) \
                        and not ({w.lower().strip(",.") for w in h.split()} & (INDIAN_PLACES | NAME_BAD_WORDS)):
                    company = h
                    break
        if company and ((PROSE_START_RE.match(company) and not COMPANY_SUFFIX_RE.search(company)) or len(company) > 80 or len(company.split()) > 8):
            company = None                                          # 'Negotiate with ...' is a sentence, not a company
        if title and len(title) > 80:
            title = None
        current = bool(re.match(r"present|current|till|now|ongoing|date|today", end_word.strip(), re.IGNORECASE))
        if start > end:
            start, end = end, start
        # what the person did in this job: the lines after the date line, up to the header of the next job
        next_date = next((k for k in range(i + 1, len(texts)) if DATE_RANGE_RE.search(texts[k])), len(texts))
        begin = i + 2 if (i + 1) in header_rows else i + 1
        stop = next_date - 1 if (next_date < len(texts) and next_date - 1 > i and len(texts[next_date - 1].split()) <= 12) else next_date
        description = shorten(" ".join(texts[begin:max(begin, stop)]))
        used = list(dict.fromkeys(L.skills[k].lower() for k, _ in find_terms(" ".join(texts[begin:max(begin, stop)]), L.skills)
                                  if is_strong_skill(k)))[:15]
        entries.append({"title": title, "company": company, "start": "%04d-%02d" % start, "end": "%04d-%02d" % end,
                        "months": months_between(start, end), "current": current, "description": description, "skills": used})
    seen, unique = set(), []
    for e in entries:
        k = (e["start"], e["end"], norm_key(e["company"] or ""), norm_key(e["title"] or ""))
        if k not in seen:
            seen.add(k)
            unique.append(e)
    if undated and not any(e["title"] or e["company"] for e in unique):
        unique += find_undated_entries(lines)
    return unique[:15]


def total_experience_years(sections, entries):
    """'6+ years of experience' if the resume says so; otherwise the length of the experience entries (overlaps counted once)."""
    top = " ".join(l.text for l in sections.get("top", [])[:12] + sections.get("summary", [])[:8])
    m = re.search(r"(\d{1,2}(?:\.\d{1,2})?)\s*\+?\s*(?:years?|yrs?)\b[^.]{0,40}?(?:experience|exp)\b|(?:experience|exp)\b[^.]{0,30}?(\d{1,2}(?:\.\d{1,2})?)\s*\+?\s*(?:years?|yrs?)", top, re.IGNORECASE)
    if m:
        return float(m.group(1) or m.group(2)), "stated in the resume"
    spans = sorted((int(e["start"][:4]) * 12 + int(e["start"][5:]), int(e["end"][:4]) * 12 + int(e["end"][5:])) for e in entries if e["start"])
    total, cursor = 0, None
    for s, e in spans:
        if cursor is None or s > cursor:
            total += e - s + 1
            cursor = e
        elif e > cursor:
            total += e - cursor
            cursor = e
    return (round(total / 12, 1), "calculated from the dates") if total else (None, None)


def find_projects(sections):
    """Project titles from the Projects section. A title is a short line that is not a sentence and not a label
    ('Description:', 'Technologies:'). 'Project Name: X' labels win. For every project the skills used are the
    dictionary skills found in its block of text."""
    L = lookups()
    lines = sections.get("projects", [])
    labels = re.compile(r"^(description|role|roles|responsibilit\w*|technolog\w*|tech stack|tools?|environment|duration|period|team size|team|"
                        r"client|customer|platform|objective|aim|abstract|outcome|summary|overview|contribution|domain|database|languages?|"
                        r"major responsibilit\w*|key responsibilit\w*|basic academic|academic credentials|roles? (?:&|and) responsibilit\w*)\b", re.I)
    titles, current = [], None

    def tidy(name):
        name = clean(name).strip(" -–:\"'“”‘’")
        return re.sub(r"[.,;]+$", "", name).strip()

    for line in lines:
        text = clean(line.text)
        m = re.match(r"^(?:(?:project|product|application)\s*(?:title|name)?\s*\d*\s*[:\-–]\s*)(.+)$", text, re.IGNORECASE)
        if m and tidy(m.group(1)):
            if current is not None and re.match(r"(?:product|application)\s*name", text, re.IGNORECASE) and not current["block"]:
                continue                                            # 'PROJECT 1: X' then 'Product Name: X' is the same project
            current = {"name": tidy(m.group(1))[:120], "block": []}
            titles.append(current)
            continue
        colon = re.match(r"^([A-Z][^:]{5,70}?)\s*:\s+(\S.*)$", text)           # 'Aladdin Product Development: Aladdin is the product of ...'
        if colon and 2 <= len(colon.group(1).split()) <= 8 and not labels.match(colon.group(1)) and not PROSE_START_RE.match(colon.group(1)) \
                and not re.search(r"\b(?:is|are|was|were|skills?|technologies|environment|responsibilit\w*|description|role|duration|client)\b", colon.group(1), re.IGNORECASE) \
                and (current is None or len(current["block"]) >= 1 or line.bold):
            current = {"name": tidy(colon.group(1))[:120], "block": [colon.group(2)]}
            titles.append(current)
            continue
        bullet = bool(re.match(r"^[>•\-\*▪●◦·]\s", line.text.strip()))
        dated = re.search(r"\s*\(\s*(?:" + _DATE_PART + r")(?:\s*(?:-|–|—|to)\s*(?:" + _END_WORDS + r"|" + _DATE_PART + r"))?\s*\)\s*$", text, re.IGNORECASE)
        shown = text[:dated.start()] if dated else text                 # 'Standard Part Cost Screening Tool (10/2018 - 07/2019)' -> the title
        is_title = (2 <= len(shown.split()) <= 12 and not shown.endswith((".", ":", "-", "–", ",")) and (dated or not DATE_RANGE_RE.search(shown))
                    and not PROSE_START_RE.match(shown) and not labels.match(shown) and not shown[:1].islower() and not re.match(r"^[\W\d]+$", shown)
                    and len(shown) >= 4 and not (shown.startswith("(") and shown.endswith(")")) and not (bullet and not line.bold)
                    and not re.search(r"\b(?:is|are|was|were|provides?|provided|able|ability|have|has|will|can)\b", shown))
        # a title is usually bold or starts a bullet; without styling we accept the first such line after a long block
        if is_title and (line.bold or current is None or len(current["block"]) >= 2):
            current = {"name": tidy(shown)[:120], "block": []}
            titles.append(current)
        elif current is not None:
            current["block"].append(text)
    if not titles:                                                  # a Projects section of plain sentences: one project per line
        for line in lines:
            text = clean(line.text)
            if len(text.split()) < 3 or text[:1].islower() or labels.match(text) or re.match(r"^[\W\d]+$", text) or heading_section(text.rstrip(":")):
                continue
            name = " ".join(re.split(r"\s+(?:at|for|in|using|with|on)\s+|[:(,]", text)[0].split()[:14])
            titles.append({"name": tidy(name), "block": [text]})
    out = []
    for p in titles:
        found = [L.skills[k] for k, _ in find_terms(" ".join(p["block"]), L.skills) if is_strong_skill(k)]
        block_text = " ".join(p["block"])
        link = re.search(r"https?://\S+|(?:github|gitlab)\.com/\S+", block_text)
        out.append({"name": p["name"], "skills": list(dict.fromkeys(found))[:12], "description": shorten(block_text, 350),
                    "link": link.group(0).rstrip(".,)") if link else ""})
    return [p for p in out if re.search(r"[A-Za-z]{3}", p["name"]) and p["name"].count(")") <= p["name"].count("(")][:10]


CERT_NOISE = re.compile(r"^(?:year of|board|university|institution|grade|validity|percentage|marks|s\.?\s?no|title|organi[sz]ation|"
                        r"mob|mobile|place|date|passing|div|dob|date of birth|gender|marital|address|personal info\w*|nationality|languages?|father|mother|"
                        r"permanent|current address|hobbies|strengths?)\b|\bmob\s*[:.\-]|^(?:pursued|worked|engineered|negotiat\w*|analy[sz]\w*|handled|managed|"
                        r"led|lead|distribut\w*|executed|conducted|collaborat\w*|coordinat\w*|performed|ensured|achieved|improved|reduced|"
                        r"honesty|creative|soft)\b|\b(?:ltd|pvt|limited)\b\.?\s*,?\s*[A-Z]?[a-z]*$|\b\d+\+?\s*(?:gb|tb|years?)\b", re.IGNORECASE)
CERT_WORDS = re.compile(r"\b(certified|certification|certificate|certificates|licen[sc]e|diploma in|nanodegree|specialization)\b", re.IGNORECASE)
CERT_VENDORS = re.compile(r"\b(aws|azure|google|microsoft|cisco|oracle|coursera|udemy|edx|udacity|linkedin learning|scrum|pmp|prince2|itil|"
                          r"ccna|ccnp|cissp|ceh|comptia|six sigma|salesforce|sap|red hat|hubspot|nptel|simplilearn)\b", re.IGNORECASE)


def find_certifications(sections, all_lines):
    """Lines of the Certifications section (job duties and bare dates removed) + lines elsewhere that clearly
    name a certificate: 'Certified X', 'X Certification', or a vendor / course platform."""
    out = []

    def keep(text, strict):
        text = clean(text)
        text = re.sub(DATE_RANGE_RE, " ", text)                         # '(09/2019 - Present)' is a date, not part of the name
        text = re.sub(r"\b(?:19|20)\d{2}\b|\((?:in progress|ongoing|in view)\)|\(\s*\)", "", text).strip(" -–:|,.()")
        if not 3 <= len(text) <= 160 or len(text.split()) > 16:
            return
        if PROSE_START_RE.match(text) or text[:1].islower() or EMAIL_RE.search(text) or CERT_NOISE.search(text):
            return
        if len(text.split()) > 9 and not CERT_WORDS.search(text) and not CERT_VENDORS.search(text):
            return                                                      # a sentence, not a certificate name
        if len(text.split()) == 1 and not CERT_VENDORS.search(text):
            return
        if re.search(r"\b(is a|is an|are|was|were|awarded to|credible|by the)\b", text, re.IGNORECASE) or text.count("(") != text.count(")"):
            return
        if re.search(r"\bpresent\b|^\d{1,2}\s*/|\bcourse curriculum\b", text, re.IGNORECASE) or (
                TITLE_HEAD.search(text.split()[-1] + " ") and not CERT_WORDS.search(text) and not CERT_VENDORS.search(text)):
            return
        if DEGREE_RE.search(text) and not CERT_WORDS.search(text) and not strict:
            return
        if text.lower() in {"certifications", "certification", "certificates", "certificate", "courses", "online", "udemy", "coursera"} \
                or re.fullmatch(r"[\d\s.%/]+", text) or re.match(r"(?:extra|accomplish|awards?|awarded|personal|trainings?\s*/|achievements?|hobbies|interests?|declaration|references?)\b", text, re.IGNORECASE):
            return
        out.append(text)

    section = sections.get("certifications", [])
    header_hits = sum(bool(re.search(r"year of|passing|percentage|cgpa|board|discipline|aggregate|institution|school", l.text, re.IGNORECASE)) for l in section)
    wrapped = []                                                    # a name that the PDF wrapped over two lines is joined first
    for l in section:
        t = clean(l.text)
        if not t:
            continue
        if wrapped and (re.search(r"(?:\b(?:and|in|of|on|for|the|by|to|with|at)|[-,&–/])$", wrapped[-1], re.IGNORECASE) or t[:1].islower()
                        or (re.match(r"(?:certification|certificate|course)\b[,\s]", t, re.IGNORECASE) and not CERT_WORDS.search(wrapped[-1]))):
            wrapped[-1] += " " + t
        else:
            wrapped.append(t)
    for text in wrapped:
        text = re.sub(r"^[o•\-\*▪►▸>]\s+", "", text)
        if header_hits >= 3 and not (CERT_WORDS.search(text) or CERT_VENDORS.search(text) or re.search(r"\b(?:course|training|workshop|bootcamp|virtual experience)\b", text, re.IGNORECASE)):
            continue                                                # an education table under a heading that looked like Certifications
        keep(text, False)
    section_ids = {id(l) for l in sections.get("certifications", [])}
    for line in all_lines:
        if id(line) in section_ids or len(line.text.split()) > 12:
            continue
        if CERT_WORDS.search(line.text) and CERT_VENDORS.search(line.text) or re.search(r"^\W*(?:o\s+)?certified\b", line.text, re.IGNORECASE):
            keep(re.sub(r"^[o•\-\*▪►▸>]\s+", "", clean(line.text)), True)
    return list(dict.fromkeys(out))[:15]


SALARY_LABEL_RE = re.compile(r"\b(ctc|salary|package|remuneration|compensation|ectc|cctc|lpa|expected|current|present|fixed|variable)\b", re.IGNORECASE)
AMOUNT_RE = re.compile(r"(?:₹|rs\.?|inr)?\s*(\d{1,3}(?:,\d{2,3})+|\d+(?:\.\d+)?)\s*(lpa|lacs?|lakhs?|l\b|crores?|cr\b|k\b|thousand|per\s*annum|pa\b|p\.a\.?|per\s*month|pm\b|/\s*month|monthly)?", re.IGNORECASE)


def to_rupees_per_year(number, unit, line):
    """'12 LPA' -> 1200000, '6,00,000' -> 600000, '30k per month' -> 360000. None if it cannot be understood."""
    unit = (unit or "").lower().replace(" ", "")
    monthly = "month" in unit or unit in ("pm", "monthly") or re.search(r"per\s*month|/\s*month|monthly|\bpm\b", line, re.IGNORECASE)
    if unit in ("lpa", "lac", "lacs", "lakh", "lakhs", "l"):
        value = number * 100000
    elif unit in ("crore", "crores", "cr"):
        value = number * 10000000
    elif unit in ("k", "thousand"):
        value = number * 1000
    elif number < 100:
        value = number * 100000                  # "CTC: 12" almost always means 12 lakhs a year
    else:
        value = number
    value = value * 12 if monthly and value < 1000000 else value
    return int(value) if 10000 <= value <= 200000000 else None


def find_salary(all_lines):
    """Current and expected salary. A line is a salary line when it has a label (CTC, salary, package, expected, current)
    and an amount. The amounts are converted to rupees per year; the raw text is kept too."""
    result = {}
    texts = [l.text for l in all_lines]
    for i, text in enumerate(texts):
        if not SALARY_LABEL_RE.search(text) or not re.search(r"ctc|salary|package|lpa|remuneration|compensation|ectc|cctc", text, re.IGNORECASE):
            continue
        window = text + " " + (texts[i + 1] if i + 1 < len(texts) and len(texts[i + 1]) < 40 else "")
        for kind, pattern in (("expected", r"expected|expecting|ectc|e-ctc|desired"), ("current", r"current|present|existing|cctc|c-ctc|drawing|fixed|last")):
            m = re.search(r"(?:" + pattern + r")[^0-9₹]{0,40}(?:₹|rs\.?|inr)?\s*(\d[\d,.]*)\s*(lpa|lacs?|lakhs?|l\b|crores?|cr\b|k\b|per\s*annum|pa\b|per\s*month|pm\b|monthly)?", window, re.IGNORECASE)
            if m and kind not in result:
                number = float(m.group(1).replace(",", "").rstrip(".") or 0)
                value = to_rupees_per_year(number, m.group(2), window)
                if value:
                    result[kind] = {"per_year_inr": value, "raw": clean(m.group(0))[:60]}
        if not result:
            m = AMOUNT_RE.search(text)
            if m and re.search(r"ctc|salary|package", text, re.IGNORECASE):
                value = to_rupees_per_year(float(m.group(1).replace(",", "")), m.group(2), text)
                if value:
                    result["current"] = {"per_year_inr": value, "raw": clean(m.group(0))[:60]}
    return result


# =====================================================================================================
# 9. CANDIDATE LISTS  (CSV / Excel files with one person per row)
# =====================================================================================================

def read_table(path):
    """Rows of a CSV / TSV / XLSX file (first sheet). Old .xls / .ods files are turned into CSV beforehand."""
    import csv
    import io
    import zipfile
    import xml.etree.ElementTree as ET
    csv.field_size_limit(10_000_000)
    with open(path, "rb") as f:
        head = f.read(4)
    if head[:2] == b"PK":                                        # .xlsx: read the XML inside (no extra library needed)
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        rows = []
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            shared = []
            if "xl/sharedStrings.xml" in names:
                for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
                    shared.append("".join(t.text or "" for t in si.iter(ns + "t")))
            sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
            for _, row in ET.iterparse(io.BytesIO(z.read(sheets[0])), events=("end",)) if sheets else []:
                if row.tag != ns + "row":
                    continue
                cells = {}
                for c in row.iter(ns + "c"):
                    ref = re.match(r"([A-Z]+)", c.get("r", ""))
                    if not ref:
                        continue
                    col = 0
                    for ch in ref.group(1):
                        col = col * 26 + ord(ch) - 64
                    v = c.find(ns + "v")
                    if c.get("t") == "s" and v is not None and v.text and v.text.isdigit():
                        cells[col] = shared[int(v.text)] if int(v.text) < len(shared) else ""
                    elif c.get("t") == "inlineStr":
                        cells[col] = "".join(t.text or "" for t in c.iter(ns + "t"))
                    elif v is not None:
                        cells[col] = v.text or ""
                if cells:
                    rows.append([cells.get(i, "") for i in range(1, max(cells) + 1)])
                row.clear()
        return rows
    raw = open(path, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    delimiter = "\t" if text[:400].count("\t") > text[:400].count(",") else ","
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    if rows and "\t" in rows[0][0]:                              # a tab-separated file saved with comma/quote junk: split every line by tabs
        rows = [re.sub(r'"\t(?=[+\d])', '"', re.sub(r",+$", "", line)).replace('"', "").split("\t") for line in text.splitlines() if line.strip()]
    return rows


COLUMN_RULES = [                      # field -> (column name must match, column name must NOT match)
    ("name", r"candidate\s*name|name\s+of\s+(the\s+)?candidate|full\s*name|first\s*name|^name$|^name\b", r"executive|recruiter|vendor|client|company|account|req\b"),
    ("email", r"e-?mail", None),
    ("phone", r"mob|phone|contact\s*n", r"alternate|alt\b"),
    ("linkedin", r"linkedin", None),
    ("skills", r"skill", None),
    ("title", r"designation|current title|job title|previous designation|^role$", None),
    ("company", r"^company$|current company|current org|previous company|employer|organi[sz]ation", None),
    ("experience", r"experience|exp\b", r"relevant|total_|notice"),
    ("education", r"degree|qualif|education|college|course|university", None),
    ("location", r"^location$|^city$|current location", None),
    ("current_ctc", r"current[_ ]?(ctc|salary)|^ctc$|cctc", None),
    ("expected_ctc", r"expected[_ ]?(ctc|salary)|ectc", None),
]


def parse_table(rows):
    """A candidate list -> {'columns', 'rows', 'candidates': [ {name, email, phone, ...} ]}. Columns are matched by name
    (any spelling of 'candidate name', 'mob.no', 'e-mail id', 'skill set' ...). Rows without a name or any contact are skipped."""
    if not rows:
        return {"columns": [], "rows": 0, "candidates": []}
    header = [str(h).strip().lower() for h in rows[0]]
    columns = collections.defaultdict(list)
    for field, want, avoid in COLUMN_RULES:
        for i, h in enumerate(header):
            if re.search(want, h) and not (avoid and re.search(avoid, h)):
                columns[field].append(i)
    candidates = []
    for row in rows[1:]:
        def get(field):
            parts = [str(row[i]).strip() for i in columns.get(field, []) if i < len(row) and str(row[i]).strip() not in ("", "-", "nan", "None")]
            return " | ".join(dict.fromkeys(parts))
        name = clean(get("name"))
        email = None
        for candidate in EMAIL_RE.findall(get("email")):
            try:
                email = validate_email(candidate, check_deliverability=False).normalized
                break
            except EmailNotValidError:
                continue
        phone = None
        for match in phonenumbers.PhoneNumberMatcher(get("phone"), "IN"):
            if phonenumbers.is_valid_number(match.number):
                phone = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.E164)
                break
        linkedin = find_linkedin([Line(get("linkedin"))])
        if len(re.sub(r"[^A-Za-z]", "", name)) < 2 and not (email or phone):
            continue
        skills = [x for x in dict.fromkeys(p.strip(" .") for p in re.split(r"[,;|]", get("skills"))) if 2 <= len(x) <= 45][:30]
        candidate = {"name": name or None, "email": email, "phoneNumber": phone,
                     "socialProfiles": [{"platform": "linkedin", "url": linkedin}] if linkedin else None,
                     "skills": [x.lower() for x in skills],
                     "currentCompanyDetails": {"companyName": get("company"), "designation": get("title")} if (get("company") or get("title")) else None,
                     "totalExperienceText": get("experience") or None, "educationText": get("education") or None,
                     "location": get("location") or None, "currentCTC": get("current_ctc") or None, "expectedCTC": get("expected_ctc") or None}
        candidates.append({k: v for k, v in candidate.items() if v not in (None, [], "")})
    return {"columns": [h for h in header if h][:30], "rows": len(rows) - 1, "candidates": candidates}


# =====================================================================================================
# 8. PARSE ONE FILE + VALIDATION LAYER
# =====================================================================================================

# =====================================================================================================
# 8. THE PRODUCTION FORMAT
#    The output of the production system (futurejobsProd.resumes.json) is the model: the same field names, the same
#    shapes. Everything below turns what was found into that shape.
# =====================================================================================================

EDU_TYPE_RULES = [                    # production educationType values, tested in this order
    ("phd", r"\bph\.?\s?d\b|doctorate"),
    ("masters", r"\bm\.?\s?tech\b|\bm\.?\s?sc\b|\bm\.?\s?com\b|\bmba\b|\bmca\b|\bpgdm\b|\bpgdca\b|\bmaster|post[- ]?graduat|\bllm\b|"
                r"\bm\.?\s?pharm\b|\bm\.?\s?arch\b|\bmsw\b|\bm\.?e\b|\bm\.?a\b"),
    ("bachelors", r"\bb\.?\s?tech\b|\bb\.?\s?sc\b|\bb\.?\s?com\b|\bbba\b|\bbca\b|\bbachelor|\bb\.?e\b|\bb\.?a\b|\bb\.?\s?pharm\b|"
                  r"\bb\.?\s?arch\b|\bllb\b|\bmbbs\b|\bbsw\b|graduat(?:e|ion)\b"),
    ("diploma", r"diploma|\biti\b"),
    ("intermediate", r"intermediate|\bhsc\b|\bpuc\b|12th|xiith|senior secondary|higher secondary|\bpre[- ]?university"),
    ("high_school", r"10th|\bssc\b|\bsslc\b|matriculation|secondary school|high school|\bxth\b"),
]
AUTHORITIES = [                       # certificate issuers, in the spelling the production data uses
    (r"amazon web services|\baws\b", "Amazon Web Services"), (r"microsoft|\bazure\b", "Microsoft"), (r"google", "Google"),
    (r"coursera", "Coursera"), (r"udemy", "Udemy"), (r"oracle", "Oracle"), (r"cisco", "Cisco"), (r"\bibm\b", "IBM"),
    (r"nptel", "NPTEL"), (r"linkedin", "LinkedIn Learning"), (r"\bedx\b", "edX"), (r"udacity", "Udacity"),
    (r"hubspot", "HubSpot"), (r"salesforce", "Salesforce"), (r"simplilearn", "Simplilearn"), (r"great learning", "Great Learning"),
    (r"hackerrank", "HackerRank"), (r"freecodecamp", "freeCodeCamp"), (r"red hat", "Red Hat"), (r"comptia", "CompTIA"),
    (r"project management institute|\bpmi\b", "PMI"), (r"scrum", "Scrum"), (r"infosys", "Infosys"), (r"nasscom", "NASSCOM"),
]
PROFILE_SITES = [("github", r"github\.com/[A-Za-z0-9_.\-]+"), ("leetcode", r"leetcode\.com/(?:u/)?[A-Za-z0-9_.\-]+"),
                 ("hackerrank", r"hackerrank\.com/(?:profile/)?[A-Za-z0-9_.\-]+"), ("kaggle", r"kaggle\.com/[A-Za-z0-9_.\-]+"),
                 ("behance", r"behance\.net/[A-Za-z0-9_.\-]+"), ("medium", r"medium\.com/@?[A-Za-z0-9_.\-]+"),
                 ("twitter", r"(?:twitter|x)\.com/[A-Za-z0-9_]+"), ("stackoverflow", r"stackoverflow\.com/users/[0-9]+/[A-Za-z0-9_\-]+"),
                 ("codechef", r"codechef\.com/users/[A-Za-z0-9_.\-]+"), ("geeksforgeeks", r"geeksforgeeks\.org/user/[A-Za-z0-9_.\-]+")]


def education_type(degree):
    low = degree.lower()
    for kind, pattern in EDU_TYPE_RULES:
        if re.search(pattern, low):
            return kind
    return ""


def field_of_study(degree):
    """'B.Tech in Mechanical Engineering' -> 'Mechanical Engineering'; 'MBA (Finance)' -> 'Finance'."""
    m = DEGREE_RE.match(degree.strip())
    rest = degree[m.end():] if m else ""
    rest = re.sub(r"^\s*(?:(?:degree\s+)?(?:in|of)\s+|[-–:(,]\s*)", "", rest, flags=re.IGNORECASE).strip(" ()-,.")
    rest = re.sub(r"\s+(?:in|of|and|&)$", "", rest, flags=re.IGNORECASE).strip()
    return rest if rest and len(rest.split()) <= 7 and not INSTITUTION_RE.search(rest) and not DEGREE_RE.fullmatch(rest) else ""


def build_education(entries):
    out = []
    for e in entries:
        out.append({"educationType": education_type(e["degree"]), "institutionName": e["institution"] or "",
                    "startDate": e.get("start_year") or "", "endDate": e.get("end_year") or "",
                    "fieldOfStudy": field_of_study(e["degree"]), "scoreAchieved": e.get("score") or ""})
    return out


def build_job(entry):
    return {"companyName": entry["company"] or "", "designation": entry["title"] or "", "startDate": entry["start"] + "-01" if entry["start"] else "",
            "endDate": "current" if entry["current"] else (entry["end"] + "-01" if entry["end"] else ""),
            "workExperienceDescription": entry["description"], "skillsUsedInRole": entry["skills"]}


def build_certifications(names):
    """Certificate names -> {name, issuingAuthority, certificateLink, certificateCredentialID}."""
    out = []
    for text in names:
        link = re.search(r"https?://\S+", text)
        credential = re.search(r"(?:credential|certificate|cert)\s*(?:id|no\.?|number)\s*[:#\-]?\s*([A-Za-z0-9\-_/]{5,})", text, re.IGNORECASE)
        name = re.sub(r"https?://\S+|(?:credential|certificate|cert)\s*(?:id|no\.?|number)\s*[:#\-]?\s*[A-Za-z0-9\-_/]{5,}", "", text, flags=re.IGNORECASE)
        parts = [clean(p) for p in re.split(r"\s*[-–—|]\s+|\s+[-–—]\s*|\s*\|\s*|,\s+|\s+(?:by|from|issued by|offered by)\s+|\s*\(\s*|\s*\)\s*", name) if clean(p)]
        authority = ""
        for part in parts[1:]:
            for pattern, canonical in AUTHORITIES:
                if re.search(pattern, part, re.IGNORECASE):
                    authority = part if len(part.split()) <= 4 else canonical
                    break
            if authority:
                break
        if not authority:
            for pattern, canonical in AUTHORITIES:
                if re.search(pattern, name, re.IGNORECASE):
                    authority = canonical
                    break
        main = parts[0] if parts else name
        if authority and authority.lower() in main.lower() and len(parts) == 1:
            main = clean(re.sub(re.escape(authority), "", main, flags=re.IGNORECASE)).strip(" -–:,") or main
        out.append({"name": main.strip(" -–:,"), "issuingAuthority": authority, "certificateLink": link.group(0).rstrip(".,)") if link else "",
                    "certificateCredentialID": credential.group(1) if credential else ""})
    return [c for c in out if len(c["name"]) >= 3]


def find_social_profiles(lines, linkedin):
    """linkedin, github, leetcode ... found anywhere in the text, plus a 'Portfolio:' website."""
    text = "\n".join(l.text for l in lines)
    out = []
    if linkedin:
        out.append({"platform": "linkedin", "url": linkedin})
    for platform, pattern in PROFILE_SITES:
        for m in re.finditer(r"(?:https?://)?(?:www\.)?" + pattern, text, re.IGNORECASE):
            url = "https://" + re.sub(r"^https?://(?:www\.)?", "", m.group(0)).rstrip(".,)/")
            if url.lower() not in (p["url"].lower() for p in out) and len(url) > len("https://github.com/") + 1:
                out.append({"platform": platform, "url": url})
    for line in lines:
        m = re.search(r"(?:portfolio|website|personal site|web)\s*[:\-]\s*((?:https?://)?[A-Za-z0-9.\-]+\.[a-z]{2,}(?:/\S*)?)", line.text, re.IGNORECASE)
        if m and not any(site in m.group(1).lower() for site in ("linkedin", "github", "leetcode", "hackerrank", "kaggle", "gmail")):
            out.append({"platform": "portfolio", "url": m.group(1).rstrip(".,)")})
    return out[:8]


def find_languages(sections, lines):
    """Spoken languages (english, hindi, ...) from the Languages part of the resume, as lower-case names."""
    L = lookups()
    if not hasattr(L, "language_table"):
        L.language_table = {norm_key(name): name for name in json.load(open(os.path.join(LOOKUPS, "languages.json"), encoding="utf-8"))}
    text = [l.text for l in sections.get("other", [])]
    for line in lines:
        m = re.match(r"\s*languages?(?:\s+known|\s+spoken)?\s*[:\-]\s*(.+)", line.text, re.IGNORECASE)
        if m:
            text.append(m.group(1))
    return list(dict.fromkeys(L.language_table[k].lower() for k, _ in find_terms(" ".join(text), L.language_table, max_words=1)))[:8]


def find_notice_period(lines):
    """Notice period in days: '30 days', '2 months', 'immediate joiner' = 0."""
    for line in lines:
        m = re.search(r"notice\s*period\s*[:\-]?\s*(?:is\s*)?(\d{1,3})\s*(days?|weeks?|months?)", line.text, re.IGNORECASE)
        if m:
            return int(m.group(1)) * {"d": 1, "w": 7, "m": 30}[m.group(2)[0].lower()]
        if re.search(r"immediate(?:ly)?\s+(?:joiner|joining|available)|can join immediately", line.text, re.IGNORECASE):
            return 0
    return 0


def build_ctc(salary):
    def block(kind, with_duration):
        value = (salary.get(kind) or {}).get("per_year_inr") or 0
        out = {"min": value, "max": value, "currency": "INR"}
        if with_duration:
            out["duration"] = "ANNUALLY"
        return out
    return {"currentCTC": block("current", False), "expectedCTC": block("expected", True)}


def rank_skills(skill_scores, lines, jobs):
    """analysedSkills = {advanced, medium, basic}. A skill is 'advanced' when the resume gives strong evidence of it:
         +2    it is in the Skills list                       +0.5 per time it appears in the text (at most 4 times)
         +1.5  it was used in the current / latest job         +1   it was used in two or more jobs
       score >= 4 advanced, >= 2.5 medium, otherwise basic."""
    L = lookups()
    text = "\n".join(l.text for l in lines)
    counts = collections.Counter(L.skills[k] for k, _ in find_terms(text, L.skills) if k not in L.whole_item_only)
    levels = {"advanced": [], "medium": [], "basic": []}
    for skill, confidence in skill_scores.items():
        lower = skill.lower()
        in_jobs = [lower in j["skillsUsedInRole"] for j in jobs]
        score = (2 if confidence >= 0.9 else 0) + 0.5 * min(counts.get(skill, 0), 4) + (1.5 if in_jobs[:1] == [True] else 0) + (1 if sum(in_jobs) >= 2 else 0)
        levels["advanced" if score >= 4 else "medium" if score >= 2.5 else "basic"].append(lower)
    return levels


def build_red_flags(jobs, education, text_lines, issues):
    """The rule-based version of the production 'redFlags'. Every flag has a plain-language description."""
    flags = {}
    # --- employment gap: a month or more between the end of one job and the start of the next ---
    ordered = sorted((j for j in jobs if j["startDate"]), key=lambda j: j["startDate"])
    gap = None
    for before, after in zip(ordered, ordered[1:]):
        if before["endDate"] != "current" and after["startDate"][:7] > before["endDate"][:7]:
            by, bm = int(before["endDate"][:4]), int(before["endDate"][5:7])
            ay, am = int(after["startDate"][:4]), int(after["startDate"][5:7])
            if (ay * 12 + am) - (by * 12 + bm) > 1:
                gap = (before["endDate"][:7], after["startDate"][:7])
                break
    flags["employmentGap"] = gap is not None
    flags["employmentGapDescription"] = (f"There is a gap between {gap[0]} and {gap[1]} that is not explained." if gap
                                         else "No gap between jobs was found in the dates.")
    # --- frequent job switcher: three or more short jobs (a year or less) ---
    short = [j for j in jobs if j["startDate"] and j["endDate"] and j["endDate"] != "current" and (int(j["endDate"][:4]) * 12 + int(j["endDate"][5:7]))
             - (int(j["startDate"][:4]) * 12 + int(j["startDate"][5:7])) <= 12]
    flags["frequentJobSwitcher"] = len(jobs) >= 3 and len(short) >= 3
    flags["frequentJobSwitcherDescription"] = (f"{len(short)} of {len(jobs)} jobs lasted a year or less." if flags["frequentJobSwitcher"]
                                               else "Job tenures do not show frequent switching.")
    # --- no role progression: three or more jobs with the same designation ---
    titles = [norm_key(j["designation"]) for j in jobs if j["designation"]]
    repeated = len(titles) >= 3 and collections.Counter(titles).most_common(1)[0][1] >= 3
    flags["lackOfRoleProgression"] = repeated
    flags["lackOfRoleProgressionDescription"] = ("The same designation is repeated in three or more jobs." if repeated
                                                 else "The designations change between jobs or there are too few jobs to judge.")
    # --- incomplete education ---
    missing = [e for e in education if e["educationType"] in ("bachelors", "masters", "phd", "diploma") and not e["institutionName"]]
    flags["incompleteEducation"] = (not education) or bool(missing)
    flags["incompleteEducationDescription"] = ("No education was found." if not education else
                                               f"{len(missing)} degree entries have no institution." if missing
                                               else "Education details include the degree and the institution.")
    # --- mistakes: words that lost their spaces, or key contact details missing ---
    words = [w for l in text_lines for w in l.text.split()]
    glued = sum(1 for w in words if len(w) > 24 and not w.startswith(("http", "www")) and "@" not in w)
    mistakes = (words and glued / len(words) > 0.01) or any("no valid email" in i or "no name" in i for i in issues)
    flags["containsMistakes"] = bool(mistakes)
    flags["containsMistakesDescription"] = ("The text has words run together or key details (name, email) are missing." if mistakes
                                            else "No obvious formatting problems were found.")
    flags["overstatedAchievements"] = False
    flags["overstatedAchievementsDescription"] = "Not assessed: this parser does not judge the claims in a resume."
    return {k: flags[k] for k in ("containsMistakes", "containsMistakesDescription", "frequentJobSwitcher", "frequentJobSwitcherDescription",
                                  "overstatedAchievements", "overstatedAchievementsDescription", "employmentGap", "employmentGapDescription",
                                  "incompleteEducation", "incompleteEducationDescription", "lackOfRoleProgression", "lackOfRoleProgressionDescription")}


def parse_lines(lines, filename):
    """Everything above, in order. Returns the JSON-ready dict in the production format."""
    sections = split_sections(lines)
    email, phone, linkedin = find_email(lines), find_phone(lines), find_linkedin(lines)
    name, name_confidence = find_name(lines, email, filename)
    skills = find_skills(sections)
    titles = find_titles(lines, sections)
    industries = find_industries(lines, sections)
    education = find_education(sections, lines)
    experience_lines = sections.get("experience") or [l for name_ in ("top", "summary", "other", "skills") for l in sections.get(name_, [])]
    experience = [e for e in find_entries(experience_lines, undated=bool(sections.get("experience"))) if not e.get("_education")]
    internships = find_entries(sections.get("internships", []))
    internships += [e for e in experience if e["title"] and re.search(r"\bintern|trainee|apprentice", e["title"], re.IGNORECASE)]   # written inside experience
    experience = [e for e in experience if e not in internships]
    years, years_source = total_experience_years(sections, experience)
    projects = find_projects(sections)
    name_words = {w for w in re.split(r"[^a-z]+", (name or "").lower()) if len(w) >= 3}
    certifications = [c for c in find_certifications(sections, lines)
                      if norm_key(c) != norm_key(name or "-") and not re.search(r"year of passing|percentage|cgpa|marks", c, re.IGNORECASE)
                      and not (INSTITUTION_RE.search(c) and not CERT_WORDS.search(c))
                      and not (name_words and len(c.split()) <= 4 and len(name_words & {w for w in re.split(r"[^a-z]+", c.lower()) if len(w) >= 3}) >= 1
                               and not CERT_WORDS.search(c) and not CERT_VENDORS.search(c))]
    salary = find_salary(lines)

    # ---- validation layer: check every field, never write a value that fails a check ----
    issues = []
    if not email:
        issues.append("no valid email found")
    if not phone:
        issues.append("no valid phone number found")
    if not name:
        issues.append("no name found")
    elif not 2 <= len(name.split()) <= 4:
        issues.append("name has one word (first name only?)")
    education = [e for e in education if e["degree"]]
    if education and not any(e["institution"] for e in education):
        issues.append("education found but no institution")
    for e in experience + internships:
        if not (e["title"] or e["company"]):
            issues.append("an experience entry has neither title nor company")
            break
    confidence = {
        "name": round(name_confidence, 2),
        "contact": round(sum(bool(x) for x in (email, phone, linkedin)) / 3, 2),
        "skills": round(sum(skills.values()) / len(skills), 2) if skills else 0,
        "education": round(sum(bool(e["institution"]) + bool(e["year"]) + 1 for e in education) / (3 * len(education)), 2) if education else 0,
        "experience": round(sum(bool(e["title"]) + bool(e["company"]) + 1 for e in experience) / (3 * len(experience)), 2) if experience else 0,
    }

    # ---- the production format ----
    named = [e for e in experience if e["title"] or e["company"]]        # an entry with dates only still counts for the years, not as a job
    jobs = [build_job(e) for e in named]
    current = next((e for e in named if e["current"]), None) or (max(named, key=lambda e: e["end"]) if named else None)
    education_out = build_education(education)
    whole_years = int(years) if years is not None else 0
    return {
        "name": name, "email": email, "phoneNumber": phone,
        "languagesKnown": find_languages(sections, lines),
        "relocationDetails": [],
        "currentCompanyDetails": {"companyName": (current or {}).get("company") or "", "designation": (current or {}).get("title") or ""},
        "employmentHistory": jobs,
        "ctcDetails": build_ctc(salary),
        "totalExperience": {"years": whole_years, "months": int(round(((years or 0) - whole_years) * 12))},
        "skills": [s.lower() for s in skills],
        "analysedSkills": rank_skills(skills, lines, jobs),
        "educationalBackground": education_out,
        "personalProjects": [{"projectName": p["name"], "projectDescription": p["description"], "projectLink": p["link"],
                              "skillsUsedInProject": [x.lower() for x in p["skills"]]} for p in projects],
        "certifications": build_certifications(certifications),
        "socialProfiles": find_social_profiles(lines, linkedin),
        "redFlags": build_red_flags(jobs, education_out, lines, issues),
        "currentNoticePeriod": find_notice_period(lines),
        # extras the production record does not have
        "titles": titles, "industries": industries, "internships": [build_job(e) for e in internships],
        "totalExperienceSource": years_source, "confidence": confidence, "issues": issues,
        "sections_found": [x for x in ("summary", "skills", "experience", "education", "projects", "internships", "certifications") if sections.get(x)],
    }


def split_bundle(lines):
    """A file that holds several resumes (a recruiter's compiled PDF) -> a list of line lists, one per candidate.
    Heuristic: every different e-mail address that first appears far from the previous one starts a new resume; the
    short lines right above it (the name / contact block) go with it. A piece without at least two real sections
    (skills, experience, education ...) is not a resume of its own and is joined to the piece before it.
    Returns [lines] (no split) when fewer than two resumes are found."""
    firsts, seen = [], set()
    for i, line in enumerate(lines):
        for e in EMAIL_RE.findall(line.text):
            if e.lower() not in seen:
                seen.add(e.lower())
                firsts.append(i)
    starts = [0]
    for i in firsts[1:]:
        begin = i
        while begin > 0 and i - begin < 3 and lines[begin - 1].text.strip() and len(lines[begin - 1].text.split()) <= 6 and not lines[begin - 1].text.rstrip().endswith("."):
            begin -= 1
        if begin - starts[-1] >= 15:
            starts.append(begin)
    if len(starts) < 2:
        return [lines]
    pieces = [lines[a:b] for a, b in zip(starts, starts[1:] + [len(lines)])]
    kept = [pieces[0]]
    for piece in pieces[1:]:
        found = split_sections(piece)
        if sum(bool(found.get(x)) for x in ("skills", "experience", "education", "projects")) >= 2:
            kept.append(piece)
        else:
            kept[-1] = kept[-1] + piece
    return kept


def run_job(job):
    """Worker: one file -> (job, JSON dict, parts). Never raises: a file that cannot be read still gets a JSON with the reason.
    parts = [(candidate number, JSON dict)] with one entry per candidate when --split-candidates found several, else []."""
    parts = []
    data = {"status": job["status"], "reason": job["reason"], "category": job["category"],
            "source_file": job["name"], "source_folder": job["folder"], "parsed": False}
    try:
        if job["mode"] == "table":
            data.update({"type": "candidate_list", **parse_table(read_table(long_path(job["read"])))})
            data["parsed"] = True
            if job.get("split") and len(data["candidates"]) > 1:
                parts = [(n, {"source_file": job["name"], "source_folder": job["folder"], "candidateNumber": n, "parsed": True, **c})
                         for n, c in enumerate(data["candidates"], 1)]
        elif job["mode"] == "parse":
            lines = read_lines(long_path(job["read"]))
            if not lines:
                raise ValueError("no text could be read from the file")
            pieces = split_bundle(lines) if job.get("split") else [lines]
            if len(pieces) > 1:
                parts = [(n, {"source_file": job["name"], "source_folder": job["folder"], "candidateNumber": n, "parsed": True,
                              **parse_lines(piece, job["name"])}) for n, piece in enumerate(pieces, 1)]
                data["type"] = "bundle"
            else:
                data.update(parse_lines(lines, job["name"]))
            data["parsed"] = True
        else:
            data["note"] = job["note"]
    except Exception as problem:
        data["parse_error"] = f"{type(problem).__name__}: {str(problem)[:200]}"
    return job, data, parts


def init_worker():
    lookups()                       # load the word lists once per worker process


# =====================================================================================================
# THE RUN
# =====================================================================================================

TABLE_EXT = (".csv", ".tsv", ".xlsx", ".xlsm", ".xls", ".ods")
SKIP_CATEGORIES = {                 # invalid files that cannot be parsed as a resume: JSON with the reason only
    "not_a_document": "not a document (video, program or unknown binary file): nothing to parse",
    "web_or_code_file": "web or code file: nothing to parse",
    "spreadsheet_not_candidates": "a spreadsheet / CSV that is not a list of candidates: not parsed as a resume",
    "empty_file": "empty file (0 bytes): nothing to parse",
}


MEDIA_EXT = {".webm", ".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".m4v", ".mpg", ".mpeg", ".3gp", ".mp3", ".wav", ".m4a", ".aac", ".ogg",
             ".flac", ".exe", ".dll", ".msi", ".iso", ".dmg", ".apk", ".bin", ".dat"}
MAX_PARSE_MB = 60                    # a resume is never this big; reading a 300 MB video hangs a worker for hours


def never_read(path):
    """A reason text if this file must not be read as a resume (video / audio / program / huge file), else None.
    This matters for 'duplicate' files: a duplicate of a video lands in a parse category, but must not be opened."""
    ext = os.path.splitext(path)[1].lower()
    try:
        size_mb = os.path.getsize(long_path(path)) / (1024 * 1024)
    except OSError:
        return None
    if ext in MEDIA_EXT:
        return f"a video, audio or program file ({ext}): nothing to parse"
    if size_mb > MAX_PARSE_MB:
        return f"file too large ({size_mb:.0f} MB) to be a resume: not read"
    return None


def json_name(folder, file_name, used):
    """'Rahul CV.pdf' -> 'Rahul CV.json'. If that name is taken by another input file: 'Rahul CV (docx).json'."""
    stem, ext = os.path.splitext(file_name)
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem) or "file"        # only characters Windows does not allow are replaced
    candidates = [f"{stem}.json", f"{stem} ({ext.lstrip('.').lower() or 'file'}).json"]
    candidates += [f"{stem} ({ext.lstrip('.').lower()}) ({n}).json" for n in range(2, 50)]
    for name in candidates:
        if (folder, name.lower()) not in used:
            used.add((folder, name.lower()))
            return name
    return f"{stem}_{len(used)}.json"


def collect_jobs(classification, only, limit):
    """One job per file of classification/Valid and classification/Invalid (the report files are skipped)."""
    import csv
    reasons = {}                                                  # (category, file name) -> reason
    report = os.path.join(classification, "Invalid", "invalid_reasons.csv")
    if os.path.exists(report):
        with open(report, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                reasons[(row["category"], row["placed_as"])] = row["reason"]
    jobs = []
    if only in ("both", "valid"):
        folder = os.path.join(classification, "Valid")
        names = sorted(n for n in os.listdir(folder) if os.path.isfile(os.path.join(folder, n)) and n != "valid_resumes.csv")
        for n in (names[:limit] if limit else names):
            path = os.path.join(folder, n)
            mode = "table" if n.lower().endswith(TABLE_EXT) else "parse"
            blocked = never_read(path) if mode == "parse" else None
            jobs.append({"path": path, "read": path, "name": n, "status": "valid", "reason": None, "category": None,
                         "folder": "Valid", "mode": "skip" if blocked else mode, "note": blocked or ""})
    if only in ("both", "invalid"):
        base = os.path.join(classification, "Invalid")
        found = []
        for category in sorted(os.listdir(base)):
            sub = os.path.join(base, category)
            if os.path.isdir(sub):
                found += [(category, n) for n in sorted(os.listdir(sub)) if os.path.isfile(os.path.join(sub, n))]
        for category, n in (found[:limit] if limit else found):
            path = os.path.join(base, category, n)
            blocked = never_read(path)
            skip = category in SKIP_CATEGORIES or blocked is not None
            jobs.append({"path": path, "read": path, "name": n, "status": "invalid",
                         "reason": reasons.get((category, n), category.replace("_", " ")), "category": category,
                         "folder": "Invalid/" + category, "mode": "skip" if skip else "parse",
                         "note": SKIP_CATEGORIES.get(category) or blocked or ""})
    return jobs


def main():
    ap = argparse.ArgumentParser(description="Parse every classified file into its own JSON (parsing only).")
    ap.add_argument("--classification", default=DEFAULT_CLASSIFICATION, help="folder with Valid and Invalid (default: ..\\classification)")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="folder that gets Valid and Invalid sub-folders of JSON files (default: ..\\Output)")
    ap.add_argument("--only", choices=["both", "valid", "invalid"], default="both")
    ap.add_argument("--workers", type=int, default=min(os.cpu_count() or 1, 3))
    ap.add_argument("--split-candidates", action="store_true",
                    help="a candidate list (CSV/Excel) or a file with several resumes gets one JSON per candidate "
                         "('<name> - 001.json', ...); the JSON with the input's name stays as an index of them")
    ap.add_argument("--limit", type=int, default=0, help="only the first N files of each folder (to try it out)")
    args = ap.parse_args()

    jobs = collect_jobs(args.classification, args.only, args.limit)
    for j in jobs:
        j["split"] = args.split_candidates
    out_valid, out_invalid = os.path.join(args.output, "Valid"), os.path.join(args.output, "Invalid")
    os.makedirs(out_valid, exist_ok=True)
    os.makedirs(out_invalid, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="parse_")
    started = time.time()
    try:
        # old formats (DOC / RTF / ODT / web pages) -> DOCX, and old spreadsheets (XLS / ODS) -> CSV, with LibreOffice
        old_docs = [j for j in jobs if j["mode"] == "parse" and sniff_type(long_path(j["path"])) == "old"]
        old_sheets = [j for j in jobs if j["mode"] == "table" and j["name"].lower().endswith((".xls", ".ods"))]
        if old_docs:
            converted = convert_old_files([long_path(j["path"]) for j in old_docs], os.path.join(tmp, "convert"))
            for j in old_docs:
                j["read"] = converted.get(long_path(j["path"]), j["path"])
                if long_path(j["path"]) not in converted:
                    j["mode"], j["note"] = "skip", "could not convert this old-format file (DOC/RTF/ODT/HTML) to read it"
            print(f"{len(converted)} of {len(old_docs)} old-format documents converted for reading.")
        if old_sheets:
            converted = convert_old_files([long_path(j["path"]) for j in old_sheets], os.path.join(tmp, "sheets"), target="csv")
            for j in old_sheets:
                j["read"] = converted.get(long_path(j["path"]), j["path"])
                if long_path(j["path"]) not in converted:
                    j["mode"], j["note"] = "skip", "could not convert this spreadsheet to read it"
        print(f"{len(jobs)} files: {sum(j['status'] == 'valid' for j in jobs)} valid, {sum(j['status'] == 'invalid' for j in jobs)} invalid. "
              f"Parsing with {args.workers} workers ...")
        used, done = set(), collections.Counter()
        with ProcessPoolExecutor(max_workers=args.workers, initializer=init_worker) as pool:
            futures = [pool.submit(run_job, j) for j in jobs]
            for n, future in enumerate(as_completed(futures), 1):
                job, data, parts = future.result()
                folder = out_valid if job["status"] == "valid" else out_invalid
                main_name = json_name(folder, job["name"], used)
                if parts:
                    stem = main_name[:-5]
                    names = [f"{stem} - {n:03d}.json" for n, _ in parts]
                    for name, (n, part) in zip(names, parts):
                        with open(os.path.join(folder, name), "w", encoding="utf-8") as f:
                            json.dump(part, f, ensure_ascii=False, indent=1)
                    data = {k: v for k, v in data.items() if k not in ("candidates", "columns")}
                    data.update({"split": True, "candidateCount": len(parts), "candidateFiles": names})
                    done[(job["status"], "split into several JSON")] += 1
                with open(os.path.join(folder, main_name), "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=1)
                done[(job["status"], "parsed" if data["parsed"] else "reason only / not readable")] += 1
                if n % 250 == 0:
                    print(f"  {n}/{len(jobs)} done ({n / (time.time() - started):.1f} per second)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"\nDONE in {round(time.time() - started)} seconds.")
    for (status, kind), count in sorted(done.items()):
        print(f"  {status:8} {kind:30} {count}")
    print(f"  valid JSON   -> {out_valid}\n  invalid JSON -> {out_invalid}")


if __name__ == "__main__":
    main()
