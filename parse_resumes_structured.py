"""
The one resume-parsing script for this project. Reads a folder of real
resumes (PDF/DOCX/DOC/image), extracts name/email/phone/linkedin, matches
skills/titles/industries/languages against your canonical lists (exact
match only -- see match_exact below), pulls out education/certifications/
projects/internships, and uploads one document per resume into a
MeiliSearch index. Fully self-contained: no other script in this project
is needed.

Checkpointing happens two ways:
  - Against MeiliSearch itself: document IDs already in the index are
    fetched once up front, so re-running never re-uploads a file whose
    document is already there.
  - Locally, via a content-hash index (hash_store_structured/), so an
    unmodified duplicate under a different filename is recognized without
    needing to re-parse or re-upload it.

Every input file gets a matching document -- a real result, a duplicate
note, an unsupported-type note, or an error -- so the index alone
explains every file you pointed the tool at.

Usage:
    python parse_resumes_structured.py [folder_name] [--workers N]
                                        [--meili-index NAME] [--lookups-dir DIR]

Requires a running MeiliSearch server, and skills.json/titles.json/
industries.json/languages.json/normalized_category_mapping.json in
lookups/ (a missing one just means that list won't match anything).
"""

import os
import re
import time
import glob
import shutil
import hashlib
import zipfile
import tempfile
import argparse
import subprocess
import pymupdf
try:
    import re2  # optional: a faster regex engine for the big canonical-list matcher
except ImportError:
    re2 = None
import meilisearch
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool


def _load_dotenv():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


_load_dotenv()

DEFAULT_INPUT_DIR = os.path.join("input", "resumes")
DEFAULT_LOOKUPS_DIR = "lookups"
DEFAULT_HASH_DIR = "hash_store_structured"
DEFAULT_MEILI_URL = os.environ.get("MEILI_URL", "https://compacter-jubilance-unfasten.ngrok-free.dev")
DEFAULT_MEILI_INDEX = "resumes_structured"
DEFAULT_MEILI_API_KEY = os.environ.get("MEILI_API_KEY")
NORMALIZED_CATEGORY_MAPPING_FILE = "normalized_category_mapping.json"

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DOC_MIME = "application/msword"
PDF_MIME = "application/pdf"
IMAGE_MIMES = {"image/jpeg", "image/png", "image/bmp", "image/tiff"}
SUPPORTED_MIMES = {PDF_MIME, DOCX_MIME, DOC_MIME} | IMAGE_MIMES


def detect_mime_type(path):
    """Identifies a file's real type from its content ('magic bytes'), not
    its filename extension, which can lie."""
    try:
        with open(path, "rb") as f:
            header = f.read(8)
    except OSError:
        return None
    if header.startswith(b"%PDF-"):
        return PDF_MIME
    if header.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        return DOC_MIME
    if header.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(path) as z:
                if "word/document.xml" in z.namelist():
                    return DOCX_MIME
        except (zipfile.BadZipFile, OSError):
            pass
        return None
    if header.startswith(b"\xFF\xD8\xFF"):
        return "image/jpeg"
    if header.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if header[:2] == b"BM":
        return "image/bmp"
    if header[:4] in (b"II*\x00", b"MM\x00*"):
        return "image/tiff"
    return None


# ============================== Heading vocabularies ==============================

KNOWN_SECTIONS = {
    "summary", "objective", "profile", "about", "professional summary",
    "skills", "technical skills", "computer skills", "skill set", "skill", "soft skill",
    "key skills", "core skills",
    "experience", "work experience", "internship", "internship experience", "professional profile",
    "education", "academic qualifications", "educational qualifications", "education qualification",
    "academic qualification",
    "projects", "academic projects", "project",
    "certifications", "certificates", "certification",
    "languages", "languages known", "language",
    "personal details", "declaration", "interests", "involvement",
    "career objectives", "career objective", "basic details",
    "achievements", "achievement", "key achievements",
    "publications", "publication",
    "hobbies", "hobby",
    "volunteer work", "volunteering", "volunteer experience",
    "references", "reference",
    "awards", "award", "awards and recognition", "awards & recognition",
    "training", "trainings", "training and development",
    "extra curricular activities", "extracurricular activities", "co-curricular activities",
    "achievements & awards", "strengths", "hobbies and interests",
    "responsibilities", "roles & responsibilities", "roles and responsibilities",
    "personal profile", "sslc", "courses", "course",
}
SKILL_HEADING_WORDS = {"skills", "technical skills", "computer skills", "skill set", "skill",
                        "soft skill", "key skills", "core skills"}
EXPERIENCE_HEADING_WORDS = {"experience", "work experience", "internship", "internship experience"}
LANGUAGE_HEADING_WORDS = {"languages", "languages known", "language"}
EDUCATION_HEADING_WORDS = {
    "education", "academic qualifications", "educational qualifications",
    "education qualification", "academic qualification", "sslc",
}
CERTIFICATION_HEADING_WORDS = {"certifications", "certificates", "certification"}
PROJECT_HEADING_WORDS = {"projects", "academic projects", "project"}
INTERNSHIP_HEADING_WORDS = {"internship", "internship experience"}

NAME_BLOCKLIST = {
    "resume", "curriculum vitae", "curriculam vitae", "cv", "bio data",
    "biodata", "contact", "contact details", "profile", "personal details",
    "name", "full name", "candidate name", "applicant name", "your name",
}

# ============================== Name-related patterns ==============================

NAME_OF_THE_PATTERN = re.compile(r'^name\s+of\s+(the\s+)?\w+', re.IGNORECASE)
NAME_LABEL_PREFIX_PATTERN = re.compile(
    r'^(full\s+name|candidate\s+name|applicant\s+name|name)\s*[:\-]\s*', re.IGNORECASE)
_SECTION_WORDS = KNOWN_SECTIONS | NAME_BLOCKLIST | {"address", "adress"}
_SECTION_WORD_PATTERN = re.compile(
    r'\b(' + '|'.join(re.escape(s) for s in sorted(_SECTION_WORDS, key=len, reverse=True)) + r')\b',
    re.IGNORECASE,
)
COMPANY_SUFFIX_PATTERN = re.compile(
    r'\b(pvt\.?\s*ltd\.?|private\s+limited|limited|llc|inc\.?|ltd\.?|corp\.?|corporation|'
    r'llp|technologies|solutions|enterprises|industries|exim|company|co\.?)\b',
    re.IGNORECASE,
)

# ============================== Date/duration/place patterns ==============================

DATE_PATTERN = re.compile(
    r'\b(19|20)\d{2}\b|\bpresent\b|\btill\s+(now|working|date)\b|\d{1,2}/\d{4}'
    r'|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{2,4}\b',
    re.IGNORECASE,
)
DURATION_PATTERN = re.compile(r'\b\d+\+?\s*(day|days|month|months|yr|yrs|year|years)\b', re.IGNORECASE)

_INDIAN_STATES_UTS = [
    "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh",
    "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jharkhand", "Karnataka",
    "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya",
    "Mizoram", "Nagaland", "Odisha", "Punjab", "Rajasthan", "Sikkim",
    "Tamil Nadu", "Telangana", "Tripura", "Uttar Pradesh", "Uttarakhand",
    "West Bengal", "Andaman and Nicobar Islands", "Chandigarh",
    "Dadra and Nagar Haveli and Daman and Diu", "Delhi", "Jammu and Kashmir",
    "Ladakh", "Lakshadweep", "Puducherry",
]
_MAJOR_INDIAN_CITIES = [
    "Mumbai", "Delhi", "Bengaluru", "Bangalore", "Hyderabad", "Ahmedabad",
    "Chennai", "Kolkata", "Surat", "Pune", "Jaipur", "Lucknow", "Kanpur",
    "Nagpur", "Indore", "Thane", "Bhopal", "Visakhapatnam", "Vadodara",
    "Ludhiana", "Agra", "Nashik", "Faridabad", "Meerut", "Rajkot", "Varanasi",
    "Srinagar", "Amritsar", "Coimbatore", "Kochi", "Chandigarh", "Gurgaon",
    "Gurugram", "Noida", "Mysuru", "Mysore", "Mangaluru", "Mangalore",
    "Belagavi", "Hubballi", "Hubli", "Davangere", "Ballari", "Bellary",
    "Tumakuru", "Tumkur", "Shivamogga", "Shimoga", "Vijayawada", "Guntur",
    "Nellore", "Tirupati", "Warangal", "Thiruvananthapuram", "Trivandrum",
    "Kozhikode", "Calicut", "Madurai", "Salem", "Tiruchirappalli", "Trichy",
    "Bhubaneswar", "Cuttack", "Guwahati", "Ranchi", "Patna", "Raipur",
    "Dehradun", "Jodhpur", "Udaipur", "Kota", "Gwalior", "Jabalpur",
]
PLACE_PATTERN = re.compile(
    r'\b(' + '|'.join(re.escape(p) for p in _INDIAN_STATES_UTS + _MAJOR_INDIAN_CITIES) + r')\b'
    r'|\b\d{6}\b'
    r'|\b\w+\s+(nagar|road|colony|layout|stage|sector|extension|cross)\b',
    re.IGNORECASE,
)

# ============================== Education classification patterns ==============================

TENTH_PATTERN = re.compile(
    r'\b(10\s*th|xth|sslc|ssc|matriculation|class\s*[-]?\s*10|std\.?\s*x)\b', re.IGNORECASE)
TWELFTH_PATTERN = re.compile(
    r'\b(12\s*th|xiith|hsc|puc|intermediate|class\s*[-]?\s*12|std\.?\s*xii|'
    r'senior\s*secondary|pre[-\s]?university|\+2)\b', re.IGNORECASE)
DEGREE_PATTERN = re.compile(
    r'\b(b\.?\s?tech|b\.?\s?e|bachelor|m\.?\s?tech|m\.?\s?e|mba|mca|bca|'
    r'b\.?\s?sc|m\.?\s?sc|b\.?\s?com|m\.?\s?com|diploma|ph\.?d|master|'
    r'associate\s+degree|iti)\b', re.IGNORECASE)
COLLEGE_PATTERN = re.compile(r'\b(college|university|institute|polytechnic|vidyalaya|academy|school)\b', re.IGNORECASE)
DISCONTINUED_PATTERN = re.compile(r'\b(discontinued|dropped\s*out|incomplete|not\s*completed)\b', re.IGNORECASE)
GRADE_PATTERN = re.compile(
    r'\b\d{1,2}(?:\.\d+)?\s*/\s*\d{1,2}(?:\.\d+)?\s*(?:cgpa|sgpa|gpa)\b'
    r'|\b\d{1,2}(?:\.\d+)?\s*(?:cgpa|sgpa|gpa)\b'
    r'|\b(?:cgpa|sgpa|gpa)\s*[:\-]?\s*\d{1,2}(?:\.\d+)?(?:\s*/\s*\d{1,2}(?:\.\d+)?)?\b'
    r'|\b\d{1,3}(?:\.\d+)?\s*%'
    r'|\b\d{1,3}(?:\.\d+)?\s*(?:percent|percentage)\b',
    re.IGNORECASE,
)
BULLET_CHARS = "\u2022\uf0b7"
_LINE_SPLIT_PATTERN = re.compile(r'[\n' + BULLET_CHARS + r'|]+')
EDUCATION_NOISE_PATTERN = re.compile(
    r'\b(university|college|institute|polytechnic|cgpa|sgpa|percentage|'
    r'passing\s*year|year\s*of\s*passing|board|semester|degree)\b',
    re.IGNORECASE,
)
TRAILING_DATE_PATTERN = re.compile(r'\s*[-,]?\s*\(?(19|20)\d{2}\)?\s*$')
# A properly-matched certification already lives in `certifications` -- if
# a candidate still mentions "certified"/"certificate", it's leftover
# noise, not a real unmatched skill/title/industry.
CERT_WORDS_PATTERN = re.compile(r'\b(certified|certificate|certification)\b', re.IGNORECASE)
PERSONAL_DETAILS_HEADING_WORDS = {"personal details", "personal profile", "declaration", "basic details"}

# Junk that occasionally shows up as a "matched" skill/title/industry when a
# reference list happens to contain it (skills.json had ~100+ such entries).
_JUNK_TERM_WORDS = {"declaration", "summary"}
_JUNK_TERM_DOMAIN_FRAGMENTS = {"com", "net", "org", "in", "co", "www"}


# ============================== Text extraction ==============================

def extract_text_layer(path):
    doc = pymupdf.open(path)
    lines = []
    for page in doc:
        for block in page.get_text("dict")["blocks"]:
            if "lines" not in block:
                continue
            for line in block["lines"]:
                spans = line["spans"]
                if not spans:
                    continue
                text = "".join(s["text"] for s in spans).strip()
                if not text:
                    continue
                size, font = spans[0]["size"], spans[0]["font"]
                lines.append((text, round(size, 1), "bold" in font.lower(), font))
    doc.close()
    return lines if sum(len(t) for t, *_ in lines) >= 100 else None


def find_tesseract():
    found = shutil.which("tesseract") or shutil.which("tesseract.exe")
    if found:
        return found
    for candidate in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                      r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(candidate):
            return candidate
    return None


def find_poppler_bin():
    if shutil.which("pdftoppm") or shutil.which("pdftoppm.exe"):
        return None
    matches = glob.glob(os.path.expandvars(
        r"%LOCALAPPDATA%\Microsoft\WinGet\Packages\oschwartz10612.Poppler_*\poppler-*\Library\bin"
    ))
    return matches[0] if matches else None


def _configure_ocr_tools():
    import pytesseract
    tesseract_path = find_tesseract()
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    return find_poppler_bin()


def get_lines_with_style(path):
    """Text+style for a PDF/DOCX, falling back to OCR if there's no real
    text layer (e.g. a scanned page saved as PDF)."""
    lines = extract_text_layer(path)
    if lines is not None:
        return lines
    from pdf2image import convert_from_path
    import pytesseract
    poppler_path = _configure_ocr_tools()
    lines = []
    for img in convert_from_path(path, dpi=200, poppler_path=poppler_path):
        for raw_line in pytesseract.image_to_string(img).split("\n"):
            if raw_line.strip():
                lines.append((raw_line.strip(), 0.0, False, "ocr"))
    return lines


def get_lines_from_image(path):
    from PIL import Image
    import pytesseract
    _configure_ocr_tools()
    lines = []
    with Image.open(path) as img:
        for raw_line in pytesseract.image_to_string(img).split("\n"):
            if raw_line.strip():
                lines.append((raw_line.strip(), 0.0, False, "ocr"))
    return lines


def resolve_lines_getter(mime_type):
    return get_lines_from_image if mime_type in IMAGE_MIMES else get_lines_with_style


# ============================== Field extraction ==============================

def find_email(text):
    m = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', text)
    return m.group(0) if m else None


def find_phone(text):
    m = re.search(r'(?:\+?\d{1,3}[\s-]?)?\d{10}', text)
    return m.group(0) if m else None


def find_linkedin(text):
    m = re.search(r'linkedin\.com/in/[\w-]+', text)
    return m.group(0) if m else None


def _collapse_letter_spacing(text):
    """Some resumes render the name with a space after every single letter.
    With a capitalization signal (lowercase run then a capital), rebuild
    word boundaries; with none (usually ALL-CAPS), just join everything."""
    tokens = text.split()
    if len(tokens) < 4:
        return text
    single_char_ratio = sum(1 for t in tokens if len(t) == 1) / len(tokens)
    if single_char_ratio < 0.7:
        return text
    has_case_signal = any(t.islower() for t in tokens) and any(t.isupper() for t in tokens)
    words, current = [], ''
    for t in tokens:
        if len(t) != 1:
            if current:
                words.append(current)
                current = ''
            words.append(t)
            continue
        if has_case_signal and t.isupper() and current and current[-1].islower():
            words.append(current)
            current = t
        else:
            current += t
    if current:
        words.append(current)
    return ' '.join(words)


def _is_mostly_a_title(candidate, titles_matcher):
    """True when a matched job title accounts for almost all of the
    candidate's real (alphabetic) words -- 'Purchase & Sales Engineer'
    matches 'Sales Engineer' (2 of 3 real words), so the whole line is a
    job title, not a name."""
    if titles_matcher is None:
        return False
    title_hits = match_exact(candidate, titles_matcher) or []
    if not title_hits:
        return False
    real_words = [w for w in candidate.split() if any(c.isalpha() for c in w)]
    if not real_words:
        return False
    longest_hit_words = max(len(t.split()) for t in title_hits)
    return longest_hit_words >= len(real_words) - 1


def _strip_trailing_title(name_candidate, titles_matcher):
    """Some resumes glue a job title onto the name with no space
    ('SYAMSALINSP PURCHASE EXECUTIVE'). If a known title matches right at
    the end, strip it off."""
    if titles_matcher is None:
        return name_candidate
    for title in sorted(match_exact(name_candidate, titles_matcher) or [], key=len, reverse=True):
        pattern = re.compile(r'\s*' + flexible_term_pattern(title) + r'\s*$', re.IGNORECASE)
        stripped = pattern.sub('', name_candidate).strip(" -,")
        if len(stripped) >= 3 and any(c.isalpha() for c in stripped):
            return stripped
    return name_candidate


def _name_from_email(email):
    """Last-resort fallback when nothing on the page qualifies as a name."""
    if not email or '@' not in email:
        return None
    local = email.split('@')[0]
    local = re.sub(r'[\d_]+', ' ', local)
    words = [p.strip().capitalize() for p in re.split(r'[.\-\s]+', local) if p.strip()]
    if not words:
        return None
    candidate = ' '.join(words)
    return candidate if len(candidate) >= 3 else None


def find_name(lines, headers=None, titles_matcher=None, email=None):
    """The name is usually the largest text on the page. Rejects: this
    resume's own detected headings; a form label ('Name of the College',
    'Name: John Doe'); anything with a colon anywhere; a sentence (too
    many words); a phone/email; resume-jargon words anywhere in the line
    ('Skills', 'Contact', 'Address', ...); a line that's mostly a matched
    job title; a company name; a place; a date/duration. Letter-spaced
    text ('J O H N') is rejoined, and a trailing ' - '/'|' credentials
    segment is cut off. Falls back to a name derived from the email if
    nothing on the page qualifies at all."""
    header_set = {h.strip() for h in (headers or [])}
    candidates = sorted(lines, key=lambda line: line[1], reverse=True)
    for text, size, is_bold, font in candidates:
        clean = text.strip()
        if not clean or len(clean) < 3:
            continue
        if clean in header_set:
            continue
        if NAME_OF_THE_PATTERN.match(clean):
            continue
        clean = NAME_LABEL_PREFIX_PATTERN.sub('', clean).strip()
        if ':' in clean:
            continue
        clean = re.split(r'\s[-|]\s', clean)[0].strip()
        clean = _collapse_letter_spacing(clean)
        if len(clean) < 3 or len(clean.split()) > 6:
            continue
        normalized = re.sub(r'[^a-z ]', '', clean.lower()).strip()
        if normalized in NAME_BLOCKLIST or normalized in KNOWN_SECTIONS:
            continue
        if _SECTION_WORD_PATTERN.search(clean):
            continue
        if find_phone(clean) or find_email(clean):
            continue
        if not any(c.isalpha() for c in clean):
            continue
        if _is_mostly_a_title(clean, titles_matcher):
            continue
        if COMPANY_SUFFIX_PATTERN.search(clean):
            continue
        if PLACE_PATTERN.search(clean):
            continue
        if DATE_PATTERN.search(clean) or DURATION_PATTERN.search(clean):
            continue
        return _strip_trailing_title(clean, titles_matcher)
    return _name_from_email(email)


def find_headers(lines):
    if not lines:
        return []
    body_size = Counter(size for _, size, _, _ in lines).most_common(1)[0][0]
    body_font = Counter(font for _, _, _, font in lines).most_common(1)[0][0]
    squashed_known = {normalize_key(s) for s in KNOWN_SECTIONS}

    headers = []
    for text, size, is_bold, font in lines:
        if len(text) > 40:
            continue
        looks_like_header = (
            size > body_size + 0.5 or is_bold or font != body_font
            or (text.isupper() and len(text.split()) <= 4)
        )
        if not looks_like_header:
            continue
        normalized = re.sub(r'[\s:.\-]+$', '', text.strip()).lower()
        word_count = len(text.split())
        is_match = (
            normalized in KNOWN_SECTIONS
            or normalize_key(text) in squashed_known
            or (word_count <= 2 and any(k in normalized for k in KNOWN_SECTIONS))
        )
        if is_match:
            headers.append(text.strip())
    return headers


def load_skill_heading_vocab(lookups_dir):
    """Flattens normalized_category_mapping.json's {category: [heading
    variants]} into one set of normalized heading phrases -- except the
    'language' category, whose headings ('Languages Known', etc.) belong
    to the LANGUAGE heading set, not skills. Including it here was a real
    bug: it made 'Languages Known' double as a skill heading too, so
    whatever else that section's raw text contained (email/gender, in one
    resume, from a heading-merge artifact) got scanned for skill
    candidates instead of just being read for languages."""
    mapping = load_canonical_list(NORMALIZED_CATEGORY_MAPPING_FILE, lookups_dir)
    vocab = set(SKILL_HEADING_WORDS)
    if mapping:
        for category, headings in mapping.items():
            if category == "language":
                continue
            for h in headings:
                vocab.add(clean_heading(h))
    return vocab


def detect_extra_headings(lines, vocab):
    """Lines matching the category-mapping vocabulary by text alone,
    regardless of font/style, plus any short line ending in ':' -- both
    prevent an unrecognized heading's content from bleeding into whatever
    section came before it."""
    extra = []
    for text, _, _, _ in lines:
        clean = text.strip()
        if not clean or len(clean) > 40:
            continue
        if clean_heading(clean) in vocab:
            extra.append(clean)
        elif clean.endswith(':') and len(clean.split()) <= 5:
            extra.append(clean)
    return extra


def extract_sections(lines, headers):
    sections, current = {}, None
    header_set = {h.strip() for h in headers}
    for text, _, _, _ in lines:
        if text.strip() in header_set:
            current = text.strip()
            sections[current] = []
        elif current is not None:
            sections[current].append(text)
    return {h: "\n".join(v) for h, v in sections.items()}


# ============================ Canonical-list matching (exact match only) =============================

def load_canonical_list(filename, lookups_dir=None):
    lookups_dir = lookups_dir or DEFAULT_LOOKUPS_DIR
    if not os.path.isabs(lookups_dir):
        lookups_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), lookups_dir)
    path = os.path.join(lookups_dir, filename)
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            import json
            return json.load(f)
    except Exception as e:
        print(f"Warning: could not read {filename} ({e}) — treating it as missing.")
        return None


def flexible_term_pattern(term):
    """Regex for a canonical term where INTERNAL separators (space/dot/dash
    between two alnum chars) become flexible, e.g. 'Node.js' also matches
    'Node JS'/'NodeJS'. Leading/trailing punctuation stays literal."""
    parts = []
    i, n = 0, len(term)
    while i < n:
        c = term[i]
        if c in " .-":
            j = i
            while j < n and term[j] in " .-":
                j += 1
            if i > 0 and term[i - 1].isalnum() and j < n and term[j].isalnum():
                parts.append(r'[\s.\-]*')
            else:
                parts.append(re.escape(term[i:j]))
            i = j
        else:
            parts.append(re.escape(c))
            i += 1
    return "".join(parts)


def normalize_key(text):
    return re.sub(r'[^a-z0-9]', '', text.lower())


def build_matcher(terms):
    """A combined regex for a whole canonical list, tolerant only of a
    term's OWN internal punctuation (not resume-inserted spacing like
    'Auto Cad' for 'AutoCAD' -- that's deliberate, see match_exact)."""
    canonical_map = {}
    for t in terms:
        canonical_map.setdefault(normalize_key(t), t)
    sorted_terms = sorted(terms, key=len, reverse=True)
    pattern_str = r'\b(?:' + '|'.join(flexible_term_pattern(t) for t in sorted_terms) + r')\b'
    pattern = None
    if re2 is not None:
        try:
            opts = re2.Options()
            opts.case_sensitive = False
            opts.max_mem = 1 << 30
            pattern = re2.compile(pattern_str, options=opts)
        except Exception:
            pattern = None
    if pattern is None:
        pattern = re.compile(pattern_str, re.IGNORECASE)
    return pattern, canonical_map


def match_exact(text, matcher, exclude=None):
    """Only literal (or punctuation-flexible, per flexible_term_pattern)
    matches against the canonical list -- no spacing-inserted fuzzy
    matching, so 'Auto Cad' does NOT match canonical 'AutoCAD'; it stays
    unmatched instead. Precision over recall, by design."""
    if matcher is None:
        return []
    pattern, canonical_map = matcher
    exclude_keys = {normalize_key(e) for e in (exclude or [])}
    found = set()
    for m in pattern.finditer(text):
        key = normalize_key(m.group(0))
        if key not in exclude_keys:
            found.add(canonical_map.get(key, m.group(0)))
    return sorted(found)


def _lang_key(text):
    return re.sub(r'[^\w]', '', text, flags=re.UNICODE).lower()


def build_language_matcher(lang_dict):
    if not lang_dict:
        return None
    canonical_map, variants_by_canonical, all_variants = {}, {}, []
    for canonical, variants in lang_dict.items():
        variants_by_canonical[canonical] = list(variants)
        for v in variants:
            key = _lang_key(v)
            if key:
                canonical_map.setdefault(key, canonical)
                all_variants.append(v)
    if not all_variants:
        return None
    sorted_terms = sorted(set(all_variants), key=len, reverse=True)
    pattern_str = r'\b(?:' + '|'.join(flexible_term_pattern(t) for t in sorted_terms) + r')\b'
    return re.compile(pattern_str, re.IGNORECASE), canonical_map, variants_by_canonical


def match_languages(text, matcher):
    if matcher is None:
        return []
    pattern, canonical_map = matcher[0], matcher[1]
    found = set()
    for m in pattern.finditer(text):
        canonical = canonical_map.get(_lang_key(m.group(0)))
        if canonical:
            found.add(canonical)
    return sorted(found)


def _strip_languages(items, languages_matcher):
    if not items or languages_matcher is None:
        return items
    canonical_map = languages_matcher[1]
    return [i for i in items if _lang_key(i) not in canonical_map]


def _is_junk_term(term):
    stripped = term.strip()
    if re.fullmatch(r'\.?\d[\d.]*', stripped):
        return True
    if stripped.lower().strip('.') in _JUNK_TERM_DOMAIN_FRAGMENTS:
        return True
    normalized = re.sub(r'[^a-z]', '', stripped.lower())
    return normalized in _JUNK_TERM_WORDS


def _filter_junk_terms(items):
    return [i for i in items if i and not _is_junk_term(i)]


# ========================= Candidate extraction & junk filtering ========================

def clean_heading(text):
    return re.sub(r'[\s:.\-]+$', '', text.strip()).lower()


def is_junk_candidate(text, language_matcher=None):
    lower = text.lower()
    if 'cgpa' in lower or 'gpa' in lower or 'semester' in lower:
        return True
    if re.search(r'\bsem\b', lower):
        return True
    if re.fullmatch(r'[\d.\s%-]+', text):
        return True
    if DATE_PATTERN.search(text) or DURATION_PATTERN.search(text) or PLACE_PATTERN.search(text):
        return True
    if language_matcher is not None and match_languages(text, language_matcher):
        return True
    return False


def is_education_noise(text):
    return bool(EDUCATION_NOISE_PATTERN.search(text))


def split_skill_candidates(text, language_matcher=None):
    candidates = []
    for line in text.split("\n"):
        line = re.sub(r'^[A-Za-z][A-Za-z &]{1,29}:\s*', '', line.strip())
        for part in re.split(r'\s*(?:,|;|&|\band\b)\s*', line, flags=re.IGNORECASE):
            part = part.strip(" .•-")
            if 1 < len(part) <= 50 and not is_junk_candidate(part, language_matcher):
                candidates.append(part)
    return candidates


def split_title_candidates(text, language_matcher=None):
    candidates = []
    for line in text.split("\n"):
        line = line.strip(" .•-")
        word_count = len(line.split())
        if 1 < word_count <= 6 and not line.startswith("•") and not is_junk_candidate(line, language_matcher):
            candidates.append(line)
    return candidates


def split_industry_candidates(text, language_matcher=None):
    return split_skill_candidates(text, language_matcher)


def flatten_comma_candidates(candidates):
    """split_title_candidates() keeps a whole line as one candidate even
    when it's really several comma-separated items ('Attendance
    Maintenance, Filing, Leave') -- splitting here makes every item
    individually comparable against matched terms."""
    flattened = []
    for c in candidates:
        parts = re.split(r'\s*,\s*', c) if ',' in c else [c]
        flattened.extend(p.strip() for p in parts if p.strip())
    return flattened


def find_discarded(candidates, matched):
    matched_terms = [m for m in (matched or []) if m]
    seen, discarded = set(), []
    for c in candidates:
        key = normalize_key(c)
        if not key or key in seen:
            continue
        if any(re.search(r'\b' + re.escape(m) + r'\b', c, re.IGNORECASE) for m in matched_terms):
            continue
        seen.add(key)
        discarded.append(c)
    return discarded


# ================== Education / certifications / projects / internships ==================

def split_lines(raw_text):
    return [l.strip(" .,-" + BULLET_CHARS) for l in _LINE_SPLIT_PATTERN.split(raw_text) if l.strip(" .,-" + BULLET_CHARS)]


def extract_array_section(sections, heading_words):
    raw = "\n".join(v for h, v in sections.items() if clean_heading(h) in heading_words)
    return split_lines(raw)


def strip_grades(line):
    cleaned = GRADE_PATTERN.sub('', line)
    cleaned = re.sub('[' + BULLET_CHARS + ']+', ' ', cleaned)
    cleaned = re.sub(r'\s{2,}', ' ', cleaned)
    return cleaned.strip(" .,-")


def classify_education_line(line):
    if TENTH_PATTERN.search(line):
        return 'tenth'
    if TWELFTH_PATTERN.search(line):
        return 'twelfth'
    if DEGREE_PATTERN.search(line):
        return 'degree'
    if COLLEGE_PATTERN.search(line):
        return 'college'
    return None


TRAILING_CONNECTOR_PATTERN = re.compile(r'\s+(in|of|and|for|the|to|on|&)\s*$', re.IGNORECASE)


def clean_entry_tail(text):
    """Strips a stray trailing colon (a leftover label marker, e.g.
    'PUC (2006-07):') and a dangling connector word ('Specialization in'
    -> 'Specialization') that a cut-off resume line can leave behind."""
    text = text.rstrip()
    text = re.sub(r':\s*$', '', text).rstrip()
    text = TRAILING_CONNECTOR_PATTERN.sub('', text).rstrip()
    return text


def extract_educations(sections):
    """Only degree/college/10th/12th mentions, grades stripped, adjacent
    lines merged (a 'college' line right after a degree/10th/12th line is
    that qualification's institution, not a separate entry). Dates are
    kept only on degree/college entries, stripped from bare 10th/12th
    entries, which shouldn't carry one at all."""
    blocks = []
    for h, v in sections.items():
        ch = clean_heading(h)
        override = None
        if ch == 'sslc' or TENTH_PATTERN.search(h):
            override = 'tenth'
        elif ch not in EDUCATION_HEADING_WORDS and TWELFTH_PATTERN.search(h):
            override = 'twelfth'
        if ch in EDUCATION_HEADING_WORDS or override:
            blocks.append((override, v))

    entries, seen = [], set()

    def flush(entry):
        if not entry:
            return
        text = entry['text']
        if entry['type'] in ('tenth', 'twelfth'):
            text = TRAILING_DATE_PATTERN.sub('', text).strip(" .,-")
        text = clean_entry_tail(text)
        key = re.sub(r'[^a-z0-9]', '', text.lower())
        if key and key not in seen:
            seen.add(key)
            entries.append(text)

    for override, raw in blocks:
        lines = [strip_grades(l) for l in split_lines(raw)]
        lines = [l for l in lines if l]
        current = None
        for line in lines:
            if DISCONTINUED_PATTERN.search(line):
                continue
            content_type = classify_education_line(line)
            is_numeric_fragment = bool(re.fullmatch(r'[\d\s./%()-]+', line))
            ctype = content_type or (override if (current is None and not is_numeric_fragment) else None)
            if ctype in ('tenth', 'twelfth', 'degree'):
                flush(current)
                current = {'type': ctype, 'text': line}
            elif ctype == 'college':
                if current:
                    current['text'] += ', ' + line
                    flush(current)
                    current = None
                else:
                    flush({'type': 'college', 'text': line})
        flush(current)
    return entries


def extract_certifications(sections):
    """Certificate title lines only -- drops bare issuer+year lines,
    education content that leaked in, prose fragments, and strips any
    date from every kept line (trailing OR mid-line, e.g. 'Certified in
    2023 by Coursera' -> 'Certified by Coursera') -- a certificate name
    carrying its own date is rare enough that it's simpler and safer to
    just never keep one."""
    raw = "\n".join(v for h, v in sections.items() if clean_heading(h) in CERTIFICATION_HEADING_WORDS)
    kept = []
    for line in split_lines(raw):
        if len(line) > 100 or len(line.split()) > 12:
            continue
        if line[:1].islower():
            continue
        if is_education_noise(line):
            continue
        if re.fullmatch(r"[A-Za-z .&'-]{2,25}[-,]?\s*\(?\d{4}\)?", line):
            continue
        if re.fullmatch(r'\(?\d{4}\)?', line):
            continue
        line = DATE_PATTERN.sub('', line)
        line = re.sub(r'\s{2,}', ' ', line)
        line = clean_entry_tail(line).strip(" .,-()")
        if line:
            kept.append(line)
    return kept


def build_unmatched_string(unmatched_skills, unmatched_titles, unmatched_industries, all_matched,
                            max_items=15, max_len=280):
    """Combines the three unmatched candidate lists into one, dropping
    anything already present in ANY matched field -- plus company names,
    education leftovers, and certification-word noise, none of which are
    real unmatched skills/titles/industries."""
    matched_keys = {m.strip().lower() for m in all_matched if m}
    combined, seen = [], set()
    for item in unmatched_skills + unmatched_titles + unmatched_industries:
        key = item.strip().lower()
        if not key or key in seen or key in matched_keys:
            continue
        if len(item) > 60 or is_education_noise(item):
            continue
        if COMPANY_SUFFIX_PATTERN.search(item) or CERT_WORDS_PATTERN.search(item):
            continue
        seen.add(key)
        combined.append(item.strip())
    combined = combined[:max_items]
    text = ", ".join(combined)
    if len(text) > max_len:
        text = text[:max_len].rsplit(",", 1)[0].rstrip() + ", ..."
    return text


def as_csv(items):
    return ", ".join(items) if items else ""


# ======================== Per-resume parsing (worker side) ========================

_skills_matcher = _titles_matcher = _industries_matcher = _languages_matcher = None
_skill_heading_vocab = None


def _init_worker(lookups_dir):
    global _skills_matcher, _titles_matcher, _industries_matcher, _languages_matcher, _skill_heading_vocab
    skills = load_canonical_list("skills.json", lookups_dir)
    titles = load_canonical_list("titles.json", lookups_dir)
    industries = load_canonical_list("industries.json", lookups_dir)
    languages = load_canonical_list("languages.json", lookups_dir)
    _skills_matcher = build_matcher(skills) if skills else None
    _titles_matcher = build_matcher(titles) if titles else None
    _industries_matcher = build_matcher(industries) if industries else None
    _languages_matcher = build_language_matcher(languages) if languages else None
    _skill_heading_vocab = load_skill_heading_vocab(lookups_dir)


def parse_resume(path, mime_type):
    lines = resolve_lines_getter(mime_type)(path)
    full_text = "\n".join(t for t, _, _, _ in lines)

    style_headers = find_headers(lines)
    extra_headers = detect_extra_headings(lines, _skill_heading_vocab)
    headers = list(dict.fromkeys(style_headers + extra_headers))
    sections = extract_sections(lines, headers)

    email = find_email(full_text)
    phone = find_phone(full_text)
    linkedin = find_linkedin(full_text)
    name = find_name(lines, headers, _titles_matcher, email)

    def mask(text):
        for value in (email, phone, linkedin):
            if value:
                text = text.replace(value, " ")
        return re.sub(r'\b(?:www\.|https?://)\S+', ' ', text, flags=re.IGNORECASE)

    skill_heading_words = _skill_heading_vocab | SKILL_HEADING_WORDS
    raw_skill_section = "\n".join(v for h, v in sections.items() if clean_heading(h) in skill_heading_words)
    raw_title_section = "\n".join(v for h, v in sections.items() if clean_heading(h) in EXPERIENCE_HEADING_WORDS)

    # Personal Details (marital status, religion, nationality, gender) is
    # never a real skill/title/industry -- strip it wherever it could leak
    # in: the whole-resume fallback used when there's no dedicated Skills/
    # Experience section, AND the section text itself, in case it merged
    # into the wrong heading's content. Everything else (email/phone/name/
    # language detection) still scans the true full_text, unaffected.
    personal_details_text = "\n".join(v for h, v in sections.items() if clean_heading(h) in PERSONAL_DETAILS_HEADING_WORDS)
    fallback_text = full_text
    if personal_details_text:
        fallback_text = full_text.replace(personal_details_text, " ")
        raw_skill_section = raw_skill_section.replace(personal_details_text, " ")
        raw_title_section = raw_title_section.replace(personal_details_text, " ")

    skill_text = mask(raw_skill_section) if raw_skill_section else mask(fallback_text)
    title_text = mask(raw_title_section) if raw_title_section else mask(fallback_text)
    industry_text = mask(raw_title_section)  # industries: Experience section only

    matched_skills = _filter_junk_terms(_strip_languages(
        match_exact(skill_text, _skills_matcher, exclude=headers), _languages_matcher))
    matched_titles = _filter_junk_terms(_strip_languages(
        match_exact(title_text, _titles_matcher, exclude=headers), _languages_matcher))
    matched_industries = _filter_junk_terms(_strip_languages(
        match_exact(industry_text, _industries_matcher, exclude=headers) if raw_title_section else [],
        _languages_matcher))
    languages = match_languages(mask(full_text), _languages_matcher)

    raw_skill_candidates = [c for c in split_skill_candidates(raw_skill_section, _languages_matcher)
                             if not is_education_noise(c)] if raw_skill_section else []
    raw_title_candidates = [c for c in flatten_comma_candidates(split_title_candidates(raw_title_section, _languages_matcher))
                             if not is_education_noise(c)] if raw_title_section else []
    raw_industry_candidates = [c for c in split_industry_candidates(industry_text, _languages_matcher)
                                if not is_education_noise(c)] if raw_title_section else []

    not_matched_skills = _filter_junk_terms(find_discarded(raw_skill_candidates, matched_skills))
    not_matched_titles = _filter_junk_terms(find_discarded(raw_title_candidates, matched_titles))
    not_matched_industries = _filter_junk_terms(find_discarded(raw_industry_candidates, matched_industries))

    all_matched = matched_skills + matched_titles + matched_industries
    unmatched = build_unmatched_string(not_matched_skills, not_matched_titles, not_matched_industries, all_matched)

    return {
        "name": name,
        "email": email,
        "linkedin": linkedin,
        "phone": phone,
        "skills": as_csv(matched_skills),
        "titles": as_csv(matched_titles),
        "industries": as_csv(matched_industries),
        "languages": as_csv(languages),
        "educations": as_csv(extract_educations(sections)),
        "certifications": as_csv(extract_certifications(sections)),
        "projects": extract_array_section(sections, PROJECT_HEADING_WORDS),
        "internships": extract_array_section(sections, INTERNSHIP_HEADING_WORDS),
        "unmatched": unmatched,
    }


def _worker_task(path, filename, mime_type, doc_id):
    try:
        doc = parse_resume(path, mime_type)
    except Exception as e:
        return filename, {"id": doc_id, "error": str(e)}
    doc["id"] = doc_id
    return filename, doc


# ============================ .doc conversion, hashing, MeiliSearch ============================

def find_soffice():
    for candidate in ("soffice", "soffice.exe"):
        found = shutil.which(candidate)
        if found:
            return found
    for candidate in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(candidate):
            return candidate
    return None


def convert_doc_files_to_docx(doc_paths, tmp_dir):
    converted = {}
    if not doc_paths:
        return converted
    soffice = find_soffice()
    if not soffice:
        return converted
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_dir = os.path.abspath(tmp_dir)
    batch_size = 50
    for i in range(0, len(doc_paths), batch_size):
        batch = doc_paths[i:i + batch_size]
        try:
            subprocess.run(
                [soffice, "--headless", "--norestore", "--convert-to", "docx", "--outdir", tmp_dir]
                + [os.path.abspath(p) for p in batch],
                capture_output=True, timeout=60 * len(batch), check=False,
            )
        except Exception:
            continue
        for doc_path in batch:
            docx_path = os.path.join(tmp_dir, os.path.splitext(os.path.basename(doc_path))[0] + ".docx")
            if os.path.exists(docx_path):
                converted[doc_path] = docx_path
    return converted


def make_doc_id(filename):
    return hashlib.md5(filename.encode("utf-8")).hexdigest()


def _load_hash_index(hash_dir):
    path = os.path.join(hash_dir, "seen_hashes.json")
    if not os.path.exists(path):
        return {}
    try:
        import json
        with open(path) as f:
            return json.load(f)
    except Exception:
        return {}


def _save_hash_index(hash_dir, seen_hashes):
    import json
    os.makedirs(hash_dir, exist_ok=True)
    with open(os.path.join(hash_dir, "seen_hashes.json"), "w") as f:
        json.dump(seen_hashes, f, indent=2)


def fetch_existing_ids(index):
    existing = set()
    offset, limit = 0, 10000
    while True:
        batch = index.get_documents({"offset": offset, "limit": limit, "fields": ["id"]}).results
        if not batch:
            break
        existing.update(d.id for d in batch)
        offset += limit
    return existing


# ================================ Driver ================================

def main(folder_name, workers, lookups_dir, hash_dir, meili_url, meili_index, meili_api_key):
    client = meilisearch.Client(meili_url, meili_api_key)
    index = client.index(meili_index)
    try:
        existing_ids = fetch_existing_ids(index)
    except Exception as e:
        print(f"Could not reach MeiliSearch at {meili_url} ({e}). Is the server running?")
        return
    print(f"{len(existing_ids)} documents already in '{meili_index}'.")

    all_files = sorted(os.listdir(folder_name))
    mime_by_file = {}
    for f in all_files:
        path = os.path.join(folder_name, f)
        if not os.path.isfile(path):
            continue
        mime_type = detect_mime_type(path)
        if mime_type in SUPPORTED_MIMES:
            mime_by_file[f] = mime_type
    resume_files = [f for f in all_files if f in mime_by_file]
    unsupported_files = [f for f in all_files if f not in mime_by_file]

    documents = []
    for f in unsupported_files:
        doc_id = make_doc_id(f)
        if doc_id in existing_ids:
            continue
        ext = os.path.splitext(f)[1].lower() or "(no extension)"
        documents.append({
            "id": doc_id, "skipped": True,
            "message": (
                f"This file was not processed because its content (checked directly, "
                f"not just its '{ext}' extension) isn't a recognized resume format. "
                "Supported: PDF, DOCX, DOC, and image files."
            ),
        })

    seen_hashes = _load_hash_index(hash_dir)
    todo, already_done, skipped_duplicates, unreadable = [], 0, 0, 0
    for f in resume_files:
        doc_id = make_doc_id(f)
        if doc_id in existing_ids:
            already_done += 1
            continue
        try:
            with open(os.path.join(folder_name, f), "rb") as fh:
                content_hash = hashlib.sha256(fh.read()).hexdigest()
        except OSError as e:
            documents.append({"id": doc_id, "error": f"could not read file: {e}"})
            unreadable += 1
            continue
        original = seen_hashes.get(content_hash)
        if original is not None and original != f:
            documents.append({
                "id": doc_id, "duplicate_of": original,
                "message": f"The file is a duplicate of the {original} file.",
            })
            skipped_duplicates += 1
            continue
        seen_hashes[content_hash] = f
        todo.append(f)
    _save_hash_index(hash_dir, seen_hashes)

    print(f"{len(todo)} to process, {already_done} already in the index, "
          f"{skipped_duplicates} duplicate content, {len(unsupported_files)} unsupported, "
          f"{unreadable} unreadable")

    doc_files = [f for f in todo if mime_by_file[f] == DOC_MIME]
    doc_tmp_dir = tempfile.mkdtemp(prefix="structured_doc_to_docx_")
    doc_to_docx = convert_doc_files_to_docx([os.path.join(folder_name, f) for f in doc_files], doc_tmp_dir)
    if doc_files:
        print(f".doc: {len(doc_to_docx)}/{len(doc_files)} converted to .docx")

    tasks = []
    for f in todo:
        mime_type = mime_by_file[f]
        if mime_type == DOC_MIME:
            path = doc_to_docx.get(os.path.join(folder_name, f))
            if path is None:
                documents.append({
                    "id": make_doc_id(f),
                    "error": "could not convert .doc to .docx (LibreOffice missing or conversion failed)",
                })
                continue
            mime_type = DOCX_MIME
        else:
            path = os.path.join(folder_name, f)
        tasks.append((path, f, mime_type, make_doc_id(f)))

    errors = 0
    start = time.time()
    with ProcessPoolExecutor(max_workers=workers, initializer=_init_worker, initargs=(lookups_dir,)) as pool:
        futures = [pool.submit(_worker_task, path, filename, mime_type, doc_id) for path, filename, mime_type, doc_id in tasks]
        for i, future in enumerate(as_completed(futures), 1):
            try:
                filename, doc = future.result()
            except BrokenProcessPool as e:
                print(f"  a worker process crashed ({e}) -- stopping early.")
                break
            if "error" in doc:
                errors += 1
            documents.append(doc)
            if i % 200 == 0 or i == len(tasks):
                elapsed = time.time() - start
                print(f"  {i}/{len(tasks)} ({(i / elapsed if elapsed else 0):.1f}/s, {errors} errors so far)")

    shutil.rmtree(doc_tmp_dir, ignore_errors=True)

    if not documents:
        print("Nothing new to upload.")
        return

    index.update_filterable_attributes(["skills", "titles", "industries", "languages"])
    batch_size = 500
    for i in range(0, len(documents), batch_size):
        batch = documents[i:i + batch_size]
        task = index.add_documents(batch, primary_key="id")
        client.wait_for_task(task.task_uid, timeout_in_ms=120_000)
        print(f"  uploaded {min(i + batch_size, len(documents))}/{len(documents)}")

    stats = index.get_stats()
    elapsed = time.time() - start
    print(f"Total: {len(documents)} documents -> MeiliSearch index '{meili_index}' "
          f"({stats.number_of_documents} documents total there now) in {elapsed / 60:.1f} min, "
          f"{errors + unreadable} error(s)")


if __name__ == "__main__":
    arg_parser = argparse.ArgumentParser(description="Parse resumes into MeiliSearch (the one script for this project).")
    arg_parser.add_argument("folder_name", nargs="?", default=DEFAULT_INPUT_DIR,
                             help=f"Folder of resumes to parse (default: {DEFAULT_INPUT_DIR})")
    arg_parser.add_argument("--workers", type=int, default=os.cpu_count(),
                             help="Worker processes (default: CPU core count)")
    arg_parser.add_argument("--lookups-dir", default=DEFAULT_LOOKUPS_DIR,
                             help=f"Folder containing skills.json/titles.json/industries.json/languages.json/"
                                  f"normalized_category_mapping.json (default: '{DEFAULT_LOOKUPS_DIR}')")
    arg_parser.add_argument("--hash-dir", default=DEFAULT_HASH_DIR,
                             help=f"Folder for the content-hash checkpoint (default: {DEFAULT_HASH_DIR})")
    arg_parser.add_argument("--meili-url", default=DEFAULT_MEILI_URL,
                             help=f"MeiliSearch server URL (default: {DEFAULT_MEILI_URL})")
    arg_parser.add_argument("--meili-index", default=DEFAULT_MEILI_INDEX,
                             help=f"Index to upload into (default: {DEFAULT_MEILI_INDEX})")
    arg_parser.add_argument("--meili-api-key", default=DEFAULT_MEILI_API_KEY,
                             help="MeiliSearch API key (default: read from .env's MEILI_API_KEY)")
    args = arg_parser.parse_args()
    main(args.folder_name, args.workers, args.lookups_dir, args.hash_dir,
         args.meili_url, args.meili_index, args.meili_api_key)
