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
import unicodedata
import pymupdf
try:
    import re2  # optional: a faster regex engine for the big canonical-list matcher
except ImportError:
    re2 = None
import meilisearch
from meilisearch.errors import MeilisearchApiError
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
# Set MEILI_URL in .env (local Docker, or the ngrok tunnel to the shared server).
DEFAULT_MEILI_URL = os.environ.get("MEILI_URL", "http://localhost:7700")
DEFAULT_MEILI_INDEX = "resumes_structured"
DEFAULT_MEILI_API_KEY = os.environ.get("MEILI_API_KEY")
# Duplicate / unsupported / error notes go to their own index ("<index>_notes"
# by default) so the resume index holds only real candidates.
NOTES_INDEX_SUFFIX = "_notes"
# Merged into (never replacing) the index's existing filterable attributes.
# The *_list fields are real arrays, one item per element, so IN / = match a
# single skill exactly; the comma-string versions stay for display.
RESUME_FILTERABLE_ATTRIBUTES = [
    "skills_list", "titles_list", "industries_list", "languages_list",
    "skills", "titles", "industries", "languages",
    "name", "email", "phone", "linkedin", "educations", "certifications", "projects",
    "educations_list", "certifications_list",
]
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

# Every section heading is sorted into one category, so downstream code asks
# "is this an education section?" instead of matching one exact wording.
# An exact phrase counts even with no styling (plenty of resumes set headings
# in plain body text); a keyword-only match also needs a style cue (bold,
# bigger, another font, ALL CAPS, trailing colon), so a skills-list line such
# as "Project Management" is never mistaken for a heading.
SECTION_PHRASES = {
    "summary": [
        "summary", "professional summary", "career summary", "profile summary", "profile",
        "professional profile", "objective", "objectives", "career objective", "career objectives",
        "about", "about me", "executive summary", "career profile", "personal summary",
        "professional overview", "career overview", "profile overview",
    ],
    "skills": sorted(SKILL_HEADING_WORDS) + [
        "it skills", "technical skill", "key skill", "skills summary", "skill summary",
        "summary of skills", "core competencies", "core competency", "competencies",
        "key competencies", "areas of expertise", "area of expertise", "expertise",
        "technical expertise", "technical proficiency", "technical skills and tools",
        "software skills", "software proficiency", "software known", "tools", "technologies",
        "tools and technologies", "technical knowledge", "computer knowledge",
        "computer proficiency", "skills and abilities", "skill highlights", "key strengths",
        "strengths", "strength", "soft skills", "professional skills", "personal skills",
        "functional skills", "domain knowledge", "technical summary", "technical profile",
        "skills and competencies", "skills and expertise", "technical competencies",
        "it proficiency", "computer skills and knowledge", "technical skill set",
    ],
    "experience": [
        "experience", "work experience", "professional experience", "work history",
        "employment history", "employment", "career history", "experience summary",
        "professional background", "employment details", "organisational experience",
        "organizational experience", "work experience details", "experience details",
        "industrial experience", "relevant experience", "job experience", "career details",
        "professional experience summary", "work profile", "employment record",
        "responsibilities", "roles and responsibilities", "key responsibilities",
        "job responsibilities", "professional work experience", "total experience",
    ],
    "internships": [
        "internship", "internships", "internship experience", "internship details",
        "industrial training", "in-plant training", "inplant training", "in plant training",
        "summer internship", "summer training", "apprenticeship", "industrial exposure",
        "internship and training", "internships and trainings", "internship and projects",
    ],
    "education": [
        "education", "educational qualification", "educational qualifications",
        "education qualification", "education qualifications", "academic qualification",
        "academic qualifications", "academic details", "academic profile", "academics",
        "academic background", "educational background", "educational profile",
        "education details", "educational details", "academic credentials",
        "scholastic profile", "qualification", "qualifications", "academic record",
        "academic history", "academic summary", "education background", "academia",
        "education and training", "education and certifications", "educational credentials",
        "academic qualification details", "qualification details", "sslc",
    ],
    "projects": [
        "projects", "project", "academic projects", "academic project", "project details",
        "projects undertaken", "project undertaken", "key projects", "major projects",
        "major project", "mini project", "mini projects", "main project", "final year project",
        "final year projects", "project work", "project experience", "projects handled",
        "projects done", "personal projects", "professional projects", "technical projects",
        "projects worked on", "project profile", "project summary", "live projects",
        "college projects", "engineering projects", "research projects",
        "academic project details", "projects and internships", "project experience details",
    ],
    "certifications": [
        # not "certificate" alone: that's an education-table column header
        "certifications", "certification", "certificates",
        "certification courses", "certificate courses", "courses and certifications",
        "certifications and courses", "licenses and certifications",
        "licences and certifications", "professional certifications",
        "certifications and training", "training and certifications",
        "certifications and achievements", "other certifications", "technical certifications",
        "online certifications", "certified courses", "additional certifications",
    ],
    "languages": [
        "languages", "languages known", "language", "language proficiency", "languages spoken",
        "linguistic abilities", "language skills",
    ],
    "personal": [
        "personal details", "personal profile", "basic details", "personal information", "personal data", "personal particulars", "bio data",
        "biodata", "personal info", "personal dossier", "personal snapshot", "contact",
        "contact details", "contact information", "contact info", "address", "personal vitae",
    ],
    "hobbies": [
        "hobbies", "hobby", "interests", "hobbies and interests", "personal interests",
        "extra curricular activities", "extracurricular activities", "co-curricular activities",
        "co curricular activities", "extra-curricular activities", "involvement", "activities",
    ],
    "achievements": [
        "achievements", "achievement", "key achievements", "awards", "award",
        "awards and recognition", "accomplishments", "honors", "honours",
        "awards and achievements", "achievements and awards", "recognition", "publications",
        "publication", "volunteer work", "volunteering", "volunteer experience",
    ],
    "declaration": ["declaration", "references", "reference"],
    "training": [
        "training", "trainings", "training and development", "courses", "course", "workshops",
        "workshops attended", "seminars", "trainings attended", "professional training",
        "technical training", "seminars and workshops", "workshops and seminars",
    ],
}
_PHRASE_TO_SECTION = {p: cat for cat, phrases in SECTION_PHRASES.items() for p in phrases}
# Checked in order; the earliest match in the heading wins ("Skills and
# Certifications" -> skills).
_SECTION_KEYWORDS = [(re.compile(p), cat) for p, cat in [
    (r"\bprojects?\b", "projects"),
    (r"\bcertif|\blicen[cs]es?\b", "certifications"),
    (r"\bintern(ship)?s?\b|\b(industrial|in-?\s?plant|summer)\s+training\b", "internships"),
    (r"\beducat|\bacademic(s)?\b|\bscholastic\b|\bqualifications?\b", "education"),
    (r"\bexperience\b|\bemployment\b|\bwork\s+history\b|\bcareer\s+history\b|\bresponsibilit", "experience"),
    (r"\bskills?\b|\bcompetenc|\bexpertise\b|\bproficienc|\bknowledge\b|\btechnologies\b|\bstrengths?\b", "skills"),
    (r"\blanguages?\b", "languages"),
    (r"\bpersonal\b|\bbio-?\s?data\b", "personal"),
    (r"\bhobb|\binterests?\b|\bextra-?\s?curricular\b|\bco-?\s?curricular\b", "hobbies"),
    (r"\bachievement|\bawards?\b|\baccomplishment|\bhono(u)?rs\b|\bpublications?\b|\bvolunteer", "achievements"),
    (r"\bdeclaration\b|\breferences?\b", "declaration"),
    (r"\bsummary\b|\bobjectives?\b|\bprofile\b|\babout\s+me\b|\boverview\b", "summary"),
    (r"\btrainings?\b|\bcourses?\b|\bworkshops?\b|\bseminars?\b", "training"),
]]
# Inside a job/project entry these are usually sub-labels ("Tools:",
# "Responsibilities:"), not the start of a new top-level section.
_SUBLABEL_PHRASES = {
    "tools", "technologies", "responsibilities", "roles and responsibilities",
    "key responsibilities", "job responsibilities", "achievements", "key achievements",
    "accomplishments", "skills", "technical skills", "key skills", "environment",
    "technical environment", "tools used", "software used", "technologies used",
}
_ENTRY_SECTIONS = {"experience", "projects", "internships"}
# Content-bearing inline headings ("Technical Skills: Java, SQL").
_INLINE_SECTIONS = {"skills", "languages", "certifications"}
# Bullets/symbols and list numbering ("1.", "2)") before a heading, and
# colons/dashes/rules after it. "10th" keeps its digits.
_HEADING_LEAD = re.compile(r"^(?:[^\w]+|_+|\d{1,2}[.)]\s*)+")
_HEADING_TRAIL = re.compile(r"[\s:.\-–—_|*=•]+$")
_ROMAN_NUMERAL_PREFIX = re.compile(r"^(?:[ivx]{1,4}|[a-h])[.)]\s+", re.IGNORECASE)

NAME_BLOCKLIST = {
    "resume", "curriculum vitae", "curriculam vitae", "cv", "bio data",
    "biodata", "contact", "contact details", "profile", "personal details",
    "name", "full name", "candidate name", "applicant name", "your name",
    "fresher", "experienced", "resume of", "curriculum vita", "career objective",
    "organization", "organisation", "age", "sex", "dob", "nationality", "religion", "status",
    "company", "designation", "duration", "location", "place", "date", "gender",
    "contacts", "self", "pradesh", "p o", "po", "post", "dist", "district", "taluk", "tq",
    "near", "opposite", "ward", "house", "flat", "street", "mandal", "personal",
}
# Words a name never contains -- they mark a sentence or a phrase instead
# ("Support for Oracle Fusion HCM", "Area of Interest", "taking required").
_NAME_STOPWORDS = {
    "for", "the", "of", "and", "with", "in", "at", "to", "from", "by", "on", "is", "are",
    "was", "an", "as", "or", "my", "your", "our", "this", "that", "dear", "sir", "madam",
    "interest", "required", "support", "system", "systems", "driven", "result", "results",
    "operations", "experience", "seeking", "looking", "position", "application", "job",
    "summary", "objective", "career", "hobbies", "email", "mobile", "phone", "address",
    "using", "based", "process", "industry", "role", "village", "percentage", "passed", "year",
    "aggregate", "board", "qualification", "higher", "secondary", "marks", "grade", "result",
    "institute", "knowledge", "basic", "areas", "good", "learner", "overview", "exposure",
    "dept", "department", "execution", "field", "end", "section", "team", "division",
}
_NAME_PREFIX = re.compile(
    r'^(?:(?:c\.?\s?v\.?|resume|bio-?\s?data|curriculum\s+vitae)\s+of\s+|'
    r'(?:mr|mrs|ms|miss|dr|er|shri|sri|smt|kum)\.?\s+)', re.IGNORECASE)

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

_EMAIL = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)*\.[A-Za-z]{2,}')


def find_email(text):
    m = _EMAIL.search(text)
    if m:
        return m.group(0)
    # An address the PDF wrapped onto two lines ("name@gmail." / "com").
    lines = text.split("\n")
    for i, line in enumerate(lines[:-1]):
        if "@" in line or "@" in lines[i + 1]:
            m = _EMAIL.search(line.rstrip() + lines[i + 1].lstrip())
            if m:
                return m.group(0)
    return None


# Digits with spaces/dashes/brackets between groups ("+91 79037 83318",
# "(+91) 98450-12345") on ONE line -- never across a line break, which is how
# a pincode used to get glued onto the front of a number.
_PHONE_CANDIDATE = re.compile(r'(?<![\w+/])(\+?\(?\+?\d[\d \-().]{8,17}\d)(?![\w/])')


def _valid_phone_digits(raw):
    digits = re.sub(r'\D', '', raw)
    if len(digits) == 10:
        return digits[0] != '0' and not re.fullmatch(r'(19|20)\d{2}(19|20)\d{2}\d{2}', digits)
    if len(digits) == 11:
        return digits[0] == '0'
    if len(digits) == 12:
        return digits.startswith('91')
    return len(digits) <= 13 and raw.lstrip("( ").startswith('+')


def find_phone(text):
    for m in _PHONE_CANDIDATE.finditer(text):
        raw = m.group(1).strip(" -.")
        if _valid_phone_digits(raw) and not re.search(r'\d{4}\s*[-–]\s*\d{4}$', raw.replace(" ", "")[:9]):
            return re.sub(r'\s{2,}', ' ', raw)
    m = re.search(r'(?<!\d)(?:\+?91[ \-]?)?[6-9]\d{9}(?!\d)', text)
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
    # An explicit "Name: X" label is the strongest signal there is.
    labelled = [l for l in lines if NAME_LABEL_PREFIX_PATTERN.match(l[0].strip())
                and not NAME_OF_THE_PATTERN.match(l[0].strip())]
    # Then the top of the page: a bigger-font table header or heading further
    # down ("Percentage", "Higher Secondary") must not beat the name there.
    by_size = lambda ls: sorted(ls, key=lambda line: line[1], reverse=True)
    candidates = labelled + by_size(lines[:12]) + by_size(lines)
    for text, size, is_bold, font in candidates:
        clean = text.strip()
        if not clean or len(clean) < 3:
            continue
        if clean in header_set:
            continue
        if NAME_OF_THE_PATTERN.match(clean):
            continue
        clean = re.sub(r'^[^\w(]+|[\ue000-\uf8ff]', '', clean).strip()  # bullets, icon-font glyphs
        clean = NAME_LABEL_PREFIX_PATTERN.sub('', clean).strip()
        clean = _NAME_PREFIX.sub('', clean).strip()
        # "Name     +91 98...  mail@x.com" / "Name, Mobile No: ..." -- the name
        # is the first part; the rest is contact details.
        clean = re.split(r'\s{3,}|\s*\|\s*|\s*,\s*(?=.*(?:\d|@|:))', clean)[0].strip()
        clean = re.sub(r'(?<=[A-Za-z])\.$', '', clean)
        if ':' in clean:
            continue
        clean = re.split(r'\s[-|]\s', clean)[0].strip()
        clean = _collapse_letter_spacing(clean)
        if len(clean) < 3 or len(clean.split()) > 6:
            continue
        ascii_clean = unicodedata.normalize("NFKD", clean).encode("ascii", "ignore").decode()
        normalized = re.sub(r'[^a-z ]', '', ascii_clean.lower()).strip()
        if normalized in NAME_BLOCKLIST or normalized in KNOWN_SECTIONS:
            continue
        if re.search(r'[,&@/()]', clean) or not clean[:1].isupper() or len(re.sub(r'[^A-Za-z]', '', clean)) < 3:
            continue
        if any(len(w) > 20 for w in clean.split()) or _edu_level(clean) or _EDU_TABLE_HEADER.fullmatch(clean):
            continue
        if _languages_matcher and match_languages(clean, _languages_matcher):
            continue
        if heading_category(clean, styled=True, skill_vocab=_skill_heading_vocab):
            continue  # a heading, possibly letter-spaced ("P E R S O N A L")
        if set(normalized.split()) & _NAME_STOPWORDS:
            continue
        if _EDU_INSTITUTION.search(clean) or is_job_title(clean):
            continue
        if _skills_matcher and normalize_key(clean) in _skills_matcher[1]:
            continue  # a lone skill ("Hibernate") on a no-name resume
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
                # Keep only heading-like entries: the mapping also lists plain
                # skills ("Problem Solving", "Team Management"), and treating a
                # skills-list line as a heading would split the section.
                if _HEADING_LIKE_WORDS.search(h):
                    vocab.add(normalize_heading(h))
    return vocab


_HEADING_LIKE_WORDS = re.compile(
    r"\b(skills?|skillsets?|competenc\w*|expertise|proficienc\w*|knowledge|strengths?|"
    r"abilit\w*|qualities|acumen|specializations?|highlights|areas?)\b", re.IGNORECASE)


def normalize_heading(text):
    """'1. EDUCATIONAL QUALIFICATIONS :' -> 'educational qualifications'."""
    t = _HEADING_LEAD.sub("", text.strip())
    t = _ROMAN_NUMERAL_PREFIX.sub("", t)
    t = _HEADING_TRAIL.sub("", t)
    t = t.replace("&", " and ").replace("/", " and ")
    return re.sub(r"\s+", " ", t).strip().lower()


def heading_category(text, styled, skill_vocab=None):
    """(category, exact) if this whole line is a section heading, else None."""
    # "WORK EXPERIENCE: (Total 15 Years)" -- judge the heading without its note
    note = re.match(r'^(.{3,45}?)\s*[:\-–]?\s*\([^)]{1,30}\)\s*[:.]?\s*$', text.strip())
    if note:
        text = note.group(1)
    norm = normalize_heading(text)
    if not norm or len(norm) > 45 or len(norm.split()) > 6:
        return None
    if norm in _PHRASE_TO_SECTION:
        return _PHRASE_TO_SECTION[norm], True
    if skill_vocab and norm in skill_vocab:
        return "skills", True
    if not styled or len(norm.split()) > 4 or re.search(r"[\d@]", text):
        return None
    best = None
    for pattern, cat in _SECTION_KEYWORDS:
        m = pattern.search(norm)
        if m and (best is None or m.start() < best[0]):
            best = (m.start(), cat)
    return (best[1], False) if best else None


_INLINE_HEADING = re.compile(r"^\s*([A-Za-z][A-Za-z &/\-]{2,40}?)\s*[:\-–]\s+(\S.*)$")


def inline_heading(text, skill_vocab=None):
    """('skills', 'Java, SQL') for 'Technical Skills: Java, SQL', else None."""
    m = _INLINE_HEADING.match(text)
    if not m:
        return None
    found = heading_category(m.group(1), styled=True, skill_vocab=skill_vocab)
    if found and found[0] in _INLINE_SECTIONS:
        return found[0], m.group(2).strip()
    return None


def split_into_sections(lines, skill_vocab=None):
    """[(category, heading, text)] in document order (content before the first
    heading is category 'top'), plus the list of heading lines."""
    if not lines:
        return [], []
    body_size = Counter(size for _, size, _, _ in lines).most_common(1)[0][0]
    body_font = Counter(font for _, _, _, font in lines).most_common(1)[0][0]
    sections, headers = [["top", "", []]], []
    for text, size, is_bold, font in lines:
        clean = text.strip()
        current = sections[-1][0]
        big = size > body_size + 0.5 or (clean.isupper() and len(clean.split()) <= 4)
        styled = big or is_bold or font != body_font or clean.endswith(":")
        found = heading_category(clean, styled, skill_vocab)
        if found:
            category, _ = found
            is_sublabel = (current in _ENTRY_SECTIONS and not big
                           and normalize_heading(clean) in _SUBLABEL_PHRASES)
            if not is_sublabel:
                sections.append([category, clean, []])
                headers.append(clean)
                continue
        inline = inline_heading(clean, skill_vocab) if current != "skills" else None
        if inline:
            sections.append([inline[0], clean, [inline[1]]])
            sections.append([current, "", []])  # carry on with the enclosing section
            continue
        sections[-1][2].append(text)
    return [(c, h, "\n".join(ls)) for c, h, ls in sections if ls], headers


def section_text(sections, *categories):
    return "\n".join(t for c, _, t in sections if c in categories)


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


def extract_array_section(sections, category):
    return split_lines(section_text(sections, category))


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


_EDU_LEVEL_EXTRA = re.compile(
    r'\b(b\.?\s?a|m\.?\s?a|bba|bbm|b\.?\s?pharm|m\.?\s?pharm|mbbs|b\.?\s?arch|pgdm|pgdca|llb|'
    r'b\.?\s?ed|m\.?\s?ed|b\.?\s?voc|graduation|post[-\s]?graduation|high\s+school|'
    r'higher\s+secondary|secondary\s+school|s\.?\s?s\.?\s?c|h\.?\s?s\.?\s?c|madhyamik|matric)\b',
    re.IGNORECASE)
_EDU_INSTITUTION = re.compile(
    r'\b(college|university|institute|institution|polytechnic|vidyalaya|vidyalay|vidyapeeth|'
    r'mahavidyalaya|academy|school|board|cbse|icse|jntu\w*|vtu|iit|nit|iiit)\b', re.IGNORECASE)
_EDU_TABLE_HEADER = re.compile(
    r'^(s\.?\s*no|sr\.?\s*no|discipline|year|passing|year\s+of\s+passing|passing\s+year|aggregate|'
    r'percentage|marks|cgpa|gpa|grade|class|course|qualification|qualifications|degree|'
    r'examination|exam|institute|institution|board|university|school|college|specialization|'
    r'stream|result|division|duration|score|name\s+of\s+the\s+\w+|name\s+of\s+\w+|'
    r'[\s/&,%()|]|and|of|the)+$', re.IGNORECASE)
_EDU_LEAD_VERB = re.compile(r'^(completed|pursuing|passed|studying|did|done)\s+', re.IGNORECASE)
_YEAR_OR_RANGE = re.compile(
    r'\(?\b((19|20)\d{2}|(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*(19|20)?\d{2})'
    r'(\s*[-–—to]+\s*((19|20)\d{2}|present|till\s+date|(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)'
    r'[a-z]*\.?\s*(19|20)?\d{2}))?\b\)?', re.IGNORECASE)


def _edu_level(line):
    if TENTH_PATTERN.search(line):
        return 'tenth'
    if TWELFTH_PATTERN.search(line):
        return 'twelfth'
    if DEGREE_PATTERN.search(line) or _EDU_LEVEL_EXTRA.search(line):
        return 'degree'
    return None


def _clean_edu_part(text):
    text = _EDU_LEAD_VERB.sub('', text)
    text = _YEAR_OR_RANGE.sub(' ', text)
    text = re.sub(r'\b(in|from)\s*$', '', text.strip(), flags=re.IGNORECASE)
    text = re.sub(r'\(\s*\)|\[\s*\]', ' ', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return clean_entry_tail(text.strip(" .,-–—:|/"))


def extract_educations(sections):
    """One entry per qualification: the level line (degree / 12th / 10th)
    plus the institution and place lines around it, grades and years
    stripped. Layouts vary -- the institution can come before or after its
    degree line, and table layouts put every cell on its own line -- so
    lines are grouped around the level lines instead of taken one by one."""
    blocks = []
    for category, h, v in sections:
        override = None
        if TENTH_PATTERN.search(h):
            override = 'tenth'
        elif category != 'education' and TWELFTH_PATTERN.search(h):
            override = 'twelfth'
        if category == 'education' or override:
            blocks.append((override, v))

    out, seen = [], set()

    def emit(parts):
        text = ", ".join(p for p in (_clean_edu_part(x) for x in parts[:4]) if p)
        key = re.sub(r'[^a-z0-9]', '', text.lower())
        if key and key not in seen and len(text) >= 3:
            seen.add(key)
            out.append(text)

    for override, raw in blocks:
        entries, pending = [], []
        lines = [strip_grades(l) for l in split_lines(raw)]
        for line in lines:
            if not line or DISCONTINUED_PATTERN.search(line):
                continue
            if re.fullmatch(r'[\d\s./%()\-–:]+', line) or _EDU_TABLE_HEADER.fullmatch(line):
                continue
            if not _clean_edu_part(line):
                continue
            level = _edu_level(line) or (override if not entries and not pending else None)
            inst = bool(_EDU_INSTITUTION.search(line))
            if level:
                entry = {'parts': [line], 'inst': inst}
                if pending and not inst:
                    entry['parts'] += pending
                    entry['inst'] = True
                pending = []
                entries.append(entry)
            elif entries and not pending and not entries[-1]['inst']:
                entries[-1]['parts'].append(line)
                entries[-1]['inst'] = inst
            elif (entries and not pending and not inst and len(entries[-1]['parts']) < 3
                  and PLACE_PATTERN.search(line) and len(line.split()) <= 4):
                entries[-1]['parts'].append(line)  # the city after an institution
            elif inst or len(line.split()) <= 6:
                pending.append(line)
        if pending:
            if entries and not entries[-1]['inst']:
                entries[-1]['parts'] += pending
            elif any(_EDU_INSTITUTION.search(p) for p in pending):
                entries.append({'parts': pending, 'inst': True})
        for entry in entries:
            emit(entry['parts'])
    return out


_CERT_NOISE = re.compile(
    r'^(does\s+not\s+expire|credential\s+id\b|issued\b|expires?\b|valid\s+(till|until)\b|'
    r'certificate\s+no\b|id\s*[:#]|link\b|url\b)', re.IGNORECASE)
_CERT_PARTICIPATION = re.compile(r'\b(participat\w*|attended|volunteer\w*|appreciation|received|awarded|won|winner)\b',
                                 re.IGNORECASE)
_CERT_LEADING_LABEL = re.compile(r'^(?:certifications?|certificates?|courses?)\s*[:\-–]\s*', re.IGNORECASE)
# A certification named outright, for lines outside a Certifications section.
_CERT_NAMED = re.compile(
    r'^(?:certified|certification|certificate)\b|\b(?:certification|certificate)\s*$|'
    r'\b(?:green|black|yellow)\s+belt\b|\b(?:ccna|ccnp|ccie|mcsa|mcse|pmp|capm|itil|prince2|cissp|ceh|'
    r'istqb|ocjp|scjp|nptel|aws\s+certified)\b', re.IGNORECASE)
_CERT_STRONG = re.compile(
    r'\b(certified|certification|certificate\s+(in|of|course)|green\s+belt|black\s+belt|yellow\s+belt|'
    r'ccna|ccnp|ccie|mcsa|mcse|mcp|pmp|capm|itil|prince2|cissp|ceh|comptia|istqb|ocjp|ocp|scjp|'
    r'nptel|aws\s+certified|azure\s+(fundamentals|administrator)|six\s+sigma|cfa|frm|acca)\b',
    re.IGNORECASE)


def _join_wrapped(lines):
    """Re-joins a line the PDF wrapped: the continuation starts lowercase, or
    the previous line ends mid-phrase (open bracket/quote, connector word)."""
    out = []
    for line in lines:
        if out and (line[:1].islower()
                    or out[-1].count("(") > out[-1].count(")")
                    or out[-1].count("“") > out[-1].count("”")
                    or re.search(r'\b(at|by|from|of|in|and|for|with|on|the|&|-|–)$', out[-1], re.IGNORECASE)):
            out[-1] = out[-1].rstrip() + " " + line.lstrip()
        else:
            out.append(line)
    return out


def extract_certifications(sections):
    """Certificate title lines only -- drops bare issuer+year lines,
    education content that leaked in, participation notes, validity/ID
    lines, prose fragments, and strips any date from every kept line.
    With no certifications section at all, falls back to lines elsewhere
    that clearly name a certification ('Six Sigma Green Belt', 'CCNA')."""
    raw = section_text(sections, "certifications")
    candidates = [(l, False) for l in _join_wrapped(split_lines(raw))]
    other = "\n".join(t for c, _, t in sections if c not in ("certifications", "education"))
    candidates += [(l, True) for l in _join_wrapped(split_lines(other))
                   if _CERT_NAMED.search(_CERT_LEADING_LABEL.sub('', l))]
    kept = []
    for line, strict in candidates:
        line = _CERT_LEADING_LABEL.sub('', line).strip()
        if len(line) > 160 or len(line.split()) > (12 if strict else 22):
            continue
        if line[:1].islower() or _CERT_NOISE.search(line):
            continue
        if _CERT_PARTICIPATION.search(line) and not _CERT_STRONG.search(line):
            continue
        if is_education_noise(line) and not _CERT_STRONG.search(line):
            continue
        if re.fullmatch(r"[A-Za-z .&'-]{2,25}[-,]?\s*\(?\d{4}\)?", line) or re.fullmatch(r'\(?\d{4}\)?', line):
            continue
        if PLACE_PATTERN.fullmatch(line.strip()) or line.strip().lower() in ("india", "online"):
            continue
        line = re.sub(r'^(certification|certificate)s?\s+(of\s+)?(course\s+completion\s+)?(in|on|of)\s+', '', line,
                      flags=re.IGNORECASE)
        line = DATE_PATTERN.sub('', line)
        line = re.sub(r'\s{2,}', ' ', line)
        line = clean_entry_tail(line).strip(" .,-()")
        if line and len(line) >= 2:
            kept.append(line)
    return list(dict.fromkeys(kept))


# "Project Title: X" / "Name of the Project: X" (strong) and "Project 2: X" /
# "Major Project: X" (weak); group 1 says which, group 2 is the value.
_PROJECT_LABEL = re.compile(
    r'^\s*(?:(project\s*(?:title|name)|title|name\s+of\s+(?:the\s+)?project)|'
    r'(?:academic|major|minor|mini|main|final\s+year|personal|b\.?\s?tech|m\.?\s?tech|b\.?\s?e|capstone)?'
    r'\s*project(?:\s*(?:#|no\.?)?\s*\d+)?)\s*(?:name\s+is|is|[:\-–—])\s*(.*)$', re.IGNORECASE)
_PROJECT_SUBLABEL = re.compile(
    r'^\s*(description|role|roles|responsibilit\w*|technolog\w*|tech\s+stack|tools?|environment|'
    r'duration|period|team\s*size|team|client|customer|platform|software|objective|aim|abstract|'
    r'guide|link|github|company|organi[sz]ation|location|domain|modules?|front\s*end|back\s*end|'
    r'database|key\s+learnings?|outcome|result|summary|overview|contribution|designation|'
    r'languages?\s+used|os|operating\s+system|hardware|under\s+the\s+guidance)\b', re.IGNORECASE)
_PROSE_START = re.compile(
    r'^(developed|designed|created|implemented|worked|working|used|using|built|building|responsible|'
    r'involved|handled|performed|managed|prepared|analy[sz]ed|evaluated|extracted|aggregat\w+|'
    r'cleans\w*|conducted|studied|in\s+this|this|the|it|we|i|my|our|a|an|to|for|with|as|by|'
    r'successfully|actively|participated|led|leading|collaborated|coordinated|ensured|'
    r'maintained|tested|testing|wrote|writing|integrated|deployed|configured|automated|'
    r'identified|improved|reduced|achieved)\b', re.IGNORECASE)


def _unbalanced(text):
    """An opening quote or bracket with no close yet: the title wrapped."""
    # Resumes use curly quotes inconsistently (”Title” as well as “Title”),
    # so any odd count of quote marks means the title isn't closed yet.
    quotes = sum(text.count(q) for q in '"“”')
    return quotes % 2 == 1 or text.count("(") > text.count(")")


def _clean_project_title(text):
    text = text.strip().strip('"\'“”‘’*').strip()
    text = re.sub(r'\s*[:\-–—]\s*$', '', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text[:150].strip(" .,")


def _looks_like_title(line):
    words = line.split()
    if not 1 <= len(words) <= 14 or len(line) > 120:
        return False
    if line.rstrip().endswith(".") or line[:1].islower() or _PROSE_START.match(line):
        return False
    if _PROJECT_SUBLABEL.match(line) or re.match(r'^[\W\d]+$', line):
        return False
    alpha = [w for w in words if w[:1].isalpha()]
    if not alpha:
        return False
    capitalised = sum(1 for w in alpha if w[:1].isupper() or w.lower() in STOPWORDS_TITLE)
    return capitalised / len(alpha) >= 0.6 or line.rstrip().endswith(":")


STOPWORDS_TITLE = {"of", "for", "and", "in", "on", "the", "a", "an", "to", "with", "using", "by", "via", "&"}


def _project_lines(text):
    """Lines with bullets and list numbering stripped, and a value that a
    table layout put on its own line (": Tax Payment") joined to its label."""
    out = []
    for raw in text.split("\n"):
        line = raw.strip().strip(" " + BULLET_CHARS + "➢❖▪●–-*>")
        line = re.sub(r'^\d{1,2}[.)]\s+|^\d(?=[A-Z][a-z])', '', line).strip()
        if not line:
            continue
        if line.startswith(":") and out:
            out[-1] = out[-1] + " " + line
        else:
            out.append(line)
    return out


def _project_labels(lines, same_line_only=False):
    """[(strong, title)] for every labelled project in these lines."""
    found, i = [], 0
    while i < len(lines):
        m = _PROJECT_LABEL.match(lines[i])
        if not m:
            i += 1
            continue
        strong, value, j = bool(m.group(1)), m.group(2).strip(), i + 1
        if not value and not same_line_only:
            while j < len(lines) and (_PROJECT_SUBLABEL.match(lines[j]) or not re.search(r'[A-Za-z]{2}', lines[j])):
                j += 1
            value = lines[j] if j < len(lines) and _looks_like_title(lines[j]) else ""
            j += 1
        # the title wrapped: an unclosed quote/bracket, or the next line
        # carries on in lowercase
        while value and j < len(lines) and (_unbalanced(value) or lines[j][:1].islower()) \
                and not _PROJECT_LABEL.match(lines[j]) and not _PROJECT_SUBLABEL.match(lines[j]):
            value += " " + lines[j]
            j += 1
        value = _clean_project_title(re.split(r'\s+\|\s+|\s{3,}', value)[0])
        value = _YEAR_OR_RANGE.sub('', value).strip(" .,-–")
        if value and re.search(r'[A-Za-z]{2}', value) and not _PROJECT_SUBLABEL.match(value) \
                and len(value.split()) <= 20:
            found.append((strong, value))
        i = max(j, i + 1)
    return found


def _is_tech_list(line):
    parts = [x for x in re.split(r'\s*[,/]\s*', line) if x]
    return len(parts) >= 2 and sum(1 for x in parts if _skills_matcher and normalize_key(x) in _skills_matcher[1]) \
        >= len(parts) / 2


def extract_projects(sections):
    """Project titles, one per project. Explicit labels win, wherever they
    are -- 'Project Title: X' inside a job entry counts too -- and 'Project
    Title/Name' labels beat generic 'Project 2:' ones. Only with no labels at
    all does it fall back to title-looking lines in the Projects section:
    short, mostly capitalised, not a sentence, not a 'Role:' sub-label, not a
    tech list -- each followed by its description, which is skipped."""
    project_lines = _project_lines(section_text(sections, "projects"))
    other_lines = _project_lines("\n".join(t for c, _, t in sections if c not in ("projects", "education")))
    labels = _project_labels(project_lines) + _project_labels(other_lines, same_line_only=True)
    if any(strong for strong, _ in labels):
        labels = [(st, v) for st, v in labels if st]
    if labels:
        return list(dict.fromkeys(v for _, v in labels))
    titles, prev_title = [], False
    for line in project_lines:
        if _looks_like_title(line) and not _is_tech_list(line):
            if prev_title and titles and line.isupper() == titles[-1].isupper() and not titles[-1].endswith(":"):
                titles[-1] = titles[-1] + " " + line  # title wrapped across lines
            else:
                titles.append(line)
            prev_title = True
        else:
            prev_title = False
    titles = [_clean_project_title(re.split(r'\s+\|\s+|\s{3,}', t)[0]) for t in titles]
    return list(dict.fromkeys(t for t in titles if t))


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


_INTERNSHIP_LABEL = re.compile(
    r'^\s*(?:summer\s+|winter\s+|industrial\s+|virtual\s+)?internship(?:\s+program(?:me)?)?'
    r'(?:\s+project)?\s*(?:[:\-–—]|\bat\b|\bwith\b|\bin\b)\s*(.+)$', re.IGNORECASE)
_ORG_WORDS = re.compile(
    r'\b(pvt|private|ltd|limited|llp|inc|corp|corporation|company|technologies|technology|solutions|'
    r'systems|services|industries|labs?|foundation|institute|varsity|infotech|software|consultancy|'
    r'enterprises|group|bank|station|plant|organi[sz]ation)\b', re.IGNORECASE)
_INTERN_ROLE = re.compile(r'\b(intern|internship|trainee|apprentice)\b', re.IGNORECASE)


def extract_internships(sections, titles=()):
    """One entry per internship: an 'Internship: X' label anywhere, a
    company or intern-role line inside an Internships section (not the
    table labels, dates or description), or a job title that is an intern
    role ('Data Science Intern')."""
    found = []
    for line in _project_lines("\n".join(t for c, _, t in sections if c != "education")):
        m = _INTERNSHIP_LABEL.match(line)
        if m:
            value = _clean_project_title(re.split(r'\s+\|\s+|\s{3,}', m.group(1))[0])
            if re.search(r'[A-Za-z]{3}', value) and len(value.split()) <= 15:
                found.append(value)
    for line in _project_lines(section_text(sections, "internships")):
        line = re.split(r'\s{3,}|\s+\|\s+', line)[0].strip(" :-–")
        if not line or len(line.split()) > 14 or _PROSE_START.match(line) or line[:1].islower():
            continue
        if re.fullmatch(r'(company|organi[sz]ation|from|to|duration|role|period|location|date)s?', line,
                        re.IGNORECASE) or DATE_PATTERN.search(line) and len(line.split()) <= 4:
            continue
        if _ORG_WORDS.search(line) or _INTERN_ROLE.search(line) and len(line.split()) <= 8:
            found.append(_YEAR_OR_RANGE.sub('', line).strip(" ,-–"))
    found += [t for t in titles if re.search(r'\bintern\b', t, re.IGNORECASE)]
    out, seen = [], set()
    for f in found:
        key = normalize_key(f)
        if key and key not in seen and not any(key in k or k in key for k in seen):
            seen.add(key)
            out.append(f)
    return out[:6]


# ================================= Job titles =================================

# A job title ends in an occupation word: "Senior Software Engineer",
# "Procurement Head". This single rule drops the generic words the 25k-title
# list also contains ("Management", "Academic", "High School", "Lead
# Generation") while keeping real titles the list doesn't know.
_TITLE_HEAD = re.compile(
    r'^(engineers?|manager|developer|executive|officer|consultant|analyst|lead|leader|head|'
    r'trainee|intern|specialist|associate|assistant|architect|designer|supervisor|coordinator|'
    r'co-?ordinator|administrator|admin|programmer|tester|technician|accountant|director|owner|'
    r'proprietor|representative|scientist|advisor|adviser|auditor|operator|in-?charge|incharge|'
    r'foreman|planner|buyer|merchandiser|recruiter|partner|president|vp|agent|clerk|cashier|'
    r'teacher|lecturer|professor|instructor|trainer|nurse|pharmacist|chemist|surveyor|draftsman|'
    r'draughtsman|fitter|welder|electrician|machinist|mechanic|driver|controller|inspector|'
    r'estimator|founder|co-?founder|ceo|cto|cfo|coo|apprentice|fellow|researcher|editor|writer|'
    r'counsell?or|therapist|physician|attorney|advocate|secretary|receptionist|telecaller|'
    r'salesman|storekeeper|superintendent|lineman|steward|chef|captain|strategist|evangelist|'
    r'marketer|member|principal|dean|economist|statistician|actuary|'
    r'underwriter|broker|banker|investigator|translator|interpreter|photographer|animator|'
    r'artist|modeler|modeller|draughtsperson|estimator|scheduler|dispatcher|packer|picker|'
    r'handler|guard|keeper|attendant|executive|sde|swe|qa|devops|tl|pm|gm|agm|dgm|avp|svp)$',
    re.IGNORECASE)
_TITLE_LEVEL_SUFFIX = re.compile(
    r'(\s*\([^)]*\)|\s+(?:i{1,3}|iv|v|l[1-4]|[1-4]|grade\s+\w+|level\s+\w+|\(?trainee\)?))+$',
    re.IGNORECASE)
_TITLE_BAD_WORDS = re.compile(
    r'\b(with|the|to|from|by|on|is|was|were|has|have|had|will|my|our|their|this|that|these|'
    r'those|who|which|as|into|over|under|about|through|during|while|where|when)\b', re.IGNORECASE)
_TITLE_LABEL = re.compile(
    r'^\s*(?:current\s+|present\s+)?(?:designation|position|job\s+title|title|role|job\s+role|'
    r'job\s+profile|post)\s*(?:held)?\s*[:\-–—]\s*(.+)$', re.IGNORECASE)
_WORKED_AS = re.compile(
    r'\b(?:work(?:ed|ing|s)?|serv(?:ed|ing)|employed|join(?:ed)?|promoted|appointed|hired)\s+'
    r'(?:here\s+)?as\s+(?:an?\s+|the\s+)?'
    r'([A-Z][\w.&/\-]*(?:\s+(?:[A-Z(][\w.&/\-)]*|&|and|of|-)){0,6})')
_TITLE_SPLIT = re.compile(r'\s+[|–—]\s+|\s*\|\s*|\s+-\s+|,\s+|\s+at\s+|\s+@\s+|\s+in\s+(?=[A-Z])|\t|\s{3,}')
_TITLE_JOINER = re.compile(r'\s*/\s*|\s+and\s+|\s+&\s+')
_DATEISH = re.compile(r'\b(19|20)\d{2}\b|\bpresent\b|\btill\b|\bsince\b|\d{1,2}/\d{2,4}', re.IGNORECASE)


def _title_core_word(title):
    core = _TITLE_LEVEL_SUFFIX.sub('', title.strip()).strip(" -–—:,.")
    return core.split()[-1] if core.split() else ""


def is_job_title(text):
    """Short, title-cased, no sentence words, and ends in an occupation word."""
    t = text.strip(" :-–—•*.,")
    words = t.split()
    if not 1 <= len(words) <= 7 or len(t) > 70:
        return False
    if t[:1].islower() or _TITLE_BAD_WORDS.search(t) or _PROSE_START.match(t):
        return False
    if re.search(r'\d', _TITLE_LEVEL_SUFFIX.sub('', t)) or '@' in t:
        return False
    alpha = [w for w in words if w[:1].isalpha()]
    if alpha and sum(1 for w in alpha if w[:1].isupper() or w.lower() in STOPWORDS_TITLE) / len(alpha) < 0.6:
        return False
    return bool(_TITLE_HEAD.match(_title_core_word(t).strip("().,")))


def _clean_title(text):
    t = _DATEISH.split(text)[0] if _DATEISH.search(text) else text
    t = re.sub(r'\s*\(\s*\)', '', t)
    return re.sub(r'\s{2,}', ' ', t).strip(" :-–—•*.,(")


def _split_joined_titles(title):
    """'Senior Engineer / Product Owner' -> both, but only when every part is
    a title on its own ('Accounts & Internal Procurement Executive' stays)."""
    parts = [p.strip() for p in _TITLE_JOINER.split(title) if p.strip()]
    if len(parts) > 1 and all(is_job_title(p) for p in parts):
        return parts
    return [title]


def extract_titles_from_lines(text, labels_only=False):
    """Titles as written in experience/headline lines: 'Designation: X',
    'Worked as a X at Y', and short title lines ('Senior Engineer | ABC Ltd |
    2019 - Present'), even when the 25k-title list doesn't have them. With
    labels_only, just the explicit 'Designation:'/'Position:' labels."""
    found = []
    for raw in text.split("\n"):
        line = raw.strip(" \t" + BULLET_CHARS + "➢❖▪●*>-")
        if not line:
            continue
        m = _TITLE_LABEL.match(line)
        if m:
            cand = _clean_title(_TITLE_SPLIT.split(m.group(1).strip())[0])
            if is_job_title(cand):
                found.extend(_split_joined_titles(cand))
            continue
        if labels_only:
            continue
        for m in _WORKED_AS.finditer(line):
            cand = _clean_title(_TITLE_SPLIT.split(m.group(1))[0])
            if is_job_title(cand):
                found.extend(_split_joined_titles(cand))
        if len(line.split()) <= 12:
            cand = _clean_title(_TITLE_SPLIT.split(line)[0])
            if is_job_title(cand):
                found.extend(_split_joined_titles(cand))
    return found


def finalize_titles(titles):
    """Keep real titles only, and drop one contained in a longer one kept
    ('Engineer' when 'Senior Engineer' is there)."""
    kept = []
    for t in titles:
        t = t.strip()
        if t and is_job_title(t) and normalize_key(t) not in {normalize_key(k) for k in kept}:
            kept.append(t)
    keys = [(t, re.sub(r'[^a-z0-9 ]', ' ', t.lower()).split()) for t in kept]
    final = []
    for t, words in keys:
        joined = " " + " ".join(words) + " "
        if any(len(o) > len(words) and joined in " " + " ".join(o) + " " for _, o in keys):
            continue
        final.append(t)
    return final


# ================================= Industries =================================

_INDUSTRY_LABEL_LINE = re.compile(r'^\s*(?:industry|industries|domain|domains|sector|vertical)\s*[:\-–]\s*(.+)$',
                                  re.IGNORECASE)
_industry_clues = None  # (ci pattern, {key: industry}, cs pattern or None, {clue: industry})


def build_industry_matcher(mapping):
    ci, cs = {}, {}
    for industry, clues in (mapping or {}).items():
        if industry.startswith("_"):
            continue
        for clue in clues:
            (cs if clue.isupper() else ci)[clue] = industry
    ci_pattern = re.compile(r'\b(?:' + '|'.join(flexible_term_pattern(c) for c in sorted(ci, key=len, reverse=True))
                            + r')\b', re.IGNORECASE) if ci else None
    cs_pattern = re.compile(r'(?<![\w&])(?:' + '|'.join(re.escape(c) for c in sorted(cs, key=len, reverse=True))
                            + r')(?![\w&])') if cs else None
    return ci_pattern, {normalize_key(k): v for k, v in ci.items()}, cs_pattern, cs


def find_industries(text, label_text="", max_items=5):
    """Sectors from clues in job/project text -- a sector word ('thermal power
    plant', 'banking domain') or a well-known employer ('Infosys', 'Tata
    Motors') -- plus explicit 'Industry:'/'Domain:' lines, where the longer
    industries.json names count too."""
    found = []
    labels = "\n".join(m.group(1) for m in (_INDUSTRY_LABEL_LINE.match(l) for l in label_text.split("\n")) if m)
    if _industry_clues:
        ci_pattern, ci_map, cs_pattern, cs_map = _industry_clues
        hits = []
        for src in (labels, text):
            if ci_pattern:
                hits += [(m.start(), ci_map.get(normalize_key(m.group(0)))) for m in ci_pattern.finditer(src)]
            if cs_pattern:
                hits += [(m.start(), cs_map.get(m.group(0))) for m in cs_pattern.finditer(src)]
            hits.sort()
            found += [h for _, h in hits if h]
            hits = []
    if labels:
        found += match_exact(labels, _industries_matcher)
    return list(dict.fromkeys(found))[:max_items]


# =================================== Skills ===================================

_SKILL_LABEL_LINE = re.compile(
    r'^\s*(?:tech(?:nical)?\s+stack|technolog(?:y|ies)(?:\s+used)?|tools?(?:\s+used)?|'
    r'(?:technical\s+)?environment|software(?:\s+used)?|platforms?|frameworks?|languages?\s+used|'
    r'databases?|skills?\s+used|key\s+skills|tools\s+(?:and|&)\s+technologies|'
    r'technologies\s+(?:and|&)\s+tools|front[\s-]?end|back[\s-]?end)\s*[:\-–]\s*(.+)$', re.IGNORECASE)
_SKILL_PREFIX = re.compile(
    r'^(?:(?:good|basic|strong|sound|excellent|working|hands[\s-]?on|in[\s-]?depth|advanced|fair|'
    r'thorough|solid|practical)\s+)?(?:knowledge|understanding|experience|exposure|expertise|'
    r'proficiency|command|familiarity)\s+(?:of|in|on|with)\s+|'
    r'^(?:proficient|skilled|experienced|expert|familiar|good|well[\s-]versed|versed|comfortable)'
    r'\s+(?:in|with|at|on)\s+', re.IGNORECASE)
_SKILL_PERSONAL = re.compile(
    r'\b(birth|marital|father|mother|nationality|gender|religion|passport|address|'
    r'dob|d\.o\.b|hobbies|declare)\b|@|https?:|www\.', re.IGNORECASE)


def skill_label_values(text):
    """'Tech Stack: Java, SQL' / 'Environment: ...' lines inside jobs and projects."""
    return [m.group(1) for m in (_SKILL_LABEL_LINE.match(l) for l in text.split("\n")) if m]


def _is_skill_like(part):
    key = normalize_key(part)
    if key in _skill_whole_items:
        return True
    words = part.split()
    if not 1 <= len(words) <= 5 or not 2 <= len(part) <= 45 or not re.search(r'[A-Za-z]{2}', part):
        return False
    if key in _skill_block_keys or _SKILL_PERSONAL.search(part) or _PROSE_START.match(part):
        return False
    if is_junk_candidate(part, _languages_matcher) or is_education_noise(part):
        return False
    if len(words) > 1 and part == part.lower() and key not in _skills_matcher[1]:
        return False  # lowercase multi-word fragments are prose, not list items
    if len(words) > 1 and words[0].lower().endswith("ing") and key not in _skills_matcher[1]:
        return False  # "Creating Forms with Reports" -- an activity, not a skill
    if ":" in part or ". " in part:
        return False
    return True


def _is_context_skill(canonical, surface):
    """A known skill found in job/project prose counts only when written the
    way the list writes it ('Quality Control', not 'quality control'), and a
    one-word one only if it's well-known tech or tech-spelled ('JavaScript',
    'C#', 'SQL') -- prose is full of capitalised product and company words
    ('Tata Motors', 'ACC Cement') that are also in skills.json."""
    if normalize_key(canonical) in _skill_whole_items:
        return False
    if surface.isupper() and len(surface) >= 2 and surface == canonical.upper():
        return True
    if surface != canonical:
        return False
    if len(canonical.split()) > 1:
        return True
    return (normalize_key(canonical) in _skill_known_tech
            or re.search(r'[A-Z].*[A-Z]|\d|[#+.]', canonical[1:] if canonical else "") is not None)


def raw_skill_items(text):
    """Items of a skills list as written ('Tally ERP 9', 'Desktop Support'),
    including ones skills.json doesn't know: prefixes like 'Good knowledge
    of' come off, 'MS Office (Word, Excel)' opens up, 'HTML/CSS' splits
    unless the whole thing is a known term ('PL/SQL', 'TCP/IP')."""
    canonical_map = _skills_matcher[1] if _skills_matcher else {}
    items = []
    for line in text.split("\n"):
        line = line.strip(" \t" + BULLET_CHARS + "➢❖▪●✓✔*>-–")
        line = re.sub(r'^[A-Za-z][A-Za-z &/()\-]{1,35}:\s*', '', line)
        if not line or len(line.split()) > 25:
            continue
        line = re.sub(r'\(([^()]*,[^()]*)\)', r', \1', line)
        for part in re.split(r'\s*[,;|•·]\s*|\s+(?:&|and)\s+|\s+[-–]\s+', line):
            part = _SKILL_PREFIX.sub('', part.strip(" .:\"'")).strip(" .:\"'")
            part = re.sub(r'^\((.*)\)$', r'\1', part).strip()
            if part.count("(") != part.count(")"):
                part = part.replace("(", " ").replace(")", " ").strip()
            subparts = [part]
            if "/" in part and normalize_key(part) not in canonical_map:
                subparts = [x.strip() for x in part.split("/")]
            items.extend(re.sub(r'\s{2,}', ' ', x) for x in subparts if x)
    return items


def canonical_skill(part):
    key = normalize_key(part)
    if _skills_matcher and key in _skills_matcher[1]:
        return _skills_matcher[1][key]
    if key in _skill_whole_items:
        return _skill_whole_items[key]
    if _skill_acronyms and part in _skill_acronyms[1]:
        return _skill_acronyms[1][part]
    return part


def match_skills_with_surface(text):
    if not _skills_matcher:
        return []
    pattern, canonical_map = _skills_matcher
    return [(canonical_map.get(normalize_key(m.group(0)), m.group(0)), m.group(0)) for m in pattern.finditer(text)]


def collect_skills(primary, context, exclude):
    """primary: skills sections + 'Tech Stack:'-style lines -- every known
    skill there counts, and so do list items that merely look like skills.
    context: jobs/projects/summary prose -- only known skills written the way
    the list writes them ('Java', not 'software'), and never generic words."""
    exclude_keys = {normalize_key(e) for e in (exclude or [])} | _skill_block_keys
    found, seen = [], set()

    def add(term):
        key = normalize_key(term)
        if key and key not in seen and key not in exclude_keys:
            seen.add(key)
            found.append(term)

    for canonical, _ in match_skills_with_surface(primary):
        add(canonical)
    for part in raw_skill_items(primary):
        if _is_skill_like(part):
            add(canonical_skill(part))
    for canonical, surface in match_skills_with_surface(context):
        if _is_context_skill(canonical, surface):
            add(canonical)
    return found


# ======================== Per-resume parsing (worker side) ========================

_skills_matcher = _titles_matcher = _industries_matcher = _languages_matcher = None
_skill_heading_vocab = None
_skill_acronyms = None          # (case-sensitive pattern, {acronym: canonical})
_skill_block_keys = set()       # normalize_key of generic non-skills
_skill_whole_items = {}         # normalize_key -> canonical, for ambiguous words
_skill_known_tech = set()       # normalize_key of skills_extra.json 'add' (curated, common)


def build_skill_matchers(skills, extra):
    """The main skills matcher over skills.json + skills_extra.json 'add',
    plus the long names' short forms: 'Accounts Payable (AP)' is also found
    as 'Accounts Payable', and a >=3-letter acronym ('AWS', 'GST') is found
    case-sensitively -- lowercase 'ace' in prose never becomes '(ACE)'."""
    terms = list(skills or []) + list(extra.get("add", []))
    known = {normalize_key(t) for t in terms}
    aliases, acronyms = {}, {}
    for s in skills or []:
        m = re.fullmatch(r'(.+?)\s*\((.+)\)', s)
        if not m:
            continue
        base, inner = m.group(1).strip(), m.group(2).strip()
        if re.fullmatch(r'[A-Z0-9][A-Za-z0-9.&/+\- ]{1,11}', inner) and sum(c.isupper() for c in inner) >= 2:
            if len(inner.replace(" ", "")) >= 3 and normalize_key(inner) not in known:
                acronyms.setdefault(inner, s)
            if len(base.split()) >= 2 and normalize_key(base) not in known:
                aliases.setdefault(base, s)
    pattern, canonical_map = build_matcher(terms + list(aliases))
    for alias, canonical in aliases.items():
        canonical_map[normalize_key(alias)] = canonical
    acr = None
    if acronyms:
        acr = (re.compile(r'(?<![\w.])(?:' + '|'.join(re.escape(a) for a in sorted(acronyms, key=len, reverse=True))
                          + r')(?![\w])'), acronyms)
    return (pattern, canonical_map), acr


def _init_worker(lookups_dir):
    global _skills_matcher, _titles_matcher, _industries_matcher, _languages_matcher, _skill_heading_vocab
    global _skill_acronyms, _skill_block_keys, _skill_whole_items, _skill_known_tech
    skills = load_canonical_list("skills.json", lookups_dir)
    titles = load_canonical_list("titles.json", lookups_dir)
    industries = load_canonical_list("industries.json", lookups_dir)
    languages = load_canonical_list("languages.json", lookups_dir)
    extra = load_canonical_list("skills_extra.json", lookups_dir) or {}
    _skills_matcher, _skill_acronyms = build_skill_matchers(skills, extra) if (skills or extra) else (None, None)
    _skill_block_keys = {normalize_key(b) for b in extra.get("block", [])}
    _skill_whole_items = {normalize_key(w): w for w in extra.get("whole_item_only", [])}
    _skill_known_tech = {normalize_key(w) for w in extra.get("add", [])}
    global _industry_clues
    _industry_clues = build_industry_matcher(load_canonical_list("industry_keywords.json", lookups_dir))
    _titles_matcher = build_matcher(titles) if titles else None
    _industries_matcher = build_matcher(industries) if industries else None
    _languages_matcher = build_language_matcher(languages) if languages else None
    _skill_heading_vocab = load_skill_heading_vocab(lookups_dir)


def parse_resume(path, mime_type):
    lines = resolve_lines_getter(mime_type)(path)
    full_text = "\n".join(t for t, _, _, _ in lines)

    sections, headers = split_into_sections(lines, _skill_heading_vocab)

    email = find_email(full_text)
    phone = find_phone(full_text)
    linkedin = find_linkedin(full_text)
    name = find_name(lines, headers, _titles_matcher, email)

    def mask(text):
        for value in (email, phone, linkedin):
            if value:
                text = text.replace(value, " ")
        return re.sub(r'\b(?:www\.|https?://)\S+', ' ', text, flags=re.IGNORECASE)

    raw_skill_section = section_text(sections, "skills")
    raw_title_section = section_text(sections, "experience", "internships")

    # Personal Details (marital status, religion, nationality, gender) is
    # never a real skill/title/industry -- strip it wherever it could leak
    # in: the whole-resume fallback used when there's no dedicated Skills/
    # Experience section, AND the section text itself, in case it merged
    # into the wrong heading's content. Everything else (email/phone/name/
    # language detection) still scans the true full_text, unaffected.
    personal_details_text = section_text(sections, "personal", "declaration")
    fallback_text = full_text
    if personal_details_text:
        fallback_text = full_text.replace(personal_details_text, " ")
        raw_skill_section = raw_skill_section.replace(personal_details_text, " ")
        raw_title_section = raw_title_section.replace(personal_details_text, " ")

    industry_text = mask(raw_title_section)  # for the unmatched-industry candidates below

    entry_text = section_text(sections, "experience", "projects", "internships", "summary", "training", "top")
    skill_primary = mask(raw_skill_section + "\n" + "\n".join(skill_label_values(entry_text)))
    skill_context = mask(entry_text)
    if not skill_primary.strip() and not skill_context.strip():
        skill_context = mask(fallback_text)
    matched_skills = _filter_junk_terms(_strip_languages(
        collect_skills(skill_primary, skill_context, headers), _languages_matcher))

    # Titles come from where people state them -- the headline under the
    # name, the summary, and experience/internship entries -- never from
    # skills, education or hobbies. Single-word list hits ("Management",
    # "Academic") are dropped; a real one-word title still comes through as
    # its own line via extract_titles_from_lines. With no experience heading
    # found (two-column layouts scramble it), entries can sit in any
    # section, so look everywhere but the ones that never hold job titles.
    has_experience = any(c in ("experience", "internships") for c, _, _ in sections)
    title_categories = (("top", "summary", "experience", "internships") if has_experience else
                        tuple({c for c, _, _ in sections} - {"education", "projects", "certifications",
                                                             "skills", "declaration"}))
    title_source = mask(section_text(sections, *title_categories))
    label_source = mask("\n".join(t for c, _, t in sections if c != "projects"))
    list_titles = [t for t in _filter_junk_terms(_strip_languages(
        match_exact(title_source, _titles_matcher, exclude=headers), _languages_matcher))
        if len(t.split()) > 1]
    matched_titles = finalize_titles(extract_titles_from_lines(title_source)
                                     + extract_titles_from_lines(label_source, labels_only=True)
                                     + list_titles)
    industry_source = mask(section_text(sections, "experience", "internships", "summary", "top"))
    matched_industries = find_industries(industry_source, mask(full_text))
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

    educations = extract_educations(sections)
    certifications = extract_certifications(sections)
    return {
        "name": name,
        "email": email,
        "linkedin": linkedin,
        "phone": phone,
        "skills": as_csv(matched_skills),
        "titles": as_csv(matched_titles),
        "industries": as_csv(matched_industries),
        "languages": as_csv(languages),
        "skills_list": matched_skills,
        "titles_list": matched_titles,
        "industries_list": matched_industries,
        "languages_list": languages,
        # Entries contain commas themselves ("B.E, XYZ College, Pune"), so the
        # display string separates them with "; " and the list keeps them whole.
        "educations": "; ".join(educations),
        "certifications": "; ".join(certifications),
        "educations_list": educations,
        "certifications_list": certifications,
        "projects": extract_projects(sections),
        "internships": extract_internships(sections, matched_titles),
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
        try:
            batch = index.get_documents({"offset": offset, "limit": limit, "fields": ["id"]}).results
        except MeilisearchApiError as e:
            if e.code == "index_not_found":  # first run: add_documents creates it later
                return existing
            raise
        if not batch:
            break
        existing.update(d.id for d in batch)
        offset += limit
    return existing


def is_note(doc):
    """Bookkeeping documents (duplicate / unsupported / error), not resumes."""
    return "error" in doc or "duplicate_of" in doc or bool(doc.get("skipped"))


def ensure_filterable_attributes(client, index, wanted):
    """Adds any missing `wanted` attributes without dropping ones already set
    (by an earlier run, or by hand), so a re-run never undoes settings."""
    try:
        current = index.get_filterable_attributes() or []
    except MeilisearchApiError as e:
        if e.code != "index_not_found":
            raise
        current = []
    current_names = {a for a in current if isinstance(a, str)}
    missing = [a for a in wanted if a not in current_names]
    if missing:
        task = index.update_filterable_attributes(list(current) + missing)
        client.wait_for_task(task.task_uid, timeout_in_ms=300_000)


# ================================ Driver ================================

def main(folder_name, workers, lookups_dir, hash_dir, meili_url, meili_index, meili_api_key,
         notes_index_name=None):
    client = meilisearch.Client(meili_url, meili_api_key)
    index = client.index(meili_index)
    notes_index_name = notes_index_name or meili_index + NOTES_INDEX_SUFFIX
    notes_index = client.index(notes_index_name)
    try:
        existing_ids = fetch_existing_ids(index) | fetch_existing_ids(notes_index)
    except Exception as e:
        print(f"Could not reach MeiliSearch at {meili_url} ({e}). Is the server running?")
        return
    print(f"{len(existing_ids)} documents already in '{meili_index}' + '{notes_index_name}'.")

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

    resumes = [d for d in documents if not is_note(d)]
    notes = [d for d in documents if is_note(d)]
    if resumes:
        ensure_filterable_attributes(client, index, RESUME_FILTERABLE_ATTRIBUTES)
    batch_size = 500
    for target, name, docs in ((index, meili_index, resumes), (notes_index, notes_index_name, notes)):
        for i in range(0, len(docs), batch_size):
            batch = docs[i:i + batch_size]
            task = target.add_documents(batch, primary_key="id")
            client.wait_for_task(task.task_uid, timeout_in_ms=120_000)
            print(f"  uploaded {min(i + batch_size, len(docs))}/{len(docs)} to '{name}'")

    elapsed = time.time() - start
    try:
        total = index.get_stats().number_of_documents
    except MeilisearchApiError:
        total = 0
    print(f"Total: {len(resumes)} resumes -> '{meili_index}' ({total} there now), "
          f"{len(notes)} notes -> '{notes_index_name}', in {elapsed / 60:.1f} min, "
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
    arg_parser.add_argument("--meili-notes-index", default=None,
                             help=f"Index for duplicate/unsupported/error notes "
                                  f"(default: <meili-index>{NOTES_INDEX_SUFFIX})")
    args = arg_parser.parse_args()
    main(args.folder_name, args.workers, args.lookups_dir, args.hash_dir,
         args.meili_url, args.meili_index, args.meili_api_key, args.meili_notes_index)
