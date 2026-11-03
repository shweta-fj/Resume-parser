"""
parse_resumes.py  --  PART 1 of 2: READ resumes, WRITE JSON files.

This script only parses. It never talks to MeiliSearch. The second script,
upload_to_meilisearch.py, uploads what this one writes.

WHAT IT DOES, STEP BY STEP
  1. Finds every file in the input folder: subfolders too, and the inside of .zip files.
  2. Sorts the files into groups:
       - unsupported : not a resume file (CSV, web page, video, ...)
       - duplicate   : exactly the same content as an earlier file
       - to parse    : everything else
  3. Turns old .doc / .rtf / .odt files into .docx (needs LibreOffice).
  4. Reads each remaining file and pulls out name, email, phone, LinkedIn, skills,
     titles, industries, languages, education, certifications, projects, internships.
  5. Decides if the file is really a resume (see "check_if_real_resume" below).
  6. Saves everything under  output/<name>/

WHAT YOU GET
  output/<name>/real_resumes/<id>.json   one JSON file per REAL resume. Only real resumes.
  output/<name>/rejected/errors.json         files that could not be read
  output/<name>/rejected/invalid.json        read fine, but do not look like a resume
  output/<name>/rejected/unsupported.json    not a resume file type
  output/<name>/rejected/duplicates.json     same content as an earlier file
  output/<name>/file_list.csv            every file and what happened to it
  output/<name>/review/needs_review.json   real resumes worth a quick human check (the list only; nothing is removed)

WHAT IS NEW IN VERSION 2 (accuracy fixes, all checked in test_rules.py)
  - Names: a line right after an address line is not a name; "Thanking you" and other
    letter endings are not names; a first name and surname on two lines are joined.
  - Not a resume: cover letters, certificates, statements, book pages and very long
    documents with no contact details are rejected as "invalid" (with the reason).
  - Titles: an organisation such as "Bhad Engineers" is no longer taken as a job title.
  - Education: degrees written as sentences ("Completed Diploma in ... from ...") are found.
  - Certifications: junk such as "year", "10th", "Udemy" or a lone word is dropped.
  - LinkedIn: an address cut across two lines is joined.
  - Resumes saved as web pages (.htm, or HTML under a .doc name) are now read.
  - More industry clues, more heading names for projects and certifications.
  - Every JSON records "parser_version", and review/needs_review.json lists doubtful resumes.

HOW TO RUN
  python parse_resumes.py  path/to/resumes  --name my_resumes

Running it again is safe: resumes that already have a JSON file are skipped
(use --redo to parse everything again).

HOW THE CODE IS ORGANISED
  The big middle section ("THE READING ENGINE") holds the rules for finding names,
  skills and so on. It is copied unchanged from the earlier script so the results do
  not change. Everything else (this header and the SIMPLE PART at the bottom) is
  plain code: find files, sort them, parse them, save the JSON.
"""

import os
import re
import csv
import glob
import json
import time
import shutil
import hashlib
import zipfile
import zlib
import tempfile
import argparse
import subprocess
import unicodedata
import pymupdf
try:
    import re2  # optional: a faster regex engine, same results
except ImportError:
    re2 = None
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures.process import BrokenProcessPool

# ---- settings used by the reading engine ----
DEFAULT_LOOKUPS_DIR = "lookups"
NORMALIZED_CATEGORY_MAPPING_FILE = "normalized_category_mapping.json"
# The sections a resume can have. "has all sections" means all seven are filled.
CORE_SECTIONS = ("skills", "titles", "industries", "educations", "certifications", "projects", "internships")
PARSER_VERSION = "2.0"   # saved inside every resume JSON, so you can tell which rules produced it

# ---- file types ----
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DOC_MIME = "application/msword"
RTF_MIME = "application/rtf"
ODT_MIME = "application/vnd.oasis.opendocument.text"
PDF_MIME = "application/pdf"
HTML_MIME = "text/html"
IMAGE_MIMES = {"image/jpeg", "image/png", "image/bmp", "image/tiff"}
CONVERT_MIMES = {DOC_MIME, RTF_MIME, ODT_MIME, HTML_MIME}   # LibreOffice turns these into .docx first
CSV_MIME = "text/x-candidate-csv"   # only candidate lists (see detect_candidate_csv), never other CSV files
SUPPORTED_MIMES = {PDF_MIME, DOCX_MIME} | CONVERT_MIMES | IMAGE_MIMES | {CSV_MIME}


# =====================================================================
#                        THE READING ENGINE
#   (unchanged from the earlier script: file types, headings, names,
#    skills, titles, industries, education, certifications, projects)
# =====================================================================

def detect_mime_type(path):
    """Identifies a file's real type from its content ('magic bytes'), not
    its filename extension, which can lie."""
    try:
        with open(path, "rb") as f:
            beginning = f.read(512)
    except OSError:
        return None
    header = beginning[:8]
    if header.startswith(b"%PDF-"):
        return PDF_MIME
    if header.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        return DOC_MIME
    if header.startswith(b"{\\rtf"):
        return RTF_MIME
    if header.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
                if "word/document.xml" in names:
                    return DOCX_MIME
                if "mimetype" in names and z.read("mimetype").strip() == ODT_MIME.encode():
                    return ODT_MIME
        except (zipfile.BadZipFile, OSError, KeyError):
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
    # A web page (some job sites save resumes as HTML, even under a .doc name).
    start_of_text = beginning.lstrip(b"\xef\xbb\xbf \r\n\t").lower()
    if start_of_text.startswith((b"<!doctype html", b"<html", b"<head", b"<meta")):
        return HTML_MIME
    if detect_candidate_csv(beginning):
        return CSV_MIME
    return None


def detect_candidate_csv(beginning):
    """True for a CSV/TSV export that lists candidates, one per row, recognised by its header line:
    an ATS pipeline export (name, resume_match, ...), a WorkIndia export (Sr. No., Full Name, Mobile No., ...)
    or a LinkedIn Recruiter export (Name, Current Title, ..., LinkedIn URL). Company lists, statistics and
    bank statements are not."""
    first_line = beginning.decode("utf-8-sig", "ignore").split("\n", 1)[0].lower()
    if "name" not in first_line:
        return False
    return ("resume_match" in first_line and "pipeline" in first_line
            or "full name" in first_line and "mobile no" in first_line
            or "current title" in first_line and "linkedin url" in first_line)


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
        "selected projects", "relevant projects", "software projects", "project highlights",
        "work samples", "portfolio", "projects portfolio", "key project", "project handled",
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
        "courses completed", "courses attended", "online courses", "professional courses",
        "licenses", "licences", "accreditations", "credentials", "certifications and licenses",
        "certificates and courses", "certification details", "certificate details",
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
    # closing lines of a cover letter ("Thanking you", "Yours faithfully") and similar phrases
    "thanking", "thank", "thanks", "you", "yours", "faithfully", "sincerely", "regards",
    "signature", "respected", "hereby", "declare", "enclosed", "attached", "kindly",
}
# A line like these just before a candidate means the candidate is part of an address
# ("Gayatri Nagar," then "Vizianagaram"), not a name.
_ADDRESS_LINE = re.compile(
    r"\b(address|s/o|d/o|w/o|c/o|house\s*no|door\s*no|flat|street|road|nagar|colony|layout|village|"
    r"taluk|tq|dist|district|post|pin(?:code)?|near|opp(?:osite)?|apartment|cross|sector|plot)\b",
    re.IGNORECASE)
_BARE_ADDRESS_LABEL = re.compile(
    r"^(?:permanent\s+|current\s+|present\s+|postal\s+|home\s+|office\s+|mailing\s+)?address\s*:?$", re.IGNORECASE)
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
    try:
        pages = list(convert_from_path(path, dpi=200, poppler_path=poppler_path))
    except Exception:
        pages = _render_pages_with_pymupdf(path)   # poppler cannot open a damaged PDF, pymupdf can repair it
    for img in pages:
        for raw_line in pytesseract.image_to_string(img).split("\n"):
            if raw_line.strip():
                lines.append((raw_line.strip(), 0.0, False, "ocr"))
    if sum(len(t) for t, *_ in lines) < 100:
        lines = _ocr_embedded_pictures(path) or lines
    return lines


def _render_pages_with_pymupdf(path, dpi=200, max_pages=12):
    """The pages of a PDF as pictures, using pymupdf (it opens PDFs with a broken structure)."""
    import io
    from PIL import Image
    pages = []
    pdf = pymupdf.open(path)
    try:
        for number, page in enumerate(pdf):
            if number >= max_pages:
                break
            pages.append(Image.open(io.BytesIO(page.get_pixmap(dpi=dpi).tobytes("png"))))
    finally:
        pdf.close()
    return pages


def _ocr_embedded_pictures(path, max_pages=4):
    """Some scanned PDFs draw a page that renders blank while the scan itself is stored inside as a
    picture. Read those pictures directly (greyscale, at most 3000 pixels wide)."""
    import io
    from PIL import Image
    import pytesseract
    Image.MAX_IMAGE_PIXELS = None
    _configure_ocr_tools()
    lines = []
    try:
        pdf = pymupdf.open(path)
    except Exception:
        return lines
    try:
        for page_number, page in enumerate(pdf):
            if page_number >= max_pages:
                break
            for info in page.get_images(full=True):
                try:
                    picture = Image.open(io.BytesIO(pdf.extract_image(info[0])["image"])).convert("L")
                except Exception:
                    continue
                if min(picture.size) < 200:
                    continue
                if picture.size[0] > 3000:
                    picture = picture.resize((3000, int(picture.size[1] * 3000 / picture.size[0])))
                for raw_line in pytesseract.image_to_string(picture).split("\n"):
                    if raw_line.strip():
                        lines.append((raw_line.strip(), 0.0, False, "ocr"))
    finally:
        pdf.close()
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


TEXT_MIME = "text/plain"


def get_lines_from_text(path):
    """Plain text (txt, md, code, spreadsheets turned into text, ...): every line, no style."""
    raw = open(path, "rb").read()
    for encoding in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    return [(line.strip(), 0.0, False, "text") for line in text.splitlines() if line.strip()]


def resolve_lines_getter(mime_type):
    if mime_type == TEXT_MIME:
        return get_lines_from_text
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
    """'linkedin.com/in/some-name'. The address is sometimes cut by a line break
    ('linkedin.com/in/some-name-' then '12345' on the next line); the two pieces
    are joined, but only when the first piece is clearly unfinished."""
    m = re.search(r'linkedin\.com/in/([\w%-]*)', text, re.IGNORECASE)
    if not m:
        return None
    slug = m.group(1)
    next_line = re.match(r'[ \t]*\n[ \t]*([\w%-]{3,})(?:[/?#]\S*)?[ \t]*(?:\n|$)', text[m.end():])
    if next_line and (not slug or slug.endswith('-')):
        slug += next_line.group(1)
    return f"linkedin.com/in/{slug}" if len(slug) >= 3 else None


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
    labelled_ids = {id(entry) for entry in labelled}
    sizes_on_top = sorted(entry[1] for entry in lines[:12])
    typical_size = sizes_on_top[len(sizes_on_top) // 2] if sizes_on_top else 0
    email_letters = re.sub(r"[^a-z]", "", (email or "").split("@")[0].lower())
    line_index = {}   # text of a line -> where it is on the page
    for i, entry in enumerate(lines):
        line_index.setdefault(entry[0].strip(), i)
    for entry in candidates:
        text, size, is_bold, font = entry
        clean = text.strip()
        if not clean or len(clean) < 3:
            continue
        if clean in header_set:
            continue
        # a line right after an address line is part of the address ("Gayatri Nagar," / "Vizianagaram").
        # Not when that line is just the label "Address" (many templates print the label before
        # the name), and never for the biggest text on the page, which is the name.
        position = line_index.get(clean)
        if (position is not None and position > 0 and id(entry) not in labelled_ids
                and size < typical_size * 1.3
                and _ADDRESS_LINE.search(lines[position - 1][0])
                and re.search(r"[\d,]|\b(?:address|s/o|d/o|w/o|c/o)\b", lines[position - 1][0], re.IGNORECASE)
                and not _BARE_ADDRESS_LABEL.match(lines[position - 1][0].strip())):
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
        # a short ALL-CAPS word like "JLT" or "TCS" is a company, unless the email says otherwise
        if len(clean.split()) == 1 and clean.isupper() and len(clean) <= 4 and clean.lower() not in email_letters:
            continue
        # deep in the page a line is often a place, skill or sentence. There the email must agree:
        # a real name shares a word with it. (Without an email nothing can be checked, so it stays.)
        if email_letters and id(entry) not in labelled_ids and line_index.get(text.strip(), 0) >= 12:
            words = [re.sub(r"[^a-z]", "", w.lower()) for w in clean.split()]
            if not any(len(w) >= 3 and (w in email_letters or (len(w) >= 4 and w[:4] in email_letters)) for w in words):
                continue
        result = _strip_trailing_title(clean, titles_matcher)
        if len(result.split()) == 1:
            result = _join_split_name(result, line_index.get(text.strip()), size, lines)
        return result
    return _name_from_email(email)


_ONE_WORD = re.compile(r"^(?:[A-Z][A-Za-z'’.-]{2,19}|[A-Z]{3,20})$")


def _join_split_name(first_word, position, size, lines):
    """Some resumes put the first name and the surname on two lines ('Mufaddal' then
    'Tandawala'). If the line next to a one-word name is also one plain capitalised
    word, in the same font size and nothing like a place, skill, title or heading,
    the two together are the name."""
    if position is None or position > 15:
        return first_word
    for neighbour_position, word_goes_first in ((position + 1, False), (position - 1, True)):
        if not 0 <= neighbour_position < len(lines):
            continue
        text, neighbour_size, _, _ = lines[neighbour_position]
        word = text.strip()
        if not _ONE_WORD.match(word) or abs(neighbour_size - size) > 0.6:
            continue
        if word.isupper() != first_word.isupper():
            continue   # 'Karan' next to 'KAUSHIK' is probably not one name
        normalized = word.lower()
        if (normalized in NAME_BLOCKLIST or normalized in KNOWN_SECTIONS or normalized in _NAME_STOPWORDS
                or normalized in _NAME_STRUCT_WORDS or PLACE_PATTERN.search(word) or junk_level(word) > 0
                or is_job_title(word) or _edu_level(word)
                or (_skills_matcher and normalize_key(word) in _skills_matcher[1])
                or (_languages_matcher and match_languages(word, _languages_matcher))):
            continue
        return f"{word} {first_word}" if word_goes_first else f"{first_word} {word}"
    return first_word


# ------------------------- Name verification and recovery -------------------------
# find_name() guesses from layout (the biggest text near the top). When the real
# name fails one of its filters it moves on to the next line, which is often a
# heading or phrase ("Class XII", "Driving Car"). refine_name() double-checks the
# result. Precision first: a name is only replaced when it looks like junk
# (is_junk_name) AND a better candidate turns up -- a name that merely differs
# from the email address is left alone, because plenty of real names do.

_term_words = set()  # every word that appears in the skills / titles / industries lists
_NAME_STRUCT_WORDS = {
    "class", "skype", "resume", "curriculum", "vitae", "declaration", "details", "highlights",
    "married", "unmarried", "subject", "subjects", "description", "responsibilities",
    "formatting", "scenarios", "miscellaneous", "india", "frontend", "backend", "fullstack",
    "employee", "technical", "experience", "education", "profile", "contact", "address",
    "email", "mobile", "phone", "skills", "skill", "projects", "project", "certifications",
    "personal", "objective", "summary", "technology", "technologies", "development",
    "developer", "engineer", "engineering", "services", "company", "limited", "pvt",
}
_NAME_LEAD_LABEL = re.compile(
    r'^\s*(?:skype(?:\s*id)?|name|full\s*name|candidate(?:\s*name)?)\b[\s.;:,\-–]*', re.IGNORECASE)
_FILENAME_NOISE = re.compile(
    r'^(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-?|naukri[_\- ]?|iimjobs[_\- ]?|'
    r'updazz[_\- ]?|copy\s+of\s+|resume[_\- ]?|cv[_\- ]?)+', re.IGNORECASE)
_FILENAME_DROP = {"resume", "cv", "profile", "final", "updated", "new", "docx", "pdf", "doc", "copy",
                  "latest", "curriculum", "vitae", "format", "of"}


def _build_term_words(*term_lists):
    words = set()
    for terms in term_lists:
        for t in terms or []:
            words.update(re.findall(r'[a-z]{3,}', str(t).lower()))
    return words


def _alpha_tokens(text, min_len=3):
    return [t for t in re.findall(r'[a-z]+', text.lower()) if len(t) >= min_len]


def junk_level(name):
    """0 = looks like a person; 1 = might not be (half or more of its words
    appear in the skills/titles/industries lists, but real names like 'Jain' or
    'Das' do too); 2 = reads like a heading, phrase or ID."""
    if not name:
        return 0
    if re.search(r'[;:@/|]', name) or any(re.search(r'[A-Za-z]', w) and re.search(r'\d', w) for w in name.split()):
        return 2
    tokens = _alpha_tokens(name)
    if not tokens:
        return 0
    if any(t in _NAME_STRUCT_WORDS for t in tokens):
        return 2
    if not _term_words:
        return 0
    hits = sum(1 for t in tokens if t in _term_words)
    if hits == len(tokens) and (len(tokens) == 1 or hits >= 3):
        return 2
    return 1 if 2 * hits >= len(tokens) else 0


def is_junk_name(name):
    return junk_level(name) > 0


def name_from_filename(filename):
    """'Naukri_Priyanka.k[3y_6m].pdf' -> 'Priyanka K'; None if it doesn't read like a name."""
    stem = _FILENAME_NOISE.sub("", os.path.splitext(filename or "")[0])
    stem = re.sub(r'\[[^\]]*\]|\([^)]*\)', " ", stem)
    words = [w for w in re.sub(r'[\d_\-.]+', " ", stem).split()
             if w.isalpha() and w.lower() not in _FILENAME_DROP]
    if not 1 <= len(words) <= 4 or len("".join(words)) < 5:
        return None
    candidate = " ".join(w.capitalize() for w in words)
    return None if is_junk_name(candidate) else candidate


def _namey(candidate):
    letters = re.sub(r'[^A-Za-z]', '', candidate or "")
    return len(letters) >= 5 and re.search(r'[aeiouAEIOU]', letters) is not None and not is_junk_name(candidate)


def _agrees(candidate, hints):
    """True when some word of the name appears in (or starts like) the email
    local part or the filename-derived name."""
    for t in _alpha_tokens(candidate):
        for h in hints:
            if t in h or (len(h) >= 5 and h in t) or (len(t) >= 4 and t[:4] in h):
                return True
    return False


def refine_name(name, email, phone, linkedin, filename, lines, headers):
    """(name, source). source: page | page_match | filename | email.

    A name is kept as is unless it looks junk AND does not match the email
    address or the filename (matching either is strong evidence it is right,
    whatever words it contains). A replacement is the first line on the page
    that matches the email/filename; only a strongly junk name (a heading,
    phrase or ID) may be replaced by a name built from the filename or email,
    and only for a file that has contact details."""
    cleaned = _NAME_LEAD_LABEL.sub("", name).strip() if name else name
    if cleaned and cleaned != name and junk_level(cleaned) == 0:
        name = cleaned
    from_file = name_from_filename(filename)
    from_email = _name_from_email(email)
    hints = [re.sub(r'[^a-z]', '', s.lower()) for s in ((email or "").split("@")[0], from_file or "")]
    hints = [h for h in hints if len(h) >= 4]
    level = junk_level(name)
    if name and (level == 0 or (hints and _agrees(name, hints))):
        return name, "page"
    if hints:
        for text, *_ in lines[:80]:
            candidate = find_name([(text, 12.0, False, "x")], headers, _titles_matcher, None)
            if candidate and junk_level(candidate) == 0 and _agrees(candidate, hints):
                return candidate, "page_match"
    if (not name or level == 2) and (email or phone or linkedin):
        if from_file and _namey(from_file):
            return from_file, "filename"
        if from_email and _namey(from_email):
            return from_email, "email"
    return name, "page"


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


_NO_SPACE_HEADINGS = {}


def _headings_without_spaces():
    if not _NO_SPACE_HEADINGS:
        for phrase, category in _PHRASE_TO_SECTION.items():
            if " " in phrase:
                _NO_SPACE_HEADINGS.setdefault(phrase.replace(" ", ""), category)
    return _NO_SPACE_HEADINGS


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
    if len(norm) >= 8 and " " not in norm and norm in _headings_without_spaces():
        return _headings_without_spaces()[norm], True   # some PDFs lose every space: "CAREEROBJECTIVE"
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
        text = re.sub(r'\s+(?:in|of|from)\s*[,;]?$', '', text).strip()   # "MCA in" with nothing after it
        key = re.sub(r'[^a-z0-9]', '', text.lower())
        if key and key not in seen and len(text) >= 3:
            seen.add(key)
            out.append(text)

    for override, raw in blocks:
        entries, pending = [], []
        lines = [strip_grades(l) for l in split_lines(raw)]
        for line in lines:
            line = re.sub(r'[\ue000-\uf8ff\u25aa\u25cf\u2022]', '', line).strip()
            if _RESPONSIBILITY_HEADING.match(line):
                break   # what follows is a project or job description, not a qualification
            if not line or DISCONTINUED_PATTERN.search(line):
                continue
            if _looks_like_duty(line) and not _edu_level(line):
                continue
            if _CERT_NAMED.search(line) and not _edu_level(line):
                continue   # "AMFI Certified" is a certification, not a qualification
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


# Education written as a sentence ("Completed Diploma in Mechanical Engineering from XYZ
# College"). Used only when the Education section gave nothing. The short forms
# BE / ME / BA / MA must be capitals, or the ordinary word "be" would match.
_DEGREE_IN_SENTENCE = re.compile(
    r"\b(?:(?i:diploma|b\.?\s?tech|m\.?\s?tech|mba|pgdm|mca|bca|b\.?\s?sc|m\.?\s?sc|b\.?\s?com|m\.?\s?com|"
    r"bba|ph\.?\s?d|bachelor|master|post[- ]graduate)|B\.?E\.?|M\.?E\.?|B\.?A\.?|M\.?A\.?)(?![A-Za-z])")
_SENTENCE_END = re.compile(r"\.\s|;|\s[-–—]\s|\s\|\s|\s+with\s|\s+where\s|\s+and\s+(?:I|have|am|my)\b|\s+I\s+")
_PROSE_VERBS = re.compile(r"\b(seek|seeking|looking|aspire|aspiring|want|wish|willing|ability|knowledge|"
                          r"experience|worked|working|responsible|handled)\b", re.IGNORECASE)


def extract_educations_from_sentences(full_text):
    """Up to three education entries found in running text, or []."""
    found, seen = [], set()
    for raw_line in full_text.split("\n"):
        line = raw_line.strip(" \t" + BULLET_CHARS + "➢❖▪●*>-")
        if not line or len(line.split()) > 30:
            continue
        match = _DEGREE_IN_SENTENCE.search(line)
        if not match or not re.search(r"\b(from|in|at)\b|\(", line[match.end():], re.IGNORECASE):
            continue
        entry = line[match.start():]
        end = _SENTENCE_END.search(entry)
        if end:
            entry = entry[:end.start()]
        entry = _clean_edu_part(strip_grades(entry))
        if len(entry) < 8 or len(entry.split()) > 14 or _PROSE_VERBS.search(entry):
            continue
        key = re.sub(r"[^a-z0-9]", "", entry.lower())
        if key not in seen:
            seen.add(key)
            found.append(entry)
        if len(found) == 3:
            break
    return found


_CERT_NOISE = re.compile(
    r'^(does\s+not\s+expire|credential\s+id\b|issued\b|expires?\b|valid\s+(till|until)\b|'
    r'certificate\s+no\b|id\s*[:#]|link\b|url\b)', re.IGNORECASE)
_CERT_PARTICIPATION = re.compile(r'\b(participat\w*|attended|volunteer\w*|appreciation|received|awarded|won|winner)\b',
                                 re.IGNORECASE)
# A line that describes a job duty ("Prepare daily production material ...", "Handling day to day
# incidents ...") is never a certification or a qualification, even when a wrongly cut section holds it.
_DUTY_VERB = re.compile(
    r'^(process|prepare|dispatch|perform|conduct|provid\w*|handl\w*|ensur\w*|support\w*|monitor\w*|assist\w*|'
    r'maintain\w*|manag\w*|coordinat\w*|develop\w*|implement\w*|execut\w*|prepar\w*|review\w*|'
    r'resolv\w*|train\w*|lead\w*|plan\w*|generat\w*|creat\w*|analy[sz]\w*|respond\w*|verif\w*|'
    r'check\w*|updat\w*|attend\w*|follow\w*|report\w*|negotiat\w*|liais\w*|oversee\w*|drive\w*)\b',
    re.IGNORECASE)
_RESPONSIBILITY_HEADING = re.compile(r'^(?:key\s+|roles?\s*(?:&|and)\s*)?responsib\w*\s*:?$', re.IGNORECASE)
_SKILLS_LABEL_LINE = re.compile(r'^(?:[\w&/]+\s+){0,2}skills?\s*:', re.IGNORECASE)
_DATE_RANGE_LINE = re.compile(
    r'\b(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*)?(?:19|20)\d{2}\s*[-\u2013\u2014]+\s*'
    r'(?:(?:(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*)?(?:19|20)\d{2}|present|till\s+date|current)',
    re.IGNORECASE)


def _looks_like_duty(line):
    return bool(_DUTY_VERB.match(line.lstrip(" \u2022-*>")) or _RESPONSIBILITY_HEADING.match(line.strip()))


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
        if strict and _SKILLS_LABEL_LINE.match(line):
            continue
        if not strict:
            # the certifications text ran on into a job entry ("Royal Enfield  MAY 2018 - MAY 2019"):
            # everything from there on belongs to the experience section
            if _DATE_RANGE_LINE.search(line) and not _CERT_STRONG.search(line) and not re.search(
                    r'certif|course|diploma|training', line, re.IGNORECASE):
                break
            if (_looks_like_duty(line) or _SKILLS_LABEL_LINE.match(line)) and not _CERT_STRONG.search(line):
                continue
            if re.search(r'\binternship\b', line, re.IGNORECASE) and not _CERT_STRONG.search(line):
                continue
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
        if not line or len(line) < 2 or _is_junk_certification(line):
            continue
        kept.append(line)
    return list(dict.fromkeys(kept))


_CERT_JUNK_WORDS = {
    "year", "years", "month", "months", "date", "duration", "grade", "score", "percentage",
    "certificate", "certificates", "certification", "certifications", "course", "courses",
    "training", "trainings", "information", "additional information", "degree", "associate",
    "online", "udemy", "coursera", "edx", "linkedin learning", "other", "others", "none", "na", "n/a",
}


def _is_junk_certification(line):
    """True for entries that are not a certification name: a single stray word
    ('year', 'Udemy', 'Python'), a school class ('10th'), a heading. A one-word
    entry is kept only if it clearly names a certification ('CCNA', 'PMP')."""
    lowered = line.lower().strip()
    if lowered in _CERT_JUNK_WORDS or re.match(r'^\d{1,2}\s?(st|nd|rd|th)\b', lowered):
        return True
    if len(line.split()) == 1 and not _CERT_STRONG.search(line):
        return True
    return False


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
    text = re.sub(r'[-]', '', text).replace('\xa0', ' ')   # icon-font bullets and no-break spaces
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


_MONTH_YEAR_ONLY = re.compile(
    r'^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*(?:19|20)\d{2}(?:\s*[-\u2013]\s*\S+)?$', re.IGNORECASE)


def _good_project_title(title):
    """False for the leftovers of a wrapped line ("Bank)", "GenAI)"), a bare date ("Jan 2026"), or one
    generic word ("Implementation"). Applied to every project title that is saved."""
    if len(re.sub(r'[^A-Za-z]', '', title)) < 4:
        return False
    if title.count(")") > title.count("("):
        return False
    if _MONTH_YEAR_ONLY.match(title.strip()):
        return False
    words = title.split()
    if len(words) == 1 and title.lower().strip(" .") in {
            "implementation", "supporting", "support", "development", "testing", "project", "projects"}:
        return False
    return True


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
        return list(dict.fromkeys(v for _, v in labels if _good_project_title(v)))
    titles, prev_title = [], False
    for line in project_lines:
        if _looks_like_title(line) and not _is_tech_list(line):
            if prev_title and titles and ((line.isupper() == titles[-1].isupper() and not titles[-1].endswith(":"))
                                          or titles[-1].count("(") > titles[-1].count(")")):
                titles[-1] = titles[-1] + " " + line  # title wrapped across lines
            else:
                titles.append(line)
            prev_title = True
        else:
            prev_title = False
    titles = [_clean_project_title(re.split(r'\s+\|\s+|\s{3,}', t)[0]) for t in titles]
    return list(dict.fromkeys(t for t in titles if t and _good_project_title(t)))


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


# A phrase ending in one of these names an organisation ("Bhad Engineers", "ABC Consultants"),
# not a person's job: nobody's title is the plural "Engineers".
_ORGANISATION_WORDS = {
    "engineers", "consultants", "associates", "developers", "managers", "analysts", "executives",
    "officers", "partners", "services", "systems", "solutions", "technologies", "enterprises",
    "industries", "group", "consultancy", "infotech", "labs", "works", "traders", "builders",
}


def is_job_title(text):
    """Short, title-cased, no sentence words, and ends in an occupation word."""
    t = text.strip(" :-–—•*.,")
    words = t.split()
    if not 1 <= len(words) <= 7 or len(t) > 70:
        return False
    if words[-1].lower().strip("().,") in _ORGANISATION_WORDS:
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
    global _term_words
    _term_words = _build_term_words(skills, titles, industries)


def parse_resume(path, mime_type, filename=""):
    lines = resolve_lines_getter(mime_type)(path)
    full_text = "\n".join(t for t, _, _, _ in lines)

    sections, headers = split_into_sections(lines, _skill_heading_vocab)

    email = find_email(full_text)
    phone = find_phone(full_text)
    linkedin = find_linkedin(full_text)
    name = find_name(lines, headers, _titles_matcher, email)
    name, name_source = refine_name(name, email, phone, linkedin, filename, lines, headers)

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

    educations = extract_educations(sections) or extract_educations_from_sentences(full_text)
    certifications = extract_certifications(sections)
    doc = {
        "parser_version": PARSER_VERSION,
        "name": name,
        "name_source": name_source,
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
    found = [s for s in CORE_SECTIONS if doc.get(s)]
    doc["sections_found"] = found
    doc["section_count"] = len(found)
    doc["has_all_sections"] = len(found) == len(CORE_SECTIONS)
    # Private helpers for check_if_real_resume (names starting with "_"). They are
    # removed before anything is saved.
    doc["_top_text"] = full_text[:1500]
    doc["_text_length"] = len(full_text)
    doc["_heading_count"] = len(headers)
    doc["_resume_filename"] = bool(_RESUME_FILENAME.search(filename or ""))
    return doc


# A name alone isn't proof: with no real name on the page the name finder can
# settle on a phrase ("Java 8 Scenarios", "Total Positions"), so job
# descriptions and study notes would slip in. A resume normally says how to
# reach the candidate -- but job-site exports (Naukri, iimjobs) often hide the
# contact details of a perfectly real resume. Those have many sections, while
# a job description or study note has only a couple, so a document with no
# contact details still counts once it has at least this many sections.
MIN_SECTIONS_WITHOUT_CONTACT = 4


# Signs that a document is a letter, certificate, statement or book page and not a resume.
# A cover letter needs a greeting AND the body or the sign-off of a letter. One of these words alone
# ("Subject:", "Thanking you", "Respected Sir") also appears in ordinary resumes.
_LETTER_GREETING = re.compile(
    r"\b(dear\s+(sir|madam|mam|ma'?am|hiring|team|recruiter|manager)|respected\s+(sir|madam|mam|ma'?am)|"
    r"to\s*,?\s*\n\s*the\s+(hr|hiring|manager|human|recruit))", re.IGNORECASE)
_LETTER_BODY = re.compile(
    r"(\bi\s+am\s+(writing|applying)|\bi['\u2019]m\s+applying|\bi\s+(would\s+like|wish)\s+to\s+apply|"
    r"\bapplication\s+for\s+the\s+(post|position)|\binterest\s+in\s+applying|\bplease\s+find\s+(my|the)\s+(attached\s+)?(resume|cv)|"
    r"\benclosed\s+(is|my)|\bsub(ject)?\s*[:\-]+\s*.{0,60}applic|\bi\s+herewith|\bcover\s+letter\b)", re.IGNORECASE)
_LETTER_SIGNOFF = re.compile(r"\b(yours\s+(faithfully|sincerely|truly)|sincerely\s*,)", re.IGNORECASE)
# Papers that are plainly not resumes. Each pattern is specific: a word like "invoice" or "government of"
# is also in real resumes (accounts executives, government jobs), so words alone never count.
# A certificate announces itself in its first lines; deeper in the page "Certificate of Achievement" is
# just an entry in somebody's certifications list.
_CERTIFICATE_TITLE = re.compile(
    r"(certificate\s+for\s+covid|covid-?19\s+vaccination|vaccination\s+certificate|this\s+is\s+to\s+certify|"
    r"provisional\s+certificate|u[a-z]?yam\s+regist\w+|registration\s+certificate|"
    r"certificate\s+of\s+(completion|participation|appreciation|achievement|business\s+excellence|merit))", re.IGNORECASE)
_NOT_A_RESUME_SIGNS = re.compile(
    r"(account\s+statement\s+as\s+of|account\s+holder'?s?\s+name|opening\s+balance|closing\s+balance|"
    r"tax\s+invoice|invoice\s+(no|number|date)\b|\bgstin\s*:|\birn\s*:|e-?way\s+bill|\back\s+no\b|purchase\s+order|\bgstn\s*(no\.?)?\s*:|"
    r"itemi[sz]ed\s+report|payment\s+(acknowledg\w+|receipt|successful)|transaction\s+(id|no)\b|"
    r"premium\s+(receipt|payment)|billdesk|\bisbn\b|all\s+rights\s+reserved|table\s+of\s+contents)", re.IGNORECASE)
_JOB_AD_SIGNS = re.compile(
    r"(job\s+description|job\s+overview|job\s+title\s*:|about\s+(the\s+)?company|about\s+us\b|we\s+are\s+(looking|seeking|hiring)|"
    r"position\s+title|desired\s+experience|key\s+skills\s+required|number\s+of\s+positions|screening\s+questions|"
    r"walk-?in\s+interview|\bjd\b)", re.IGNORECASE)
_RESUME_FILENAME = re.compile(r"resume|\bcv\b|curriculum|naukri|bio-?data|updazz", re.IGNORECASE)
MAX_TEXT_LENGTH_WITHOUT_CONTACT = 25000   # a resume this long almost always has contact details


def validity_problem(doc):
    """Why a file is not a usable resume, or None if it is. For the classification script a file is a
    resume unless it is clearly something else: a certificate / invoice / statement, a
    job description (a document with no contact details), or it has no readable text at all.
    A cover letter counts as a resume.
    There is NO requirement for a name, a section, contact details or a minimum number of sections."""
    top_text = doc.get("_top_text", "")
    if doc.get("_text_length", 0) < 20:
        return "no readable text in the file (blank page, picture without text, or a scan that could not be read)"
    if doc.get("section_count", 0) <= 3 and (
            len({m.group(0).lower() for m in _NOT_A_RESUME_SIGNS.finditer(top_text[:600])}) >= 2
            or _CERTIFICATE_TITLE.search(top_text[:300])):
        return "looks like a certificate, statement or book page, not a resume"
    has_any_contact = doc.get("email") or doc.get("phone") or doc.get("linkedin")
    if not has_any_contact:
        # "Position Title:" and "JOB DESCRIPTION" are also headings inside a resume's experience section, so
        # one advert phrase counts only at the very top of the page; deeper down two different ones are needed
        ad_phrases = {m.group(0).lower() for m in _JOB_AD_SIGNS.finditer(top_text)}
        if _JOB_AD_SIGNS.search(top_text[:400]) or len(ad_phrases) >= 2:
            return "looks like a job description, not a resume"
    return None


# =====================================================================
#                           THE SIMPLE PART
#   Plain code: find the files, sort them, parse them, save the JSON.
# =====================================================================

SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUTPUT_FOLDER = os.path.join(SCRIPT_FOLDER, "output")
DEFAULT_WORKERS = min(os.cpu_count() or 1, 4)   # reading resumes uses a lot of memory
MAX_ZIP_LEVELS = 3                              # a zip inside a zip inside a zip


# ---------- small helpers ----------

def make_file_id(path_inside_input):
    """A fingerprint of the file's path inside the input folder.
    The same file always gets the same id. Two files with the same name in
    different folders get different ids."""
    return hashlib.md5(path_inside_input.encode("utf-8")).hexdigest()


def folder_part(path_inside_input):
    """'Java dev/cv.pdf' gives 'Java dev'. A file in the top folder gives ''."""
    if "/" in path_inside_input:
        return path_inside_input.rsplit("/", 1)[0]
    return ""


def name_part(path_inside_input):
    """'Java dev/cv.pdf' gives 'cv.pdf'."""
    return path_inside_input.rsplit("/", 1)[-1]


def check_if_real_resume(resume):
    """Returns None if this looks like a real resume, or a short reason if not.
    A real resume has a name, at least one section, and a way to reach the person
    (or, when the contact details are hidden, at least 4 sections)."""
    return validity_problem(resume)


def make_note(path_inside_input, **details):
    """A short record about a file that did NOT become a resume (error, duplicate, ...)."""
    note = {"id": make_file_id(path_inside_input)}
    note.update(details)
    folder = folder_part(path_inside_input)
    if folder:
        note["source_folder"] = folder
    note["source_file"] = name_part(path_inside_input)
    return note


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


# ---------- step 1: find every file ----------

def find_all_files(input_folder, zip_temp_folder, skip_folders, open_zips):
    """Finds every file in the input folder, subfolders included. A .zip file is
    opened and its files are treated as if the zip were a folder.

    Returns two things:
      files        : a list of (path_inside_input, real_path_on_disk)
      zip_problems : a list of (path_inside_input, reason) for zip files that would not open
    """
    folders_to_skip = {name.lower() for name in skip_folders}
    files = []
    zip_problems = []
    unpacked_count = [0]   # numbers the files we unpack, so their names never clash

    def open_zip(zip_path, zip_name_inside_input, level):
        try:
            zip_file = zipfile.ZipFile(zip_path)
        except (zipfile.BadZipFile, OSError) as problem:
            zip_problems.append((zip_name_inside_input, f"could not open zip: {problem}"))
            return
        with zip_file:
            for member in zip_file.infolist():
                if member.is_dir():
                    continue
                parts = [p for p in member.filename.replace("\\", "/").split("/") if p not in ("", ".", "..")]
                if not parts or any(p.lower() in folders_to_skip for p in parts[:-1]):
                    continue
                name_inside_input = zip_name_inside_input + "/" + "/".join(parts)
                extension = os.path.splitext(parts[-1])[1].lower()
                unpacked_count[0] += 1
                saved_path = os.path.join(zip_temp_folder, f"{unpacked_count[0]:07d}{extension}")
                try:
                    with zip_file.open(member) as source, open(saved_path, "wb") as target:
                        shutil.copyfileobj(source, target)
                except (OSError, zipfile.BadZipFile, RuntimeError, NotImplementedError) as problem:
                    zip_problems.append((name_inside_input, f"could not extract from zip: {problem}"))
                    continue
                is_inner_zip = extension == ".zip" and level < MAX_ZIP_LEVELS and zipfile.is_zipfile(saved_path)
                if open_zips and is_inner_zip:
                    open_zip(saved_path, name_inside_input, level + 1)
                else:
                    files.append((name_inside_input, saved_path))

    for folder, subfolders, file_names in os.walk(input_folder):
        subfolders[:] = sorted(n for n in subfolders if n.lower() not in folders_to_skip)
        folder_inside_input = os.path.relpath(folder, input_folder)
        for file_name in sorted(file_names):
            if folder_inside_input == ".":
                path_inside_input = file_name
            else:
                path_inside_input = folder_inside_input.replace(os.sep, "/") + "/" + file_name
            real_path = os.path.join(folder, file_name)
            if open_zips and file_name.lower().endswith(".zip") and zipfile.is_zipfile(real_path):
                open_zip(real_path, path_inside_input, 1)
            else:
                files.append((path_inside_input, real_path))
    return files, zip_problems


# ---------- step 2: sort the files into groups ----------

def sort_files_into_groups(files):
    """Looks at every file and decides what happens to it.

    Returns four lists:
      to_parse    : (path_inside_input, real_path, file_type)  files we will read
      unsupported : notes for files that are not a resume file type
      duplicates  : notes for files with the same content as an earlier file
      errors      : notes for files we could not even open
    """
    to_parse, unsupported, duplicates, errors = [], [], [], []
    first_file_with_content = {}   # content fingerprint -> path of the first file that had it
    for path_inside_input, real_path in files:
        file_type = detect_mime_type(real_path)
        if file_type not in SUPPORTED_MIMES:
            extension = os.path.splitext(path_inside_input)[1].lower() or "(no extension)"
            unsupported.append(make_note(
                path_inside_input, skipped=True,
                message=(f"This file was not processed because its content (checked directly, not just its "
                         f"'{extension}' extension) isn't a recognized resume format. "
                         "Supported: PDF, DOCX, DOC, RTF, ODT, and image files.")))
            continue
        try:
            with open(real_path, "rb") as file:
                content_fingerprint = hashlib.sha256(file.read()).hexdigest()
        except OSError as problem:
            errors.append(make_note(path_inside_input, error=f"could not read file: {problem}"))
            continue
        first_path = first_file_with_content.get(content_fingerprint)
        if first_path is not None:
            duplicates.append(make_note(
                path_inside_input, duplicate_of=first_path,
                message=f"The file is a duplicate of the {first_path} file."))
            continue
        first_file_with_content[content_fingerprint] = path_inside_input
        to_parse.append((path_inside_input, real_path, file_type))
    return to_parse, unsupported, duplicates, errors


# ---------- step 3: turn old formats into .docx ----------

def find_libreoffice():
    for name in ("soffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found
    for path in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                 r"D:\LibreOffice\program\soffice.exe"):
        if os.path.exists(path):
            return path
    return None


def convert_old_formats(paths, temp_folder):
    """Turns .doc / .rtf / .odt files into .docx with LibreOffice.
    Returns {original_path: new_docx_path} for every file that worked."""
    converted = {}
    libreoffice = find_libreoffice()
    if not paths or not libreoffice:
        return converted
    temp_folder = os.path.abspath(temp_folder)
    os.makedirs(temp_folder, exist_ok=True)

    # Work on numbered copies so two files with the same name can never overwrite each other.
    copies = {}   # original path -> (copy to convert, the .docx we expect to appear, is it a web page?)
    for number, path in enumerate(paths):
        extension = os.path.splitext(path)[1].lower() or ".bin"
        is_web_page = detect_mime_type(path) == HTML_MIME
        if is_web_page:
            extension = ".html"   # LibreOffice must see .html, even if the file is called .doc
        copy_path = os.path.join(temp_folder, f"{number:07d}{extension}")
        try:
            shutil.copy2(path, copy_path)
        except OSError:
            continue
        copies[path] = (copy_path, os.path.join(temp_folder, f"{number:07d}.docx"), is_web_page)

    def run_libreoffice(files_to_convert, seconds, web_pages):
        # Web pages only convert to .docx when opened as a Writer document.
        settings = ["--infilter=HTML (StarWriter)", "--convert-to", "docx:MS Word 2007 XML"] if web_pages \
            else ["--convert-to", "docx"]
        try:
            subprocess.run([libreoffice, "--headless", "--norestore"] + settings +
                           ["--outdir", temp_folder] + files_to_convert,
                           capture_output=True, timeout=seconds, check=False)
        except Exception:
            pass

    originals = list(copies)
    for web_pages in (False, True):                       # ordinary files first, then web pages
        group_of_this_kind = [path for path in originals if copies[path][2] == web_pages]
        for start in range(0, len(group_of_this_kind), 50):    # 50 files at a time
            group = group_of_this_kind[start:start + 50]
            run_libreoffice([copies[path][0] for path in group], 60 * len(group), web_pages)
    for path in originals:
        copy_path, docx_path, is_web_page = copies[path]
        if not os.path.exists(docx_path):
            run_libreoffice([copy_path], 120, is_web_page)    # one more try, on its own
        if os.path.exists(docx_path):
            converted[path] = docx_path
    return converted


# ---------- step 4: read the files ----------

def parse_one_file(task):
    """Runs inside a worker process. Reads one file.
    task = (path_to_read, path_inside_input, file_type, file_id)
    Returns (kind, data) where kind is 'real', 'invalid' or 'error'."""
    path_to_read, path_inside_input, file_type, file_id = task
    folder = folder_part(path_inside_input)
    file_name = name_part(path_inside_input)
    try:
        resume = parse_resume(path_to_read, file_type, file_name)
        problem = check_if_real_resume(resume)
        if problem:
            kind = "invalid"
            data = {"invalid_resume": True, "reason": problem}
            for key in ("name", "email", "phone"):
                if resume.get(key):
                    data[key] = resume[key]
        else:
            kind = "real"
            data = {key: value for key, value in resume.items() if not key.startswith("_")}
    except Exception as problem:
        kind = "error"
        data = {"error": str(problem)}
    data["id"] = file_id
    if folder:
        data["source_folder"] = folder
    if kind != "real":
        data["source_file"] = file_name
    return kind, data


def parse_all_files(tasks, workers, lookups_folder, on_result):
    """Reads every task with several worker processes. Each time a file is done,
    on_result(kind, data) is called.

    If a worker process crashes (usually out of memory) we start again with half
    as many workers, only for the files that were not finished. When only one
    worker is left and it still crashes, that one file is recorded as an error
    and we carry on with the rest."""
    waiting = list(tasks)
    done_count = 0
    start_time = time.time()
    while waiting:
        finished_ids = set()
        crashed = False
        try:
            with ProcessPoolExecutor(max_workers=workers, initializer=_init_worker,
                                     initargs=(lookups_folder,)) as pool:
                futures = [pool.submit(parse_one_file, task) for task in waiting]
                for future in as_completed(futures):
                    try:
                        kind, data = future.result()
                    except BrokenProcessPool:
                        crashed = True
                        break
                    finished_ids.add(data["id"])
                    on_result(kind, data)
                    done_count += 1
                    if done_count % 200 == 0 or done_count == len(tasks):
                        speed = done_count / max(time.time() - start_time, 0.001)
                        print(f"  {done_count}/{len(tasks)} files read ({speed:.1f} per second)")
        except BrokenProcessPool:
            crashed = True
        waiting = [task for task in waiting if task[3] not in finished_ids]
        if not crashed:
            break
        if workers > 1:
            workers = max(1, workers // 2)
            print(f"  a worker crashed - trying the {len(waiting)} unfinished files again with {workers} worker(s)")
        elif waiting:
            bad_task = waiting.pop(0)
            print(f"  a worker crashed on {name_part(bad_task[1])} - saving it as an error and carrying on")
            on_result("error", make_note(bad_task[1],
                                         error="the parser process crashed on this file (likely out of memory)"))
            done_count += 1


# ---------- a short list of resumes worth a human look ----------

# ---------- candidate lists (CSV files): one resume record per row ----------

_CSV_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)*\.[A-Za-z]{2,}")


def read_csv_rows(path):
    """(header in lower case, rows) of a CSV or TSV candidate list. A tab-separated header wrapped in
    quotes inside a comma file is split back into its columns."""
    import io
    csv.field_size_limit(10_000_000)
    raw = open(path, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    delimiter = "\t" if text[:400].count("\t") > text[:400].count(",") else ","
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    if rows and len(rows[0]) <= 2 and "\t" in rows[0][0]:
        rows = [r[0].split("\t") if r and "\t" in r[0] else r for r in rows]
    if not rows:
        return [], []
    return [c.strip().lower() for c in rows[0]], rows[1:]


def _csv_text(value):
    value = (value or "").replace("\t", " ").strip()
    return "" if value in ("-", "--", "nan", "None", "null", "NULL") else re.sub(r"\s{2,}", " ", value)


def _csv_items(text, limit=40):
    items = []
    for part in re.split(r"\s*[,;|]\s*", text or ""):
        part = re.sub(r"^(knows|know)\s+", "", part.strip(), flags=re.IGNORECASE).strip(" .>")
        if 2 <= len(part) <= 60 and re.search(r"[A-Za-z]{2}", part) and part.lower() not in {x.lower() for x in items}:
            items.append(part)
    return items[:limit]


def csv_row_to_doc(header, row, folder, file_name):
    """One candidate row -> a resume record with the same fields as a resume read from a file, or
    (None, reason) when the row is not usable (no name, no contact details, nothing to search on)."""
    def get(*names):
        for n in names:
            if n in header and header.index(n) < len(row):
                v = _csv_text(row[header.index(n)])
                if v:
                    return v
        return ""

    name = re.sub(r"\s{2,}", " ", get("name", "full name")).strip(" .-")
    if len(re.sub(r"[^A-Za-z]", "", name)) < 2:
        return None, "no name in the row"
    email = next(iter(_CSV_EMAIL.findall(get("email"))), None)
    phone_text = get("phone", "mobile no.", "phone number")
    phone = find_phone(phone_text) or (re.sub(r"\D", "", phone_text) if len(re.sub(r"\D", "", phone_text)) >= 10 else None)
    linkedin_text = get("linkedin", "linkedin url")
    linkedin = None
    m = re.search(r"linkedin\.com/in/[\w%-]+", linkedin_text, re.IGNORECASE)
    if m:
        linkedin = m.group(0).lower() if len(m.group(0)) < 200 else m.group(0)
    if not (email or phone or linkedin):
        return None, "no email, phone or LinkedIn in the row"

    skill_items = _csv_items(get("skills"))
    skills = _filter_junk_terms(_strip_languages(collect_skills("\n".join(skill_items), "", []), _languages_matcher)) \
        if skill_items else []
    for item in skill_items:      # keep what the row says even if the skill list does not know it
        if len(item.split()) <= 4 and item.lower() not in {x.lower() for x in skills}:
            skills.append(item)
    titles = []
    for value in (get("designation"), get("previous designation"), get("current title")):
        for part in re.split(r"\s*\|\s*", value):
            part = part.strip()
            if 2 <= len(part) <= 80 and part.lower() not in {x.lower() for x in titles}:
                titles.append(part)
    company = " ".join(x for x in (get("company"), get("previous company name"), get("current company")) if x)
    industries = find_industries(" ".join([company, " ".join(titles), " ".join(skill_items)]), "")
    languages = match_languages(get("languages known"), _languages_matcher) if get("languages known") else []

    degree = get("degrees")
    colleges = get("colleges")
    course = get("course")
    specialization = get("specialization")
    college = get("college name")
    educations = []
    if degree or colleges:
        educations.append(", ".join(x for x in (degree, colleges) if x))
    if course or college:
        educations.append(", ".join(x for x in (f"{course} ({specialization})" if course and specialization else course, college) if x))
    if not educations and get("qualification"):
        educations.append(get("qualification"))

    doc = {
        "parser_version": PARSER_VERSION,
        "name": name,
        "name_source": "csv",
        "email": email,
        "linkedin": linkedin,
        "phone": phone,
        "skills": as_csv(skills),
        "titles": as_csv(titles),
        "industries": as_csv(industries),
        "languages": as_csv(languages),
        "skills_list": skills,
        "titles_list": titles,
        "industries_list": industries,
        "languages_list": languages,
        "educations": "; ".join(educations),
        "certifications": "",
        "educations_list": educations,
        "certifications_list": [],
        "projects": [],
        "internships": [],
        "unmatched": "",
        "source_type": "csv_candidate_list",
        "experience": get("experience", "relevant experience", "years of experience", "level of experience"),
        "location": get("location", "city"),
    }
    _count_sections(doc)
    if not doc["section_count"]:
        return None, "nothing to search on (no skills, titles, industries or education)"
    if folder:
        doc["source_folder"] = folder
    return doc, None


def _count_sections(doc):
    found = [section for section in CORE_SECTIONS if doc.get(section)]
    doc["sections_found"] = found
    doc["section_count"] = len(found)
    doc["has_all_sections"] = len(found) == len(CORE_SECTIONS)


def _csv_person_id(doc):
    key = (doc.get("email") or "").lower() or re.sub(r"\D", "", doc.get("phone") or "")[-10:] or (doc.get("linkedin") or "")
    return hashlib.md5(("csv:" + key).encode("utf-8")).hexdigest()


def _merge_csv_docs(first, other):
    """The same person in several files (or rows): keep one record with everything known about them."""
    for key in ("email", "phone", "linkedin", "experience", "location", "source_folder"):
        if not first.get(key) and other.get(key):
            first[key] = other[key]
    for key in ("skills_list", "titles_list", "industries_list", "languages_list", "educations_list"):
        seen = {x.lower() for x in first.get(key, [])}
        first[key] = first.get(key, []) + [x for x in other.get(key, []) if x.lower() not in seen]
    first["skills"], first["titles"] = as_csv(first["skills_list"]), as_csv(first["titles_list"])
    first["industries"], first["languages"] = as_csv(first["industries_list"]), as_csv(first["languages_list"])
    first["educations"] = "; ".join(first["educations_list"])
    _count_sections(first)


def read_candidate_csvs(csv_jobs):
    """Reads every candidate-list CSV. Returns ({person id: record}, [rows that were not usable])."""
    people, invalid = {}, []
    for path_inside_input, real_path, _, file_id in csv_jobs:
        folder, file_name = folder_part(path_inside_input), name_part(path_inside_input)
        try:
            header, rows = read_csv_rows(real_path)
        except Exception as problem:
            invalid.append({"id": file_id, "source_file": file_name, "reason": f"could not read the CSV: {problem}"})
            continue
        for number, row in enumerate(rows, start=2):
            doc, reason = csv_row_to_doc(header, row, folder, file_name)
            if doc is None:
                invalid.append({"id": hashlib.md5(f"{path_inside_input}#{number}".encode()).hexdigest(),
                                "source_file": file_name, "source_folder": folder, "row": number, "reason": reason})
                continue
            doc["id"] = _csv_person_id(doc)
            if doc["id"] in people:
                _merge_csv_docs(people[doc["id"]], doc)
            else:
                people[doc["id"]] = doc
    return people, invalid


def find_doubtful_resumes(real_folder, path_of):
    """Goes through the saved real resumes and lists the ones that deserve a quick
    human check, with the reasons. Nothing is changed or removed: it is only a list."""
    doubtful = []
    for file_name in sorted(os.listdir(real_folder)):
        if not file_name.endswith(".json"):
            continue
        with open(os.path.join(real_folder, file_name), encoding="utf-8") as file:
            resume = json.load(file)
        if resume.get("source_type") == "csv_candidate_list":
            continue
        reasons = []
        name = resume.get("name") or ""
        if resume.get("name_source") in ("filename", "email"):
            reasons.append("the name was rebuilt from the file name or the email, not read from the page")
        if len(name.split()) == 1:
            reasons.append("the name is a single word (first name or surname only?)")
        if any(character.isdigit() for character in name):
            reasons.append("the name contains digits")
        if not resume.get("email") and not resume.get("phone"):
            reasons.append("no email and no phone number")
        if resume.get("section_count", 0) <= 2:
            reasons.append("only one or two sections were found")
        if reasons:
            doubtful.append({"id": resume["id"], "file": path_of.get(resume["id"], ""),
                             "name": name, "reasons": reasons})
    return doubtful


# ---------- files saved inside a web-upload body ----------

_WRAPPED_START = (b"------WebKit", b"x\x9c", b"x\x01", b"x\xda")
_EXTENSION_OF = {PDF_MIME: ".pdf", DOCX_MIME: ".docx", DOC_MIME: ".doc", RTF_MIME: ".rtf", ODT_MIME: ".odt",
                 HTML_MIME: ".html", "image/jpeg": ".jpg", "image/png": ".png"}


def _unwrapped_content(data):
    """The real file hidden inside a saved browser upload ('------WebKitFormBoundary...' around a PDF or
    DOCX) or inside a zlib stream. None if the data is neither."""
    if data.startswith(b"------WebKit"):
        start = data.find(b"\r\n\r\n")
        if start < 0:
            return None
        boundary = data.split(b"\r\n", 1)[0]
        body = data[start + 4:]
        end = body.rfind(b"\r\n" + boundary)
        return body[:end] if end > 0 else body
    try:
        return zlib.decompress(data)
    except zlib.error:
        return None


def unwrap_uploaded_files(files, temp_folder):
    """Replaces every file that is really a wrapped PDF/DOCX/... by a copy of the real file, so it is read
    like any other resume. Files that are not wrapped, or whose content is not a supported type once
    unwrapped, are left exactly as they were."""
    out, count = [], 0
    for path_inside_input, real_path in files:
        try:
            with open(real_path, "rb") as file:
                if not file.read(12).startswith(_WRAPPED_START):
                    out.append((path_inside_input, real_path))
                    continue
            with open(real_path, "rb") as file:
                inner = _unwrapped_content(file.read())
        except OSError:
            out.append((path_inside_input, real_path))
            continue
        new_path = None
        if inner:
            count += 1
            probe = os.path.join(temp_folder, f"unwrapped_{count:06d}.bin")
            with open(probe, "wb") as file:
                file.write(inner)
            kind = detect_mime_type(probe)
            if kind in SUPPORTED_MIMES:
                new_path = probe[:-4] + _EXTENSION_OF.get(kind, ".bin")
                os.replace(probe, new_path)
            else:
                os.remove(probe)
        out.append((path_inside_input, new_path or real_path))
    unwrapped = sum(1 for (_, a), (_, b) in zip(out, files) if a != b)
    if unwrapped:
        print(f"Unwrapped {unwrapped} files that were saved inside a web-upload body.")
    return out


# ---------- the whole run ----------

def run(input_folder, name, output_folder, lookups_folder, workers, skip_folders, open_zips, redo, read_csv=True):
    if not os.path.isdir(input_folder):
        print(f"Input folder not found: {input_folder}")
        return
    run_folder = os.path.join(output_folder, name)
    real_folder = os.path.join(run_folder, "real_resumes")
    rejected_folder = os.path.join(run_folder, "rejected")
    os.makedirs(real_folder, exist_ok=True)
    os.makedirs(rejected_folder, exist_ok=True)

    zip_temp_folder = tempfile.mkdtemp(prefix="resume_zip_")
    convert_temp_folder = tempfile.mkdtemp(prefix="resume_convert_")
    try:
        # 1. find the files
        files, zip_problems = find_all_files(input_folder, zip_temp_folder, skip_folders, open_zips)
        files = unwrap_uploaded_files(files, zip_temp_folder)
        print(f"Found {len(files)} files.")

        # 2. sort them into groups
        to_parse, unsupported, duplicates, errors = sort_files_into_groups(files)
        for path_inside_input, reason in zip_problems:
            errors.append(make_note(path_inside_input, error=reason))
        invalid = []
        result_of = {}   # file id -> what happened: real / invalid / error / duplicate / unsupported
        for note in unsupported:
            result_of[note["id"]] = "unsupported"
        for note in duplicates:
            result_of[note["id"]] = "duplicate"
        for note in errors:
            result_of[note["id"]] = "error"

        # 3. skip resumes that are already saved, and convert old formats for the rest
        still_to_do = []
        already_saved = 0
        for path_inside_input, real_path, file_type in to_parse:
            file_id = make_file_id(path_inside_input)
            if os.path.exists(os.path.join(real_folder, file_id + ".json")) and not redo:
                already_saved += 1
                result_of[file_id] = "real"
            else:
                still_to_do.append((path_inside_input, real_path, file_type, file_id))
        print(f"{len(still_to_do)} files to read, {already_saved} already saved, "
              f"{len(unsupported)} unsupported, {len(duplicates)} duplicates.")

        csv_jobs = [job for job in still_to_do if job[2] == CSV_MIME]
        still_to_do = [job for job in still_to_do if job[2] != CSV_MIME]
        if csv_jobs and not read_csv:
            for path_inside_input, _, _, file_id in csv_jobs:
                unsupported.append(make_note(path_inside_input, skipped=True,
                                             message="A candidate-list CSV file. Not read because --no-csv was given."))
                result_of[file_id] = "unsupported"
            csv_jobs = []
        if csv_jobs:
            print(f"{len(csv_jobs)} candidate-list CSV files will be read row by row.")

        old_format_paths = [real_path for _, real_path, file_type, _ in still_to_do if file_type in CONVERT_MIMES]
        converted = convert_old_formats(old_format_paths, convert_temp_folder)
        if old_format_paths:
            print(f"Old formats (.doc/.rtf/.odt): {len(converted)} of {len(old_format_paths)} converted to .docx.")

        tasks = []
        for path_inside_input, real_path, file_type, file_id in still_to_do:
            path_to_read = real_path
            if file_type in CONVERT_MIMES:
                path_to_read = converted.get(real_path)
                if path_to_read is None:
                    errors.append(make_note(path_inside_input,
                                            error="could not convert to .docx (LibreOffice missing or conversion failed)"))
                    result_of[file_id] = "error"
                    continue
                file_type = DOCX_MIME
            tasks.append((path_to_read, path_inside_input, file_type, file_id))

        # 4. read the files; every real resume is saved the moment it is done
        saved_now = 0

        def when_a_file_is_done(kind, data):
            nonlocal saved_now
            file_id = data["id"]
            result_of[file_id] = kind
            if kind == "real":
                save_json(os.path.join(real_folder, file_id + ".json"), data)
                saved_now += 1
            elif kind == "invalid":
                invalid.append(data)
            else:
                errors.append(data)

        if tasks:
            parse_all_files(tasks, workers, lookups_folder, when_a_file_is_done)

        # 4b. candidate-list CSV files: one record per person. A person who already has a resume FILE in this
        #     run is skipped (the file is richer); the same person in several CSV files becomes one record.
        csv_saved = csv_skipped = 0
        csv_invalid = []
        if csv_jobs:
            _init_worker(lookups_folder)
            people, csv_invalid = read_candidate_csvs(csv_jobs)
            have_file = set()
            for file_name in os.listdir(real_folder):
                if file_name.endswith(".json"):
                    with open(os.path.join(real_folder, file_name), encoding="utf-8") as handle:
                        other = json.load(handle)
                    if other.get("source_type") != "csv_candidate_list":
                        if other.get("email"):
                            have_file.add(other["email"].lower())
                        if other.get("phone"):
                            have_file.add(re.sub(r"\D", "", other["phone"])[-10:])
            skipped_people = []
            for person_id, doc in people.items():
                phone_digits = re.sub(r"\D", "", doc.get("phone") or "")[-10:]
                if (doc.get("email") or "").lower() in have_file or (len(phone_digits) >= 10 and phone_digits in have_file):
                    csv_skipped += 1
                    skipped_people.append({"id": person_id, "name": doc["name"], "email": doc.get("email"),
                                           "reason": "this person already has a resume file"})
                    continue
                save_json(os.path.join(real_folder, person_id + ".json"), doc)
                csv_saved += 1
            save_json(os.path.join(rejected_folder, "csv_skipped_already_have_resume.json"), skipped_people)
            save_json(os.path.join(rejected_folder, "csv_invalid_rows.json"), csv_invalid)
            for _, _, _, file_id in csv_jobs:
                result_of[file_id] = "csv"

        # 5. save the rejected files (these are NOT resumes and are never uploaded)
        save_json(os.path.join(rejected_folder, "errors.json"), errors)
        save_json(os.path.join(rejected_folder, "invalid.json"), invalid)
        save_json(os.path.join(rejected_folder, "unsupported.json"), unsupported)
        save_json(os.path.join(rejected_folder, "duplicates.json"), duplicates)

        # a file that used to be a resume but is rejected now must not stay in real_resumes
        for note in errors + invalid + unsupported + duplicates:
            old_json = os.path.join(real_folder, note["id"] + ".json")
            if os.path.exists(old_json):
                os.remove(old_json)

        # 6. a list of every file and what happened to it
        path_of = {make_file_id(path): path for path, _ in files}
        for path, _ in zip_problems:
            path_of[make_file_id(path)] = path
        with open(os.path.join(run_folder, "file_list.csv"), "w", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file, quoting=csv.QUOTE_ALL)
            writer.writerow(["id", "file", "result"])
            for file_id, path in path_of.items():
                writer.writerow([file_id, path, result_of.get(file_id, "")])

        # 7. a short list of real resumes that deserve a human look (nothing is removed)
        review_folder = os.path.join(run_folder, "review")
        os.makedirs(review_folder, exist_ok=True)
        doubtful = find_doubtful_resumes(real_folder, path_of)
        save_json(os.path.join(review_folder, "needs_review.json"), doubtful)
    finally:
        shutil.rmtree(zip_temp_folder, ignore_errors=True)
        shutil.rmtree(convert_temp_folder, ignore_errors=True)

    real_total = sum(1 for result in result_of.values() if result == "real") + csv_saved
    print()
    print(f"DONE. Results are in: {run_folder}")
    print(f"  real resumes     : {real_total}  ({saved_now} new, {already_saved} already saved)  -> real_resumes/")
    print(f"  invalid          : {len(invalid)}  -> rejected/invalid.json")
    print(f"  errors           : {len(errors)}  -> rejected/errors.json")
    print(f"  unsupported      : {len(unsupported)}  -> rejected/unsupported.json")
    print(f"  duplicates       : {len(duplicates)}  -> rejected/duplicates.json")
    print(f"  worth a look     : {len(doubtful)} of the real resumes  -> review/needs_review.json")
    if csv_jobs:
        print(f"  from CSV lists   : {csv_saved} people saved in real_resumes/ (included in the real resumes total), "
              f"{csv_skipped} skipped because they already have a resume file, {len(csv_invalid)} unusable rows "
              f"-> rejected/csv_invalid_rows.json")


def main():
    parser = argparse.ArgumentParser(
        description="Part 1 of 2: read resumes and save them as JSON files. Never touches MeiliSearch.")
    parser.add_argument("input_folder",
                        help="folder with the resumes (subfolders and .zip files are read too)")
    parser.add_argument("--name", default=None,
                        help="name of this batch; results go to output/<name>/ (default: the input folder's name)")
    parser.add_argument("--output-folder", default=DEFAULT_OUTPUT_FOLDER,
                        help="where the results are saved (default: the 'output' folder next to this script)")
    parser.add_argument("--lookups-folder", default=DEFAULT_LOOKUPS_DIR,
                        help="folder with skills.json, titles.json, industries.json, languages.json ... (default: lookups)")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS,
                        help=f"how many files to read at once (default: {DEFAULT_WORKERS})")
    parser.add_argument("--skip-folder", action="append", default=[], metavar="NAME",
                        help="ignore every folder with this name (can be used several times). "
                             "Nothing is skipped unless you ask.")
    parser.add_argument("--no-zips", action="store_true",
                        help="do not open .zip files (they are then counted as unsupported)")
    parser.add_argument("--redo", action="store_true",
                        help="read every file again, even if its JSON is already saved")
    parser.add_argument("--no-csv", action="store_true",
                        help="do not read candidate-list CSV files (they are then counted as unsupported). "
                             "By default each row of a candidate list becomes one resume record.")
    args = parser.parse_args()
    name = args.name or os.path.basename(os.path.normpath(args.input_folder))
    run(args.input_folder, name, args.output_folder, args.lookups_folder, args.workers,
        args.skip_folder, not args.no_zips, args.redo, not args.no_csv)


if __name__ == "__main__":
    main()
