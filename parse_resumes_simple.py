"""
parse_resumes_simple.py  --  read every file the classifier sorted and write ONE JSON FILE PER FILE.

PARSING ONLY: it does not decide whether a file is a resume (that is classify_resumes.py) and it never touches
MeiliSearch. It reads both folders made by the classifier:

    classification/Valid/            ->  Output/Valid/      one SLIM search JSON per valid resume / list person (what is uploaded to MeiliSearch)
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
    python parse_resumes_simple.py                        # Valid AND Invalid  ->  ../Output/Valid (slim search JSON) and ../Output/Invalid
    python parse_resumes_simple.py --only valid           # only the valid resumes
    python parse_resumes_simple.py --limit 50             # only the first 50 files of each folder, to try it
    python parse_resumes_simple.py --workers 3
"""

import os
import html
import re
import sys
import json
import time
import shutil
import hashlib
import tempfile
import argparse
import subprocess
import functools
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
    __slots__ = ("text", "size", "bold", "list_id")

    def __init__(self, text, size=0.0, bold=False, list_id=None):
        self.text, self.size, self.bold, self.list_id = text, size, bold, list_id


def long_path(path):
    """Windows needs the \\\\?\\ prefix for very long paths."""
    path = os.path.abspath(path)
    return "\\\\?\\" + path if os.name == "nt" and not path.startswith("\\\\?\\") else path


def fix_mojibake(text):
    """Text that was UTF-8 but got read as Windows-1252 ('Dr. K.K.R.â€™s' instead of 'Dr. K.K.R.’s', 'JNTUH â€“ Hyderabad') is put right.
    Only when the whole text turns back into valid UTF-8; real accents ('Ãngela', 'Mãhêßh') are never changed."""
    if not isinstance(text, str) or not re.search("[âÃÂ]", text):
        return text
    raw = bytearray()
    for ch in text:                                    # each character back to its byte: cp1252 first, else the plain byte (U+0081, U+009D ...)
        try:
            raw += ch.encode("cp1252")
        except UnicodeEncodeError:
            if ord(ch) > 255:
                return text
            raw.append(ord(ch))
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return text


def decode_word_escape(match):
    """'_x0092_' (how Word writes a symbol) -> the character it stands for."""
    code = int(match.group(1), 16)
    if 0x80 <= code <= 0x9f:
        return bytes([code]).decode("cp1252", "ignore")
    return chr(code)


def decode_control_character(match):
    """A control character U+0080..U+009F that is really a Windows-1252 symbol -> that symbol."""
    return bytes([ord(match.group())]).decode("cp1252", "ignore")


def clean(text):
    """Collapse whitespace and drop private-use / bullet symbols that PDFs and Word leave in the text."""
    text = fix_mojibake(text)
    text = re.sub(r"_x([0-9A-Fa-f]{4})_", decode_word_escape, text)                # Word writes an apostrophe as 'Aug_x0092_12'
    text = re.sub(r"[\x80-\x9f]", decode_control_character, text)                  # U+0092 -> ’ , U+0096 -> –
    text = re.sub(r"[\u200b\u200c\u200d\u2060\ufeff\u00ad]", "", text)                           # zero-width characters ('name\u200b@gmail.com')
    text = (text.replace("\u019f", "ti").replace("\ufb01", "fi").replace("\ufb02", "fl").replace("\ufb00", "ff").replace("\ufb03", "ffi").replace("\ufb04", "ffl")
            .replace("\ufb05", "st").replace("\ufb06", "st"))
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
    return [Line(clean(text)) for text in pytesseract.image_to_string(img).splitlines() if clean(text)]


def block_starts_with_heading(block_lines):
    """True when the first line of a block is a short section heading ('EXPERIENCE', 'Skills')."""
    first = block_lines[0]
    return heading_section(first.text, first.bold, first.size, 0.0) is not None and len(first.text.split()) <= 5


def block_position(item):
    """Sort key of a (bounding box, lines) block: top to bottom, then left to right."""
    bbox = item[0]
    return (bbox[1], bbox[0])


def column_order(pages):
    """pages = [(page width, [(bbox, [Line, ...]), ...]), ...]  ->  the line lists of the blocks in reading order.
    A resume with a main column and a sidebar is stored page by page (page 1 main, page 1 sidebar, page 2 main ...): a job that starts on page 1 and
    goes on on page 2 ends up after the sidebar. Here: the header first (everything above the first heading), then the left column of ALL pages, then the
    right column of all pages. Only done when a vertical gutter exists that almost no block crosses and both columns have headings of their own."""
    blocks_of_all_pages = [(page_index, bbox, block_lines) for page_index, (page_width, blocks) in enumerate(pages) for bbox, block_lines in blocks]
    plain_order = [block_lines for _, _, block_lines in blocks_of_all_pages]
    total_lines = sum(len(block_lines) for block_lines in plain_order)
    if total_lines < 25 or not pages:
        return plain_order
    page_width = pages[0][0]
    best = None
    for gutter_x in range(int(0.30 * page_width), int(0.82 * page_width), 3):
        crossing = sum(len(block_lines) for _, bbox, block_lines in blocks_of_all_pages if bbox[0] < gutter_x - 2 and bbox[2] > gutter_x + 2)
        left = sum(len(block_lines) for _, bbox, block_lines in blocks_of_all_pages if bbox[2] <= gutter_x + 2)
        right = sum(len(block_lines) for _, bbox, block_lines in blocks_of_all_pages if bbox[0] >= gutter_x - 2)
        if left >= 0.15 * total_lines and right >= 0.15 * total_lines and crossing <= 0.10 * total_lines:
            if best is None or crossing < best[0]:
                best = (crossing, [gutter_x])
            elif crossing == best[0]:
                best[1].append(gutter_x)
    if best is None:
        return plain_order
    gutter_x = best[1][len(best[1]) // 2]

    left_heads = sum(1 for _, bbox, block_lines in blocks_of_all_pages if bbox[2] <= gutter_x + 2 and block_starts_with_heading(block_lines))
    right_heads = sum(1 for _, bbox, block_lines in blocks_of_all_pages if bbox[0] >= gutter_x - 2 and block_starts_with_heading(block_lines))
    if left_heads < 1 or right_heads < 1:
        return plain_order                                                # a table or a list in two columns, not two columns of the resume
    first_page = [(bbox, block_lines) for page_index, bbox, block_lines in blocks_of_all_pages if page_index == 0]
    heading_tops = [bbox[1] for bbox, block_lines in first_page if block_starts_with_heading(block_lines)]
    header_bottom = min(heading_tops) if heading_tops else 0
    ordered = [block_lines for bbox, block_lines in sorted(first_page, key=block_position) if bbox[3] <= header_bottom + 2]
    already_placed = {id(block_lines) for block_lines in ordered}
    for side in ("L", "R"):
        for page_number in range(len(pages)):
            part = [(bbox, block_lines) for page_index, bbox, block_lines in blocks_of_all_pages
                    if page_index == page_number and id(block_lines) not in already_placed and ((bbox[0] >= gutter_x - 2) == (side == "R"))]
            ordered += [block_lines for bbox, block_lines in sorted(part, key=block_position)]
    return ordered


def read_pdf(path):
    """Text layer of a PDF (with font size and bold). If there is hardly any text it is a scan: use OCR.
    Hyperlinks are added as lines too, because a 'LinkedIn' word is often a link without visible URL."""
    import io
    from PIL import Image
    pdf = pymupdf.open(path)
    lines, links = [], []
    try:
        pages_blocks = []
        for page in pdf:
            blocks = []
            for block in page.get_text("dict")["blocks"]:
                got = []
                for line in block.get("lines", []):
                    spans = line["spans"]
                    text = clean("".join(s["text"] for s in spans))
                    if text and not re.fullmatch(r"page\s*\d+(?:\s*(?:of|/)\s*\d+)?", text, re.IGNORECASE):
                        got.append(Line(text, round(spans[0]["size"], 1), "bold" in spans[0]["font"].lower()))
                if got:
                    blocks.append((tuple(block["bbox"]), got))
            pages_blocks.append((page.rect.width, blocks))
            links += [text_line["uri"] for text_line in page.get_links() if text_line.get("uri")]
        for got in column_order(pages_blocks):
            lines += got
        if sum(len(text_line.text) for text_line in lines) < 100:           # a scanned PDF: read the pages as pictures
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


def add_docx_paragraph(paragraph, lines):
    """One Word paragraph -> a Line (text, font size, bold, list id) appended to lines. Empty paragraphs are skipped."""
    text = clean(paragraph.text)
    if not text:
        return
    runs = [run for run in paragraph.runs if run.text.strip()]
    size = 0.0
    for run in runs:
        try:
            if run.font.size:
                size = run.font.size.pt
                break
        except (ValueError, TypeError):               # a malformed size in the file: ignore it
            continue
    bold = bool(runs) and all(run.bold or (paragraph.style is not None and paragraph.style.font.bold) for run in runs)
    list_id = None
    try:
        numbering = paragraph._p.pPr.numPr if paragraph._p.pPr is not None else None
        if numbering is not None and numbering.numId is not None:
            list_id = (numbering.numId.val, numbering.ilvl.val if numbering.ilvl is not None else 0)
    except (AttributeError, ValueError):
        list_id = None
    lines.append(Line(text, size, bold, list_id))


def read_docx(path):
    """Paragraphs AND table cells (many resumes are tables), plus header text and hyperlinks."""
    document = docx.Document(path)
    lines = []
    for section in document.sections:                       # the header often holds name and contact data
        for part in section.header.paragraphs:
            add_docx_paragraph(part, lines)
    for item in document.iter_inner_content():              # paragraphs and tables in document order
        if hasattr(item, "rows"):
            seen = set()
            for row in item.rows:
                for cell in row.cells:
                    if cell._tc in seen:                              # merged cells repeat; the element itself is kept in the set, so its id can never be reused by another cell
                        continue
                    seen.add(cell._tc)
                    for part in cell.paragraphs:
                        add_docx_paragraph(part, lines)
        else:
            add_docx_paragraph(item, lines)
    for rel in document.part.rels.values():                 # hyperlinks (LinkedIn)
        if rel.is_external and "linkedin.com" in rel.target_ref.lower():
            lines.append(Line(rel.target_ref))
    xml = document.element.xml
    existing = " ".join(line.text for line in lines)
    boxes = []
    for box in re.findall(r"<w:txbxContent>(.*?)</w:txbxContent>", xml, flags=re.DOTALL):      # text boxes / frames
        for paragraph in re.findall(r"<w:p[ >].*?</w:p>", box, flags=re.DOTALL):
            text = clean("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", paragraph)))
            if text and text not in existing:
                boxes.append(Line(text))
    lines[0:0] = boxes                                       # in reading order (insert(0) for every paragraph reversed them)
    if sum(len(line.text) for line in lines) < 60:                # nothing read at all: take every piece of text in the XML
        lines += [Line(clean(part_text)) for part_text in re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml) if clean(part_text)]
    return lines


def read_docx_raw(path):
    """A DOCX that python-docx cannot open (a picture with a bad checksum, a malformed part): the text of word/document.xml and the headers,
    paragraph by paragraph, without styles."""
    import zipfile
    lines = []
    with zipfile.ZipFile(path) as z:
        for name in sorted(n for n in z.namelist() if re.fullmatch(r"word/(?:header\d*|document)\.xml", n)):
            try:
                xml = z.read(name).decode("utf-8", errors="replace")
            except Exception:
                continue
            for paragraph in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.DOTALL):
                text = clean("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", paragraph)))
                if text:
                    lines.append(Line(text))
    if not lines:
        raise ValueError("no text could be read from the DOCX")
    return lines


def read_text_file(path):
    raw = open(path, "rb").read()
    for encoding in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            return [Line(clean(text)) for text in raw.decode(encoding).splitlines() if clean(text)]
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
        import zipfile
        try:
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
            if "word/document.xml" in names:
                return "docx"
        except Exception:
            return "old"                    # a damaged zip: LibreOffice can often repair and read it
        return "old"                        # .odt (OpenDocument) and other zip documents
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
    for part in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              r"D:\LibreOffice\program\soffice.exe"):
        if os.path.exists(part):
            return part
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


def read_hyperlinks(path):
    """Every hyperlink target of a PDF or a DOCX (the address behind a clickable word such as 'Portfolio' or 'GitHub')."""
    out = []
    try:
        kind = sniff_type(path)
        if kind == "pdf":
            pdf = pymupdf.open(path)
            try:
                for page in pdf:
                    out += [line["uri"] for line in page.get_links() if line.get("uri")]
            finally:
                pdf.close()
        elif kind == "docx":
            document = docx.Document(path)
            out += [rel.target_ref for rel in document.part.rels.values() if rel.is_external and rel.reltype.endswith("/hyperlink")]
    except Exception:
        pass
    return out


def read_lines(path):
    """Any file -> list of Line. Raises an error with a clear message if nothing can be read."""
    kind = sniff_type(path)
    if kind == "pdf":
        return read_pdf(path)
    if kind == "docx":
        try:
            return read_docx(path)
        except Exception:
            return read_docx_raw(path)
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

def load_lookup_file(name):
    """The content of one JSON file of the lookups folder."""
    with open(os.path.join(LOOKUPS, name), encoding="utf-8") as handle:
        return json.load(handle)


class Lookups:
    """The word lists: skills.json, titles.json, industries.json, industry clues, section headings."""

    def __init__(self):
        extra = load_lookup_file("skills_extra.json")
        # skills: key -> nice name.  'add' = extra known skills, 'block' = words that are never skills,
        # 'whole_item_only' = ambiguous words (C, R, Go, Excel, Word) accepted only when they are a whole list item
        self.skills = {}
        for name in load_lookup_file("skills.json") + extra.get("add", []):
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
        self.titles = {norm_key(text): text for text in load_lookup_file("titles.json") if len(norm_key(text)) >= 4}
        self.industries = load_lookup_file("industries.json")
        clues = load_lookup_file("industry_keywords.json")
        self.industry_of = {}                                      # clue (lower) -> industry
        for industry, words in clues.items():
            if not industry.startswith("_"):
                for word in words:
                    self.industry_of.setdefault(word.lower(), industry)
        self.industry_regex = re.compile(r"\b(?:" + "|".join(re.escape(word) for word in sorted(self.industry_of, key=len, reverse=True)) + r")\b",
                                         re.IGNORECASE)
        # section headings: our own list + the skill-heading variants from 'extra skill.json'
        skill_headings = [heading for group in load_lookup_file("extra skill.json").values() for heading in group]
        self.headings = {
            "summary": ["summary", "professional summary", "profile", "profile summary", "career objective", "objective", "career objectives",
                        "about me", "career summary", "executive summary", "professional profile", "personal summary"],
            "skills": ["skills", "key skills", "core skills", "skill set", "areas of expertise", "core competencies", "expertise"] + skill_headings,
            "experience": ["experience", "work experience", "professional experience", "employment history", "career history", "work history",
                           "employment", "professional background", "relevant experience", "experience summary", "employment details",
                           "organizational experience", "organisational experience", "consultancy experience", "consulting experience",
                           "entrepreneurial assignment", "entrepreneurial experience", "previous experience", "freelance experience"],
            "education": ["education", "academic qualification", "academic qualifications", "educational qualification", "educational qualifications",
                          "qualification", "qualifications", "academics", "academic background", "education and training", "educational background",
                          "academic details", "education details", "scholastic record", "academic profile", "academic record", "academic records"],
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
            for word in words:
                self.heading_of[norm_key(word)] = section
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

WEAK_HEADINGS = {"projects", "project", "profile", "summary", "objective", "languages", "interests", "activities", "expertise", "courses", "employment",
                 "qualification", "qualifications", "portfolio", "credentials", "licenses", "hobbies", "references", "strengths"}
# inside a job these words label a sub-block, they do not start a new section of the resume
WEAK_INSIDE_JOB = {"project", "projects", "project name", "expertise", "techstack", "tech stack", "technologies", "technologies used", "tools", "tools and technologies",
                   "tools and technologies used", "tools used", "environment", "skills used", "skills", "key skills", "responsibilities", "key responsibilities",
                   "roles and responsibilities", "role and responsibility", "description", "team size", "client", "client name"}
WEAK_BARE_INSIDE_JOB = {"expertise", "environment", "technologies", "tools", "responsibilities", "key responsibilities", "roles and responsibilities",
                        "role and responsibility", "description", "techstack", "tech stack"}
KEYWORD_SECTIONS = [("certifications", r"certif|licen[sc]e|credential"),
                    ("certifications", r"curricul"),
                    ("education", r"coursework|course work"),
                    ("other", r"achievement|accomplishment|award|honou?r|volunteer|hobb|interest|activit|extra|publication|reference|declaration|personal"),
                    ("internships", r"\bintern(?:s|ship|ships)?\b|summer training"),
                    ("certifications", r"\bcourses?\b|training"),
                    ("skills", r"skill|technolog|tools|competenc|expertise|proficienc"),
                    ("experience", r"experience|employment|work history|career history|professional background"),
                    ("education", r"education|academic|qualification|scholastic"),
                    ("projects", r"project"), ("summary", r"summary|objective|profile|about me")]


def keyword_section(text):
    """An unknown heading ('Volunteering Experience', 'Interests and Hobbies', 'Accomplishments') -> the section it belongs to."""
    low = text.lower()
    for name, pattern in KEYWORD_SECTIONS:
        if re.search(pattern, low):
            return name
    return "other"


SECTION_NOUNS = {"curriculum", "curricular", "extracurricular", "certification", "certifications", "certificate", "certificates", "credential", "credentials", "training", "trainings", "workshop", "workshops",
                 "education", "educational", "academic", "academics", "qualification", "qualifications", "experience", "experiences", "employment", "project", "projects",
                 "skill", "skills", "competencies", "competency", "expertise", "achievement", "achievements", "accomplishment", "accomplishments", "activities", "activity",
                 "summary", "profile", "objective", "objectives", "interests", "hobbies", "languages", "personal", "declaration", "references", "internship", "internships",
                 "publications", "awards", "development", "particulars", "chronicle", "background", "history", "courses", "course", "strengths", "licenses", "licences", "dissertation", "dossier", "trait", "traits", "strength", "information", "statement", "mini"}
HEADING_FILLER = {"and", "&", "of", "the", "in", "for", "to", "professional", "technical", "key", "core", "work", "present", "current", "previous", "other", "additional",
                  "relevant", "selected", "major", "notable", "across", "career", "details", "detail", "extra", "co", "curricular", "co-curricular", "extra-curricular",
                  "attended", "involved", "undertaken", "handled", "completed", "computer", "it", "related", "general", "areas", "set", "known", "information", "data",
                  "my", "our", "curriculum", "curricular", "extracurricular", "software", "functional", "summary", "profile", "organisational", "organizational", "industrial", "academic", "scholastic", "scholastics",
                  "carrer", "carreer", "personality", "educational", "vocational", "on", "job", "highlights", "overview", "record", "records", "status", "level"}


def vocab_heading(text):
    """A section heading recognised by its words alone, whatever its font: a short line made only of heading vocabulary ('Academic Credentials',
    'Training and Workshop Attended:', 'Professional Training & Certifications', 'Present Work Experience') -> the section it names."""
    part_text = clean(text).rstrip(":-–— ").strip()
    if not part_text or len(part_text) > 50 or re.search(r"[@\d.,;()]", part_text) or COMPANY_SUFFIX_RE.search(part_text) or not part_text[:1].isalpha() or not (part_text[:1].isupper() or part_text.isupper()):
        return None
    words = [word.lower() for word in re.split(r"[\s/]+", part_text.replace("&", " & ")) if word]
    if not 1 <= len(words) <= 6 or not any(word in SECTION_NOUNS for word in words):
        return None
    if any(word not in SECTION_NOUNS and word not in HEADING_FILLER for word in words):
        return None
    if any(word in ("academic", "academics", "educational", "education", "qualification", "qualifications", "scholastic") for word in words) and not any(word.startswith("certif") for word in words):
        return "education"                                          # 'Academic Credentials' is the education, although 'credentials' is also a certificate word
    if "technical" in words and "profile" in words:
        return "skills"                                             # 'Technical Profile' lists tools and projects, it is not the career summary
    if any(word in ("summary", "profile", "objective", "objectives") for word in words) and not any(word in ("project", "projects", "experience", "experiences", "skills", "skill", "internship", "internships", "education") for word in words):
        return "summary"                                            # 'Expertise Summary', 'Career Profile': the summary, not the skills
    return keyword_section(part_text)


LABEL_CELL_RE = re.compile(r"^(?:permanent |current |present |residential |correspondence |local |home )?(?:address|phone(?: no\.?| number)?|mobile(?: no\.?)?|contact(?: no\.?| number)?|"
                           r"e-?mail(?: id)?|linkedin|location|city|date of birth|dob|nationality|languages?(?: known)?|marital status|gender|father'?s? name|hobbies|skype|github|"
                           r"project(?: name| title)?|company(?: name)?|employer(?: name)?|organi[sz]ation(?: name)?|designation|duration|period|tenure|job title|position held)$",
                           re.IGNORECASE)


def merge_label_cells(lines):
    """A table written as three cells ('Address' / ':' / 'R/703 Velvet Tower ...') or as a bold label line followed by its value ('Address:' / 'Pune, MH')
    becomes one line 'Address: value', so labels are read the same way as in a plain resume."""
    out, i = [], 0
    while i < len(lines):
        text = lines[i].text.strip()
        label = text.rstrip(":").strip()
        if LABEL_CELL_RE.match(label):
            if i + 2 < len(lines) and lines[i + 1].text.strip() == ":" and lines[i + 2].text.strip():
                out.append(Line(f"{label}: {lines[i + 2].text.strip()}", lines[i].size, lines[i].bold))
                i += 3
                continue
            job_label = re.match(r"(?:company|employer|organi[sz]ation|designation|duration|period|tenure|job title|position)", label, re.IGNORECASE)
            if (text.endswith(":") or (job_label and lines[i].bold is lines[i].bold and i + 3 < len(lines))) and i + 1 < len(lines) and lines[i + 1].text.strip() \
                    and not LABEL_CELL_RE.match(lines[i + 1].text.strip().rstrip(":")) and not heading_section(lines[i + 1].text):
                out.append(Line(f"{label}: {lines[i + 1].text.strip()}", lines[i].size, lines[i].bold))
                i += 2
                continue
        out.append(lines[i])
        i += 1
    return out


GLUED_LABEL_RE = re.compile(r"(?<=[a-z0-9])(?=(?:contact|e-?mail|mobile|phone|address)\s*(?:no\.?|number|id)?\s*:)", re.IGNORECASE)


HEADING_TYPOS = {"certificatons": "certifications", "certifcations": "certifications", "certificatios": "certifications", "certfications": "certifications",
                 "certificaton": "certification", "certifcation": "certification", "certificatoins": "certifications", "experiance": "experience",
                 "experiances": "experience", "experiense": "experience", "expirience": "experience", "experince": "experience", "experiene": "experience",
                 "educaton": "education", "eduction": "education", "qualifcation": "qualification", "qualifications": "qualifications", "acheivements": "achievements",
                 "achivements": "achievements", "achievments": "achievements", "projcts": "projects", "proejcts": "projects", "skils": "skills", "sklls": "skills"}


def tidy_line_text(text):
    """An HTML entity left in the text ('TRAINING &amp; CERTIFICATIONS') is read as the character it stands for; a heading-like line (up to five words) that
    has a known misspelling of a heading word ('CERTIFICATONS', 'EXPERIANCE') gets the word spelled right, so the section is found."""
    if "&" in text and re.search(r"&(?:#\d+|[A-Za-z]+);", text):
        text = html.unescape(text)
    text = re.sub(r"^\s*e\s+(?=[A-Z0-9(])", "", text)                    # a PDF bullet read as the letter 'e': 'e AWS Developer - Associate' -> 'AWS Developer - Associate'
    words = text.split()
    if 1 <= len(words) <= 5 and len(text) <= 60:
        fixed = []
        for word in words:
            bare = re.sub(r"[^A-Za-z]", "", word)
            right = HEADING_TYPOS.get(bare.lower())
            if right:
                right = right.upper() if bare.isupper() else right.capitalize() if bare[:1].isupper() else right
                word = word.replace(bare, right, 1)
            fixed.append(word)
        text = " ".join(fixed)
    return text


def unglue_labels(lines):
    """'Raja Mohan Reddy Thippareddycontact: +91-9490691806Email: x@y.com' (a PDF that lost the space or line break before a label) is split in front of
    each label, so the name, the phone and the e-mail are read from lines of their own."""
    out = []
    for line in lines:
        fixed = tidy_line_text(line.text)
        if fixed != line.text:
            line = Line(fixed, line.size, line.bold, line.list_id)
        parts = GLUED_LABEL_RE.split(line.text) if ":" in line.text else [line.text]
        if len(parts) > 1:
            out += [Line(part.strip(), line.size, line.bold, line.list_id) for part in parts if part.strip()]
        else:
            out.append(line)
    return out


def join_split_headings(lines):
    """A heading that the page wrapped over two lines ('ACHIEVEMENTS AND' / 'CERTIFICATES', 'Tools and' / 'Technologies') is one heading."""
    out, i = [], 0
    while i < len(lines):
        a = lines[i]
        if i + 1 < len(lines):
            b = lines[i + 1]
            if (a.bold or a.text.isupper()) and (b.bold or b.text.isupper()) and re.search(r"(?:\band|&|\bof)\s*$", a.text, re.IGNORECASE) \
                    and len(a.text.split()) <= 4 and len(b.text.split()) <= 3 and a.size == b.size and not re.search(r"[.@\d]", a.text + b.text):
                out.append(Line(a.text + " " + b.text, a.size, a.bold))
                i += 2
                continue
        out.append(a)
        i += 1
    return out


def heading_section(text, bold=False, size=0.0, body_size=0.0):
    """If this line is a section heading, which section? Exact match first, then fuzzy for typos and
    lost spaces ('CAREEROBJECTIVE'). Only short lines can be headings."""
    word_lists = lookups()
    text = clean(text).rstrip(":-–— ").strip()
    if re.match(r"project", text, re.IGNORECASE) and re.search(r"\(\s*intern", text, re.IGNORECASE) and len(text) <= 45:
        return "internships"                                  # 'PROJECT EXPERIENCE(Internship)' lists internships
    text = re.sub(r"^\(?(?:\d{1,2}|[ivx]{1,4})[.)]\s+", "", text, flags=re.IGNORECASE)              # '1. Education', 'II) Skills'
    text = re.sub(r"\s*[(\[][^)\]]*[)\]]\s*$", "", text).strip()                                   # 'Experience (5 Years)'
    if not text or len(text) > 45 or len(text.split()) > 6 or re.search(r"[@\d]", text):
        return None
    key = norm_key(text.replace("&", " and "))
    styled = bold or text.isupper() or text.istitle() or (body_size and size > body_size + 0.5)
    if key in ("address", "permanentaddress", "contactdetails", "contactinformation") and key not in word_lists.heading_of and styled:
        return "personal"                                     # 'ADDRESS' on its own line starts the personal details, it is not part of the project above
    if key in word_lists.heading_of:
        if key in WEAK_HEADINGS and not styled:
            return None                                       # 'projects' at the end of a wrapped sentence is not a heading
        return word_lists.heading_of[key]
    if styled and len(text.split()) <= 4 and len(key) >= 6 and not (len(text.split()) >= 2 and COMPANY_SUFFIX_RE.search(text)):
        match = process.extractOne(key, word_lists.heading_keys, scorer=fuzz.ratio, score_cutoff=88)
        if match:
            return word_lists.heading_of[match[0]]
    return None


def heading_style(line):
    """The look of a line (bold or not, font size rounded to half a point): headings of one resume share one look."""
    return (bool(line.bold), round(line.size * 2) / 2)


def position_in_text(line, positions):
    """Sort key: where a line stands in the resume (0 when unknown)."""
    return positions.get(id(line), 0)


def first_of_pair(pair):
    """Sort key: the first element of a pair."""
    return pair[0]


def split_sections(lines):
    """{'top': [...], 'skills': [...], 'experience': [...], ...}. The lines before the first heading are 'top'.
    'Technical Skills: Java, SQL' (a heading with text after the colon) starts a section with that text.
    A line that has the style of the known headings of this resume (same bold / size) and is short is a heading too, even when its words are not in the
    list ('Volunteering Experience'): its section comes from its words, else 'other'. Inside a job, 'Project: X', 'Techstack: ...' or 'Expertise:' are
    labels of the job, they do not start a new section."""
    word_lists = lookups()
    sizes = collections.Counter(text_line.size for text_line in lines if text_line.size)
    body_size = sizes.most_common(1)[0][0] if sizes else 0.0
    heads = [heading_section(text_line.text, text_line.bold, text_line.size, body_size) for text_line in lines]
    exact = collections.Counter(heading_style(text_line) for text_line, heading in zip(lines, heads)
                                if heading and norm_key(clean(text_line.text).rstrip(":-–— ").strip().replace("&", " and ")) in word_lists.heading_of)
    others = collections.Counter(heading_style(text_line) for text_line, heading in zip(lines, heads) if not heading)
    dynamic = {style_key for style_key, n in exact.items() if n >= 2 and others.get(style_key, 0) <= 3 * n}
    sections = collections.defaultdict(list)
    heading_idx = set()
    current, seen_heading = "top", False
    for index, (line, found) in enumerate(zip(lines, heads)):
        text = clean(line.text)
        list_item = current == "skills" and len(text.split()) == 1 and not line.bold and not text.isupper() and 0 < index < len(lines) - 1             and (not body_size or line.size <= body_size + 0.1) and not heads[index - 1] and not heads[index + 1]             and heading_style(lines[index - 1]) == heading_style(line) == heading_style(lines[index + 1]) and len(lines[index - 1].text.split()) <= 6 and len(lines[index + 1].text.split()) <= 6
        if found and list_item:
            found = None                                          # 'Training' in the middle of a plain list of competencies is one more item, not a heading
        if found:
            seen_heading = True
            key = norm_key(text.rstrip(":-–— ").replace("&", " and "))
            if current in ("experience", "internships", "projects") and key in WEAK_BARE_INSIDE_JOB and found != current:
                found = None                                      # 'Expertise:' / 'Responsibilities:' under a job
        elif vocab_heading(text) and not list_item and not (not line.bold and text.rstrip().endswith(":") and not heading_section(text) and index + 1 < len(lines) and len(lines[index + 1].text.split()) >= 6 and not vocab_heading(lines[index + 1].text)) \
                and not (current in ("experience", "internships", "projects") and norm_key(text.rstrip(": ").replace("&", " and ")) in WEAK_INSIDE_JOB):
            found = vocab_heading(text)                           # the words say it is a heading, the font does not matter
        elif seen_heading and heading_style(line) in dynamic and 1 <= len(text.split()) <= 5 and not re.search(r"[.@\d:]", text.rstrip(": ")) and (line.size > 0 or text.isupper()) \
                and not COMPANY_SUFFIX_RE.search(text) and any(re.search(pat, text.lower()) for _, pat in KEYWORD_SECTIONS):
            found = keyword_section(text)                         # an unknown heading with the style of the known ones, and a section word in it ('Volunteering Experience')
        if found:
            current = found
            heading_idx.add(index)
            continue
        if ":" in line.text:                                      # heading with the first item on the same line
            head, rest = line.text.split(":", 1)
            inline = heading_section(head, True)
            if inline and rest.strip() and current == "education" and norm_key(head) in ("course", "courses"):
                inline = None                                         # '-Course: Bachelor of Engineering' inside the education is a label of the degree
            if inline and rest.strip():
                if current in ("experience", "internships", "projects") and inline != current and norm_key(head.replace("&", " and ")) in {norm_key(word) for word in WEAK_INSIDE_JOB}:
                    sections[current].append(line)                # 'Project: Aruba' inside a job is a label, not the Projects section
                    continue
                current = inline
                sections[current].append(Line(rest.strip(), line.size, line.bold))
                continue
        sections[current].append(line)
    return reclassify_sections(sections, lines, heading_idx)


def reclassify_sections(sections, lines, heading_idx):
    """A heading can have any name ('OTHER QUALIFICATION', 'Profile'): the values under it decide. An education section with no degree, school or year
    that is a list of software names is the skills; a summary / other section that is a dated list of certificates is the certificates."""
    word_lists = lookups()
    owner = {id(line): name for name, ls in sections.items() for line in ls}
    segments, segment = [], []
    for i, line in enumerate(lines):                                    # the lines between two headings
        if i in heading_idx and segment:
            segments.append(segment)
            segment = []
        segment.append(line)
    if segment:
        segments.append(segment)
    cert_heading = any((heading_section(clean(lines[i].text).rstrip(":"), True) or vocab_heading(lines[i].text)) == "certifications" for i in heading_idx)
    if cert_heading:                                                 # the resume has a Certification heading: a certified-line that the column order left under another heading belongs to it
        for name in ("skills", "other", "summary"):
            moved = [line for line in sections.get(name, []) if len(clean(line.text).split()) <= 8 and re.search(r"\bcertified\b|\bcertification\b|\bcertificate\b", line.text, re.IGNORECASE)
                     and CERT_VENDORS.search(line.text) and not re.search(r"\b(?:experience|knowledge|worked|working|looking)\b", line.text, re.IGNORECASE)]
            if moved:
                ids = {id(line) for line in moved}
                sections[name] = [line for line in sections[name] if id(line) not in ids]
                sections["certifications"] = sections.get("certifications", []) + moved
                owner.update({id(line): "certifications" for line in moved})
    if not any(DEGREE_RE.search(clean(line.text)) for line in sections.get("education", [])):         # the Education heading holds no degree: degree lines written elsewhere are the education
        for name in ("top", "summary", "other"):
            moved = [line for line in sections.get(name, []) if len(clean(line.text).split()) <= 20 and DEGREE_RE.search(clean(line.text)) and INSTITUTION_RE.search(clean(line.text))
                     and YEAR_RE.search(clean(line.text)) and not re.search(r"@|https?://", line.text)]
            if moved:
                ids = {id(line) for line in moved}
                sections[name] = [line for line in sections[name] if id(line) not in ids]
                sections["education"] = sections.get("education", []) + moved
                owner.update({id(line): "education" for line in moved})
    if not any(DEGREE_RE.search(clean(line.text)) for line in sections.get("education", [])):         # still no degree: a short degree line written under another heading (or under none) is the education
        pos = {id(line): k for k, line in enumerate(lines)}
        allowed = ("top", "summary", "other", "skills", "projects", "personal", "education")
        field_word = re.compile(r"engineering|science|commerce|arts|management|technology|computer|business|mechanical|electrical|electronics|civil|information|administration", re.IGNORECASE)
        picked = {}
        for name in allowed[:-1]:
            for line in sections.get(name, []):
                text = clean(line.text)
                k = pos.get(id(line))
                if k is None or not DEGREE_RE.search(text) or len(text.split()) > 14 or DATE_RANGE_RE.search(text) \
                        or re.search(r"@|https?://|\b(?:experience|worked|working|responsible|developed|project|years?)\b", text, re.IGNORECASE):
                    continue
                for j in range(max(0, k - 2), min(len(lines), k + 3)):
                    u = clean(lines[j].text)
                    if owner.get(id(lines[j])) not in allowed or re.search(r"@|https?://", u) or DATE_RANGE_RE.search(u) or (COMPANY_SUFFIX_RE.search(u) and not INSTITUTION_RE.search(u) and not DEGREE_RE.search(u)):
                        continue
                    if j == k or (len(u.split()) <= 10 and (INSTITUTION_RE.search(u) or YEAR_RE.search(u) or SCORE_RE.search(u) or DEGREE_RE.search(u)
                                                             or (j == k + 1 and len(u.split()) <= 6 and field_word.search(u)))):
                        picked[id(lines[j])] = (j, lines[j])
        if picked:
            moved = sorted(picked.values(), key=first_of_pair)
            for name in list(sections):
                if name != "education":
                    sections[name] = [line for line in sections[name] if id(line) not in picked]
            sections["education"] = sorted(sections.get("education", []) + [line for _, line in moved], key=functools.partial(position_in_text, positions=pos))
            owner.update({id(line): "education" for _, line in moved})
    for segment in segments:
        names = [owner.get(id(line)) for line in segment if id(line) in owner]
        name = names[0] if names and all(n == names[0] for n in names) else None
        if name == "certifications":
            body = [line for line in segment if id(line) in owner and clean(line.text)]
            items = [clean(line.text) for line in body]
            if len(items) >= 4 and not any(CERT_WORDS.search(text) or CERT_VENDORS.search(text) for text in items) and sum(1 for text in items if len(text.split()) >= 7) >= 0.6 * len(items) \
                    and any(DATE_RANGE_RE.search(text) for text in items):
                ids = {id(line) for line in body}
                sections["certifications"] = [line for line in sections["certifications"] if id(line) not in ids]
                sections["experience"] = sections.get("experience", []) + body             # duties and a dated job under a certificates heading
            continue
        if name not in ("education", "other", "summary"):
            continue
        body = [line for line in segment if id(line) in owner and clean(line.text)]
        items = [clean(line.text) for line in body]
        if len(items) < 2:
            continue
        target = None
        if name == "education" and not any(DEGREE_RE.search(text) or INSTITUTION_RE.search(text) or YEAR_RE.search(text) for text in items) and all(len(text.split()) <= 6 for text in items) \
                and sum(1 for text in items if find_terms(text, word_lists.skills, max_words=3)) >= 2:
            target = "skills"                                        # a list of software names under a heading like 'OTHER QUALIFICATION'
        dated = sum(1 for text in items if re.match(r"^\W*(?:19|20)\d{2}\s*[|:\-–]\s*\S", text))
        if name in ("other", "summary") and dated >= 3 and dated >= 0.4 * len([text for text in items if len(text) > 2]):
            target = "certifications"                                # '2021 | Certified Accounts Payable Professional' ... under a heading like 'Profile'
        if name == "education":                                      # the page ran on: sentences after the last degree / school / year are the experience of the job above
            edu_at = [k for k, text in enumerate(items) if len(text.split()) <= 8 and (DEGREE_RE.search(text) or INSTITUTION_RE.search(text) or YEAR_RE.search(text) or SCORE_RE.search(text))]
            last_edu = edu_at[0] if edu_at else -1
            for k in edu_at[1:]:
                if k - last_edu > 6:
                    break
                last_edu = k                                          # the education lines are one cluster; what follows a gap is another section's text
            edu_at = [last_edu] if edu_at else []
            tail = body[(edu_at[-1] + 1 if edu_at else 0):]
            texts_ = [clean(line.text) for line in tail]
            if edu_at and len(tail) >= 3 and sum(1 for text in texts_ if len(text.split()) >= 6 or text.endswith(":")) >= 0.6 * len(tail) \
                    and not any(DEGREE_RE.search(text) or INSTITUTION_RE.search(text) for text in texts_):
                ids = {id(line) for line in tail}
                sections["education"] = [line for line in sections["education"] if id(line) not in ids]
                sections["experience"] = sections.get("experience", []) + tail
                continue
        if name == "certifications":
            pass
        if target:
            ids = {id(line) for line in body}
            sections[name] = [line for line in sections[name] if id(line) not in ids]
            sections[target] = sections.get(target, []) + body
    return sections


# =====================================================================================================
# 3. CONTACT DATA  (regex finds candidates, a library validates them)
# =====================================================================================================

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
LINKEDIN_RE = re.compile(r"(?:https?://)?(?:[a-z]{2,3}\.)?(?:www\.)?linkedin\.com/in/([A-Za-z0-9_%\-]+)", re.IGNORECASE)


_TLDS = r"(?:co\.in|ac\.in|edu\.in|gov\.in|org\.in|net\.in|com|in|org|net|edu|gov|io|info|biz|me|us|uk|ai|dev|tech|online|co|xyz|app|pro|ae|sa|sg)"
EMAIL_GLUED_RE = re.compile(r"^(.+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*?\." + _TLDS + r")(?=[A-Z][a-z]|[A-Z]{2,}|\d|(?:phone|mobile|mob|contact|tel|cell|linkedin|address|dob|skype|github)|$)")
EMAIL_OTHER_PERSON = re.compile(r"\b(?:reference|referee|referees|ref\.?|guardian|father|mother|spouse|husband|wife|hr|manager|supervisor|recruiter|reporting)\b", re.IGNORECASE)


def cut_domain_end(text, following):
    """'Way2nazir@gmail.co' + 'm,': the line ends inside the domain ending and the next line holds the rest of it (the two make .com, .net, .org ...)."""
    end = re.search(r"@[A-Za-z0-9.\-]+\.([A-Za-z]{1,2})\s*[|,;]?\s*$", text)
    rest = re.fullmatch(r"\s*([A-Za-z]{1,2})\s*[|,;.]?\s*", following or "")
    return bool(end and rest and (end.group(1) + rest.group(1)).lower() in ("com", "net", "org", "edu", "gov"))


def find_email(lines):
    """The candidate's e-mail. Every valid address is scored: earlier in the document is better, the header (first 15 lines) is
    better still, and an address next to 'Reference', 'Manager', 'HR' ... is somebody else's. A PDF may wrap an address over two
    lines ('name@gmail.' + 'com': joined first) or glue the next word to it ('name@gmail.comPhone': the glued word is cut off)."""
    texts = [line.text for line in lines]
    joined = [part_text + texts[i + 1] if i + 1 < len(texts) and ("@" in part_text) and (not EMAIL_RE.search(part_text) or cut_domain_end(part_text, texts[i + 1])) else part_text for i, part_text in enumerate(texts)]
    best = {}                                                   # e-mail -> best score
    for i in range(len(texts)):
        for text in {texts[i], joined[i]}:
            joined_only = text != texts[i]
            flat = text.replace(" @", "@").replace("@ ", "@")
            for candidate in EMAIL_RE.findall(flat):
                candidate = re.sub(r"^(?:e-?mail(?:-?id)?[:\-_]*|mail(?:-?id)?[:\-_]+)", "", candidate.strip(".,;:-"), flags=re.IGNORECASE)
                glued = EMAIL_GLUED_RE.match(candidate)
                if glued:
                    candidate = glued.group(1)
                try:
                    email = validate_email(candidate, check_deliverability=False).normalized
                except EmailNotValidError:
                    continue
                score = -0.04 * i + (0.5 if i < 15 else 0) - (2.0 if EMAIL_OTHER_PERSON.search(text) else 0) - (0.6 if joined_only else 0)
                if email.lower() not in best or score > best[email.lower()][0]:
                    best[email.lower()] = (score, email)
    return max(best.values())[1] if best else None


def is_timestamp_like(raw):
    """'1730530800' (or '911730530800') written as a bare run of digits (no +, no leading 0, no spaces or hyphens) that is a Unix timestamp
    (1.3 to 2.2 billion = 2011-2039) is not a phone: web pages and exports are full of them, and phonenumbers accepts them as an Ambala
    landline. Real mobile numbers start with 6-9, a landline is written with a leading 0, +91 or separators."""
    raw = str(raw).strip()
    if re.fullmatch(r"91\d{10}", raw):
        raw = raw[2:]
    return re.fullmatch(r"\d{10}", raw) is not None and 1_300_000_000 <= int(raw) <= 2_200_000_000


PHONE_LABEL = re.compile(r"\b(?:mobile|mob|phone|ph|contact|cell|tel|call|whatsapp)\b\.?\s*(?:no\.?|number|num)?\s*[:\-]?\s*\(?\+?$", re.IGNORECASE)
PHONE_OTHER = re.compile(r"\b(?:father|mother|guardian|spouse|husband|wife|reference|referee|referees|emergency|office|landline|aadhaar|aadhar|pan|uan|esic|"
                         r"passport|pincode|pin|account|a/c|roll|reg|enrol\w*|ifsc)\b\.?\s*(?:no\.?|number|name)?\s*[:\-]?\s*\(?\+?$", re.IGNORECASE)


def find_phone(lines):
    """The candidate's mobile number. phonenumbers reads many formats (+91, spaces, hyphens, brackets) and checks that the number can
    really exist (default country India). Every valid number is scored: a mobile number beats a landline, a 'Mobile:' / 'Phone:' label
    in front of it helps, the header (first 15 lines) helps, and 'Father:' / 'Reference:' / 'Aadhaar' / 'PIN' in front of it
    disqualify it. Numbers inside links and e-mail addresses are ignored. Returned in international format (+919876543210)."""
    found = {}                                                  # e164 -> best score
    for i, line in enumerate(lines):
        text = re.sub(r"https?://\S+|www\.\S+|\S+@\S+", " ", line.text)
        for match in phonenumbers.PhoneNumberMatcher(text, "IN"):
            number = match.number
            if not phonenumbers.is_valid_number(number) or is_timestamp_like(match.raw_string):
                continue
            before = text[max(0, match.start - 30):match.start]
            kind = phonenumbers.number_type(number)
            score = -0.03 * i + (0.5 if i < 15 else 0)
            score += 1.0 if kind in (phonenumbers.PhoneNumberType.MOBILE, phonenumbers.PhoneNumberType.FIXED_LINE_OR_MOBILE) else -0.5
            score += 1.0 if PHONE_LABEL.search(before) else 0
            score -= 2.5 if PHONE_OTHER.search(before) else 0
            e164 = phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.E164)
            if e164 not in found or score > found[e164]:
                found[e164] = score
    if found:
        return max(found, key=found.get)
    text = "\n".join(text_line.text for text_line in lines)                      # a number split over two lines
    for match in phonenumbers.PhoneNumberMatcher(text, "IN"):
        if phonenumbers.is_valid_number(match.number) and not is_timestamp_like(match.raw_string):
            return phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.E164)
    return None


def find_phones(lines, main):
    """The other numbers of the candidate ('Phone: +91 9799914013, 8079069645'): valid numbers in the first 25 lines that are not labelled as somebody
    else's (father, reference, Aadhaar ...). The main number is not repeated."""
    out = []
    for line in lines[:25]:
        text = re.sub(r"https?://\S+|www\.\S+|\S+@\S+", " ", line.text)
        for match in phonenumbers.PhoneNumberMatcher(text, "IN"):
            if not phonenumbers.is_valid_number(match.number) or is_timestamp_like(match.raw_string):
                continue
            if PHONE_OTHER.search(text[max(0, match.start - 30):match.start]):
                continue
            e164 = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.E164)
            if e164 != main and e164 not in out:
                out.append(e164)
    return out


def find_linkedin(lines):
    """LinkedIn profile URL, normalised to https://www.linkedin.com/in/username. A URL cut over two lines is joined."""
    texts = [line.text for line in lines]
    for i, text in enumerate(texts):
        match = LINKEDIN_RE.search(text)
        if not match:
            continue
        slug = match.group(1)
        next_text = texts[i + 1].strip() if i + 1 < len(texts) else ""
        if (not slug or slug.endswith("-")) and re.fullmatch(r"[A-Za-z0-9_%\-]{3,}", next_text):
            slug += next_text
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
EXTRA_NOT_NAME = {"add", "winning", "dob", "sex", "male", "female", "gender", "nadu", "tamil", "andhra", "pradesh", "karnataka", "maharashtra",
                  "kerala", "telangana", "gujarat", "rajasthan", "punjab", "haryana", "bihar", "odisha", "assam", "bengal", "prime", "minister",
                  "indian", "division", "medical", "parts", "brand", "mgmt", "pin", "code", "gyan", "vihar", "dialogue", "flow", "strength",
                  "strengths", "hobby", "hobbies", "student", "intern", "trainee", "husband", "wife", "passport", "aadhaar", "aadhar", "uan",
                  "ctc", "lpa", "signature", "exam", "passed", "employee", "navy", "internet", "application", "cost", "analysis", "cover", "letter",
                  "airport", "international", "army", "drdo", "isro", "tunnel", "enclave", "apartment", "apartments", "society", "colony", "west",
                  "east", "north", "south", "project", "projects", "soc", "year", "pass", "passing", "having", "educational", "history", "documentation",
                  "miscellaneous", "databases", "materials", "align", "around", "nov"}
COMPANY_WORDS = re.compile(r"\b(pvt|ltd|limited|inc|llc|llp|corp|corporation|company|technologies|solutions|services|systems|industries|"
                           r"consultancy|enterprises|bank|university|college|institute|school|hospital)\b", re.IGNORECASE)
TITLE_HEAD = re.compile(r"\b(engineer|developer|manager|executive|officer|consultant|analyst|lead|head|specialist|associate|assistant|architect|"
                        r"designer|supervisor|coordinator|administrator|programmer|tester|technician|accountant|director|scientist|"
                        r"recruiter|representative|intern|trainee|president|vp|owner|founder|partner|auditor|advisor|planner|buyer|"
                        r"operator|clerk|cashier|agent|teacher|lecturer|professor|trainer|writer|editor|nurse|doctor|pharmacist)s?\b", re.IGNORECASE)
INDIAN_PLACES = {"bangalore", "bengaluru", "mumbai", "delhi", "hyderabad", "chennai", "pune", "kolkata", "noida", "gurgaon", "gurugram", "ahmedabad",
                 "jaipur", "lucknow", "kochi", "indore", "bhopal", "nagpur", "india", "karnataka", "maharashtra", "tamilnadu", "telangana", "kerala",
                 "gujarat", "rajasthan", "uttar", "pradesh", "dubai", "riyadh", "saudi", "arabia", "uae", "usa", "london", "singapore"}


_NAME_WORDS = None


def name_words():
    """name_words.json (built by build_name_words.py from the parsed names): word -> how many candidates have it."""
    global _NAME_WORDS
    if _NAME_WORDS is None:
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "name_words.json"), encoding="utf-8") as f:
                _NAME_WORDS = json.load(f)
        except Exception:
            _NAME_WORDS = {}
    return _NAME_WORDS


def build_name_words(valid_dir):
    """Rewrites name_words.json from the names in Output\\Valid (clean 2-4 word names; a word counts when 2 or more candidates have it), so
    the next parse cuts glued letter-spaced names better. Runs at the end of every parse. Returns the number of words."""
    global _NAME_WORDS
    heading_words = {"contact", "resume", "curriculum", "vitae", "personal", "details", "objective", "location", "preference", "dotcom", "autodesk",
                     "fusion", "mail", "email", "career", "profile", "summary", "tech", "dear", "hiring", "manager", "engineer", "developer", "format", "india"}
    counts = collections.Counter()
    for file_name in os.listdir(valid_dir):
        try:
            with open(os.path.join(valid_dir, file_name), encoding="utf-8") as f:
                person = json.load(f).get("name") or ""
        except Exception:
            continue
        words = person.split()
        if not 2 <= len(words) <= 4 or any(len(word.strip(".")) <= 1 for word in words) or re.search(r"[^A-Za-z .'\-]", person):
            continue                                                    # only clean, ordinary names teach the list
        for word in words:
            word = re.sub(r"[^a-z]", "", word.lower())
            if 3 <= len(word) <= 12 and word not in heading_words:            # a longer word is a glued name, not a name word
                counts[word] += 1
    known = {word: n for word, n in counts.items() if n >= 2}
    if len(known) < 500:                                                # a small or failed parse must not shrink the list
        return len(name_words())
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "name_words.json"), "w", encoding="utf-8") as f:
        json.dump(dict(sorted(known.items())), f, ensure_ascii=False)
    _NAME_WORDS = known
    return len(known)


def cut_glued_name(glued, email="", filename=""):
    """'ABHISHEKTIWARI' -> 'Abhishek Tiwari'. The e-mail / file name may show the break ('abhishek.tiwari@'); else the cut that uses the most known
    name words (name_words.json), with a single initial allowed in front or behind ('JAKUMAR' -> 'J A Kumar' is not tried: no cut, no guess)."""
    s = re.sub(r"[^a-z]", "", glued.lower())
    for source in ((email or "").split("@")[0], os.path.splitext(os.path.basename(filename or ""))[0]):
        parts = [part for part in re.split(r"[^A-Za-z]+", source) if len(part) >= 2]
        if len(parts) >= 2 and "".join(parts).lower() == s:
            return " ".join(part.capitalize() for part in parts)
    known = name_words()
    best = {0: (0.0, [])}                                               # position -> (score, words) of the best cut of s[:position]
    for end in range(1, len(s) + 1):
        for start in range(max(0, end - 14), end):
            word, before = s[start:end], best.get(start)
            if before is None:
                continue
            if word in known:
                score = before[0] + 1.0 + min(known[word], 50) / 100.0
            elif len(word) == 1 and (start == 0 or end == len(s)):      # one initial at either end
                score = before[0] + 0.2
            else:
                continue
            if end not in best or score > best[end][0]:
                best[end] = (score, before[1] + [word])
    cut = best.get(len(s))
    if cut and 1 <= len(cut[1]) <= 4 and sum(len(one_word) >= 3 for one_word in cut[1]) >= 1:
        return " ".join(one_word.capitalize() for one_word in cut[1])
    return s.capitalize()


def unspace_name(text, email="", filename=""):
    """A name set in letter-spaced type: 'A B H I S H E K T I WA R I', 'R U C H I', 'AB H I S H EK T I WAR I' -> 'Abhishek Tiwari', 'Ruchi'.
    Only when most pieces are one or two letters. Anything else comes back unchanged."""
    tokens = text.split()
    if len(tokens) >= 4 and all(re.fullmatch(r"[A-Za-z]{1,2}", part_text) or re.fullmatch(r"[A-Z]{3,6}", part_text) for part_text in tokens) \
            and sum(len(part_text) <= 2 for part_text in tokens) >= 0.6 * len(tokens) and sum(len(part_text) == 1 for part_text in tokens) >= 2:
        return cut_glued_name("".join(tokens), email, filename)
    return text


def name_candidates(lines, email, filename):
    """METHOD A  - scoring.  Look at the first meaningful lines, drop what is clearly not a name (email, phone, URL,
    headings, job titles, companies, places) and score the rest by position, font size, capitalisation, number of
    words, and whether the words also appear in the email address or file name."""
    word_lists = lookups()
    local = re.sub(r"[^a-z]", "", (email or "").split("@")[0].lower())
    base = re.sub(r"[^a-z]", "", os.path.splitext(os.path.basename(filename))[0].lower())
    sizes = [text_line.size for text_line in lines[:70] if text_line.size]
    biggest = max(sizes) if sizes else 0
    typical = sorted(sizes)[len(sizes) // 2] if sizes else 0
    out, meta = [], []                                                 # meta: (line number, font size, one word?) of every entry of out
    for i, line in enumerate(lines[:70]):                              # lines 25-70 (a name in a side column) only count when the e-mail / file name agrees or a label says so
        text = clean(line.text)
        text = unspace_name(text, email, filename)                          # letter-spaced: 'M U R A L I' -> 'Murali', 'A B H I S H E K T I WA R I' -> 'Abhishek Tiwari'
        if re.fullmatch(r"(?:[A-Za-z]\s){3,}[A-Za-z]", text):
            text = text.replace(" ", "")
        # a line that is only the label: 'NAME' / 'Name :' -> the name is on the next line (': BATTULA NARENDER KUMAR')
        labelled = bool(re.match(r"^(?:candidate\s+|full\s+|applicant\s+|student\s+)?name\s*[:\-]", text, re.IGNORECASE)) or (i > 0 and re.fullmatch(r"name\s*:?", clean(lines[i - 1].text), re.IGNORECASE) is not None)
        text = re.sub(r"^[:\-\s]+", "", text)
        text = re.sub(r"^(?:(?:candidate\s+|full\s+|applicant\s+|student\s+)?name\s*[:\-]|resume of|cv of|curriculum vitae of)\s*", "", text, flags=re.IGNORECASE)
        bare_label = re.match(r"^Name\s+(?=[A-Z][a-z]|[A-Z]\.|[A-Z]{2})", text)       # 'Name M.R.Chandrasekaran'
        if bare_label:
            text, labelled = text[bare_label.end():], True
        # 'Rakesh V Email: x@y.com'  /  'Sasi .K Email-id: ...'  -> keep what is before the contact label
        text = re.split(r"\s+[EMP]\s*[:\-](?=\s|$)", text)[0]                         # 'Ravi Kumar Bonthala E : ravikumar@...' -> the name
        text = re.split(r"\s+(?:e-?mail|mail|email-?id|mobile|mob\.?|phone|ph\.?|contact|tel|cell|address|room|flat|house|h\.?\s?no\.?|door|plot|street|road|no\.)(?=\W|$)", text, flags=re.IGNORECASE)[0]
        text = re.sub(r"^(?:mr|mrs|ms|miss|dr|shri|smt)\.?\s+", "", text, flags=re.IGNORECASE)
        text = re.split(r"\s*[|,]\s*|\s+[-–—]\s+|\s{3,}", text)[0].strip(" .-")      # "Name | Software Engineer" -> "Name"
        text = re.sub(r"\s*\([^)]*\)", "", text).strip(" .-")                          # 'AMRUTA SHAH (MALAGI)' -> 'AMRUTA SHAH'
        text = re.sub(r"(?:\s+(?:B\.?\s?E|B\.?\s?Tech|M\.?\s?Tech|MBA|MCA|BCA|B\.?\s?Sc|M\.?\s?Sc|B\.?\s?Com|M\.?\s?Com|Ph\.?\s?D|PMP|PGDM))+\.?$", "", text, flags=re.IGNORECASE).strip(" .-")   # 'Vishal Thadka B.E' -> 'Vishal Thadka'
        text = re.split(r"\s+[A-Za-z]{0,2}[-/]?\d", text)[0].strip(" .-")             # 'VINAY KUMAR MENDIRATTA C-196/6A, PARK' -> the name before the address
        text = re.sub(r"(?<=[A-Za-z])\.(?=[A-Za-z])", ". ", text)                     # 'LOKESHWARAN.P' -> 'LOKESHWARAN. P' (two words)
        if len(text) < 3 or re.search(r"[@\d:/\\()]|www|http", text) or heading_section(text) \
                or (PROSE_START_RE.match(text) and not re.match(r"^[A-Za-z]\.\s*[A-Za-z]", text)) or any(len(word) > 22 for word in text.split()):
            continue
        words = text.split()
        max_words = 5 if any(len(word.strip(".")) <= 2 for word in words) else 4      # 'Rakesh Sonnayla Tenkayla K': initials do not count against the limit
        if not 1 <= len(words) <= max_words or not all(re.fullmatch(r"\.?[A-Za-z][A-Za-z.'\-]*", word) for word in words):
            continue
        letters = [re.sub(r"[^a-z]", "", word.lower()) for word in words]
        agrees = any(len(word) >= 3 and (word in local or word in base) for word in letters)           # the e-mail / file name has one of the words
        big = bool(biggest and line.size >= biggest - 0.5 and line.size > typical)
        if i >= 25 and not (agrees or labelled):                    # big font alone is not enough far down the page ('Indian Prime Minister')
            continue
        # every word starts with a capital letter (except name particles): 'A versatile', 'Second Class sti' are not names.
        # An all-lowercase line ('peddikuppa sai swetha') is a name only with a sign: the e-mail / file name agrees, big font, or a label.
        if not all(word.lstrip(".")[:1].isupper() or word.lower() in {"van", "de", "der", "bin", "al", "la", "di", "da", "von", "bhai", "ben"} for word in words):
            # 'Sumit kumar Kattimani' (one word written in lower case) is a name when the e-mail / file name agrees and it is among the first lines
            if not ((text.islower() and (agrees or big or labelled)) or (agrees and i < 8)):
                continue
        if {word.lower().strip(".") for word in re.split(r"[\s\-]+", text)} & (NAME_BAD_WORDS | EXTRA_NOT_NAME):
            continue
        if norm_key(text) in word_lists.skills:                               # 'JavaScript', 'Python', 'Java Programming', 'Azure Blob Storage' are skills
            continue
        if any(word.islower() and word in {"is", "a", "an", "the", "and", "or", "with", "in", "at", "for", "to", "who", "has", "have", "are", "was", "as", "by",
                                     "on", "from", "that", "this", "he", "she"} for word in words):
            continue                                                 # a sentence, not a name
        score = 1.0 - i * 0.06                                        # earlier is better
        score += 1.2 if labelled else 0                               # 'Name: X'  is as sure as it gets
        score += 0.5 if big else 0
        score += 0.1 if line.bold else 0
        score += 0.2 if text.isupper() or text.istitle() else -0.2
        score += {1: -0.3, 2: 0.4, 3: 0.4, 4: 0.1}.get(len(words), -0.2)
        if agrees:
            score += 0.7                                             # the email / file name agrees
        if COMPANY_WORDS.search(text) or TITLE_HEAD.search(text) or norm_key(text) in word_lists.titles:
            score -= 1.0
        if norm_key(text) in word_lists.skills:
            score -= 0.8
        if {word.lower() for word in words} & INDIAN_PLACES:
            score -= 0.8
        out.append((round(score, 2), text))
        meta.append((i, line.size, len(words) == 1, big))
    # first name and surname on two lines of the same big font ('Karan' / 'Kaushik'): join them
    for (s1, t1), (s2, t2), (i1, z1, one1, big1), (i2, z2, one2, big2) in zip(out, out[1:], meta, meta[1:]):
        if i2 == i1 + 1 and one1 and one2 and z1 and z1 >= biggest - 0.5 and abs(z1 - z2) < 0.5:
            out.append((round(max(s1, s2) + 0.8, 2), f"{t1} {t2}"))
    return sorted(out, reverse=True)


def name_by_scoring(lines, email, filename):
    candidates = name_candidates(lines, email, filename)
    return (candidates[0][1], candidates[0][0]) if candidates else (None, 0.0)


def name_by_spacy(lines):
    """METHOD B - spaCy named-entity recognition: the first PERSON in the top lines."""
    text = "\n".join(clean(line.text) for line in lines[:12] if not EMAIL_RE.search(line.text))
    for entity in lookups().nlp()(text).ents:
        words = entity.text.split()
        if entity.label_ == "PERSON" and 1 <= len(words) <= 4 and not re.search(r"[\d@]", entity.text) \
                and not ({word.lower() for word in words} & NAME_BAD_WORDS) and not COMPANY_WORDS.search(entity.text):
            return clean(entity.text), 0.6
    return None, 0.0


def name_from_email_or_file(email, filename):
    """Last resort: 'rahul.sharma92@gmail.com' -> 'Rahul Sharma'."""
    for source in ((email or "").split("@")[0], re.sub(r"(?i)resume|cv|naukri|updazz|iimjobs|hirist|\d+|\[.*?\]|\(.*?\)", " ", os.path.splitext(os.path.basename(filename))[0])):
        parts = [part for part in re.split(r"[._\-\s]+", re.sub(r"\d+", "", source)) if len(part) >= 2 and part.lower() not in ("email", "mail", "id", "gmail", "yahoo", "com", "dotcom", "in", "co")]
        if 1 <= len(parts) <= 3 and all(part.isalpha() for part in parts):
            return " ".join(part.capitalize() for part in parts), 0.3
    return None, 0.0


FILE_NOISE = re.compile(r"(?i)\b(?:resume|resumes|cv|curriculum|vitae|updated|final|new|latest|copy|profile|download|docx?|pdf|"
                        r"jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b")


def drop_trailing_id(match):
    """A trailing generated id of a file name ('_AE4F1E4977', letters and digits mixed) is removed; a plain word stays."""
    text = match.group(0)
    if re.search(r"\d", text) and re.search(r"[A-Za-z]", text):
        return ""
    return text


def name_from_file(filename):
    """A fuller name from the FILE NAME, for resumes whose own text gives only a first name or a glued word.
    'Rutvik CV - Rutwik Dave.docx' -> 'Rutwik Dave';  'Naukri_RashmiJori[11y_0m].pdf' -> 'Rashmi Jori';
    'PalavesKumar_3.2yrs_iPhone_Bangalore.pdf' -> 'Palaves Kumar'.  Returns None when no 2-4 word name can be read."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = re.sub(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-", "", stem)       # upload id in front
    stem = re.sub(r"^\d{10,}", "", stem)
    if " - " in stem:                                                  # 'Resume1 - Shivu Kumar h m' : the part after the dash is the name
        tail = stem.rsplit(" - ", 1)[1].split()
        if 2 <= len(tail) <= 4 and all(re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*", word) for word in tail):
            return " ".join(word.capitalize() if (word.islower() or word.isupper()) else word for word in tail)
    stem = re.sub(r"\[.*?\]|\(.*?\)", " ", stem)
    stem = re.sub(r"[_\-\s][A-Za-z0-9]{6,}$", drop_trailing_id, stem)   # a trailing id like '_AE4F1E4977'
    stem = re.sub(r"(?i)^(?:naukri|updazz|iimjobs|hirist|resume|cv)[_\- .]*", "", stem)
    stem = re.sub(r"\d+(?:\.\d+)?\s*(?:yrs?|years?)\b.*$|\d+y[_ ]?\d+m.*$", " ", stem, flags=re.IGNORECASE)
    stem = FILE_NOISE.sub(" ", re.sub(r"[_.+\-]+", " ", stem))
    stem = re.sub(r"\d+", " ", stem)
    stem = re.sub(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ", stem)                       # CamelCase -> words
    words = [word for word in stem.split() if re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*", word)]
    if 2 <= len(words) <= 4 and sum(len(word.strip(".")) < 3 for word in words) <= 1 and any(len(word.strip(".")) >= 3 for word in words) \
            and not ({word.lower().strip(".") for word in words} & (NAME_BAD_WORDS | INDIAN_PLACES | {"iphone", "android", "java", "net"})):
        return " ".join(word.capitalize() if (word.islower() or word.isupper()) else word for word in words)
    return None


def reliable_file_name(filename):
    """File names made by a job site or by the upload ('Naukri_X', 'updazz_X', 'iimjobs_X', 'hirist_X', 'Resume - X Y') carry the candidate's name."""
    base = os.path.splitext(os.path.basename(filename))[0]
    return bool(re.match(r"(?i)(?:naukri|updazz|iimjobs|hirist)[_\- ]", base)) or bool(re.search(r" - [A-Za-z][A-Za-z.'\-]* [A-Za-z][A-Za-z.' \-]*$", base))


def looks_like_person(name):
    """A 2-4 word name made only of ordinary name words: no company / job / place / skill / heading word. Such a name from the
    top of the resume is never replaced by the file name."""
    word_lists = lookups()
    words = name.split()
    low = {word.lower().strip(".,") for word in words}
    return (2 <= len(words) <= 4 and not (low & (NAME_BAD_WORDS | INDIAN_PLACES)) and not COMPANY_WORDS.search(name) and not TITLE_HEAD.search(name)
            and norm_key(name) not in word_lists.skills and norm_key(name) not in word_lists.titles and not PROSE_START_RE.match(name))


def better_with_file(name, confidence, filename):
    """A one-word or e-mail-derived name is replaced by the file name's fuller name when the two share a word."""
    full = name_from_file(filename)
    if not full:
        return name, confidence
    if not name:
        return full, 0.3
    clean_words = [word.strip(".") for word in re.split(r"[\s.]+", name.strip()) if word.strip(".")]
    initials = [word for word in clean_words if len(word) == 1]
    rest = [word for word in clean_words if len(word) > 1]
    file_words = full.split()
    if initials and rest and all(any(r.lower() == f.lower() for f in file_words) for r in rest) \
            and any(f[0].lower() == i.lower() and len(f) >= 4 and f.lower() not in {r.lower() for r in rest} for i in initials for f in file_words):
        return full, max(confidence, 0.6)
    joined = re.sub(r"[^a-z]", "", name.lower())
    parts = [re.sub(r"[^a-z]", "", word.lower()) for word in full.split()]
    if len(name.split()) == 1 and (joined == "".join(parts) or any(len(word) >= 3 and word in joined for word in parts)):
        return full, max(confidence, 0.5)
    return name, confidence


def agrees_with_contact(name, email, filename):
    """Do the words of this name also appear in the email address or the file name? (Real names usually do.)"""
    local = re.sub(r"[^a-z]", "", (email or "").split("@")[0].lower())
    base = re.sub(r"[^a-z]", "", os.path.splitext(os.path.basename(filename))[0].lower())
    words = [re.sub(r"[^a-z]", "", word.lower()) for word in (name or "").split()]
    return any(len(word) >= 3 and (word in local or word in base or (len(word) >= 5 and (word[:4] in local or word[:4] in base))) for word in words)


def name_from_linkedin(url):
    """'https://www.linkedin.com/in/pavithra-b-r-205a82169' -> 'Pavithra B R'. The trailing id (letters and digits mixed) is dropped."""
    slug = str(url or "").rstrip("/").split("/in/")[-1].split("?")[0]
    parts = [part for part in re.split(r"[-_]+", slug) if part]
    while parts and (re.fullmatch(r"\d+", parts[-1]) or (len(re.findall(r"\d", parts[-1])) >= 3 and len(parts[-1]) >= 5)):
        parts.pop()                                              # the id: 'b18bab17a', '205a82169', '12345678' (but 'poojaanand26' is a name + 2 digits)
    parts = [re.sub(r"\d+", "", part) for part in parts]
    parts = [part for part in parts if part]
    if not 1 <= len(parts) <= 4 or not all(re.fullmatch(r"[A-Za-z]+", part) for part in parts):
        return None
    return " ".join(part.capitalize() for part in parts)


def capitalise_name_word(word):
    """'sumit' -> 'Sumit', 'anne-marie' -> 'Anne-Marie', "o'neil" -> "O'Neil"."""
    return "-".join("'".join(part.capitalize() for part in piece.split("'")) for piece in word.split("-"))


def tidy_name(name):
    """The name as a recruiter wants to read it:  'ILAKKIYA .A' -> 'Ilakkiya A',  'PRADEEP. S' -> 'Pradeep S',
    'Harish.S' -> 'Harish S',  'peddikuppa sai swetha' -> 'Peddikuppa Sai Swetha'.  A name that is already mixed case
    ('Prithviraj U M', 'McDonald') only loses the stray dots."""
    if not name:
        return name
    words = [word for word in re.split(r"[\s.]+", name.strip()) if word]
    while len(words) > 1 and words[0].lower() in ("mr", "mrs", "ms", "miss", "dr", "shri", "smt", "prof", "engr"):
        words.pop(0)                                            # 'Mr.SUDHARSANK' -> 'SUDHARSANK'
    if not words:
        return name
    particles = {"van", "de", "der", "bin", "al", "la", "di", "da", "von"}
    if all(word.isupper() or word.islower() for word in words):
        words = [capitalise_name_word(word) for word in words]
    else:                                                       # 'Sumit kumar Kattimani' -> 'Sumit Kumar Kattimani'; 'McDonald' stays
        words = [capitalise_name_word(word) if (word.islower() and word not in particles) else word for word in words]
    return " ".join(words)


NOT_A_NAME_PARTS = ("contact", "mailid", "emailid", "location", "preference", "hiring", "dotcom", "autodesk", "fusion", "dearsir", "dearmadam",
                    "btech", "mtech", "dotnet", "curriculum", "vitae", "resume", "personal", "details", "objective")


def plausible_name(name):
    """False for a heading / label / fragment that was read as a name: 'Mail Id', 'Location Preference', 'C O Ntact', 'B Tech',
    'Dear Hiring Manager', 'Aditkr Dotcom', 'Ab H I S H Ek T I War I'."""
    if not name:
        return False
    words = name.split()
    squashed = re.sub(r"[^a-z]", "", name.lower())
    if any(part in squashed for part in NOT_A_NAME_PARTS) or re.match(r"(?i)dear\b", name):
        return False
    short = [word for word in words if len(word.strip(".")) <= 2]
    return not (len(words) >= 4 and len(short) >= 4 and sum(len(word.strip(".")) == 1 for word in words) >= 3)      # a letter-spaced word broken into pieces


def find_name(lines, email, filename):
    name, confidence = find_name_and_score(lines, email, filename)
    return tidy_name(name), confidence


OTHER_PERSON_LABEL = re.compile(r"\b(?:father|mother|spouse|husband|wife|guardian|referee|reference|brother|sister|parent|nominee|manager|supervisor|reporting|company|organi[sz]ation|employer|project|course|institute|college|school|university|bank|account|hr)\b", re.IGNORECASE)
PERSONAL_NAME_RE = re.compile(r"^\W*(?:candidate\W?s?\s+|full\s+|applicant\W?s?\s+|student\W?s?\s+|my\s+)?name(?:\s+of\s+(?:the\s+)?(?:candidate|applicant))?\s*[:\-–]\s*(?P<n>[A-Za-z][A-Za-z.\'' \-]{2,40})\s*$")


def name_from_personal(lines):
    """'Name: Priyanka Joshi' in a personal-details block anywhere in the resume. Father's / mother's / spouse's / reference's names are somebody else."""
    for i, line in enumerate(lines):
        text = clean(line.text)
        if OTHER_PERSON_LABEL.search(text.split(":")[0]) or (i and re.search(r"(?:father|mother|spouse|husband|wife|guardian|reference|referee)", clean(lines[i - 1].text), re.IGNORECASE) and len(lines[i - 1].text.split()) <= 4):
            continue
        match = PERSONAL_NAME_RE.match(text)
        if match:
            cand = match.group("n").strip(" .-")
            words = cand.split()
            if 1 <= len(words) <= 5 and all(word[:1].isupper() or word.isupper() for word in words) and not ({word.lower().strip(".") for word in words} & (NAME_BAD_WORDS | EXTRA_NOT_NAME)):
                return cand
    return None


def genuine_name(name):
    """A name a person can have: not a heading, a skill, a job title, a company, a place or any other word of a resume."""
    if not name or not plausible_name(name):
        return False
    word_lists = lookups()
    words = [word.strip(".,") for word in name.split()]
    low = {word.lower() for word in words}
    letters = re.sub(r"[^a-z]", "", name.lower())
    return (len(letters) >= 3 and 1 <= len(words) <= 5 and not (low & (NAME_BAD_WORDS | EXTRA_NOT_NAME | INDIAN_PLACES))
            and not COMPANY_WORDS.search(name) and not TITLE_HEAD.search(name) and norm_key(name) not in word_lists.skills and norm_key(name) not in word_lists.titles
            and not (PROSE_START_RE.match(name) and not (len(words[0]) == 1 and len(words) > 1 and words[1][:1].isupper()))      # 'A versatile ...' is prose, 'A K Joshi' a name
            and not all(norm_key(word) in word_lists.skills for word in words if len(word) >= 3 and word.lower() not in ("kumar", "singh"))
            and not any(norm_key(word) in word_lists.skills for word in words if len(word) >= 4))


def name_is_backed(text, *, email, filename, known):
    """Evidence that the words are a person's name: a word that other candidates carry as a name (name_words.json), or the e-mail / file name has it."""
    return agrees_with_contact(text, email, filename) or not known or any(re.sub(r"[^a-z]", "", word.lower()) in known for word in text.split())


def find_name_and_score(lines, email, filename):
    """The name is read ONLY from the resume's own text: the top of the page (scored by position, size and capitals), a labelled 'Name:' in the
    personal details, or the first PERSON spaCy finds in the top lines. The e-mail and the file name never make a name: they only help to choose
    between two names that were really written in the resume. A name that is not a genuine person's name (a skill, a place, a title, a heading)
    is not used; then there is no name (None) and the field stays empty."""
    known = name_words()

    candidates = [(score, text) for score, text in name_candidates(lines, email, filename) if genuine_name(text) and name_is_backed(text, email=email, filename=filename, known=known)]
    personal = name_from_personal(lines)
    personal = personal if genuine_name(personal) and name_is_backed(personal, email=email, filename=filename, known=known) else None
    if candidates and candidates[0][0] >= 1.0:
        top = candidates[0]
        if personal and not agrees_with_contact(top[1], email, filename) and agrees_with_contact(personal, email, filename):
            return personal, 0.7                                      # the page top does not fit the e-mail / file name, the 'Name:' in the personal details does
        return top[1], min(1.0, top[0] / 2)
    if personal:
        return personal, 0.7
    spacy_name, spacy_score = name_by_spacy(lines)
    if spacy_name and genuine_name(spacy_name) and agrees_with_contact(spacy_name, email, filename):             # (spaCy's name must match the e-mail / file name)
        return spacy_name, spacy_score
    for score, text in candidates:                                   # not sure (score < 1.0): only a name that the e-mail / file name also contains
        if agrees_with_contact(text, email, filename):
            return text, min(1.0, max(score, 0.1) / 2) * 0.5
    return None, 0.0                                                 # nothing in the resume is a sure name: blank, never a guess


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
                nice = table[key] if isinstance(table[key], str) else ""
                if n > 1 and nice and not re.search(r"[\s\-./+#&]", nice.strip()) and (len(key) <= 4 or any(len(word) == 1 for word in words)):
                    continue                                        # 'is a' is not the skill ISA, 'E m a' / 'S S I' (letter-spaced text) are not skills
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


LISTED_NOT_SKILL = {"english", "hindi", "marathi", "tamil", "telugu", "kannada", "malayalam", "gujarati", "bengali", "punjabi", "urdu", "french", "german", "spanish",
                    "january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december", "present", "current",
                    "india", "yes", "no", "na", "none", "nil", "good", "excellent", "fluent", "basic", "intermediate", "advanced", "beginner", "expert", "etc", "others", "other"}


def listed_tool(item):
    """An item of a skills list that the dictionary does not know, but that reads like a tool or technology: at most three words, starts with a capital
    letter or has a digit / symbol in it ('Angular 8', 'ASP.Net', 'Dapper'), and is not a place, a job title, a language, a month or a sentence."""
    text = clean(item).strip(" .:-")
    words = text.split()
    if not text or len(words) > 3 or len(text) > 30 or len(re.findall(r"[A-Za-z]", text)) < 2 or not (text[:1].isupper() or re.search(r"\d|[#+.]", text)):
        return False
    if not re.fullmatch(r"[A-Za-z0-9+#./&\-]+(?: [A-Za-z0-9+#./&\-]+){0,2}", text) or re.match(r"\d+[.\-]?[A-Za-z]*$", text) or text.endswith("."):
        return False                                                # a fragment ('3.months', 'general). sip', '& inverters. electrical')
    if len(words) > 1 and not all(word[:1].isupper() or word[:1].isdigit() or word.isupper() for word in words):
        return False                                                # a lower-case phrase is a piece of a sentence
    low = {word.lower().strip(".,") for word in words}
    if low & (LISTED_NOT_SKILL | INDIAN_PLACES | NAME_BAD_WORDS) or SENTENCE_START_RE.match(text) or TITLE_HEAD.search(words[-1] + " ") or heading_section(text) or DEGREE_RE.search(text) \
            or re.search(r"@|https?://|\d{4}|\b(?:ltd|pvt|inc|llp)\b", text, re.IGNORECASE) or text.endswith(("ing", "ed")) and len(words) == 1:
        return False
    return True


def negative_score(item):
    """Sort key of a (name, score) pair: the best score first."""
    return -item[1]


def add_skill_score(name, value, *, scores):
    """Keeps the best score of a skill; a score of 1.0 always wins."""
    scores[name] = max(scores.get(name, 0), value) if value < 1.0 else value


def find_skills(sections, use_fuzzy=True):
    """Skill -> confidence.  Where a skill was found decides how sure we are:
         in the Skills section            0.9      (the list is meant to hold skills)
         in experience / projects text    0.4  (+0.3 if it is a multi-word skill or well-known tech, +0.2 if repeated)
         fuzzy match of a skills-list item 0.8 x similarity
       Only skills with confidence >= 0.5 are kept. Ambiguous one-letter / common words (C, R, Go, Excel ...)
       count only when they are a whole item of the skills list."""
    word_lists = lookups()
    scores = {}


    section_lines = sections.get("skills", [])
    items = skill_items(section_lines)
    exact_keys = set()
    for item in items:                                               # 1. skills section, item by item
        key = norm_key(item)
        if key in word_lists.skills:
            add_skill_score(word_lists.skills[key], 0.9, scores=scores)
            exact_keys.add(key)
        else:
            for k, words in find_terms(item, word_lists.skills):
                nice = word_lists.skills[k]
                if k in word_lists.whole_item_only or not (" " in nice.strip() or k in word_lists.known_tech or re.search(r"[A-Z].*[A-Z]|\d|[#+.]", nice[1:])):
                    continue                                        # a generic single word cut out of a longer item ('Cyber' of 'Cyber Crime')
                add_skill_score(nice, 0.9, scores=scores)
                exact_keys.add(k)
            else:
                if listed_tool(item) and not find_terms(item, word_lists.skills):
                    add_skill_score(item, 0.9, scores=scores)                                  # a tool the dictionary does not know ('Microsoft Orleans', 'Dapper'), written as an item of the skills list
    if use_fuzzy:                                                    # 2. near matches: 'MS-Excell' -> 'MS Excel'
        for item in items[:40]:
            key = norm_key(item)
            if key in word_lists.skills or len(key) < 8 or not 2 <= len(item.split()) <= 4:     # multi-word items only: 'Pivot Table' -> 'Pivot Tables'
                continue
            match = process.extractOne(key, word_lists.skill_buckets.get(key[:1], []), scorer=fuzz.ratio, score_cutoff=92)
            if match and match[0] not in word_lists.whole_item_only:
                add_skill_score(word_lists.skills[match[0]], round(0.8 * match[1] / 100, 2), scores=scores)
    from_section = set(scores)
    body = []                                                        # 3. skills mentioned in the rest of the resume
    for name in ("experience", "projects", "internships", "summary", "top"):
        body += [line.text for line in sections.get(name, [])]
    counts = collections.Counter()
    for k, words in find_terms("\n".join(body), word_lists.skills):
        if k in word_lists.whole_item_only:
            continue
        counts[k] += 1
    for k, n in counts.items():
        nice = word_lists.skills[k]
        score = 0.4
        techlike = k in word_lists.known_tech or re.search(r"[A-Z].*[A-Z]|\d|[#+.]", nice[1:]) is not None
        if not techlike and " " not in nice.strip() and k not in exact_keys:
            continue                                                # 'Deals', 'Roller', 'Case' of a job description are words, not skills
        if techlike:                                                 # well-known tech, or written like one (AWS, C++, Node.js)
            score += 0.3
        if n >= 2:
            score += 0.2
        if k in exact_keys:
            score = max(score, 0.9)
        if nice not in scores or scores[nice] < score:
            scores[nice] = round(score, 2)
    if len(from_section) >= 3:                                       # the resume has a skills section: its items are the skills, not every word of the jobs
        scores = {k: v for k, v in scores.items() if k in from_section}
    return {name: s for name, s in sorted(scores.items(), key=negative_score) if s >= 0.5}


# =====================================================================================================
# 6. TITLES and INDUSTRIES
# =====================================================================================================

ROLE_LABEL_RE = re.compile(r"^\W*(?:role|roll|designation|position|job title)\s*[:\-–]\s*(?P<t>[A-Za-z][^,;()|]{2,60}?)\s*[.,;|]?\s*$", re.IGNORECASE)
ROLE_SENTENCE_RE = re.compile(r"\b(?:working|worked|work|serving|served|employed)\s+(?:in\s+\S+(?:\s+\S+)?\s+)?(?:as|with)\s+(?:an?\s+|the\s+)?"
                              r"(?P<t>[A-Za-z][A-Za-z&/. -]{2,50}?)(?=\s+(?:in|at|with|for|from|since|till|until|of)\b|[.,;:(]|$)", re.IGNORECASE)
ROLE_AS_RE = re.compile(r"\bas\s+(?:an?\s+|the\s+)(?P<t>[A-Za-z][A-Za-z&/. -]{2,50}?)(?=\s+(?:in|at|with|for|from|since|till|until|of|company|organi[sz]ation)\b|[.,;:(]|$)", re.IGNORECASE)


def sentence_titles(lines):
    """Job titles written as a plain sentence or a label: 'Role : Java Developer', 'Roll: Education counsellor.', 'Working as Purchase & Stores Executive in X',
    'Work in Axis bank as a sale executive (BDE).', 'Working since 2020 with X at Y as a sales associate.'. The phrase must end with an occupation word (or be a
    known title), so a sentence about the work is not a title."""
    word_lists = lookups()
    out = {}
    for line in lines:
        text = clean(line.text)
        if len(text.split()) > 30 or EMAIL_RE.search(text) or PROSE_START_RE.match(text[:0] or "x") and False:
            continue
        candidates = []
        match = ROLE_LABEL_RE.match(text)
        if match:
            candidates.append(match.group("t"))
        for pattern in (ROLE_SENTENCE_RE, ROLE_AS_RE):
            for second_match in pattern.finditer(text):
                candidates.append(second_match.group("t"))
        for c in candidates:
            c = clean(c).strip(" .,-:")
            words = c.split()
            if not (1 <= len(words) <= 6) or DEGREE_RE.search(c) or COMPANY_SUFFIX_RE.search(c) or INSTITUTION_RE.search(c):
                continue
            key = norm_key(c)
            if key in word_lists.titles or (re.search(r"counsel+ors?$", words[-1], re.IGNORECASE) and len(words) <= 3) or (TITLE_HEAD.search(words[-1] + " ") and not any(word.lower().strip(".,") in NAME_BAD_WORDS - {"executive", "engineer", "developer", "manager", "analyst", "consultant", "software", "data", "system", "systems", "business", "development", "management", "marketing", "sales", "operations", "technology", "technologies", "project"} for word in words)):
                out.setdefault(key, c.title() if c.islower() or c.isupper() else c)
    return list(out.values())


def pool_lines_for_titles(sections):
    return list(sections.get("top", [])[:12]) + list(sections.get("summary", [])[:8]) + list(sections.get("experience", [])) + list(sections.get("other", [])) + list(sections.get("projects", [])) + list(sections.get("skills", []))


def find_titles(lines, sections, use_fuzzy=False):
    """Job titles: match the titles list in the headline (top lines), the summary and experience lines.
    A title must end with an occupation word (Engineer, Manager, Executive ...). With use_fuzzy a near match
    ('Sr. Softwere Engineer') is accepted too."""
    word_lists = lookups()
    pool = [line.text for line in sections.get("top", [])[:10]] + [line.text for line in sections.get("summary", [])[:6]]
    pool += [line.text for line in sections.get("experience", [])]
    found = {}
    for text in pool:
        if len(text.split()) > 14 or EMAIL_RE.search(text):
            continue
        for key, words in find_terms(text, word_lists.titles, max_words=6):
            name = word_lists.titles[key]
            if TITLE_HEAD.search(name.split()[-1] + " ") and (len(name.split()) >= 2 or text in [line.text for line in sections.get("top", [])[:6]]):
                found.setdefault(key, name)
        if use_fuzzy and len(text.split()) <= 8 and TITLE_HEAD.search(text):
            key = norm_key(text)
            match = process.extractOne(key, [k for k in word_lists.titles if k[:1] == key[:1]], scorer=fuzz.ratio, score_cutoff=93)
            if match:
                found.setdefault(match[0], word_lists.titles[match[0]])
    for c in sentence_titles(pool_lines_for_titles(sections)):
        if norm_key(c) not in {norm_key(v) for v in found.values()}:
            found.setdefault(norm_key(c), c)
    # drop a title that is only a part of a longer one found in the same resume
    names = list(found.values())
    return [n for n in names if not any(n != regex_match and n.lower() in regex_match.lower() for regex_match in names)]


def find_industries(lines, sections, use_fuzzy=False):
    """Industries from clue words ('banking', 'thermal power plant', 'Infosys', 'Tata Motors' ...) in the experience /
    summary / project text -- never in education, where a college name would read as 'Education'. The industry that
    is mentioned most comes first. (A fuzzy match against industries.json was tested and added nothing.)"""
    word_lists = lookups()
    text = "\n".join(text_line.text for name in ("experience", "projects", "internships", "summary", "top") for text_line in sections.get(name, []))
    counts = collections.Counter(word_lists.industry_of[regex_match.group(0).lower()] for regex_match in word_lists.industry_regex.finditer(text))
    if use_fuzzy:                                           # kept only so compare_methods.py can show it adds nothing
        for line in lines:
            regex_match = re.match(r"\s*(?:industry|domain|sector)\s*[:\-]\s*(.+)", line.text, re.IGNORECASE)
            if regex_match:
                for part in re.split(r"[,;/|]", regex_match.group(1)):
                    match = process.extractOne(part.strip(), word_lists.industries, scorer=fuzz.WRatio, score_cutoff=88)
                    if match:
                        counts[match[0]] += 2
    return [name for name, _ in counts.most_common(5)]


# =====================================================================================================
# 7. EDUCATION, EXPERIENCE, INTERNSHIPS, PROJECTS, CERTIFICATIONS, SALARY
# =====================================================================================================

# Degrees. The short forms BE / ME / BA / MA / CA must be capitals (otherwise they are ordinary words).
DEGREE_RE = re.compile(
    r"\b(?:(?i:b[.\-]?\s?tech|m[.\-]?\s?tech|b[.\-]?\s?sc|m[.\-]?\s?sc|b[.\-]?\s?com|m[.\-]?\s?com|b\.?\s?pharm|m\.?\s?pharm|b\.?\s?arch|b\.?\s?b\.?\s?a|m\.?\s?b\.?\s?a|b\.?\s?c\.?\s?a|m\.?\s?c\.?\s?a|"
    r"llb|llm|mbbs|pgdm|pgdca|pgdba|pgdbm|pgpm|high\s+school(?=\s*$)|ph\.?\s?d|bachelors?(?=\s+(?:of|in|degree|\())|bachelor['’]s|masters?(?=\s+(?:of|in|degree|\())|master['’]s|doctorate|diploma|post[- ]?graduat\w*|graduat(?:e|ion)|"
    r"intermediate|h\.?s\.?c|s\.?s\.?c|sslc|puc|10th|12th|x{1,2}th|iti|b\.?\s?e\.?\s?d|m\.?\s?s\.?\s?w|b\.?\s?s\.?\s?w|cma|"
    r"matriculation|secondary school|senior secondary|higher secondary|class\s+(?:xii|xi|x)|std\.?\s*(?:xii|xi|x)|standard\s+(?:xii|x)|xii)|B\.?E\.?|M\.?E\.?|B\.?A\.?|M\.?A\.?|C\.?A\.?)(?![A-Za-z\-])")
INSTITUTION_RE = re.compile(r"\b(universit(?:y|ies)|colleges?|collage|institutes?|institue|schools?|polytechnic|academy|iit|nit|iiit|iim|vidyalaya|vidyapeeth|mahavidyalaya|"
                            r"campus|board|institution|govt|government|centre|center|cbse|icse|junior|"
                            r"high school|sr\.? sec|public school|deemed|autonomous|symbiosis|amity|manipal|bits)\b", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
SCORE_RE = re.compile(r"(\d{2}(?:\.\d+)?\s?%|(?:cgpa|gpa|sgpa)\s*[:\-]?\s*\d(?:\.\d+)?(?:\s*/\s*\d+)?|\d(?:\.\d+)?\s*(?:/|of|out of)\s*10|"
                      r"first class(?: with distinction)?|distinction|second class|merit)", re.IGNORECASE)
MONTHS = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
MONTHS_STRICT = (r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
                 r"nov(?:ember)?|dec(?:ember)?)\b\.?")
_DATE_PART = (r"(?:" + MONTHS + r"[\s,'’‘\-]*)?(?:19|20)\d{2}(?:[\s,]*" + MONTHS_STRICT + r")?|\d{1,2}[/.\-](?:19|20)\d{2}|" + MONTHS + r"\s*['’‘]\s*\d{2}\b|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\d{2}\b")
_END_WORDS = r"present|current|up\s?to\s+(?:now|date)|up\s?to|till date|till now|till|now|ongoing|date|today"
DATE_RANGE_RE = re.compile(
    r"(?P<start>" + _DATE_PART + r")\s*(?:-|–|—|−|~|to|till|until|through|thru)\s*(?P<end>" + _END_WORDS + r"|" + _DATE_PART + r")",
    re.IGNORECASE)
SINCE_RE = re.compile(r"\b(?:since|from|w\.?e\.?f\.?|joined(?:\s+in)?)\s+((?:\d{1,2}(?:st|nd|rd|th)?\s+)?(?:" + MONTHS_STRICT + r"\s*,?\s*)?(?:19|20)\d{2}|" + MONTHS + r"\s*['’‘]\s*\d{2})\b", re.IGNORECASE)
ONE_DATE_RE = re.compile(r"^\W*(?:" + _DATE_PART + r"|" + _END_WORDS + r")\W*$", re.IGNORECASE)
LABEL_COMPANY_RE = re.compile(r"^\W*(?:(?:current|present)\s+)?(?:name of (?:the )?)?(?:employer|company|organi[sz]ation|firm)(?:\s+(?:name|\d+))?\s*[:\-–]\s*(.+)$", re.IGNORECASE)
LABEL_TITLE_RE = re.compile(r"^\W*(?:(?:functional\s+)?designation|position(?:\s+(?:held|title))?|role|job title|job profile|title)\s*[:\-–]\s*(.+)$", re.IGNORECASE)
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


GENERIC_INSTITUTION = re.compile(r"^(?:name of (?:the )?)?(?:(?:high|higher secondary|higher|senior secondary|secondary|primary|junior)\s+)?(?:colleges?|schools?|universit(?:y|ies)|institutes?|institutions?|boards?)(?:\s*[/&,]\s*(?:colleges?|schools?|universit(?:y|ies)|institutes?|institutions?|boards?))*"
                                 r"(?:\s+name)?$|\b(?:year of|passing|percentage|completion)\b", re.IGNORECASE)
AMBIGUOUS_DEGREE = {"BE", "ME", "BA", "MA", "CA", "B.E", "M.E", "B.A", "M.A", "C.A"}
AMBIGUOUS_OK = re.compile(r"^\s*(?:\(|[-–:,|.]|$|in\b|of\b|from\b|\d|(?:computer|electronics|electrical|mechanical|civil|information|chemical|instrumentation|"
                          r"production|industrial|automobile|biotech\w*|aeronautical|telecom\w*|it|cse|ece|eee|mech|economics|english|history|arts|"
                          r"commerce|political|sociology|psychology|accountancy)\b)", re.IGNORECASE)
ACRONYM_SKIP = {"CGPA", "SGPA", "GPA", "CSE", "ECE", "EEE", "MECH", "CIVIL", "MBA", "BCA", "MCA", "PGDM", "HSC", "SSC", "SSLC", "PUC", "CBSE", "ICSE"}


def real_institution(text):
    """An institution line: has a school word and is more than the generic table header ('College/University', 'Board')."""
    text = clean(text).strip(" .,-:/")
    return bool(text) and 3 < len(text) <= 90 and bool(INSTITUTION_RE.search(text)) and not GENERIC_INSTITUTION.search(text) \
        and not DEGREE_RE.fullmatch(text) and not re.search(r"@|https?://", text)


EDU_LABELS = {
    "inst": re.compile(r"^\W*(?:university|college|institute|institution|school|board|university\s*/\s*board|board\s*/\s*university)\s*(?:name)?\s*[:\-–]\s*(.+)$", re.IGNORECASE),
    "degree": re.compile(r"^\W*(?:degree|qualification|course|stream|programme|program|examination|exam)\s*(?:name)?\s*[:\-–]\s*(.+)$", re.IGNORECASE),
    "field": re.compile(r"^\W*(?:specialization|specialisation|branch|major|discipline|field of study|subject)\s*[:\-–]\s*(.+)$", re.IGNORECASE),
    "years": re.compile(r"^\W*(?:duration(?: of (?:the )?course)?|period|year|years|year of passing|passing year|completed in|passed in|session)\s*[:\-–]\s*(.+)$", re.IGNORECASE),
    "score": re.compile(r"^\W*(?:percentage|percent|cgpa|gpa|grade|marks|score|aggregate)\s*[:\-–]\s*(.+)$", re.IGNORECASE),
}


def edu_years(window):
    """(start year, end year or 'Present') of one education entry from its lines: a range ('07/2018 - Present', "Aug ’11 - Jun ‘16", '2012-2016'),
    else a single year ('in 2016', '2011: B.Tech'). The line of the degree is read first, then the lines below it."""
    for text in window:
        match = DATE_RANGE_RE.search(text)
        if match:
            start, end = parse_date(match.group("start")), parse_date(match.group("end"))
            if start and end:
                present = re.match(r"present|current|up\s?to|till|now|ongoing|date|today", match.group("end").strip(), re.IGNORECASE)
                return str(min(start[0], end[0])), ("Present" if present else str(max(start[0], end[0])))
        years = YEAR_RE.findall(text)
        if years:
            return (years[0], years[-1]) if len(set(years)) >= 2 else (None, years[-1])
    return None, None


def edu_field(rest, window):
    """The field of study of a degree: '(Computer Science Engineering)', '- Civil Engineering', 'in Mechanical', 'Mechanical' on the next line."""
    rest = clean(rest)
    match = DEGREE_RE.match(rest)
    after = rest[match.end():] if match else rest
    after = re.sub(MONTH_YEAR_RE, " ", after, flags=re.IGNORECASE)
    after = re.sub(r"\b(?:19|20)\d{2}\b", " ", after)
    for segment in re.split(r"[()\[\]]|\s[-–—]\s|[–—:;|,]|\s(?:in|from|at|with)\s", after):
        segment = re.sub(r"^(?:of|in|the)\s+", "", segment.strip(" ."), flags=re.IGNORECASE).strip()
        if segment and not DEGREE_RE.fullmatch(segment):
            field = canon_field(segment)
            if field:
                return field
    paren = re.match(r"\s*\(\s*([A-Za-z][A-Za-z&/ \-]{3,80}?)\s*\)", (rest[match.end():] if match else rest))
    if paren and 1 <= len(paren.group(1).split()) <= 8 and not re.search(r"\b(?:percentage|cgpa|gpa|marks|grade|pursuing|distance|regular|full|part|time|appearing|completed)\b", paren.group(1), re.IGNORECASE) \
            and not real_institution(paren.group(1)):
        return clean(paren.group(1))                                # 'PGDBA (Operations Research & Customer Relationship Management) from SCDL'
    if len(window) > 1:
        under = re.match(r"\s*\(\s*([A-Za-z][A-Za-z&/ \-]{3,60}?)\s*\)\s*(?:from\b|at\b|,|$)", clean(window[1]), re.IGNORECASE)
        if under and not real_institution(under.group(1)) and not re.search(r"\b(?:percentage|cgpa|gpa|marks|grade|pursuing|distance|regular|full|part|time)\b", under.group(1), re.IGNORECASE):
            return clean(under.group(1))                            # 'MBA' / '(Sales and Marketing) From Indain Institue of Management'
        frag = clean(window[1])
        k = 2
        while k < min(len(window), 4) and re.search(r"(?:\band|&|\bof|\bin)\s*$", frag, re.IGNORECASE) or (
                k < min(len(window), 4) and re.fullmatch(r"[A-Z][A-Za-z&. ]{2,25}", clean(window[k])) and not YEAR_RE.search(window[k]) and len(frag.split()) <= 3
                and not real_institution(window[k]) and not DEGREE_RE.match(clean(window[k]))):
            frag += " " + clean(window[k])                           # a field wrapped over lines: 'Electronics and' / 'Telecommunication' / 'Engineering'
            k += 1
            if k >= min(len(window), 4):
                break
        next_text = re.match(r"^\(?\s*([A-Za-z&. ]{2,60}?)\s*\)?[,.]?$", frag)
        if next_text and not real_institution(next_text.group(1)):
            got = canon_field(next_text.group(1))
            if got:
                return got
    written = re.match(r"\s*(?:of\s+\w+\s+)?in\s+([A-Za-z][A-Za-z &/]{3,50}?)\s*(?:\(|,|$)", (rest[match.end():] if match else rest), re.IGNORECASE)
    if written and 1 <= len(written.group(1).split()) <= 5 and not real_institution(written.group(1)) and not DEGREE_RE.search(written.group(1)):
        return clean(written.group(1)).title()                      # 'BACHELOR OF COMMERCE in COMPUTER APPLICATIONS (B. Com C.A)': the field is written after 'in'
    return ""


def find_labelled_education(lines):
    """Education written as labelled lines: 'University: X' / 'Stream: Bachelor's in Y' / 'Duration: 2011-2014' / 'Percentage: 57%' (one block per degree)."""
    texts = []
    for line in lines:                                                 # '-Course: B.E. and Branch: Computer Engineering' / '66.87% [DISTINCTION] and Duration of the course: ...' : two labels on a line
        texts += [item for item in re.split(r"\s*(?:\[[^\]]*\]\s*)?\band\s+(?=(?:branch|specialization|specialisation|stream|duration|percentage|university|college|institute)\b[^:]{0,25}:)", clean(line.text), flags=re.IGNORECASE) if item]
    if sum(1 for text in texts if EDU_LABELS["inst"].match(text) or EDU_LABELS["degree"].match(text)) < 2:
        return []
    entries, current_job = [], None
    for text in texts:
        key = match = None
        for k, pattern in EDU_LABELS.items():
            second_match = pattern.match(text)
            if second_match:
                key, match = k, second_match
                break
        if not key:
            continue
        value = clean(match.group(1)).strip(" .")
        if key == "inst" and current_job is not None and current_job["institution"] and not (current_job["end_year"] or current_job["score"]):
            continue                                                # 'College: X' then 'University: Y' : the same degree, the college is the institute
        if current_job is None or (key == "inst" and current_job["institution"]) or (key == "degree" and current_job["degree"]):
            current_job = {"degree": "", "institution": None, "institution_labelled": False, "year": None, "start_year": None, "end_year": None, "score": None, "field": ""}
            entries.append(current_job)
        if key == "inst":
            current_job["institution"], current_job["institution_labelled"] = value, True
        elif key == "degree":
            current_job["degree"] = value
        elif key == "field":
            current_job["field"] = canon_field(value)
        elif key == "years":
            current_job["start_year"], current_job["end_year"] = edu_years([value])
            current_job["year"] = current_job["end_year"]
        elif key == "score":
            current_job["score"] = value
    out = []
    for entry in entries:
        if not (DEGREE_RE.search(entry["degree"]) or entry["institution"]):
            continue
        entry["field"] = entry["field"] or edu_field(entry["degree"], [])
        out.append(entry)
    return out


def find_education(sections, all_lines, use_spacy=False):
    """The Education section when it holds degrees; when the section found is empty of them (a heading placed far from its table), every line of the resume."""
    return find_education_with_labels(sections, all_lines, use_spacy)


def find_education_with_labels(sections, all_lines, use_spacy=False):
    """Education = degree + institution + year (+ field of study, score). Look in the Education section; if the resume has none, scan every line for degree
    words. The lines of one entry are the degree line and the lines below it up to the next degree:
        degree       the degree words and up to 6 words after them ('B.Tech in Computer Science')
        field        in brackets, after a dash, 'in X', or alone on the next line ('(E&TC)'), as a degree on two lines ('Post Graduate Diploma in' / 'Data Science')
        institution  the part of the line after 'from' / ',' / '-' that has a word like University / College / School, else the nearest such line, else the
                     short capitalised line right under the degree ('MSIT Jnakpuri')
        years        a range or a year in those lines;  score: '8.5 CGPA', '74.2%'"""
    lines = sections.get("education") or []
    labelled = find_labelled_education(lines)
    if labelled and any(entry["degree"] for entry in labelled):             # 'Board/University- WBUT' labels alone (the degrees are on unlabelled lines) are not an education table
        plain = find_education_in_lines([line for line in lines if not any(pattern.match(clean(line.text)) for pattern in EDU_LABELS.values())], sections, use_spacy)
        have = {canon_degree(entry["degree"]) or norm_key(entry["degree"]) for entry in labelled}
        return (labelled + [entry for entry in plain if (canon_degree(entry["degree"]) or norm_key(entry["degree"])) not in have])
    return find_education_in_lines(lines, sections, use_spacy)


def is_school_name(text):
    """'PRG Higher Secondary School' : the 'secondary' there belongs to the name of a school, it is not a qualification."""
    degree = DEGREE_RE.search(text)
    if not degree:
        return False
    has_name_before = bool(re.search(r"[A-Za-z]", text[:degree.start()]))
    is_school_level = bool(re.match(r"(?:higher |senior |secondary |sr\.? )?(?:secondary|sec)\b", degree.group(0), re.IGNORECASE))
    return has_name_before and is_school_level


def find_education_in_lines(lines, sections, use_spacy=False):
    texts = []
    for line in lines:                                                 # an institution name the PDF wrapped: 'Marathwada Mitra Mandal College Of' / 'Engineering'
        if texts and INSTITUTION_RE.search(texts[-1]) and re.search(r"(?:\b(?:of|and|for)|&)\s*$", texts[-1], re.IGNORECASE) and not DEGREE_RE.match(line.text) and not YEAR_RE.search(line.text):
            texts[-1] = clean(texts[-1] + " " + line.text)
        else:
            texts.append(line.text)
    entries, seen, consumed, used_inst = [], set(), set(), set()
    for i, text in enumerate(texts):
        if i in consumed:
            continue
        match = DEGREE_RE.search(text)
        if match and is_school_name(text):
            continue                                                # 'PRG Higher Secondary School,' is the name of a school, not a qualification
        if not match or (PROSE_START_RE.match(text) and len(text.split()) > 12):
            continue
        if re.match(r"^\W*\d+(?:st|nd|rd|th)\s+of\s+(?!the\b|std\b|standard\b)[A-Z]", text) and not re.search(r"board|school|cbse|icse|ssc|class", text, re.IGNORECASE):
            continue                                                # '10th of Ramadan' is a city, not a '10th' certificate
        if match.group(0).rstrip(".") in AMBIGUOUS_DEGREE and (not AMBIGUOUS_OK.match(text[match.end():]) or (
                len(text.split()) > 9 and not INSTITUTION_RE.search(text) and not YEAR_RE.search(text))):
            continue                                                # 'ME Thailand and Power Line Connection between ...'
        stop = min(len(texts), i + 6)
        next_text = next((k for k in range(i + 1, stop) if DEGREE_RE.search(texts[k]) and not is_school_name(texts[k]) and not PROSE_START_RE.match(texts[k])), stop)
        window = texts[i:next_text]
        rest = text[match.start():]
        degree_text = re.split(r"[,|;(:\[]|\s/\s|\s{2,}|\s[-–—]\s|\b(?:from|at|with|in the year|in\s+(?:19|20)\d{2})\b|\b(?:" + MONTHS + r"\s*)?(?:19|20)\d{2}\b",
                               rest, flags=re.IGNORECASE)[0]
        institution = None
        inst_above = None
        year_between = None
        words = degree_text.split()                                 # 'B.Sc. ACHARYA NAGARJUNA UNIVERSITY' / 'B.Tech in CSE in Andhra University'
        key_at = next((k for k, word in enumerate(words) if k > 0 and INSTITUTION_RE.search(word)), None)
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
        degree = re.sub(r"\s+(?:passed|completed|complete|cleared)$", "", degree, flags=re.IGNORECASE)          # 'Intermediate Passed' / '10th completed' is the degree, the rest only says it is done
        if len(degree) < 2 or (INSTITUTION_RE.fullmatch(degree) and not DEGREE_RE.fullmatch(degree)):
            continue
        wrap_field = ""                                             # 'Post Graduate Diploma in' / 'Data Science'
        if i + 1 < len(texts) and re.search(r"\b(?:in|of|and|&)\s*:?\s*$", text.strip(), re.IGNORECASE) and (i + 1) not in consumed:
            nline = clean(texts[i + 1]).strip(" .,()")
            if nline and len(nline.split()) <= 6 and not real_institution(nline) and not DEGREE_RE.search(nline) and not YEAR_RE.search(nline) and canon_field(nline):
                wrap_field = canon_field(nline)
                degree = clean(degree + " " + nline)
                consumed.add(i + 1)
                window = [text] + texts[i + 2:next_text]
        # 1. the institution may be on the same line: 'B.Tech from Avanthi College of Engineering, Vizag in 2014'
        tail = re.sub(DATE_RANGE_RE, " ", rest[len(" ".join(degree_text.split())):] if institution else rest[len(degree_text):])
        tail = re.sub(r"\([^)]*\)", " ", tail)
        tail = re.sub(r"\b(?:" + MONTHS + r"\s*)?(?:19|20)\d{2}\b", " ", tail, flags=re.IGNORECASE)
        segments = [re.sub(r"^(?:in|at|from|of)\s+|\s+(?:in|of|at)$", "", clean(re.split(r"\bwith\b|\(|\d{2}(?:\.\d+)?\s?%|\bcgpa\b", seg, flags=re.IGNORECASE)[0]).strip(" .,-:[]/"), flags=re.IGNORECASE)
                    for seg in re.split(r"\bfrom\b|\s[-–—|]\s|[:|]|\s/\s|\]|\s{2,}|,\s*(?=[A-Z])", tail, flags=re.IGNORECASE)]
        for segment in ([] if institution else segments):
            if real_institution(segment):
                institution = segment
                break
        if not institution:                                         # '10th from UP Board in 2013' / 'BCA from MMHAPU PATNA in 2018' / '[Nashik Board]'
            fm = re.search(r"\bfrom\s+(?:M/s\s+)?([A-Z][^,()\[\]]*?)(?=\s+in\s+(?:19|20)\d{2}|\s+with\b|\s*[,(\[]|\s+(?:19|20)\d{2}|\s*$)", rest)
            if fm:
                cand = re.sub(r"\s+(?:in|of|at)$", "", clean(fm.group(1)).strip(" .,-:"))
                if 2 <= len(cand) <= 60 and len(cand.split()) <= 6 and not DEGREE_RE.match(cand) and not YEAR_RE.search(cand) \
                        and not (set(word.lower().strip(".") for word in cand.split()) <= INDIAN_PLACES) and not SCORE_RE.search(cand):
                    institution = cand
        if not institution:                                         # 'B.E. (CSE) | RGPV, Bhopal | 2016': an acronym is the institution
            for segment in segments:
                if re.fullmatch(r"[A-Z]{3,10}", segment) and segment not in ACRONYM_SKIP and not DEGREE_RE.fullmatch(segment):
                    institution = segment
                    break
        # 2. otherwise the nearest line that has an institution word and no degree
        if not institution:
            for j in (i - 1, i - 2, i + 1, i + 2, i + 3, i + 4):           # a table: 'MBA' / 'Graphic Era Hill' / 'University' / 'Graphic Era Hill University, Bhimtal' / '2021'
                if j > i + 2 and (j >= next_text or YEAR_RE.search(texts[j - 1])):
                    break
                if 0 <= j < len(texts) and j not in used_inst and real_institution(texts[j]) and (not DEGREE_RE.search(texts[j]) or is_school_name(texts[j])):
                    used_inst.add(j)
                    if j == i - 1:
                        inst_above = j
                    elif j == i - 2 and re.fullmatch(r"\W*(?:19|20)\d{2}\W*", clean(texts[i - 1])):
                        inst_above, year_between = j, texts[i - 1]          # 'Government College' / '2014' / 'B.E(I.T)' / '62.48%': the year above the degree is its year
                    institution = clean(re.split(r"\s[-–—|]\s|\b(?:19|20)\d{2}\b|\(", re.sub(r"^\s*\([^)]*\)\s*(?:from\s+|at\s+)?", "", texts[j], flags=re.IGNORECASE))[0]).strip(" .,-:")[:90] or None
                    break
        if not institution:                                         # 3. the short capitalised line right under the degree ('MSIT Jnakpuri')
            for part_text in window[1:3]:
                text_part = clean(part_text).strip(" .,-:|")
                if not text_part or dates_of_text(text_part)[0] or SCORE_RE.search(text_part) or DEGREE_RE.search(text_part) or re.match(r"^(?:cgpa|gpa|score|percentage|marks)\b", text_part, re.IGNORECASE):
                    continue
                if 1 <= len(text_part.split()) <= 8 and not YEAR_RE.search(text_part) and not canon_field(text_part) \
                        and (CORE_INSTITUTION.search(text_part) or INSTITUTION_RE.search(text_part) or all(word[:1].isupper() or word.lower() in ("of", "and", "for", "the", "&") for word in text_part.split())):
                    institution = text_part
                    break
        if not institution and use_spacy:                           # optional spaCy layer (tested: mostly noise, so it is off by default)
            for entity in lookups().nlp()(" . ".join(texts[max(0, i - 1):i + 3])).ents:
                if entity.label_ in ("ORG", "FAC") and 2 < len(entity.text) <= 80 and not DEGREE_RE.fullmatch(entity.text.strip()):
                    institution = clean(entity.text)
                    break
        if institution:
            institution = re.split(r"\s+(?:affiliated|approved|recogni[sz]ed|accredited)\b", institution, flags=re.IGNORECASE)[0].strip(" .,-:")
        if institution and norm_key(institution) in norm_key(degree):
            institution = None                                      # the 'institution' is just the degree text again
        if inst_above is not None and not wrap_field:               # 'United Business Institutes - 2012' / 'Masters of Business Administration' / 'Mumbai University - 2009' / 'Bachelors ...': the institution and year above belong to this degree
            cut = next((k for k in range(1, len(window)) if real_institution(window[k]) and not DEGREE_RE.search(window[k])), len(window))
            window = [texts[inst_above]] + ([year_between] if year_between else []) + window[:cut]
        start_year, end_year = edu_years(window if not wrap_field else [text] + texts[i + 2:next_text])
        if not sections.get("education") and not institution and not end_year:
            continue                                                # no Education section: a bare degree word alone is not enough
        key = (norm_key(degree), norm_key(institution or ""))
        if key in seen:
            continue
        seen.add(key)
        score = SCORE_RE.search(" ".join(window[:4]))
        entries.append({"degree": degree, "institution": institution, "year": end_year, "start_year": start_year, "end_year": end_year,
                        "score": clean(score.group(0)) if score else None, "field": wrap_field or edu_field(text[match.start():], window)})
    merged = []
    for entry in entries:                                               # 'GRADUATION' (college, years, score) then 'BE' + 'Electronics and Telecommunication Engineering'
        prev = merged[-1] if merged else None
        if prev and norm_key(prev["degree"]) in ("graduationdegree", "postgraduationdegree", "graduation", "postgraduation") and not prev["institution"] and entry["institution"] \
                and not entry["end_year"] and DEGREE_RE.search(entry["degree"]):
            prev["degree"], prev["institution"] = entry["degree"], entry["institution"]       # 'Graduation Degree [2012-2016]' / 'Bachelor of Engineering in Mechanical' / 'MRK Institute ...': one entry
            prev["field"] = prev.get("field") or entry.get("field")
            continue
        prev = merged[-1] if merged else None
        if prev and not (entry["institution"] and norm_key(entry["institution"]) not in norm_key(entry.get("field") or "")) and not entry["end_year"] and prev["institution"] and (
                norm_key(prev["degree"]) in ("graduation", "graduate", "diploma", "postgraduation", "postgraduate") or norm_key(prev["degree"]) == norm_key(entry["degree"])):
            if norm_key(prev["degree"]) in ("graduation", "graduate", "postgraduation", "postgraduate"):
                prev["degree"] = entry["degree"]
            prev["field"] = prev.get("field") or entry.get("field")
            continue
        if prev and not entry["institution"] and not entry["end_year"] and canon_degree(entry["degree"]) and canon_degree(entry["degree"]) == canon_degree(prev["degree"]):
            prev["field"] = prev.get("field") or entry.get("field") or edu_field(entry["degree"], [])          # the same degree written twice: the second line adds the field
            continue
        merged.append(entry)
    entries = merged
    # tables read column by column leave the institutions far from their degrees: give the unused ones, in order, to the entries without
    if any(not entry["institution"] for entry in entries):
        taken = {norm_key(entry["institution"]) for entry in entries if entry["institution"]}
        pool = []
        for part_text in texts:
            c = clean(re.split(r"\s[-–—|]\s|\(|\b(?:19|20)\d{2}\b", part_text)[0]).strip(" .,-:")
            if real_institution(c) and not DEGREE_RE.search(c) and norm_key(c) not in taken and norm_key(c) not in {norm_key(item) for item in pool}:
                pool.append(c)
        for entry in entries:
            if not entry["institution"] and pool and sections.get("education"):
                entry["institution"] = pool.pop(0)
    return entries


def enrich_from_header(education, top):
    """'Vishal Thadka B.E (Information Technology)' at the top: the field of study of the Bachelor of Engineering that has none."""
    for line in top[:6]:
        for match in re.finditer(r"\b(B\.?\s?E\.?|B\.?\s?Tech\.?|M\.?\s?Tech\.?|M\.?\s?E\.?|MBA|MCA|BCA|BBA|B\.?\s?Sc\.?|M\.?\s?Sc\.?|B\.?\s?Com\.?)\s*\(([^)]{3,50})\)", line.text):
            degree, field = canon_degree(match.group(1)), canon_field(match.group(2))
            if degree and field:
                for entry in education:
                    if canon_degree(entry["degree"]) == degree and not entry.get("field"):
                        entry["field"] = field
    return education


MONTH_FULL = {"jan": "january", "feb": "february", "mar": "march", "apr": "april", "may": "may", "jun": "june", "jul": "july", "aug": "august",
              "sep": "september", "oct": "october", "nov": "november", "dec": "december"}
MONTH_YEAR_RE = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s*,?\s*(?:19|20)\d{2}"


def full_month_name(match):
    """'feburary' (any spelling that starts with the 3 letters) -> 'february'."""
    return MONTH_FULL[match.group(1)]


def parse_date(text):
    """'Jan 2020' / '03/2021' / '2019' -> (year, month) or None; 'Present' -> today."""
    text = text.strip().lower()
    text = re.sub(r"([a-z]{3,9})\.?\s*['’‘]\s*(\d{2})$", r"\1 20\2", text)                     # Jan'20 -> Jan 2020
    text = re.sub(r"^((?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec))(\d{2})$", r"\1 20\2", text)       # Jul11 -> Jul 2011
    text = re.sub(r"^((?:19|20)\d{2})[\s,]*([a-z]{3,9})\.?$", r"\2 \1", text)                   # 2018 December -> December 2018
    if re.match(r"present|current|up\s?to|till|now|ongoing|date|today", text):
        now = time.localtime()
        return now.tm_year, now.tm_mon
    match = re.match(r"(\d{1,2})[/.\-]((?:19|20)\d{2})$", text)
    if match:
        return int(match.group(2)), min(12, int(match.group(1)))
    text = re.sub(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b", full_month_name, text)          # 'feburary' -> 'february'
    doc = dateparser.parse(text, settings={"PREFER_DAY_OF_MONTH": "first", "REQUIRE_PARTS": ["year"]})
    return (doc.year, doc.month) if doc else None


def months_between(start, end):
    return max(0, (end[0] - start[0]) * 12 + end[1] - start[1] + 1)


def shorten(text, limit=3000):
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
    word_lists = lookups()
    nice = word_lists.skills[key]
    return key not in word_lists.whole_item_only and (" " in nice.strip() or key in word_lists.known_tech or re.search(r"[A-Z].*[A-Z]|\d|[#+.]", nice[1:]) is not None)


def prepare_date_lines(texts):
    """Text of the lines with the PDF's lost dash (U+FFFD) put back, brackets removed, and two lines that hold only one date each
    ('October 2018' / 'January 2020') joined into one range."""
    out = [re.sub(r"[\[\]]", " ", text.replace("\ufffd", "-")) for text in texts]
    out = [re.sub(r"\b(" + MONTHS + r")\s+(\d{2})\b(?=\s*(?:[-\u2013\u2014]|to\b|till\b|until\b|up\s*to\b)\s*(?:present|current|now|date|today|ongoing))", r"\1 20\2", text, flags=re.IGNORECASE)
           for text in out]                                            # 'FEB 24 - PRESENT' / 'June 24 to present': the two digits are the year 2024
    out = [re.sub(r"^(\W*)(" + MONTHS + r")[\s']*((?:0\d|1\d|2\d|3\d))\b\s*(?:[-–—]|\bto\b)\s*(" + MONTHS + r")[\s']*((?:0\d|1\d|2\d|3\d))\b(\W*)$", r"\1\2 20\3 - \4 20\5\6", text, flags=re.IGNORECASE)
           for text in out]                                            # a line that is only 'Nov 23 - jan24' / "Jun'22 to Aug'22": two digits on both sides are the years 2023 - 2024
    out = [re.sub(r"(" + MONTH_YEAR_RE + r")\s+(?=" + MONTH_YEAR_RE + r")", r"\1 - ", text, flags=re.IGNORECASE) for text in out]       # 'October 2020 December 2020'
    
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


def job_line_head(text):
    """The text of a line when it can be a title / company line of a job, else ''."""
    text = text.strip(" .-:|*>")
    if not text or len(text.split()) > 9 or PROSE_START_RE.match(text) or re.search(r"@|https?|\d{4}", text) or heading_section(text) \
            or text[:1].islower() or text.endswith((",", ";")) or LABEL_COMPANY_RE.match(text) or LABEL_TITLE_RE.match(text) or SKIP_LABEL_RE.match(text) \
            or re.fullmatch(r"(?:company|employer|designation|position|role|fresher|experience|total experience)\W*", text, re.IGNORECASE):
        return ""
    return text


def is_title_text(text, *, word_lists):
    """True when the text is a job title."""
    return norm_key(text) in word_lists.titles or (TITLE_HEAD.search(text) is not None and not COMPANY_SUFFIX_RE.search(text))


def is_company_text(text):
    """True when the text has a company word (Pvt, Ltd, Technologies ...)."""
    return COMPANY_SUFFIX_RE.search(text) is not None


def is_company_like(text, *, word_lists):
    """True when the text can be a company: it has a company word, or it is a short capitalised phrase that is no title, skill or place."""
    return bool(text) and not is_title_text(text, word_lists=word_lists) and (is_company_text(text) or (len(text.split()) <= 5 and text[:1].isupper() and norm_key(text) not in word_lists.skills
                                                              and not ({word.lower().strip(",") for word in text.split()} & INDIAN_PLACES)))


def find_undated_entries(lines):
    """A resume that gives no dates: a title line next to a company line (either order), or 'Title at Company' / 'Title - Company'."""
    word_lists = lookups()
    texts = [clean(line.text) for line in lines]




    entries, used = [], set()
    for i, raw in enumerate(texts):                                 # 'Company: X' / 'Designation: Y' written as labels
        company_match = LABEL_COMPANY_RE.match(raw)
        if company_match and clean(company_match.group(1)):
            title = None
            for j in range(i + 1, min(len(texts), i + 5)):
                title_match = LABEL_TITLE_RE.match(texts[j])
                if title_match:
                    title = clean(title_match.group(1))
                    used.add(j)
                    break
            entries.append({"title": title, "company": clean(company_match.group(1))[:80], "company_source": "label", "start": "", "end": "", "months": 0, "current": False,
                            "description": "", "skills": []})
            used.add(i)
    for i, raw in enumerate(texts):
        text = job_line_head(raw)
        if not text or i in used:
            continue
        parts = [part.strip() for part in re.split(r"\s*[|@,]\s*|\s+[-–—]\s+|\s+at\s+", text) if part.strip()]
        title = company = None
        if is_title_text(text, word_lists=word_lists) or any(is_title_text(part, word_lists=word_lists) for part in parts):
            title = next((part for part in parts if is_title_text(part, word_lists=word_lists)), text)
            company = next((part for part in parts if part != title and is_company_text(part)), None)
            if company is None and len(parts) > 1:
                company = next((part for part in parts if part != title and is_company_like(part, word_lists=word_lists)), None)
            if company is None:
                for j in (i + 1, i - 1):
                    if 0 <= j < len(texts) and j not in used and is_company_like(job_line_head(texts[j]), word_lists=word_lists):
                        company = job_line_head(texts[j])
                        used.add(j)
                        break
        elif is_company_text(text):
            company = text
            if i + 1 < len(texts) and is_title_text(job_line_head(texts[i + 1]), word_lists=word_lists):
                title = job_line_head(texts[i + 1])
                used.add(i + 1)
        else:
            continue
        if title and company and norm_key(title) == norm_key(company):
            company = None
        company_source = ("suffix" if is_company_text(company) else "caps") if company else ""
        stop = next((k for k in range(i + 1, len(texts)) if k not in used and job_line_head(texts[k]) and (is_title_text(job_line_head(texts[k]), word_lists=word_lists) or is_company_text(job_line_head(texts[k])))), len(texts))
        body = " ".join(texts[i + 1:stop])
        entries.append({"title": title, "company": company, "company_source": company_source, "start": "", "end": "", "months": 0, "current": False,
                        "description": shorten(body) if not body.startswith(company or "\0") else "",
                        "skills": list(dict.fromkeys(word_lists.skills[k].lower() for k, _ in find_terms(body, word_lists.skills) if is_strong_skill(k)))[:15]})
        used.add(i)
    seen, unique = set(), []
    for entry in entries:
        k = (norm_key(entry["company"] or ""), norm_key(entry["title"] or ""))
        if k not in seen and (entry["title"] or entry["company"]):
            seen.add(k)
            unique.append(entry)
    return unique


STATE_ABBREV_RE = re.compile(r"^\(?(?:[A-Z]{1,3}|[A-Z](?:\.[A-Z]){1,2}\.?)\)?$")
LEGAL_END_RE = re.compile(r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation|co)\b\.?\s*$", re.IGNORECASE)
JOB_LABEL_RE = re.compile(r"^\W*(?:responsibilit|description|duties|key|role|roles|project|client|environment|technolog|designation|position|employer|company|"
                          r"organi[sz]ation|duration|period|location|tech|tools|skills)\w*\b", re.IGNORECASE)


def is_place_text(part):
    """'Pune, MH', 'Gurgaon,Haryana', 'Mumbai': a text made of places (and at most one other word)."""
    part = clean(part).strip(" .,()-")
    if not part:
        return False
    if STATE_ABBREV_RE.match(part):
        return True
    hits = find_places(part, trusted=True)
    if not hits:
        return False
    return len(re.findall(r"[A-Za-z]+", part)) - sum(len(heading[1].split()) for heading in hits) <= 1


def is_title_part(part):
    """A job title ends with an occupation word ('Software Developer', 'Systems Engineer', 'Data Engineer Intern'), even when it holds a company word."""
    part = re.sub(r"\s*\([^)]*\)\s*$", "", clean(part)).strip(" .,:-")
    if not part or LEGAL_END_RE.search(part):
        return False
    last = part.split()[-1].lower().strip(".,)(")
    return norm_key(part) in lookups().titles or bool(TITLE_HEAD.search(last + " ")) or bool(re.search(r"\bintern(?:ship)?s?\b", part, re.IGNORECASE))


def is_company_part(part):
    part = clean(part)
    return COMPANY_SUFFIX_RE.search(part) is not None and not is_title_part(part)


def trim_name(part):
    """Strip the dashes and dots around a name, but keep a bracket pair that belongs to it: 'Ameyo acquired by (exotel)'."""
    part = part.strip(" .-:")
    while part.endswith(")") and part.count(")") > part.count("("):
        part = part[:-1].rstrip(" .-:")
    while part.startswith("(") and part.count("(") > part.count(")"):
        part = part[1:].lstrip(" .-:")
    return part


def split_header_parts(text):
    """'Accenture, Pune, MH' -> ['Accenture'];  'HR Intern - Honda Pvt Ltd, Bangalore' -> ['HR Intern', 'Honda Pvt Ltd'];  a comma between two non-places stays a split."""
    parts = []
    for chunk in re.split(r"\s*[|@]\s*|\s+[-–—]\s+|\s+at\s+|\s+as\s+(?:an?\s+)?|\s{3,}", text):
        pieces = [item.strip() for item in re.split(r"\s*,\s*", chunk) if item.strip()]
        parts += [piece for k, piece in enumerate(pieces) if not (k > 0 and is_place_text(piece))]
    return parts


SKIP_LABEL_RE = re.compile(r"^\W*(?:project|domain|client|customer|description|technolog\w*|environment|team size|role|tools?|industry|sector)\b[^:]{0,20}:", re.IGNORECASE)


def usable_header_text(text):
    """True when a line near a date can hold the title or the company of the job."""
    return bool(text) and len(text.split()) <= 12 and not PROSE_START_RE.match(text) and not DATE_RANGE_RE.search(text) and not heading_section(text) \
        and "@" not in text and not JOB_LABEL_RE.match(text) and not text.endswith(".") and not re.match(r"^\d+(?:\.\d+)?\+?\s*(?:years?|yrs?|months?)\b", text, re.IGNORECASE)


def tidy_header_text(text):
    """A header line without bullets, 'Experience:' style labels and empty brackets."""
    return re.sub(r"\(\s*\)", "", re.sub(r"^(?:experience|exp|duration|period|tenure)\s*[:\-]\s*", "", re.sub(r"^[^\w(]+", "", clean(text)), flags=re.IGNORECASE)).strip(" .-:|*>")


def is_sentence_line(raw):
    """True when the line is a sentence of a description, not a title or a company."""
    raw = raw.strip()
    return raw.endswith(".") and len(raw.split()) >= 5 and not COMPANY_SUFFIX_RE.search(raw.split()[-1])


def read_job_header(texts, i, same_line):
    """Title and company of the job whose date is on line i. Looks at the line itself (the date cut out), the line above, the line above that (a title
    written over the company) and the line below (a place). A part is a title when it ends with an occupation word, a company when it ends with a company
    word; a short capitalised line next to the title is the company even without a company word ('LineupX', 'Ameyo acquired by (exotel)')."""
    word_lists = lookups()
    sentence_job = re.match(r"^\W*(?:(?:working|worked|serving)\s+as\s+[^,]{2,40},\s*(?P<place>[^,]{2,25}?)\s+|(?:with|at|in)\s+(?P<head>.+?)\s+)as\s+(?:an?\s+)?(?P<title>[A-Za-z][\w &/.\-]{2,50}?)[\s.]*$",
                            same_line or "", re.IGNORECASE)         # 'Since Jul'16 Working as Consultant, Jaipur as Manager General' / 'Since Oct'16 with Charvik Education, Jaipur as Entrepreneur'
    said_job = re.match(r"^\W*(?:currently\s+|presently\s+|previously\s+)?(?:working|worked|serving|served|employed)\s+as\s+(?:an?\s+)?(?P<title>[A-Za-z][A-Za-z0-9 &/.\-]{2,50}?)\s+(?:at|with|in|for)\s+(?P<company>[A-Z][^,;]{2,60}?)(?:,[^,;]{2,30}?)?\s*(?:\b(?:from|since|during|w\.?e\.?f)\b.*)?$",
                        same_line or "", re.IGNORECASE)         # 'Worked as Software Engineer at Evam Technologies Pvt Ltd, Bengaluru from  to .'
    if said_job:
        return clean(said_job.group("title")), trim_name(said_job.group("company")) or None, "sentence"
    if sentence_job:
        head = sentence_job.group("head")
        company = trim_name(re.split(r"\s*,\s*", head)[0]) if head else None
        return clean(sentence_job.group("title")), (company or None), ("sentence" if company else "")


    candidates = []
    same = tidy_header_text(same_line)
    if usable_header_text(same):
        candidates.append(("same", same))
    above = i - 1
    while above >= 1 and i - above <= 3 and SKIP_LABEL_RE.match(texts[above]):
        above -= 1                                                  # 'Project: FinNova' written between the company line and its date
    prev_owned = above >= 1 and bool(DATE_RANGE_RE.search(texts[above - 1])) and not DATE_RANGE_RE.search(texts[above])       # '(1) ADGM | Sep 2025 – Now' / 'Lead | UAE' / '(2) Raya Gulf | Sep 2022 – Sep 2025': the line above is the title of the job before
    prev = tidy_header_text(texts[above]) if above >= 0 and not is_sentence_line(texts[above]) and not (prev_owned and same_line and same_line.strip(" |-–:")) else ""
    if usable_header_text(prev):
        candidates.append(("prev", prev))
        prev2 = tidy_header_text(texts[above - 1]) if above >= 1 and not is_sentence_line(texts[above - 1]) else ""
        if usable_header_text(prev2) and len(prev.split()) <= 10:
            candidates.append(("prev2", prev2))
    next_text = tidy_header_text(texts[i + 1]) if i + 1 < len(texts) and not is_sentence_line(texts[i + 1]) else ""
    if usable_header_text(next_text) and not SKIP_LABEL_RE.match(next_text) and (is_place_text(next_text) or len(next_text.split()) <= 6):
        candidates.append(("next", next_text))
    title = company = None
    source = ""
    unknown = []
    for name, text in candidates:
        for part in split_header_parts(text):
            part = re.sub(r"^(?:from|since|w\.?e\.?f\.?)\b.*?\b(?:with|at)\s+(?=[A-Z])", "", part, flags=re.IGNORECASE)
            part = trim_name(re.sub(r"^(?:\d+[.)]\s*|with\s+|at\s+|in\s+)", "", part, flags=re.IGNORECASE))
            if not part or len(part.split()) > 8 or re.search(r"[@\d]{4,}", part) or is_place_text(part):
                continue
            glued = None
            if title is None and company is None and 3 <= len(part.split()) <= 8:
                words_ = part.split()
                for k in range(len(words_) - 1, 0, -1):
                    if TITLE_HEAD.search(words_[k - 1] + " ") and words_[k][:1].isupper() and not TITLE_HEAD.search(words_[k] + " ") and norm_key(words_[k]) not in word_lists.titles:
                        glued = (" ".join(words_[:k]), " ".join(words_[k:]))
                        break
                if glued and is_title_part(glued[0]) and not is_place_text(glued[1]):
                    title, company, source = glued[0], glued[1], "suffix"
                    continue
            if title is None and is_title_part(part):
                title = part
            elif company is None and is_company_part(part):
                company, source = part, "suffix"
            else:
                unknown.append((name, part))
    if company is None:
        for name, part in unknown:
            if (name in ("same", "prev", "next") or title is not None) and not is_title_part(part) and 1 <= len(part.split()) <= 6 \
                    and (part[:1].isupper() or part[:1].isdigit() or part.startswith("(")) and (norm_key(part) not in word_lists.skills or name == "prev") and norm_key(part) not in word_lists.titles \
                    and not heading_section(part) and not re.search(r"[:;]$", part) and not ({word.lower().strip(",.") for word in part.split()} & NAME_BAD_WORDS) \
                    and not re.search(r"\b(?:responsib|developed|worked|managed|handled|worked)\w*", part, re.IGNORECASE):
                company, source = part, "structure"
                break
    if title is None:
        for name, part in unknown:
            if part != company and 1 <= len(part.split()) <= 8 and part[:1].isupper() and (name == "prev2" or norm_key(part) not in word_lists.skills) \
                    and not heading_section(part):
                title = part                                        # 'Data Engineering' written over the company is the role, though it is also a skill
                break
    company = re.sub(r"^\(\d+\)\s*", "", company) if company else company
    company = re.sub(r"\s+(?:remote|hybrid|on-?site)$", "", company, flags=re.IGNORECASE) if company else company       # '(3) Oracle': the number is the list position
    return title, company, source


LABEL_DURATION_RE = re.compile(r"^\W*(?:duration|period|tenure|working period|dates?|experience)\s*[:\-–]\s*(.+)$", re.IGNORECASE)
LABEL_STATUS_RE = re.compile(r"^\W*status\s*[:\-–]\s*(?P<title>[A-Za-z][^()\[\]]*?)?\s*[(\[]?(?P<rest>.*?)[)\]]?\s*$", re.IGNORECASE)
GROWTH_RE = re.compile(r"^\W*(?P<start>" + _DATE_PART + r")\s*(?:-|–|—|to)\s*(?P<end>" + _END_WORDS + r"|" + _DATE_PART + r")\s+as\s+(?:an?\s+)?(?P<title>.+?)[\s.]*$", re.IGNORECASE)
SINCE_AS_RE = re.compile(r"^\W*(?:since|from)\s+(?P<start>" + _DATE_PART + r")\s+(?:(?:with|at|in)\s+(?P<company>.+?)\s+)?as\s+(?:an?\s+)?(?P<title>.+?)[\s.]*$", re.IGNORECASE)
TABLE_COLUMNS = {"srno": "n", "sno": "n", "no": "n", "slno": "n", "number": "n", "organization": "company", "organisation": "company", "company": "company",
                 "companyname": "company", "employer": "company", "location": "place", "place": "place", "duration": "dates", "period": "dates", "tenure": "dates",
                 "role": "title", "designation": "title", "position": "title", "title": "title", "jobtitle": "title"}


def new_job(title=None, company=None, source="", start="", end="", months=0, current=False, description=""):
    return {"title": title, "company": company, "company_source": source, "start": start, "end": end, "months": months, "current": current,
            "description": description, "skills": []}


def dates_of_text(text):
    """(start 'YYYY-MM', end 'YYYY-MM', months, current) of a text that holds a date range or 'Since Jan 2020'; else empty values."""
    text = text or ""
    match = DATE_RANGE_RE.search(text)
    if match:
        start, end = parse_date(match.group("start")), parse_date(match.group("end"))
        if start and end:
            if start > end:
                start, end = end, start
            return "%04d-%02d" % start, "%04d-%02d" % end, months_between(start, end), bool(re.match(r"present|current|up\s?to|till|now|ongoing|date|today", match.group("end").strip(), re.IGNORECASE))
    sm = SINCE_RE.search(text)
    if sm and parse_date(sm.group(1)):
        start, end = parse_date(sm.group(1)), parse_date("present")
        return "%04d-%02d" % start, "%04d-%02d" % end, months_between(start, end), True
    return "", "", 0, False


TIMELINE_DATE_RE = re.compile(r"^(?P<m1>[A-Za-z]{3,9})[\s'-]*(?P<y1>(?:19|20)?\d{2})\s*(?:[-–—]|to|till|until)?\s*(?:(?P<m2>[A-Za-z]{3,9})[\s'-]*(?P<y2>(?:19|20)?\d{2})|(?P<now>date|present|now|current|today))?$", re.IGNORECASE)


def is_month_word(word):
    """True for 'May', 'Sept', 'march' ... (a word whose first 3 letters are a month)."""
    return parse_date(word[:3] + " 2000") is not None


def full_year(year_text):
    """'15' -> 2015, '98' -> 1998, '2015' -> 2015."""
    year = int(year_text)
    return year if year >= 100 else (2000 + year if year < 60 else 1900 + year)


def find_timeline_entries(lines):
    """A career-timeline table that a PDF read column by column: the companies come first, then the roles, and the date ranges in the same order
    ('MONDELEZ INDIA FOODS PVT LTD' / 'S.S.E - Nasik' / 'May 2015 till date' / 'EMAMI LTD.' / ...). When there are as many companies as date ranges, the n-th company has the n-th range."""
    texts = [clean(line.text) for line in lines]
    dates = []
    for text in texts:
        match = TIMELINE_DATE_RE.match(text.strip())
        if match and is_month_word(match.group("m1")) and (match.group("now") or (match.group("m2") and is_month_word(match.group("m2")))):
            start = parse_date("%s %d" % (match.group("m1")[:3], full_year(match.group("y1"))))
            end = parse_date("present") if match.group("now") else parse_date("%s %d" % (match.group("m2")[:3], full_year(match.group("y2"))))
            if start and end:
                dates.append((start, end, bool(match.group("now"))))
    companies = [k for k, text in enumerate(texts) if COMPANY_SUFFIX_RE.search(text) and 1 <= len(text.split()) <= 8 and not TIMELINE_DATE_RE.match(text) and not DEGREE_RE.search(text)
                 and not INSTITUTION_RE.search(text)]
    if len(companies) < 2 or len(companies) != len(dates):
        return []
    entries = []
    for n, k in enumerate(companies):
        stop = companies[n + 1] if n + 1 < len(companies) else len(texts)
        role = next((text for text in texts[k + 1:stop] if not TIMELINE_DATE_RE.match(text) and 1 <= len(text.split()) <= 6 and re.search(r"[A-Za-z]", text)), "")
        title = clean(re.split(r"\s[-–—]\s*|,", role)[0]).strip(" .-") if role else None
        start, end, now = dates[n]
        entries.append(new_job(title or None, clean(texts[k]).strip(" .,-"), "structure", "%04d-%02d" % start, "%04d-%02d" % end, months_between(start, end), now))
    return entries


def close_labelled_job(*, word_lists, cur, desc):
    """Writes the description and skills collected for the job that is being read."""
    if cur is not None:
        cur["description"] = shorten(" ".join(desc))
        cur["skills"] = list(dict.fromkeys(word_lists.skills[k].lower() for k, _ in find_terms(" ".join(desc), word_lists.skills) if is_strong_skill(k)))


def find_labelled_entries(lines):
    """Jobs written as labelled lines: 'Employer: X' / 'Designation: Y' / 'Duration: May 2019 - Mar 2021' / 'Client Name: Z' (one block per job)."""
    word_lists = lookups()
    texts = [clean(line.text) for line in lines]
    if sum(1 for text in texts if LABEL_COMPANY_RE.match(text)) < 2 and not (any(LABEL_COMPANY_RE.match(text) for text in texts) and any(LABEL_DURATION_RE.match(text) for text in texts)):
        return []
    entries, current_job, description_lines = [], None, []

    for text in texts:
        company_match, title_match, duration_match = LABEL_COMPANY_RE.match(text), LABEL_TITLE_RE.match(text), LABEL_DURATION_RE.match(text)
        if company_match or (title_match and current_job is not None and current_job["title"]):
            close_labelled_job(word_lists=word_lists, cur=current_job, desc=description_lines)
            current_job = new_job(company=clean(company_match.group(1)).strip(" -–:.") if company_match else None, source="label" if company_match else "")
            if title_match:
                current_job["title"] = clean(title_match.group(1))
            entries.append(current_job)
            description_lines = []
            continue
        if current_job is None:
            continue
        if title_match and not current_job["title"]:
            current_job["title"] = clean(title_match.group(1))
        elif duration_match:
            current_job["start"], current_job["end"], current_job["months"], current_job["current"] = dates_of_text(duration_match.group(1))
        elif LABEL_STATUS_RE.match(text) and dates_of_text(text)[0] and current_job is not None:
            status_match = LABEL_STATUS_RE.match(text)
            current_job["start"], current_job["end"], current_job["months"], current_job["current"] = dates_of_text(text)
            if not current_job["title"] and status_match.group("title") and clean(status_match.group("title")).strip(" .-"):
                current_job["title"] = clean(status_match.group("title")).strip(" .-")
        elif current_job is not None and not current_job["start"] and len(text.split()) <= 8 and (dates_of_text(text)[0] or current_date_span(prepare_date_lines([text])[0])):
            found = dates_of_text(text)                                    # '(Feb 2022 - Present)' on its own line under 'Company: X' belongs to that job
            if found[0]:
                current_job["start"], current_job["end"], current_job["months"], current_job["current"] = found
            else:
                start, _span = current_date_span(prepare_date_lines([text])[0])
                current_job["start"], current_job["end"], current_job["months"], current_job["current"] = "%04d-%02d" % start, "", 0, True
        elif text and not re.match(r"^\W*(?:client name|location|responsibilit\w*|expertise)\b", text, re.IGNORECASE):
            description_lines.append(text)
    close_labelled_job(word_lists=word_lists, cur=current_job, desc=description_lines)
    return [entry for entry in entries if entry["title"] or entry["company"]]


def find_table_entries(lines):
    """Jobs written as a table: header cells 'Sr No. / Organization / Location / Duration / Role', then one row per job (each cell is a line)."""
    texts = [clean(line.text) for line in lines]
    start = None
    for a in range(len(texts)):
        run, j = [], a
        while j < len(texts) and norm_key(texts[j]) in TABLE_COLUMNS:
            run.append(TABLE_COLUMNS[norm_key(texts[j])])
            j += 1
        if len(run) >= 3 and "company" in run and "title" in run:
            start, columns = j, run
            break
    if start is None:
        return []
    fields = [c for c in columns if c != "n"]
    rows, row = [], None
    for text in texts[start:]:
        if re.match(r"^\W*[A-Za-z ]+:\s*$", text) or heading_section(text):
            break                                                       # 'Key Responsibilities:' ends the table
        if "n" in columns and re.fullmatch(r"\d{1,2}\.?", text):
            row = []
            rows.append(row)
        elif row is not None:
            row.append(text)
        elif "n" not in columns:
            if not rows or len(rows[-1]) >= len(fields):
                rows.append([])
            rows[-1].append(text)
    entries = []
    for cells in rows:
        entry = new_job(source="structure")
        cells = [c for c in cells if c]
        if len(cells) == len(fields):
            for field, c in zip(fields, cells):
                if field == "company":
                    entry["company"] = c
                elif field == "title":
                    entry["title"] = c
                elif field == "dates":
                    entry["start"], entry["end"], entry["months"], entry["current"] = dates_of_text(c)
        else:
            for c in cells:
                if dates_of_text(c)[0]:
                    entry["start"], entry["end"], entry["months"], entry["current"] = dates_of_text(c)
                elif entry["company"] is None and not is_place_text(c):
                    entry["company"] = c
                elif is_title_part(c):
                    entry["title"] = c
        if entry["title"] or entry["company"]:
            entries.append(entry)
    return entries


def find_growth_entries(lines):
    """'Since Jul11 with Secure Meters Ltd., Udaipur as an Executive (R&D)' followed by 'Jul11-Jul12 as Graduate Engineer Trainee', 'Aug12-Aug15 as Executive
    (Level 1)' ...: one job per line, all at the company of the first line."""
    texts = [clean(line.text) for line in lines]
    if sum(1 for text in texts if GROWTH_RE.match(text)) < 2:
        return []
    company, entries = None, []
    for text in texts:
        match = SINCE_AS_RE.match(text)
        g = GROWTH_RE.match(text)
        if match:
            if match.group("company"):
                company = (split_header_parts(match.group("company")) or [match.group("company")])[0]
            start = parse_date(match.group("start"))
            if start:
                end = parse_date("present")
                entries.append(new_job(clean(match.group("title")), company, "sentence", "%04d-%02d" % start, "%04d-%02d" % end, months_between(start, end), True))
        elif g:
            start, end = parse_date(g.group("start")), parse_date(g.group("end"))
            if start and end:
                entries.append(new_job(clean(g.group("title")), company, "sentence", "%04d-%02d" % start, "%04d-%02d" % end, months_between(start, end),
                                      bool(re.match(r"present|current|till|now", g.group("end").strip(), re.IGNORECASE))))
    return entries


def find_numbered_entries(lines):
    """Jobs written as numbered blocks: '1. Motion Auto Spares' / 'Royal Enfield Authorized Distributor' / 'Sales Manager' / 'Aurangabad, Maharashtra' / 'February 2018 – up to'.
    The numbered line is the company, the dated line is the period, and the line with an occupation word is the title."""
    texts = [clean(line.text) for line in lines]
    starts = [i for i, text in enumerate(texts) if re.match(r"^\d{1,2}[.)]\s+[A-Z]", text) and len(text.split()) <= 10]
    if len(starts) < 2:
        return []
    entries = []
    for n, i in enumerate(starts):
        block = texts[i:(starts[n + 1] if n + 1 < len(starts) else len(texts))]
        date_at = next((k for k, text in enumerate(block) if DATE_RANGE_RE.search(text)), None)
        if date_at is None:
            continue
        start, end, months, now = dates_of_text(block[date_at])
        if not start:
            continue
        company = re.sub(r"^\d{1,2}[.)]\s+", "", block[0]).strip(" .:-")
        title = next((text for text in block[1:date_at] if is_title_part(text) and not is_place_text(text)), None)
        body = [text for k, text in enumerate(block[1:], 1) if k > date_at]
        entry = new_job(title, company, "structure", start, end, months, now, shorten(" ".join(body)))
        entries.append(entry)
    return entries if len(entries) >= 2 else []


def find_entries(lines, undated=False):
    """Jobs of the Experience / Internships section: labelled blocks, a table, a 'growth path', numbered blocks, else every date range."""
    for parser in (find_table_entries, find_labelled_entries, find_growth_entries, find_numbered_entries):
        try:
            found = parser(lines)
        except Exception:
            found = []
        if found:
            return found
    return find_dated_entries(lines, undated)


def find_dated_entries(lines, undated=False):
    """Experience / internship entries. Every date range ('Jan 2020 - Present', '2018-2021') starts an entry. The title
    and company come from that line and the line before it: a company has a suffix (Pvt Ltd, Technologies ...),
    a title ends with an occupation word. When no suffix is found spaCy's ORG label is tried."""
    word_lists = lookups()
    entries = []
    texts = prepare_date_lines([line.text for line in lines])
    for i, text in enumerate(texts):
        match = DATE_RANGE_RE.search(text)
        if match:
            start, end, end_word, span = parse_date(match.group("start")), parse_date(match.group("end")), match.group("end"), match.span()
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
        company_source = ""
        around = (texts[i - 1] + " " + text) if i and not re.match(r"\W*(?:currently\s+|presently\s+)?(?:working|worked|serving|employed)\b", text, re.IGNORECASE) else text      # the sentence may start on the line before the date, unless this line starts its own
        sentence = re.search(r"(?:working|worked|serving|employed|currently)\s+as\s+(?:an?\s+)?(.+?)\s+(?:in|at|with|for)\s+(.+?)(?:\s+(?:from|since|for|till)\b|\.|,|\[|$)", around, re.IGNORECASE)
        if sentence:                                                # 'Working as Associate Consultant in Cognizant from Oct 2019'
            title, company = clean(sentence.group(1)), clean(sentence.group(2))
            company_source = "sentence"
        else:
            sentence = re.search(r"(?:working|worked|employed|serving)\s+(?:in|at|with|for)\s+(.+?)\s+since\b.*?\bas\s+(?:an?\s+)?(.+?)(?:\.|,|$)", around, re.IGNORECASE)
            if sentence:                                            # 'Working in Infosys Ltd. since 14th February, 2011 as Senior Systems Engineer.'
                company, title = clean(sentence.group(1)).strip(" ."), clean(sentence.group(2))
                company_source = "sentence"
        if title is None or company is None:                        # title and company from the lines around the date
            h_title, h_company, h_source = read_job_header(texts, i, header[0])
            title = title or h_title
            if company is None and h_company:
                company, company_source = h_company, h_source
        if title is None:                                           # a title in the lines around the date
            for j in range(max(0, i - 3), min(len(texts), i + 4)):
                found = [word_lists.titles[k] for k, _ in find_terms(texts[j], word_lists.titles, max_words=6)
                         if TITLE_HEAD.search(word_lists.titles[k].split()[-1] + " ") and len(word_lists.titles[k].split()) >= 2]
                if found:
                    title = found[0]
                    break
        near_after = [texts[j] for j in range(i + 1, min(len(texts), i + 4))] + [texts[j] for j in (i - 1, i - 2) if j >= 0]       # nearest first
        label_title = next((clean(LABEL_TITLE_RE.match(part_text).group(1)) for part_text in near_after if LABEL_TITLE_RE.match(part_text)), None)
        if label_title:                                             # 'Designation: Application Development Senior Analyst' is the official title
            title = label_title
        if company is None:
            label_company = next((clean(LABEL_COMPANY_RE.match(part_text).group(1)) for part_text in near_after if LABEL_COMPANY_RE.match(part_text)), None)
            if label_company:
                company, company_source = label_company, "label"
        if company is None:                                         # last chance: spaCy organisation
            for entity in word_lists.nlp()(" . ".join(header)).ents:
                if entity.label_ == "ORG" and entity.text.strip() != title and len(entity.text.split()) <= 6 \
                        and norm_key(entity.text) not in word_lists.skills and norm_key(entity.text) not in word_lists.titles and not re.search(r"compliance|management|analysis|testing|support", entity.text, re.I):
                    company = clean(entity.text)
                    company_source = "spacy"
                    break
        if company and ((PROSE_START_RE.match(company) and not COMPANY_SUFFIX_RE.search(company)) or len(company) > 80 or len(company.split()) > 8):
            company = None                                          # 'Negotiate with ...' is a sentence, not a company
            company_source = ""
        if title and len(title) > 80:
            title = None
        current = bool(re.match(r"present|current|up\s?to|till|now|ongoing|date|today", end_word.strip(), re.IGNORECASE))
        if start > end:
            start, end = end, start
        # what the person did in this job: the lines after the date line, up to the header of the next job
        next_date = next((k for k in range(i + 1, len(texts)) if DATE_RANGE_RE.search(texts[k])), len(texts))
        begin = i + 2 if (i + 1) in header_rows else i + 1
        stop = next_date - 1 if (next_date < len(texts) and next_date - 1 > i and len(texts[next_date - 1].split()) <= 12) else next_date
        body_lines = texts[begin:max(begin, stop)]
        while body_lines and len(body_lines[0].split()) <= 3 and (is_place_text(body_lines[0]) or re.match(r"^(?:remote|hybrid|onsite|on-site|work from home)$", body_lines[0].strip(), re.IGNORECASE)):
            body_lines = body_lines[1:]
        description = shorten(" ".join(body_lines))
        used = list(dict.fromkeys(word_lists.skills[k].lower() for k, _ in find_terms(" ".join(texts[begin:max(begin, stop)]), word_lists.skills)
                                  if is_strong_skill(k)))[:15]
        entries.append({"title": title, "company": company, "company_source": company_source, "start": "%04d-%02d" % start, "end": "%04d-%02d" % end,
                        "months": months_between(start, end), "current": current, "description": description, "skills": used})
    seen, unique = set(), []
    for entry in entries:
        entry["title"] = entry["title"].lstrip(":;|- ") if entry["title"] else entry["title"]               # 'DESIGNATION: :Graduate Apprentice Training.'
        entry["company"] = entry["company"].lstrip(":;|- ") if entry["company"] else entry["company"]
        k = (entry["start"], entry["end"], norm_key(entry["company"] or ""), norm_key(entry["title"] or ""))
        if k not in seen:
            seen.add(k)
            unique.append(entry)
    if undated and not any(entry["title"] or entry["company"] for entry in unique):
        unique += find_undated_entries(lines)
    return unique


def total_experience_years(sections, entries):
    """'6+ years of experience' if the resume says so; otherwise the length of the experience entries (overlaps counted once)."""
    top = " ".join(line.text for line in sections.get("top", [])[:12] + sections.get("summary", [])[:8])
    match = re.search(r"(\d{1,2}(?:\.\d{1,2})?)\s*\+?\s*(?:years?|yrs?)\b[^.]{0,40}?(?:experience|exp)\b|(?:experience|exp)\b[^.]{0,30}?(\d{1,2}(?:\.\d{1,2})?)\s*\+?\s*(?:years?|yrs?)", top, re.IGNORECASE)
    if match:
        return float(match.group(1) or match.group(2)), "stated in the resume"
    return None, None                                               # not written in the resume: left blank, never calculated from the dates


SMALL_WORDS = {"a", "an", "the", "of", "for", "in", "on", "at", "to", "and", "or", "with", "by", "from", "using", "via", "as", "&", "-", "–", "is", "vs", "per", "into", "over"}


def title_like(text):
    """A name is written with capitals ('Online Store Sales Analysis', 'OFFICE MANAGEMENT APPLICATION'); a description is a sentence with small-letter
    words ('Designed a fully functional BI App', 'Agenda of the project'). True when at least 60% of the content words start with a capital, a digit or
    are upper-case, so the decision does not depend on bold or font size."""
    words = [word.strip("()[],.:;\"'“”‘’") for word in re.split(r"[\s/]+", clean(text)) if word.strip("()[],.:;\"'“”‘’")]
    content = [word for word in words if word.lower() not in SMALL_WORDS]
    if not content:
        return False
    big = sum(1 for word in content if word[:1].isupper() or word[:1].isdigit() or word.isupper() or re.match(r"[a-z]+[A-Z]", word))
    return big >= 0.6 * len(content)


def tidy_project_name(name):
    name = clean(name).strip(" -–:\"'“”‘’")
    return re.sub(r"[.,;]+$", "", name).strip()


def find_projects(sections):
    """Project titles from the Projects section. A title is a short line that is not a sentence and not a label
    ('Description:', 'Technologies:'). 'Project Name: X' labels win. For every project the skills used are the
    dictionary skills found in its block of text."""
    word_lists = lookups()
    lines = sections.get("projects", [])
    labels = re.compile(r"^(description|organi[sz]ation|company|employer|details?|role|roles|responsibilit\w*|technolog\w*|tech stack|tools?|environment|duration|period|team size|team|"
                        r"client|customer|platform|objective|aim|abstract|outcome|summary|overview|contribution|domain|database|languages?|"
                        r"major responsibilit\w*|key responsibilit\w*|basic academic|academic credentials|roles? (?:&|and) responsibilit\w*)\b", re.I)
    titles, current, title_sizes = [], None, []
    industry_left = 0
    bold_titles = False                                             # bold is only a hint: a project name is recognised by its words (capitals, no verb, no label), plain or bold


    ids = [text_line.list_id for text_line in lines if text_line.list_id]                   # a Word list: the project titles are one list, their points another
    if len(set(ids)) >= 2 and ids.count(ids[0]) >= 2 and lines and lines[0].list_id and len(ids) >= 0.8 * len(lines):
        for line in lines:
            if line.list_id == ids[0]:
                current = {"name": tidy_project_name(clean(line.text))[:120], "block": []}
                titles.append(current)
            elif current is not None:
                current["block"].append(clean(line.text))
        lines = []

    group_ids = {id(text_line) for text_line in lines if text_line.bold and (re.match(r"at\s", clean(text_line.text), re.IGNORECASE) or re.search(r"\b(?:ltd|pvt|limited|inc|llc|corp)\b", text_line.text, re.IGNORECASE)) and len(text_line.text.split()) <= 8}
    if group_ids:
        bold_titles = False                                         # bold 'At Samsung Heavy Industries' lines group the projects: the plain lines under them are the projects
    for line_no, line in enumerate(lines):
        text = clean(line.text)
        nxt_line = clean(lines[line_no + 1].text) if line_no + 1 < len(lines) else ""
        if id(line) in group_ids:
            current = None
            continue
        if re.match(r"^(?:be|b\.?e\.?|diploma|academic|final year|college|school|major|minor|mini|main|live|personal|industrial|training)\s+projects?\s*:?$", text, re.IGNORECASE) \
                and nxt_line and not labels.match(nxt_line):
            current = None                                          # 'Diploma Project' heads the projects under it, it is not one itself
            continue
        if group_ids and not line.bold and not labels.match(text) and not text[:1].islower() and len(text.split()) >= 2 and not re.match(r"^[\W\d]+$", text):
            year_prefixed = re.match(r"^(?:19|20)\d\d\s*[:\-–]\s*(.+)$", text)
            body = year_prefixed.group(1) if year_prefixed else text
            parts = re.split(r"\s+[-–—]\s+", body, maxsplit=1) if year_prefixed else [body]
            current = {"name": tidy_project_name(parts[0])[:120], "block": [parts[1]] if len(parts) > 1 else []}
            titles.append(current)
            continue
        if current is not None and re.search(r"\bindustry\s*:?\s*$", text, re.IGNORECASE):
            current["block"].append(text)
            industry_left = 2                                       # 'Location: Industry:' is a label, the two lines under it are its values
            continue
        if current is not None and industry_left > 0 and not re.match(r"project\s*(?:name|title)", text, re.IGNORECASE):
            industry_left -= 1
            current["block"].append(text)
            continue
        marker = re.fullmatch(r"\W*(?:#\s*\d+\s*(?:project|proj)|(?:project|proj)\s*(?:#|no\.?)?\s*\d*)\s*[:\-–.]?\s*", text, re.IGNORECASE)
        if marker:                                                  # '# 4 Project' / 'Project 2': only a number, the name comes with the next lines (or inside the description)
            current = {"name": "", "block": [], "pending": True}
            titles.append(current)
            continue
        if current is not None and current.get("pending") and not current["block"] and not current["name"]:
            cand = tidy_project_name(text)
            if cand and 1 <= len(cand.split()) <= 14 and not labels.match(cand) and not PROSE_START_RE.match(cand) and not text.rstrip().endswith(":") and not cand[:1].islower()                     and (line.bold or title_like(cand)) and not re.search(r"\b(?:is|are|was|were|will|can)\b", cand) and len(cand) >= 4:
                current["name"] = cand[:120]                       # the line under the marker is the project name
                continue
        titled = re.match(r"^\W*title\s*[:\-–]\s*(?P<n>[A-Za-z0-9][^:]{2,100})$", text, re.IGNORECASE)
        if titled and not labels.match(titled.group("n")) and not PROSE_START_RE.match(titled.group("n")):
            current = {"name": tidy_project_name(titled.group("n"))[:120], "block": []}                       # 'Title : CCTNS' under 'PROJECT # 1'
            titles.append(current)
            continue
        inline = re.match(r"^\W*(?:in\s+|the\s+)?(?:(?:mini|main|major|minor|final|academic|live|college|personal)\s+)?project\s*[:\-–]?\s+(?P<n>(?-i:[A-Z0-9])[^.:]{5,100})$", text, re.IGNORECASE)
        if inline and title_like(inline.group("n")) and not PROSE_START_RE.match(inline.group("n")) and len(inline.group("n").split()) <= 14 and not labels.match(inline.group("n")):
            current = {"name": tidy_project_name(inline.group("n"))[:120], "block": [], "wrap": True}          # 'In main project DESIGN AND ANALYSIS OF LEAF SPRING BY USING'
            titles.append(current)
            continue
        if current is not None and current.get("wrap") and not current["block"] and len(text.split()) <= 8 and (line.bold or text.isupper()) and not labels.match(text)                 and not text.endswith(":") and title_like(text) and len(current["name"].split()) + len(text.split()) <= 18:
            current["name"] = tidy_project_name(current["name"] + " " + text)[:140]      # the project name wrapped over two lines
            continue
        match = re.match(r"^(?:(?:(?:project|product|application)\s*(?:title|name)?|title)\s*\d*\s*[:\-–]\s*)(.+)$", text, re.IGNORECASE)
        if match and tidy_project_name(match.group(1)):
            if current is not None and re.match(r"title\b", text, re.IGNORECASE) and not current["block"] and current["name"]:
                current["block"].append(tidy_project_name(match.group(1)))               # 'Pharma Distributor Project' / 'Title : Identifying Scope for Pharma Distributor': the sub-title of the same project
                continue
            if current is not None and re.match(r"(?:product|application)\s*name", text, re.IGNORECASE) and not current["block"]:
                continue                                            # 'PROJECT 1: X' then 'Product Name: X' is the same project
            current = {"name": tidy_project_name(match.group(1))[:120], "block": []}
            titles.append(current)
            continue
        colon = re.match(r"^([A-Z0-9][^:]{5,70}?)\s*:\s+(\S.*)$", text)           # 'Aladdin Product Development: Aladdin is the product of ...'
        if colon and 2 <= len(colon.group(1).split()) <= 8 and not labels.match(colon.group(1)) and not re.match(r"at\s", colon.group(1), re.IGNORECASE) and not PROSE_START_RE.match(colon.group(1)) \
                and not re.search(r"\b(?:is|are|was|were|skills?|technologies|environment|responsibilit\w*|description|role|duration|client)\b", colon.group(1), re.IGNORECASE) \
                and (current is None or len(current["block"]) >= 1 or line.bold):
            current = {"name": tidy_project_name(colon.group(1))[:120], "block": [colon.group(2)]}
            titles.append(current)
            continue
        bullet = bool(re.match(r"^[>•\-\*▪●◦·]\s", line.text.strip()))
        dated = re.search(r"\s*\(\s*(?:" + _DATE_PART + r")(?:\s*(?:-|–|—|to)\s*(?:" + _END_WORDS + r"|" + _DATE_PART + r"))?\s*\)\s*$", text, re.IGNORECASE)
        shown = text[:dated.start()] if dated else text                 # 'Standard Part Cost Screening Tool (10/2018 - 07/2019)' -> the title
        next_label = current is None and bool(labels.match(nxt_line)) and not labels.match(shown) and not shown.endswith(":")
        is_title = (2 <= len(shown.split()) <= 12 and not re.match(r"^at\s", shown, re.IGNORECASE) and not shown.endswith((".", ":", "-", "–", ",")) and (dated or not DATE_RANGE_RE.search(shown))
                    and (line.bold or next_label or ((not PROSE_START_RE.match(shown) or shown.isupper()) and title_like(shown))) and not (shown.count(",") >= 2 and "(" not in shown) and not labels.match(shown) and not re.search(r"^(?:front|back)[\s-]*end\b|\btechnolog(?:y|ies)\b|\btech stack\b|\bframeworks?\b|\bdatabase\b|\bplatform\s*:", shown, re.IGNORECASE) and not shown[:1].islower() and not re.match(r"^[\W\d]+$", shown)
                    and len(shown) >= 4 and not (shown.startswith("(") and shown.endswith(")")) and not (bullet and not line.bold)
                    and not re.search(r"\b(?:is|are|was|were|provides?|provided|able|ability|have|has|will|can)\b", shown))
        # a title is usually bold or starts a bullet; without styling we accept the first such line after a long block
        small = bool(title_sizes) and line.size and not line.bold and line.size < min(title_sizes) - 1.0
        relaxed = ((group_ids or not any(text_line.bold for text_line in lines)) and not current["block"]) if current is not None else False
        if re.search(r"\b(?:overview|skills|responsibilit\w*|technologies)\s*$", shown, re.IGNORECASE):
            is_title = False                                        # 'SALESFORCE SUMMARY' heads a block of skills, it is not a project
        if is_title and not small and (line.bold or (not bold_titles and (current is None or len(current["block"]) >= 1 or relaxed))):
            current = {"name": tidy_project_name(shown)[:120], "block": []}
            titles.append(current)
            if line.size and not line.bold:
                title_sizes.append(line.size)
        elif current is not None:
            current["block"].append(text)
    if not titles:                                                  # a Projects section of plain sentences: one project per line
        for line in sections.get("projects", []):
            text = clean(line.text)
            if len(text.split()) < 3 or text[:1].islower() or labels.match(text) or re.match(r"^[\W\d]+$", text) or heading_section(text.rstrip(":")):
                continue
            name = " ".join(re.split(r"\s+(?:at|for|in|using|with|on)\s+|[:(,]", text)[0].split()[:14])
            titles.append({"name": tidy_project_name(name), "block": [text]})
    for part in titles:
        if not part["name"]:                                           # 'The project Vehicle Tracking & Ticketing System (VTTS) mainly deals with ...'
            said = re.search(r"\b(?:the|this)\s+project\s+(?:named\s+|called\s+|titled\s+)?(?P<n>[A-Z][A-Za-z0-9 &/\-]{3,70}?(?:\s*\([A-Z]{2,8}\))?)\s+(?:mainly|manly|mostly|is|was|deals|aims|which|that|will|has|provides)\b", " ".join(part["block"]))
            if said:
                part["name"] = tidy_project_name(said.group("n"))
    titles = [part for part in titles if part["name"]]
    titles = [part for part in titles if not re.search(r"\b(?:industrial|field|site)\s+(?:visits?|trips?|tours?)\b", part["name"], re.IGNORECASE)]
    out = []
    for part in titles:
        found = [word_lists.skills[k] for k, _ in find_terms(" ".join(part["block"]), word_lists.skills) if is_strong_skill(k)]
        head_lines = list(part["block"])
        while head_lines and len(head_lines[0].split()) <= 12 and not head_lines[0].rstrip().endswith(".") and (head_lines[0].count(",") >= 2 or re.search(r"\bproject$", head_lines[0].strip(), re.IGNORECASE)):
            head_lines.pop(0)                                       # the sub-title / technology list under the project name is not its description
        block_text = " ".join(head_lines)
        link = re.search(r"https?://\S+|(?:github|gitlab)\.com/\S+", block_text)
        out.append({"name": part["name"], "skills": list(dict.fromkeys(found))[:12], "description": shorten(block_text, 2500),
                    "link": link.group(0).rstrip(".,)") if link else ""})
    out = [part for part in out if re.search(r"[A-Za-z]{3}", part["name"]) and part["name"].count(")") <= part["name"].count("(")]
    known = {norm_key(part["name"]) for part in out}
    for name in ("experience", "internships"):                       # 'Project: FinNova Cloud 360 Platform' written inside a job is a project of the candidate
        block_lines = sections.get(name, [])
        starts = [k for k, line in enumerate(block_lines) if re.match(r"^\W*project(?:\s*(?:name|title))?\s*[:\-–]\s*(.{3,100})$", clean(line.text), re.IGNORECASE)]
        for n, k in enumerate(starts):
            match = re.match(r"^\W*project(?:\s*(?:name|title))?\s*[:\-–]\s*(.{3,100})$", clean(block_lines[k].text), re.IGNORECASE)
            pname = tidy_project_name(match.group(1))
            if pname and norm_key(pname) not in known and len(pname.split()) <= 14:
                known.add(norm_key(pname))
                body = [clean(text_line.text) for text_line in block_lines[k + 1:(starts[n + 1] if n + 1 < len(starts) else len(block_lines))]]
                found = [word_lists.skills[part_text] for part_text, _ in find_terms(" ".join(body), word_lists.skills) if is_strong_skill(part_text)]
                out.append({"name": pname, "skills": list(dict.fromkeys(found))[:12], "description": shorten(" ".join(body), 2500), "link": ""})
    return out


CERT_NOISE = re.compile(r"^(?:duration|other education|certification (?:name|body)|co-?curricul\w*|extra-?curricul\w*|year of|board|university|institution|grade|validity|percentage|marks|s\.?\s?no|title|organi[sz]ation|"
                        r"mob|mobile|place|date|passing|div|dob|date of birth|gender|marital|address|personal info\w*|nationality|languages?|father|mother|"
                        r"permanent|current address|hobbies|strengths?)\b|\bmob\s*[:.\-]|^(?:pursued|worked|engineered|negotiat\w*|analy[sz]\w*|handled|managed|"
                        r"led|lead|distribut\w*|executed|conducted|collaborat\w*|coordinat\w*|performed|ensured|achieved|improved|reduced|"
                        r"honesty|creative|soft)\b|\b(?:ltd|pvt|limited)\b\.?\s*,?\s*[A-Z]?[a-z]*$|\b\d+\+?\s*(?:gb|tb|years?)\b", re.IGNORECASE)
COURSE_WORDS = re.compile(r"program|course|certif|bootcamp|training|exam|diploma|nanodegree|workshop", re.IGNORECASE)
CERT_DETAIL_RE = re.compile(r"^(?:[A-Z]{1,4}[- ]?\d{2,4}[A-Za-z]?|\(?[A-Z]{2,5}[- ]\d{2,4}\)?)$")          # 'AZ 900', 'AZ-104': the code of the line above
CERT_WORDS = re.compile(r"\b(certified|certification|certificate|certificates|licen[sc]e|diploma in|nanodegree|specialization)\b", re.IGNORECASE)
CERT_VENDORS = re.compile(r"\b(aws|azure|google|microsoft|cisco|oracle|coursera|udemy|edx|udacity|linkedin learning|scrum|pmp|prince2|itil|"
                          r"ccna|ccnp|cissp|ceh|comptia|six sigma|salesforce|sap|red hat|hubspot|nptel|simplilearn)\b", re.IGNORECASE)


def keep_certification(text, strict, in_section=False, *, out):
    text = clean(text)
    text = re.sub(r"^(?:" + MONTHS + r")\.?\s+\d{1,2},?\s+(?:19|20)\d{2}\s*[:\-–]\s*", "", text, flags=re.IGNORECASE)     # 'September 18, 2015: Certified in CSWP'
    text = re.sub(DATE_RANGE_RE, " ", text)                         # '(09/2019 - Present)' is a date, not part of the name
    text = re.sub(r"\b(?:" + MONTHS + r")\.?,?\s+(?:19|20)\d{2}\b", "", text, flags=re.IGNORECASE)      # 'July 2014' after the name
    text = re.sub(r"\b(?:19|20)\d{2}\b|\((?:in progress|ongoing|in view)\)|\(\s*(?:\d{1,2}\s*[/\-]?)?\s*\)", "", text)
    text = re.sub(r"\(\s*\d{0,2}\s*[/\-]?\s*\)", "", text)                       # '(09/)' left where the year was cut out
    text = re.sub(r"\s+,|,(?=\s*(?:from|at|by)\b)", "", re.sub(r"\s{2,}", " ", text)).strip(" -–:|,.")
    taught = in_section and re.match(r"(?:successfully\s+)?(?:completed|attended|done|passed|cleared|pursued|undergone|participated|participating|presented|won|received|secured|awarded|represented|organi[sz]ed|candidate|successful(?:ly)?\s+completion)\b", text, re.IGNORECASE)
    if not 3 <= len(text) <= 160 or len(text.split()) > (22 if in_section else 16):
        return
    if text.startswith("(") and text.endswith(")"):
        return                                                      # '(Certified in above courses)' is a note about the list, not one more certificate
    noise = CERT_NOISE.search(text)
    if noise and in_section and COURSE_WORDS.search(text) and re.search(r"\b(?:ltd|pvt|limited)\b", noise.group(0), re.IGNORECASE):
        noise = None                                                # 'Full Stack Data Science Program at Business Toys Pvt Ltd'
    if (PROSE_START_RE.match(text) and not taught) or text[:1].islower() or EMAIL_RE.search(text) or noise or re.search(r"\brank\b", text, re.IGNORECASE):
        return
    if len(text.split()) > 9 and not CERT_WORDS.search(text) and not CERT_VENDORS.search(text) and not (in_section and (taught or COURSE_WORDS.search(text))):
        return                                                      # a sentence, not a certificate name
    if len(text.split()) == 1 and not CERT_VENDORS.search(text):
        return
    if re.search(r"\b(is a|is an|are|was|were|awarded to|credible|by the)\b", text, re.IGNORECASE) or text.count("(") != text.count(")"):
        return
    if re.search(r"\bpresent\b|^\d{1,2}\s*/|\bcourse curriculum\b", text, re.IGNORECASE) or (
            TITLE_HEAD.search(text.split()[-1] + " ") and not CERT_WORDS.search(text) and not CERT_VENDORS.search(text) and not in_section):
        return
    if DEGREE_RE.search(text) and not CERT_WORDS.search(text) and not strict and not (in_section and taught):
        return
    if text.lower() in {"certifications", "certification", "certificates", "certificate", "courses", "online", "udemy", "coursera"} \
            or re.fullmatch(r"[\d\s.%/]+", text) or re.match(r"(?:extra|accomplish|awards?|awarded|personal|trainings?\s*/|achievements?|hobbies|interests?|declaration|references?)\b", text, re.IGNORECASE):
        return
    out.append(text)


def find_certifications(sections, all_lines):
    """Lines of the Certifications section (job duties and bare dates removed) + lines elsewhere that clearly
    name a certificate: 'Certified X', 'X Certification', or a vendor / course platform."""
    out = []


    section = sections.get("certifications", [])
    dated_lines = [line for line in section if re.match(r"^\W*(?:19|20)\d{2}\s*[|:\-–]\s*\S", clean(line.text))]
    if len(dated_lines) >= 3:                                       # '2021 | Certified Accounts Payable Professional': a dated list; the sentences that follow it (a page column of another section) are not certificates
        section = dated_lines
    joined = " ".join(clean(line.text) for line in section)
    said = re.search(r"\b(?:gone through|various|several|attended|done)\s+(?:\w+\s+){0,2}certifications?\s+(?:inv\w*|in|on|of|like|such as)\s+(.+?)(?:\s+etc\b\.?|\.\s|$)", joined, re.IGNORECASE)
    if said:                                                        # 'Gone through various certification involving Python, Machine Learning, Data Structure and Algorithm etc.'
        for part in re.split(r"\s*,\s*|\s+and\s+(?=[A-Z])", said.group(1)):
            part = part.strip(" .")
            if 2 <= len(part) <= 60:
                out.append(part)
        return list(dict.fromkeys(out))
    bold_names = sum(1 for line in section if line.bold) >= 2 and any(not line.bold for line in section)
    kept_size = None
    header_hits = sum(bool(re.search(r"year of|passing|percentage|cgpa|board|discipline|aggregate|institution|school", line.text, re.IGNORECASE)) for line in section)
    wrapped, last_bold = [], False                                  # a name that the PDF wrapped over two lines is joined first
    for index, line in enumerate(section):
        part_text = clean(line.text)
        if not part_text:
            continue
        if bold_names and not line.bold and not (CERT_WORDS.search(part_text) and CERT_VENDORS.search(part_text)):
            continue                                                # bold = the certificate, plain = its issuer / code line below it
        if bold_names and wrapped and last_bold and index > 0 and section[index - 1].bold and index + 1 < len(section) and not section[index + 1].bold \
                and section[index + 1].text.strip().startswith("("):
            wrapped[-1] += " " + part_text                                  # a bold name wrapped over two lines: its issuer line (plain) comes after the second
            continue
        if bold_names and wrapped and len(part_text.split()) <= 1 and last_bold:
            wrapped[-1] += " " + part_text                                  # 'Azure DevOps Engineer' / 'Expert'
            continue
        if wrapped and re.match(r"(?:internship|training|workshop|course)\b", part_text, re.IGNORECASE) and re.search(r"certificate\s*$", wrapped[-1], re.IGNORECASE) and len(wrapped[-1].split()) <= 3:
            wrapped[-1] += " : " + part_text                                # 'Golden Certificate' / 'Internship of HR in (Notes pickup)': one certificate
            continue
        if wrapped and CERT_DETAIL_RE.match(part_text):
            wrapped[-1] += " " + part_text                                  # 'Microsoft Azure Fundamentals' / 'AZ 900'
            continue
        if not bold_names and wrapped and kept_size and line.size and kept_size - line.size >= 1.5 and len(part_text.split()) <= 4 and not CERT_WORDS.search(part_text):
            continue                                                # a smaller line under a certificate is its issuer (Coursera, GeeksforGeeks)
        if line.size and (not bold_names):
            kept_size = line.size
        last_bold = line.bold
        if wrapped and (re.search(r"(?:\b(?:and|in|of|on|for|the|by|to|with|at)|[-,&–/])$", wrapped[-1], re.IGNORECASE) or part_text[:1].islower()
                        or (re.match(r"(?:certification|certificate|course)\b[,\s]", part_text, re.IGNORECASE) and not CERT_WORDS.search(wrapped[-1]))):
            wrapped[-1] += " " + part_text
        else:
            wrapped.append(part_text)
    for text in wrapped:
        text = re.sub(r"^[o•\-\*▪►▸>]\s+", "", text)
        if header_hits >= 3 and not (CERT_WORDS.search(text) or CERT_VENDORS.search(text) or re.search(r"\b(?:course|training|workshop|bootcamp|virtual experience)\b", text, re.IGNORECASE)):
            continue                                                # an education table under a heading that looked like Certifications
        keep_certification(text, False, True, out=out)
    return list(dict.fromkeys(out))


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
    texts = [line.text for line in all_lines]
    for i, text in enumerate(texts):
        if not SALARY_LABEL_RE.search(text) or not re.search(r"ctc|salary|package|lpa|remuneration|compensation|ectc|cctc", text, re.IGNORECASE):
            continue
        window = text + " " + (texts[i + 1] if i + 1 < len(texts) and len(texts[i + 1]) < 40 else "")
        for kind, pattern in (("expected", r"expected|expecting|ectc|e-ctc|desired"), ("current", r"current|present|existing|cctc|c-ctc|drawing|fixed|last")):
            match = re.search(r"(?:" + pattern + r")[^0-9₹]{0,40}(?:₹|rs\.?|inr)?\s*(\d[\d,.]*)\s*(lpa|lacs?|lakhs?|l\b|crores?|cr\b|k\b|per\s*annum|pa\b|per\s*month|pm\b|monthly)?", window, re.IGNORECASE)
            if match and kind not in result:
                number = float(match.group(1).replace(",", "").rstrip(".") or 0)
                value = to_rupees_per_year(number, match.group(2), window)
                if value:
                    result[kind] = {"per_year_inr": value, "raw": clean(match.group(0))[:60]}
        if not result:
            match = AMOUNT_RE.search(text)
            if match and re.search(r"ctc|salary|package", text, re.IGNORECASE):
                value = to_rupees_per_year(float(match.group(1).replace(",", "")), match.group(2), text)
                if value:
                    result["current"] = {"per_year_inr": value, "raw": clean(match.group(0))[:60]}
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
                    shared.append("".join(part_text.text or "" for part_text in si.iter(ns + "t")))
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
                        cells[col] = "".join(part_text.text or "" for part_text in c.iter(ns + "t"))
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
    ("phone", r"mob|phone|contact\s*n|contant\s*n", r"alternate|alt\b"),
    ("linkedin", r"linkedin", None),
    ("skills", r"skill", None),
    ("title", r"designation|current title|job title|^role$", r"previous|past|last|applied"),
    ("past_title", r"(?:previous|past|last)\s*designation", None),
    ("company", r"^company$|current company|current org|current employer|current firm|employer|organi[sz]ation", r"previous|past|last|\bjr\b|with jr"),
    ("past_company", r"(?:previous|past|last)\s*(?:company|employer|organi[sz]ation)", None),
    ("experience", r"experience|\bexp\b|yrs of exp", r"relevant|\brel\b|total_|notice|skill|salary|ctc|\bsal\b|level|expected|comment"),
    ("experience_alt", r"relevant\s*(?:exp|yrs)|\brel\b\s*exp", None),
    ("education", r"degree|qualif|education|college|university", r"start|end|time|date|percent|%|\bpg\b|\bug\b|10th|12th"),
    ("course", r"^course$|course name", None),
    ("specialization", r"special|branch|stream|discipline|field of study", None),
    ("languages", r"language", None),
    ("links", r"link|url|website|portfolio|social|github|^resume$", r"match|similar|linkedin"),
    ("location", r"^location$|^city$|current location", None),
    ("current_ctc", r"current[_ ]?(ctc|salary)|^ctc$|cctc", None),
    ("expected_ctc", r"expected[_ ]?(ctc|salary)|ectc", None),
]

COURSE_CODES = {"be": "B.E", "me": "M.E", "btech": "B.Tech", "mtech": "M.Tech", "bba": "BBA", "mba": "MBA", "bca": "BCA", "mca": "MCA", "bsc": "B.Sc",
                "msc": "M.Sc", "bcom": "B.Com", "mcom": "M.Com", "ba": "B.A", "ma": "M.A", "barch": "B.Arch", "bpharm": "B.Pharm", "mpharm": "M.Pharm",
                "diploma": "Diploma", "iti": "ITI", "phd": "PhD", "llb": "LLB", "bed": "B.Ed"}


def degrees_from_course(course, specialization):
    """The Course and Specialization columns of a sheet ('be' + 'Civil', 'btech' + 'Mechanical Engineering', 'diploma_in_tool_and_die_making',
    'pg_diploma') -> (degrees, fields of study). Course codes are written in lower case in these sheets, so they are read from a code table."""
    degrees, fields = [], []
    for segment in re.split(r"\s*[|;,/]\s*", course or ""):
        segment = segment.strip()
        if not segment:
            continue
        degree, extra = canon_degree(segment), ""
        if not degree:
            code = segment.lower().replace(" ", "_")
            if code in ("pg_diploma", "pgdiploma"):
                degree = "PG Diploma"
            else:
                head, _, rest = code.partition("_")
                if head in COURSE_CODES:
                    degree, extra = COURSE_CODES[head], re.sub(r"^in_", "", rest).replace("_", " ")
        if degree:
            degrees.append(degree)
            if extra and canon_field(extra):
                fields.append(canon_field(extra))
    if degrees:
        for segment in re.split(r"\s*[|;,/]\s*", specialization or ""):
            if canon_field(segment):
                fields.append(canon_field(segment))
    return list(dict.fromkeys(degrees)), list(dict.fromkeys(fields))


def links_from_cell(cell):
    """The web links of the link columns of a sheet ('Resume Link', 'Profile Link', 'Portfolio', 'Other links' ...), also a bare 'johndoe.dev/work'.
    LinkedIn links are left out: they are in the 'linkedin' field."""
    tokens = re.split(r"[\s,;|]+", cell or "")
    text = " ".join("https://" + part_text if (re.fullmatch(r"(?=[^\s]*[A-Za-z])[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+(?:/\S*)?", part_text) and not part_text.lower().startswith(("http", "www."))) else part_text
                    for part_text in tokens)
    return find_links([Line(text)])


def languages_from_text(text):
    """'english, hindi, oriya, odia' -> the languages that are in languages.json (lower case)."""
    word_lists = lookups()
    if not hasattr(word_lists, "language_table"):
        word_lists.language_table = {norm_key(name): name for name in json.load(open(os.path.join(LOOKUPS, "languages.json"), encoding="utf-8"))}
    return list(dict.fromkeys(word_lists.language_table[k].lower() for k, _ in find_terms(text or "", word_lists.language_table, max_words=1)))


SENTENCE_START_RE = re.compile(
    r"^(?:ability|able|willing\w*|good|excellent|strong|great|hard[\s-]?working|self[\s-]?motivated|team player|honest\w*|keen|eager|passionate|"
    r"enthusiastic|motivated|dedicated|seeking|looking|having|possess\w*|knowledge of|experience (?:in|of|with|on)|proficient (?:in|with)|"
    r"expertise (?:in|with)|responsible for|worked|working|handled|managed|developed|designed|implemented|my|we|the|to(?!-)|an)\b", re.IGNORECASE)


def plausible_item(text, max_words=14):
    """A value of a list section (title, certificate, education, project, company ...) must be something a person would write
    as a name, not a number, a date, a phone number or a sentence."""
    part_text = clean(str(text or "")).strip(" .,:;|-–—•")
    letters = len(re.findall(r"[^\W\d_]", part_text))
    if letters < 3 or letters < 0.5 * len(re.sub(r"\s", "", part_text)):
        return False                                                   # '2012', '06-05', '+91', '9265583411'
    if re.search(r"\d{7,}|\+\d|\d{1,4}[-/.]\d{1,2}[-/.]\d{2,4}", part_text):
        return False                                                   # a phone number or a date inside the text
    return len(part_text.split()) <= max_words and not SENTENCE_START_RE.match(part_text)


def listed_item(text):
    """A value taken from its own section is kept whenever it has words: only numbers, dates and phone numbers are refused."""
    part_text = clean(str(text or "")).strip(" .,:;|-–—•")
    return len(re.findall(r"[^\W\d_]", part_text)) >= 3 and len(part_text) <= 300 and not re.fullmatch(r"[\d\s+()./-]+", part_text)


def skill_ok(text):
    """A skill is a short name: never a bare number, a date or a sentence ('ability to work in a team')."""
    part_text = clean(str(text or "")).strip(" .,:;|-–—•")
    if part_text.lower() in SKILL_NOT_A_SKILL or re.search(r"\b(?:yrs?|years?|months?|k|lpa|ctc)\b", part_text, re.IGNORECASE) or re.match(r"^\d", part_text):
        return False                                                # 'problem solving', 'english', '5+ yr', '4k' are not skills
    return bool(re.search(r"[^\W\d_]", part_text)) and not re.fullmatch(r"[\d\s.,/+\-]+", part_text) and len(part_text.split()) <= 6 and not SENTENCE_START_RE.match(part_text)


SKILL_NOT_A_SKILL = {"problem solving", "english", "hindi", "marathi", "telugu", "tamil", "kannada", "malayalam", "gujarati", "bengali", "punjabi",
                     "urdu", "good communication", "communication", "communication skills", "team work", "teamwork", "team player", "time management",
                     "analytical skills", "leadership", "leadership skills", "interpersonal skills", "hard working", "hardworking", "self motivated",
                     "self-motivated", "attention to detail", "adaptability", "multitasking", "critical thinking", "decision making", "system development",
                     "application development", "requirements analysis", "business analysis", "responsible for", "other", "others", "etc", "and"}


def titles_from_text(text):
    """Job titles of a spreadsheet cell ('SDET Lead | Senior Software Development Engineer in Test'): only titles that are in
    titles.json, and the same occupation-word rule as for a resume."""
    word_lists = lookups()
    out = {}
    for part in re.split(r"\s*[|;/\n]\s*|,\s+(?=[A-Z])", text or ""):
        part = clean(part).strip(" .,-:")
        if not part:
            continue
        key = norm_key(part)
        if key in word_lists.titles:
            out.setdefault(key, word_lists.titles[key])
            continue
        for k, _ in find_terms(part, word_lists.titles, max_words=6):
            name = word_lists.titles[k]
            if TITLE_HEAD.search(name.split()[-1] + " ") and len(name.split()) >= 2:
                out.setdefault(k, name)
    names = list(out.values())
    return [n for n in names if not any(n != match and n.lower() in match.lower() for match in names)]


def clean_list_education(text):
    """A spreadsheet's education cell ('Graduate | btech | Dr.MGR UNIVERSITY | 2025-12') -> only the parts that are a degree or an
    institution, without years and without the sheet's own category words."""
    out = []
    for segment in re.split(r"\s*[|;]\s*|,\s+", text or ""):
        segment = re.sub(DATE_RANGE_RE, " ", segment)
        segment = re.sub(r"\b(?:19|20)\d{2}(?:-\d{1,2})?\b", " ", segment)
        segment = clean(segment).strip(" .,-:/")
        if segment.count("(") > segment.count(")"):
            segment = segment.replace("(", "", 1).strip()
        elif segment.count(")") > segment.count("("):
            segment = segment[::-1].replace(")", "", 1)[::-1].strip()
        if not segment or re.fullmatch(r"(?:graduate|post[- ]?graduate|under[- ]?graduate|n/?a|none|not specified)", segment, re.IGNORECASE) \
                or re.search(r"pass or above", segment, re.IGNORECASE):
            continue
        if (DEGREE_RE.search(segment) or real_institution(segment)) and plausible_item(segment, 14):
            out.append(segment[:90])
    return list(dict.fromkeys(out))


def table_cell(field, *, columns, row):
    parts = [fix_mojibake(str(row[i]).strip()) for i in columns.get(field, []) if i < len(row) and str(row[i]).strip() not in ("", "-", "nan", "None")]
    return " | ".join(dict.fromkeys(parts))


def first_number_in_row(cols, *, row):
    for i in cols:
        match = re.search(r"\d+(?:\.\d+)?", str(row[i]) if i < len(row) else "")
        if match:
            return match.group(0)
    return ""


def parse_table(rows):
    """A candidate list -> {'columns', 'rows', 'candidates': [ {name, email, phone, ...} ]}. Columns are matched by name
    (any spelling of 'candidate name', 'mob.no', 'e-mail id', 'skill set' ...). Rows without a name or any contact are skipped."""
    if not rows:
        return {"columns": [], "rows": 0, "candidates": []}
    header = [str(heading).strip().lower() for heading in rows[0]]
    columns = collections.defaultdict(list)
    for field, want, avoid in COLUMN_RULES:
        for i, heading in enumerate(header):
            if re.search(want, heading) and not (avoid and re.search(avoid, heading)):
                columns[field].append(i)
    candidates = []
    for row in rows[1:]:
        name = clean(table_cell("name", columns=columns, row=row))
        email = None
        for candidate in EMAIL_RE.findall(table_cell("email", columns=columns, row=row)):
            try:
                email = validate_email(candidate, check_deliverability=False).normalized
                break
            except EmailNotValidError:
                continue
        name = unspace_name(name, email or "")                     # 'AB H I S H EK T I WAR I' -> 'Abhishek Tiwari'
        phone = None
        for match in phonenumbers.PhoneNumberMatcher(table_cell("phone", columns=columns, row=row), "IN"):
            if phonenumbers.is_valid_number(match.number) and not is_timestamp_like(match.raw_string):
                phone = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.E164)
                break
        linkedin = find_linkedin([Line(table_cell("linkedin", columns=columns, row=row))])
        if len(re.sub(r"[^A-Za-z]", "", name)) < 2 and not (email or phone):
            continue
        skill_cell = [Line(part.strip()) for part in re.split(r"[,;|\n]", table_cell("skills", columns=columns, row=row)) if part.strip()]
        skills = list(find_skills({"skills": skill_cell}, use_fuzzy=False))[:30] if skill_cell else []      # skills.json only
        title, past_title, company, past_company = table_cell("title", columns=columns, row=row), table_cell("past_title", columns=columns, row=row), table_cell("company", columns=columns, row=row), table_cell("past_company", columns=columns, row=row)
        titles = list(dict.fromkeys(titles_from_text(title) + titles_from_text(past_title)))
        around = " | ".join(item for item in (company, past_company, title, past_title) if item)
        industries = find_industries([], {"experience": [Line(around)]}) if around else []
        experience = table_cell("experience", columns=columns, row=row)
        years_cols = [i for i in columns.get("experience", []) if re.search(r"year|yrs", header[i]) and not re.search(r"month", header[i])]
        month_cols = [i for i in columns.get("experience", []) if re.search(r"month", header[i])]
        if years_cols and month_cols:                                # 'Work experience - years' + 'Work experience - months'
            y, mo = first_number_in_row(years_cols, row=row), first_number_in_row(month_cols, row=row)
            experience = ((f"{y}y " if y else "") + (f"{mo}mo" if mo else "")).strip()
        experience = experience or table_cell("experience_alt", columns=columns, row=row)
        if re.fullmatch(r"\d+(?:\.\d+)?", experience or ""):
            experience += " years"                                   # a bare number in an experience column is years
        list_degrees, list_fields = degrees_from_course(table_cell("course", columns=columns, row=row), table_cell("specialization", columns=columns, row=row))
        place = compose_location([find_places(table_cell("location", columns=columns, row=row), trusted=True)])
        links_cell = links_from_cell(table_cell("links", columns=columns, row=row))
        candidate = {"name": name or None, "email": email, "phoneNumber": phone,
                     "socialProfiles": [{"platform": "linkedin", "url": linkedin}] if linkedin else None,
                     "skills": [item.lower() for item in skills], "titles": titles, "industries": industries,
                     "languages": languages_from_text(table_cell("languages", columns=columns, row=row)),
                     "educationItems": clean_list_education(table_cell("education", columns=columns, row=row)),
                     "listDegrees": list_degrees, "listFields": list_fields, "locationDetails": place, "links": links_cell,
                     "currentCompanyDetails": {"companyName": company, "designation": title} if (company or title) else None,
                     "pastCompanyDetails": {"companyName": past_company, "designation": past_title} if (past_company or past_title) else None,
                     "totalExperienceText": experience or None, "educationText": table_cell("education", columns=columns, row=row) or None,
                     "location": table_cell("location", columns=columns, row=row) or None, "currentCTC": table_cell("current_ctc", columns=columns, row=row) or None, "expectedCTC": table_cell("expected_ctc", columns=columns, row=row) or None}
        candidates.append({k: v for k, v in candidate.items() if v not in (None, [], "")})
    return {"columns": [heading for heading in header if heading][:30], "rows": len(rows) - 1, "candidates": candidates}


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
    ("masters", r"\bm\.?\s?tech\b|\bm\.?\s?sc\b|\bm\.?\s?com\b|\bmba\b|\bmca\b|\bpgdm\b|\bpgdca\b|\bpgdba\b|\bpgdbm\b|\bpgpm\b|\bmaster|post[- ]?graduat|\bllm\b|"
                r"\bm\.?\s?pharm\b|\bm\.?\s?arch\b|\bmsw\b|\bm\.?e\b|\bm\.?a\b"),
    ("bachelors", r"\bb\.?\s?tech\b|\bb\.?\s?sc\b|\bb\.?\s?com\b|\bbba\b|\bbca\b|\bbachelor|\bb\.?e\b|\bb\.?a\b|\bb\.?\s?pharm\b|"
                  r"\bb\.?\s?arch\b|\bllb\b|\bmbbs\b|\bbsw\b|graduat(?:e|ion)\b"),
    ("diploma", r"diploma|\biti\b"),
    ("intermediate", r"intermediate|\bhsc\b|\bpuc\b|12th|xiith|senior secondary|higher secondary|\bpre[- ]?university|\bxii\b|class\s+xii"),
    ("high_school", r"10th|\bssc\b|\bsslc\b|matriculation|secondary school|high school|\bxth\b|class\s+x\b|std\.?\s*x\b"),
]
AUTHORITIES = [                       # certificate issuers, in the spelling the production data uses
    (r"\badobe\b", "Adobe"), (r"amazon web services|\baws\b", "Amazon Web Services"), (r"microsoft|\bazure\b", "Microsoft"), (r"google", "Google"),
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
    match = DEGREE_RE.match(degree.strip())
    rest = degree[match.end():] if match else ""
    rest = re.sub(r"^\s*(?:(?:degree\s+)?(?:in|of)\s+|[-–:(,]\s*)", "", rest, flags=re.IGNORECASE).strip(" ()-,.")
    rest = re.sub(r"\s+(?:in|of|and|&)$", "", rest, flags=re.IGNORECASE).strip()
    rest = re.sub(r"^(?:Technology|Engineering|Science|Arts|Commerce|Pharmacy|Education|Laws|Business Administration)\s+in\s+", "", rest, flags=re.IGNORECASE).strip()   # 'Bachelor of Technology in X' -> X
    return rest if rest and len(rest.split()) <= 7 and not INSTITUTION_RE.search(rest) and not DEGREE_RE.fullmatch(rest) else ""


def build_education(entries):
    out = []
    for entry in entries:
        kind = education_type(entry["degree"])
        field = "" if kind in ("intermediate", "high_school") else (entry.get("field") or canon_field(field_of_study(entry["degree"])))
        degree_shown = entry["degree"]
        if field and re.search(r"\s+in\s+" + re.escape(field) + r"\s*$", degree_shown, re.IGNORECASE):
            degree_shown = re.sub(r"\s+in\s+" + re.escape(field) + r"\s*$", "", degree_shown, flags=re.IGNORECASE)      # 'BACHELOR OF COMMERCE in COMPUTER APPLICATIONS': the field has its own column
        out.append({"educationType": kind, "degreeText": degree_shown, "institutionName": entry["institution"] or "", "institutionLabelled": bool(entry.get("institution_labelled")),
                    "startDate": entry.get("start_year") or "", "endDate": entry.get("end_year") or "",
                    "fieldOfStudy": field, "scoreAchieved": entry.get("score") or ""})
    return out


def build_job(entry):
    return {"companyName": entry["company"] or "", "companySource": entry.get("company_source", ""), "designation": entry["title"] or "", "startDate": entry["start"] + "-01" if entry["start"] else "",
            "endDate": "current" if entry["current"] else (entry["end"] + "-01" if entry["end"] else ""),
            "workExperienceDescription": entry["description"], "skillsUsedInRole": entry["skills"]}


def build_certifications(names):
    """Certificate names -> {name, issuingAuthority, certificateLink, certificateCredentialID}. The name is the whole text; only an issuer that comes
    after a dash / bar / 'by' / 'from' is cut off it ('Python Bootcamp - Udemy' -> 'Python Bootcamp', issuer Udemy)."""
    out = []
    for text in names:
        link = re.search(r"https?://\S+", text)
        credential = re.search(r"(?:credential|certificate|cert)\s*(?:id|no\.?|number)\s*[:#\-]?\s*([A-Za-z0-9\-_/]{5,})", text, re.IGNORECASE)
        name = clean(re.sub(r"https?://\S+|(?:credential|certificate|cert)\s*(?:id|no\.?|number)\s*[:#\-]?\s*[A-Za-z0-9\-_/]{5,}", "", text, flags=re.IGNORECASE))
        authority, main = "", name
        for match in re.finditer(r"\s+[-–—|]\s+|\s*\|\s*|,\s+|\s+(?:by|from|issued by|offered by)\s+", name):
            tail = clean(name[match.end():])
            hit = next((canonical for pattern, canonical in AUTHORITIES if re.fullmatch(r"(?:" + pattern + r")(?:\W.*)?", tail, re.IGNORECASE) or re.match(pattern, tail, re.IGNORECASE)), None)
            if hit and len(tail.split()) <= 5:
                authority, main = (tail if len(tail.split()) <= 4 else hit), clean(name[:match.start()])
                break
        if not authority:
            for pattern, canonical in AUTHORITIES:
                if re.search(pattern, name, re.IGNORECASE):
                    authority = canonical
                    break
        out.append({"name": main.strip(" -–:,"), "issuingAuthority": authority, "certificateLink": link.group(0).rstrip(".,)") if link else "",
                    "certificateCredentialID": credential.group(1) if credential else ""})
    return [c for c in out if len(c["name"]) >= 3]


def find_social_profiles(lines, linkedin):
    """linkedin, github, leetcode ... found anywhere in the text, plus a 'Portfolio:' website."""
    text = "\n".join(text_line.text for text_line in lines)
    out = []
    if linkedin:
        out.append({"platform": "linkedin", "url": linkedin})
    for platform, pattern in PROFILE_SITES:
        for match in re.finditer(r"(?:https?://)?(?:www\.)?" + pattern, text, re.IGNORECASE):
            url = "https://" + re.sub(r"^https?://(?:www\.)?", "", match.group(0)).rstrip(".,)/")
            if url.lower() not in (part["url"].lower() for part in out) and len(url) > len("https://github.com/") + 1:
                out.append({"platform": platform, "url": url})
    for line in lines:
        match = re.search(r"(?:portfolio|website|personal site|web)\s*[:\-]\s*((?:https?://)?[A-Za-z0-9.\-]+\.[a-z]{2,}(?:/\S*)?)", line.text, re.IGNORECASE)
        if match and not any(site in match.group(1).lower() for site in ("linkedin", "github", "leetcode", "hackerrank", "kaggle", "gmail")):
            out.append({"platform": "portfolio", "url": match.group(1).rstrip(".,)")})
    return out


def find_languages(sections, lines):
    """Spoken languages (english, hindi, ...) from the Languages part of the resume, as lower-case names."""
    word_lists = lookups()
    if not hasattr(word_lists, "language_table"):
        word_lists.language_table = {norm_key(name): name for name in json.load(open(os.path.join(LOOKUPS, "languages.json"), encoding="utf-8"))}
    text = [text_line.text for text_line in sections.get("other", [])]
    for line in lines:
        match = re.match(r"\s*languages?(?:\s+known|\s+spoken)?\s*[:\-]\s*(.+)", line.text, re.IGNORECASE)
        if match:
            text.append(match.group(1))
    return list(dict.fromkeys(word_lists.language_table[k].lower() for k, _ in find_terms(" ".join(text), word_lists.language_table, max_words=1)))


def find_notice_period(lines):
    """Notice period in days: '30 days', '2 months', 'immediate joiner' = 0."""
    for line in lines:
        match = re.search(r"notice\s*period\s*[:\-]?\s*(?:is\s*)?(\d{1,3})\s*(days?|weeks?|months?)", line.text, re.IGNORECASE)
        if match:
            return int(match.group(1)) * {"d": 1, "w": 7, "m": 30}[match.group(2)[0].lower()]
        if re.search(r"immediate(?:ly)?\s+(?:joiner|joining|available)|can join immediately", line.text, re.IGNORECASE):
            return 0
    return 0


def ctc_block(kind, with_duration, *, salary):
    value = (salary.get(kind) or {}).get("per_year_inr") or 0
    out = {"min": value, "max": value, "currency": "INR"}
    if with_duration:
        out["duration"] = "ANNUALLY"
    return out


def build_ctc(salary):
    return {"currentCTC": ctc_block("current", False, salary=salary), "expectedCTC": ctc_block("expected", True, salary=salary)}


def rank_skills(skill_scores, lines, jobs):
    """analysedSkills = {advanced, medium, basic}. A skill is 'advanced' when the resume gives strong evidence of it:
         +2    it is in the Skills list                       +0.5 per time it appears in the text (at most 4 times)
         +1.5  it was used in the current / latest job         +1   it was used in two or more jobs
       score >= 4 advanced, >= 2.5 medium, otherwise basic."""
    word_lists = lookups()
    text = "\n".join(line.text for line in lines)
    counts = collections.Counter(word_lists.skills[k] for k, _ in find_terms(text, word_lists.skills) if k not in word_lists.whole_item_only)
    levels = {"advanced": [], "medium": [], "basic": []}
    for skill, confidence in skill_scores.items():
        lower = skill.lower()
        in_jobs = [lower in j["skillsUsedInRole"] for j in jobs]
        score = (2 if confidence >= 0.9 else 0) + 0.5 * min(counts.get(skill, 0), 4) + (1.5 if in_jobs[:1] == [True] else 0) + (1 if sum(in_jobs) >= 2 else 0)
        levels["advanced" if score >= 4 else "medium" if score >= 2.5 else "basic"].append(lower)
    return levels


def job_start_date(job):
    """Sort key: the start date of a job."""
    return job["startDate"]


def build_red_flags(jobs, education, text_lines, issues):
    """The rule-based version of the production 'redFlags'. Every flag has a plain-language description."""
    flags = {}
    # --- employment gap: a month or more between the end of one job and the start of the next ---
    ordered = sorted((job for job in jobs if job["startDate"]), key=job_start_date)
    gap = None
    for before, after in zip(ordered, ordered[1:]):
        if before["endDate"] and before["endDate"] != "current" and after["startDate"][:7] > before["endDate"][:7]:
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
    missing = [entry for entry in education if entry["educationType"] in ("bachelors", "masters", "phd", "diploma") and not entry["institutionName"]]
    flags["incompleteEducation"] = (not education) or bool(missing)
    flags["incompleteEducationDescription"] = ("No education was found." if not education else
                                               f"{len(missing)} degree entries have no institution." if missing
                                               else "Education details include the degree and the institution.")
    # --- mistakes: words that lost their spaces, or key contact details missing ---
    words = [word for line in text_lines for word in line.text.split()]
    glued = sum(1 for word in words if len(word) > 24 and not word.startswith(("http", "www")) and "@" not in word)
    mistakes = (words and glued / len(words) > 0.01) or any("no valid email" in i or "no name" in i for i in issues)
    flags["containsMistakes"] = bool(mistakes)
    flags["containsMistakesDescription"] = ("The text has words run together or key details (name, email) are missing." if mistakes
                                            else "No obvious formatting problems were found.")
    flags["overstatedAchievements"] = False
    flags["overstatedAchievementsDescription"] = "Not assessed: this parser does not judge the claims in a resume."
    return {k: flags[k] for k in ("containsMistakes", "containsMistakesDescription", "frequentJobSwitcher", "frequentJobSwitcherDescription",
                                  "overstatedAchievements", "overstatedAchievementsDescription", "employmentGap", "employmentGapDescription",
                                  "incompleteEducation", "incompleteEducationDescription", "lackOfRoleProgression", "lackOfRoleProgressionDescription")}


CITY_TABLE = {                         # state -> cities; 'A/B' = the same city written two ways, the first is the name that is stored
    "Andhra Pradesh": "Visakhapatnam/Vizag/Visakhapatam, Vijayawada, Guntur, Nellore, Kurnool, Tirupati, Rajahmundry/Rajamahendravaram, Kakinada, Kadapa/Cuddapah, Anantapur/Anantapuram, Eluru, Ongole, Chittoor, Machilipatnam, Srikakulam, Vizianagaram, Amaravati, Proddatur, Madanapalle, Nandyal, Hindupur, Tenali, Bhimavaram, Guntakal, Dharmavaram, Adoni, Narasaraopet, Tadepalligudem",
    "Telangana": "Hyderabad, Secunderabad, Warangal, Karimnagar, Nizamabad, Khammam, Mahbubnagar, Nalgonda, Adilabad, Siddipet, Ramagundam, Suryapet, Medchal, Sangareddy",
    "Karnataka": "Bengaluru/Bangalore/Bengalore/Banglore, Mysuru/Mysore, Hubballi/Hubli, Dharwad, Mangaluru/Mangalore, Belagavi/Belgaum, Davanagere/Davangere, Ballari/Bellary, Tumakuru/Tumkur, Shivamogga/Shimoga, Udupi, Hassan, Mandya, Raichur, Bidar, Kalaburagi/Gulbarga, Vijayapura/Bijapur, Chitradurga, Kolar, Ramanagara, Chikkamagaluru/Chikmagalur, Hosapete/Hospet, Chikkaballapur, Koppal, Gadag, Haveri, Karwar, Sirsi, Bagalkot",
    "Tamil Nadu": "Chennai/Madras, Coimbatore, Madurai, Tiruchirappalli/Trichy/Tiruchirapalli, Salem, Tirunelveli, Tiruppur/Tirupur, Erode, Vellore, Thoothukudi/Tuticorin, Dindigul, Thanjavur, Hosur, Nagercoil, Kanchipuram, Karur, Cuddalore, Kumbakonam, Tiruvannamalai, Pollachi, Namakkal, Krishnagiri, Ooty, Sivakasi, Ambattur, Avadi, Tambaram, Chengalpattu, Virudhunagar, Karaikudi, Villupuram, Pudukkottai, Perambalur, Ariyalur, Theni, Ramanathapuram, Sivagangai, Nagapattinam, Tiruvallur",
    "Kerala": "Thiruvananthapuram/Trivandrum, Kochi/Cochin, Kozhikode/Calicut, Thrissur/Trichur, Kollam/Quilon, Kannur/Cannanore, Alappuzha/Alleppey, Palakkad/Palghat, Kottayam, Malappuram, Ernakulam, Kasaragod, Pathanamthitta, Idukki, Wayanad, Muvattupuzha, Perinthalmanna, Thalassery, Kayamkulam, Aluva, Angamaly",
    "Maharashtra": "Mumbai/Bombay, Navi Mumbai, Thane, Pune/Poona, Nagpur, Nashik/Nasik, Aurangabad/Chhatrapati Sambhajinagar, Solapur, Kolhapur, Amravati, Nanded, Sangli, Jalgaon, Akola, Latur, Dhule, Ahmednagar/Ahilyanagar, Satara, Ratnagiri, Kalyan, Vasai, Virar, Pimpri-Chinchwad/Pimpri Chinchwad/Pimpri/Chinchwad, Panvel, Bhiwandi, Mira Road, Ulhasnagar, Chandrapur, Parbhani, Beed, Jalna, Wardha, Yavatmal, Malegaon, Ichalkaranji, Badlapur, Ambernath, Dombivli, Lonavala, Baramati",
    "Delhi": "Delhi/Delhi NCR, New Delhi",
    "Uttar Pradesh": "Lucknow, Kanpur, Agra, Varanasi/Benares/Banaras, Prayagraj/Allahabad, Meerut, Bareilly, Aligarh, Moradabad, Gorakhpur, Saharanpur, Jhansi, Mathura, Firozabad, Muzaffarnagar, Ayodhya/Faizabad, Raebareli, Sultanpur, Etawah, Mirzapur, Basti, Jaunpur, Rampur, Shahjahanpur, Hapur, Sitapur, Budaun, Bulandshahr, Orai, Lakhimpur, Ballia, Noida, Greater Noida, Ghaziabad, Gautam Buddha Nagar, Unnao, Hardoi, Azamgarh, Deoria",
    "Haryana": "Gurugram/Gurgaon, Faridabad, Panipat, Ambala, Karnal, Rohtak, Hisar, Sonipat/Sonepat, Yamunanagar, Panchkula, Bhiwani, Sirsa, Bahadurgarh, Jind, Kurukshetra, Rewari, Palwal, Manesar, Fatehabad",
    "Punjab": "Ludhiana, Amritsar, Jalandhar, Patiala, Bathinda, Mohali/SAS Nagar, Pathankot, Hoshiarpur, Moga, Firozpur, Kapurthala, Sangrur, Barnala, Rupnagar",
    "Chandigarh": "Chandigarh",
    "Rajasthan": "Jaipur, Jodhpur, Udaipur, Kota, Ajmer, Bikaner, Alwar, Bhilwara, Sikar, Bharatpur, Pali, Sri Ganganagar, Tonk, Chittorgarh, Jhunjhunu, Hanumangarh",
    "Gujarat": "Ahmedabad, Surat, Vadodara/Baroda, Rajkot, Bhavnagar, Jamnagar, Gandhinagar, Junagadh, Anand, Nadiad, Morbi, Mehsana, Bharuch, Vapi, Navsari, Porbandar, Surendranagar, Bhuj, Valsad, Ankleshwar, Gandhidham, Palanpur, Himmatnagar",
    "Madhya Pradesh": "Bhopal, Indore, Gwalior, Jabalpur, Ujjain, Sagar, Dewas, Satna, Ratlam, Rewa, Katni, Chhindwara, Khandwa, Burhanpur, Singrauli, Morena, Vidisha, Mandsaur, Neemuch",
    "Chhattisgarh": "Raipur, Bhilai, Bilaspur, Korba, Durg, Rajnandgaon, Raigarh, Jagdalpur",
    "West Bengal": "Kolkata/Calcutta, Howrah, Durgapur, Asansol, Siliguri, Bardhaman/Burdwan, Kharagpur, Haldia, Barrackpore, Darjeeling, Malda, Berhampore, Krishnanagar, Bankura, Purulia, Cooch Behar, Jalpaiguri, Bidhannagar, Salt Lake, Dum Dum, Hooghly, Chandannagar",
    "Odisha": "Bhubaneswar, Cuttack, Rourkela, Berhampur/Brahmapur, Sambalpur, Puri, Balasore, Bhadrak, Baripada, Jharsuguda, Angul, Koraput, Paralakhemundi, Jeypore, Rayagada, Kendrapara, Jajpur",
    "Bihar": "Patna, Gaya, Bhagalpur, Muzaffarpur, Darbhanga, Purnia, Arrah, Begusarai, Katihar, Munger, Chapra, Bettiah, Sasaram, Hajipur, Samastipur, Bihar Sharif, Buxar, Banka, Siwan, Motihari, Nalanda",
    "Jharkhand": "Ranchi, Jamshedpur, Dhanbad, Bokaro, Hazaribagh, Deoghar, Giridih, Ramgarh, Adityapur, Phusro",
    "Assam": "Guwahati, Dibrugarh, Silchar, Jorhat, Nagaon, Tinsukia, Tezpur, Bongaigaon, Dhubri, Karimganj",
    "Uttarakhand": "Dehradun, Haridwar, Roorkee, Haldwani, Rishikesh, Rudrapur, Kashipur, Nainital",
    "Himachal Pradesh": "Shimla, Solan, Dharamshala, Mandi, Baddi, Kullu, Manali, Una, Hamirpur",
    "Jammu and Kashmir": "Srinagar, Jammu, Anantnag, Baramulla, Udhampur, Kathua",
    "Goa": "Panaji/Panjim, Margao, Vasco da Gama/Vasco, Mapusa, Ponda",
    "Puducherry": "Puducherry/Pondicherry, Karaikal",
    "Manipur": "Imphal", "Meghalaya": "Shillong", "Mizoram": "Aizawl", "Nagaland": "Kohima, Dimapur", "Tripura": "Agartala",
    "Arunachal Pradesh": "Itanagar", "Sikkim": "Gangtok", "Ladakh": "Leh",
}
STATE_NAMES = {"Andhra Pradesh": [], "Arunachal Pradesh": [], "Assam": [], "Bihar": [], "Chhattisgarh": [], "Goa": [], "Gujarat": [], "Haryana": [],
               "Himachal Pradesh": [], "Jharkhand": [], "Karnataka": [], "Kerala": [], "Madhya Pradesh": [], "Maharashtra": [], "Manipur": [],
               "Meghalaya": [], "Mizoram": [], "Nagaland": [], "Odisha": ["Orissa"], "Punjab": [], "Rajasthan": [], "Sikkim": [],
               "Tamil Nadu": ["Tamilnadu"], "Telangana": [], "Tripura": [], "Uttar Pradesh": [], "Uttarakhand": ["Uttaranchal"], "West Bengal": [],
               "Jammu and Kashmir": ["Jammu & Kashmir", "J&K"], "Ladakh": [], "Puducherry": [], "Chandigarh": [],
               "Andaman and Nicobar Islands": ["Andaman & Nicobar"], "Lakshadweep": [], "Delhi": ["NCT of Delhi"]}
FOREIGN_CITIES = {"United Arab Emirates": "Dubai, Abu Dhabi, Sharjah, Ajman, Al Ain", "Qatar": "Doha", "Saudi Arabia": "Riyadh, Jeddah, Dammam, Khobar, Jubail",
                  "Oman": "Muscat", "Kuwait": "Kuwait City", "Bahrain": "Manama", "Singapore": "Singapore", "Malaysia": "Kuala Lumpur, Penang",
                  "United Kingdom": "London, Manchester, Birmingham, Edinburgh, Glasgow, Leeds",
                  "United States": "New York, San Francisco, San Jose, Seattle, Dallas, Houston, Chicago, Boston, Atlanta, Los Angeles, Washington DC, Santa Clara, Sunnyvale",
                  "Canada": "Toronto, Vancouver, Montreal, Ottawa, Calgary", "Australia": "Sydney, Melbourne, Brisbane, Perth", "Germany": "Berlin, Munich, Frankfurt, Hamburg",
                  "Ireland": "Dublin", "Netherlands": "Amsterdam", "France": "Paris", "Sri Lanka": "Colombo", "Nepal": "Kathmandu", "Bangladesh": "Dhaka",
                  "South Africa": "Johannesburg, Cape Town", "Japan": "Tokyo", "New Zealand": "Auckland", "Switzerland": "Zurich", "Sweden": "Stockholm"}
COUNTRY_NAMES = {"India": [], "United States": ["USA", "U.S.A", "United States of America", "America"], "United Kingdom": ["UK", "U.K", "England", "Great Britain"],
                 "United Arab Emirates": ["UAE", "U.A.E"], "Saudi Arabia": ["KSA"], "Qatar": [], "Oman": [], "Kuwait": [], "Bahrain": [], "Singapore": [],
                 "Malaysia": [], "Canada": [], "Australia": [], "Germany": [], "France": [], "Netherlands": [], "Ireland": [], "Sri Lanka": [], "Nepal": [],
                 "Bangladesh": [], "Pakistan": [], "South Africa": [], "Japan": [], "China": [], "Hong Kong": [], "New Zealand": [], "Switzerland": [],
                 "Sweden": [], "Italy": [], "Spain": [], "Nigeria": [], "Kenya": [], "Egypt": [], "Ghana": [], "Indonesia": [], "Philippines": [],
                 "Thailand": [], "Vietnam": [], "Russia": [], "Brazil": [], "Mexico": [], "Poland": [], "Belgium": [], "Denmark": [], "Norway": [], "Finland": []}
AMBIGUOUS_PLACES = {"anand", "sagar", "kota", "salem", "hassan", "pali", "una", "mandi", "rewa", "jind", "orai", "basti", "banka", "puri", "tonk", "moga",
                    "leh", "ponda", "kolar", "nalanda", "wayanad", "ariyalur"}
_PLACES = {}                                        # lower-case written form -> (kind, stored name, state or country)
for _state, _cities in CITY_TABLE.items():
    for _entry in _cities.split(","):
        _names = [n.strip() for n in _entry.split("/") if n.strip()]
        for _n in _names:
            _PLACES.setdefault(_n.lower(), ("city", _names[0], _state))
for _state, _aliases in STATE_NAMES.items():
    for _n in [_state] + _aliases:
        _PLACES.setdefault(_n.lower(), ("state", _state, ""))
for _country, _cities in FOREIGN_CITIES.items():
    for _n in [x.strip() for x in _cities.split(",") if x.strip()]:
        _PLACES.setdefault(_n.lower(), ("city", _n, _country))
for _country, _aliases in COUNTRY_NAMES.items():
    for _n in [_country] + _aliases:
        _PLACES.setdefault(_n.lower(), ("country", _country, ""))
_PLACES["delhi"] = ("city", "Delhi", "Delhi")
_PLACE_RE = re.compile(r"(?<![A-Za-z])(" + "|".join(re.escape(n) for n in sorted(_PLACES, key=len, reverse=True)) + r")(?![A-Za-z])", re.IGNORECASE)
_PLACE_ABBR_RE = re.compile(r"(?<![A-Za-z.])(USA|U\.S\.A|UK|U\.K|UAE|U\.A\.E|KSA)(?![A-Za-z])")
LOCATION_LABEL_RE = re.compile(r"^\W*(?P<kind>permanent|present|current|residential|correspondence|local|home|mailing|native)?\s*"
                               r"(?:address|location|city|residence|native place|hometown|place of residence)\s*(?:\([^)]*\))?\s*[:\-–]\s*(?P<rest>.*)$", re.IGNORECASE)
PIN_RE = re.compile(r"\b[1-9]\d{2}\s?\d{3}\b")
ADDR_WORD_RE = re.compile(r"\b(?:h\.?\s?no|house|flat|door|plot|street|road|nagar|colony|layout|sector|village|near|opp|opposite|post|dist|district|taluk|mandal|floor|apartment|apt|cross|main)\b", re.IGNORECASE)


def match_start(match):
    """Sort key: where a regex match starts in the text."""
    return match.start()


def find_places(text, trusted=False):
    """The places written in a text, in order: [(kind, stored name, state or country)]. A city that is also a common name or word ('Sagar', 'Anand')
    counts only when the text is trusted (a labelled address, a PIN code, a spreadsheet cell)."""
    text = re.sub(r"\S+@\S+|https?://\S+|www\.\S+", " ", text or "")
    hits = []
    for match in sorted(list(_PLACE_RE.finditer(text)) + list(_PLACE_ABBR_RE.finditer(text)), key=match_start):
        key = match.group(1).lower()
        if key in AMBIGUOUS_PLACES and not trusted:
            continue
        if re.match(r"\s*(?:road|rd\b|nagar|layout|colony|street|st\b|cross|main\b|junction|bazar|bazaar|market|circle|extension|enclave|puram|palya|halli|gate|chowk)",
                    text[match.end():match.end() + 14], re.IGNORECASE):
            continue
        hits.append(_PLACES.get(key) or ("country", {"usa": "United States", "u.s.a": "United States", "uk": "United Kingdom", "u.k": "United Kingdom", "uae": "United Arab Emirates",
                                                    "u.a.e": "United Arab Emirates", "ksa": "Saudi Arabia"}.get(key, match.group(1)), ""))
    return hits


def compose_location(blocks):
    """blocks = [hits of a text, ...] in priority order -> {'city', 'state', 'country', 'cities', 'states', 'countries', 'text'}."""
    cities, states, countries = [], [], []
    city = state = country = ""
    for hits in blocks:
        for kind, name, extra in hits:
            if kind == "city":
                if name not in cities:
                    cities.append(name)
                if extra and extra in STATE_NAMES and extra not in states:
                    states.append(extra)
                if extra and extra not in STATE_NAMES and extra not in countries:
                    countries.append(extra)                          # a city outside India: its country
                elif extra in STATE_NAMES and "India" not in countries:
                    countries.append("India")
            elif kind == "state":
                if name not in states:
                    states.append(name)
                if "India" not in countries:
                    countries.append("India")
            elif kind == "country" and name not in countries:
                countries.append(name)
    for hits in blocks:                                              # the first block with a city decides city, state and country together
        first_city = next(((n, entry) for k, n, entry in hits if k == "city"), None)
        if first_city:
            city = first_city[0]
            state = next((n for k, n, entry in hits if k == "state"), "") or (first_city[1] if first_city[1] in STATE_NAMES else "")
            country = "India" if state else (next((n for k, n, entry in hits if k == "country"), "") or (first_city[1] if first_city[1] not in STATE_NAMES else ""))
            break
    if not city:
        state = states[0] if states else ""
        country = countries[0] if countries else ""
    parts = [item for item in (city, state, country) if item]
    return {"city": city, "state": state, "country": country, "cities": cities[:4], "states": states[:4], "countries": countries[:4], "text": ", ".join(parts)}


def location_from_pin(lines, name_words):
    """A PIN code is written where an address ends: the place names just before it ('Ariyalur(DT)Tamilnadu-612903') are the city and the state."""
    texts = [clean(line.text) for line in lines]
    states = {re.sub(r"[^a-z]", "", style_key.lower()): style_key for style_key in STATE_NAMES}
    for i, part_text in enumerate(texts):
        match = PIN_RE.search(part_text)
        if not match or re.search(r"\b(?:19|20)\d{2}\b", match.group(0)) or DATE_RANGE_RE.search(part_text):
            continue
        before = part_text[:match.start()]
        if len(re.findall(r"[A-Za-z]{3,}", before)) < 2 and i:
            before = texts[i - 1] + " , " + before                       # 'Tamilnadu-612903' alone on its line: the district is on the line above
        before = re.sub(r"\((?:po|tk|dt|dist|district|tq|taluk|tehsil|ps|via)\)", ", ", before, flags=re.IGNORECASE)
        parts = [piece.strip(" .:-") for piece in re.split(r"[,;/()]|\s[-–]\s|-(?=$)", before) if piece.strip(" .:-")]
        parts = [re.sub(r"^(?:permanent|present|current|residential|correspondence)?\s*address\s*[:\-]*\s*", "", piece, flags=re.IGNORECASE).strip() for piece in parts]
        parts = [piece for piece in parts if piece and piece.lower() not in ("india",)]
        state = city = ""
        for k, part in enumerate(parts):
            key = re.sub(r"[^a-z]", "", part.lower())
            if key in states:
                state = states[key]
                city = parts[k - 1] if k else ""
                break
            tail = part.split()[-1] if part.split() else ""
            if re.sub(r"[^a-z]", "", tail.lower()) in states and tail:
                state = states[re.sub(r"[^a-z]", "", tail.lower())]
                city = " ".join(part.split()[:-1]) or (parts[k - 1] if k else "")
                if city.lower() in ("new", "north", "south", "east", "west", "old", "greater"):
                    city = part.title()                             # 'NEW DELHI-110045': the city is New Delhi, not 'NEW'
                break
        if not city and parts and not state:
            city = parts[-1]
        city = city.strip(" .,-")
        if not city or len(city.split()) > 3 or not re.fullmatch(r"[A-Za-z][A-Za-z .'-]+", city) or ADDR_WORD_RE.search(city) or city.lower() in name_words                 or re.search(r"\b(?:road|rd|street|st|lane|nagar|colony)\b", city, re.IGNORECASE):
            if not state:
                continue
            city = ""
        if not state:                                               # no state written: the city must be one the place list knows, else it may be a street or a village
            hits = find_places(" ".join(parts[-3:]), trusted=True)
            return compose_location([hits]) if any(k == "city" for k, _, _ in hits) else None
        if re.search(r"\b(?:dt|tq|po|ps|dist|tk)\b|[.0-9]", city, re.IGNORECASE) or any(word[:1].islower() for word in city.split()):
            city = ""                                               # 'Shirol. Dt.Gadag Tq.Nargund' is not one city name
        text = ", ".join(item for item in (city, state, "India") if item)
        return {"city": city, "state": state, "country": "India", "cities": [city] if city else [], "states": [state] if state else [], "countries": ["India"], "text": text}
    return None


def find_location(lines, sections, name):
    """City, state and country of a resume. Only the places of an address are used: the contact block at the top and lines written as
    'Address:' / 'Location:' / 'Current location:' (with the line or two after them). The cities of jobs and colleges are not used."""
    name_words = {word.lower() for word in re.split(r"[^A-Za-z]+", name or "") if len(word) >= 3}
    texts = [text_line.text for text_line in lines]
    labelled, header = [], []
    for i, part_text in enumerate(texts):
        match = LOCATION_LABEL_RE.match(part_text)
        if match:
            block = match.group("rest")
            for next_text in texts[i + 1:i + 3]:
                if LOCATION_LABEL_RE.match(next_text) or heading_section(next_text) or len(next_text.split()) > 12 or DATE_RANGE_RE.search(next_text):
                    break
                block += " " + next_text
            weight = 2 if (match.group("kind") or "").lower() in ("permanent", "home", "native") else 0
            labelled.append((weight, block))
    for line in sections.get("top", [])[:25]:
        part_text = line.text
        if heading_section(part_text) or DATE_RANGE_RE.search(part_text) or len(part_text.split()) > 14 or re.match(r"^\W*place\s*[:\-–]", part_text, re.IGNORECASE):
            continue                                                # a heading, a date, a long line, the 'Place:' of a declaration
        if (INSTITUTION_RE.search(part_text) or COMPANY_SUFFIX_RE.search(part_text)) and not PIN_RE.search(part_text):
            continue
        if re.search(r"\b(?:19|20)\d{2}\b|\b\d{2}\s*[-–]\s*\d{2}\b", part_text) and not PIN_RE.search(part_text):
            continue                                                # a line with a year belongs to a job or a study
        if TITLE_HEAD.search(part_text) and not PIN_RE.search(part_text):
            continue                                                # a job title line ('... Noida India Lead Consultant')
        if re.match(r"^\W*[A-Z][a-z]+ed\b", part_text):
            continue                                                # a sentence ('Visited ... at Ahmedabad')
        header.append((1, part_text))
    blocks = []
    for weight, text in sorted(labelled + header, key=first_of_pair):
        words = text.split()
        for word in set(words):
            if word.lower().strip(",.") in name_words:
                text = re.sub(r"\b" + re.escape(word) + r"\b", " ", text)       # the candidate's own name is not a place ('Sagar Patil')
        addr_like = bool(PIN_RE.search(text)) or bool(ADDR_WORD_RE.search(text))
        trusted = addr_like or weight != 1
        hits = find_places(text, trusted=trusted)
        if hits and weight == 1 and not addr_like:
            tokens = re.findall(r"[A-Za-z]+", re.sub(r"\S+@\S+|https?://\S+|www\.\S+", " ", text))
            if not (re.search(r"[,|•]", text) or sum(len(heading[1].split()) for heading in hits) * 2 > len(tokens)):
                hits = []                                            # a place word inside a sentence or a title, not an address
        if hits:
            blocks.append(hits)
    loc = compose_location(blocks)
    if not loc.get("city") and not loc.get("state"):
        by_pin = location_from_pin(lines, name_words)
        if by_pin:
            loc = by_pin
    if not loc.get("city"):                                          # a place the list does not know: 'Nainwa, Bundi, India' written alone in the header
        for line in sections.get("top", [])[:12]:
            match = re.fullmatch(r"\s*([A-Z][A-Za-z.\-]+(?:\s[A-Z][A-Za-z.\-]+)?)\s*,\s*([A-Z][A-Za-z.\-]+(?:\s[A-Z][A-Za-z.\-]+)?)\s*,\s*India\s*", line.text)
            if match and not INSTITUTION_RE.search(line.text) and match.group(1).lower() not in name_words:
                state = match.group(2) if match.group(2) in STATE_NAMES else ""
                loc = dict(loc, city=match.group(1), state=state or loc.get("state", ""), country="India", cities=[match.group(1)],
                           states=[state] if state else loc.get("states", []), countries=["India"], text=clean(line.text))
                break
    return loc


URL_RE = re.compile(r"(?:https?://|www\.)[^\s<>\"'\]\)]+|\b(?:[a-z0-9\-]+\.)*(?:linkedin\.com|github\.com|gitlab\.com|behance\.net|dribbble\.com|medium\.com|kaggle\.com|"
                    r"leetcode\.com|hackerrank\.com|codechef\.com|stackoverflow\.com|twitter\.com|x\.com|facebook\.com|instagram\.com|youtube\.com|youtu\.be|"
                    r"bit\.ly|drive\.google\.com|notion\.site|github\.io|netlify\.app|vercel\.app|herokuapp\.com)/[^\s<>\"'\]\)]*", re.IGNORECASE)
EMAIL_HOSTS = ("gmail.com", "yahoo.com", "yahoo.co.in", "outlook.com", "hotmail.com", "rediffmail.com", "live.com")


def split_link_parts(u):
    host, _, path = re.sub(r"^https?://(?:www\.)?", "", u.lower()).rstrip("/").partition("/")
    return host, path


def find_links(lines, extra=()):
    """Every web link of a resume, once each: written in the text (also when the PDF wrapped it over two lines) or behind a clickable word."""
    texts = [line.text for line in lines]
    pool = list(texts) + [a + b for a, b in zip(texts, texts[1:]) if a.endswith(("/", "-", "=", "_")) and " " not in a.split()[-1]] + [str(u) for u in extra]
    out, seen = [], set()
    for text in pool:
        for match in URL_RE.finditer(text):
            if match.start() and text[match.start() - 1] == "@":
                continue
            url = match.group(0).strip().rstrip(".,;:)]}'\">")
            if not re.match(r"https?://", url, re.IGNORECASE):
                url = "https://" + url
            host = re.sub(r"^https?://([^/]+).*", r"\1", url, flags=re.IGNORECASE).lower()
            if "." not in host or host.endswith(EMAIL_HOSTS) or len(url) < 12 or re.search(r"(?:^|\.)linkedin\.com$", host):
                continue                                             # LinkedIn has its own field ('linkedin'), it is not in 'links'
            key = re.sub(r"^https?://(?:www\.)?", "", url.lower()).rstrip("/")
            if re.fullmatch(r"(?:[a-z0-9\-]+\.)*(?:linkedin\.com/in|linkedin\.com|github\.com)", key):
                continue
            if key not in seen:
                seen.add(key)
                out.append(url)
    out = [u for u in out if not (split_link_parts(u)[1].endswith(("-", "_", "=")) and any(split_link_parts(v)[0] == split_link_parts(u)[0] and split_link_parts(v)[1] != split_link_parts(u)[1]
                                                                               and split_link_parts(v)[1].startswith(split_link_parts(u)[1]) for v in out))]       # the first half of a wrapped link
    return out


CURRENT_STATEMENT_RE = re.compile(r"\b(?i:currently|presently)\s+(?:(?i:i|we)\s+(?i:am|was|have been|had been)\s+)?(?i:working|employed|serving|associated)\s+(?:(?i:as)\s+(?:(?i:an?)\s+)?(?P<title>[^,;()|]+?)\s+)?(?i:with|at|in|for)\s+(?P<company>[A-Z0-9][^,;()|]*?)"
                                  r"(?=\s*(?:\(|,|;|\||\s(?i:from|since|for\s+\d|till|w\.?e\.?f)\b|\s[-–—]\s|$))")


PAST_STATEMENT_RE = re.compile(r"\b(?i:worked|served)\s+(?:(?i:as)\s+(?:(?i:an?)\s+)?(?P<title>[^,;()|]+?)\s+)?(?i:with|at|in|for)\s+(?P<company>[A-Z0-9][^,;()|]*?)"
                               r"(?=\s*(?:\(|,|;|\||\s(?i:from|since|during)\b|\s[-–—]\s|\s(?i:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*['’]?\d|\s(?:19|20)\d{2}|$))")
CURRENT_LABEL_RE = re.compile(r"^\W*(?i:current|present)\s+(?i:company|employer|organi[sz]ation)(?:\s+(?i:name))?\s*[:\-–]\s*(?P<company>[A-Z0-9][^()|;]*?)\s*(?:\(|\||;|$)")
TILL_DATE_WITH_RE = re.compile(r"(?i:till date|to date|till now|present)\s*[:\-–]\s*(?i:with|at)\s+(?P<company>[A-Z0-9][^,;()|]*?)(?=\s+(?i:as)\s|\s*[,;(|]|$)")
CURRENT_NO_YEAR_RE = re.compile(r"(?P<company>[A-Z0-9][^,;()|]*?)\s+(?i:from|since)\s+(?i:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(?i:till|to|until|upto|up to)\s+(?i:present|date|now|today)")


def apply_past_statements(experience, lines):
    """'Worked with METALLIC EQUIPMENT L.L.C Aug 2015 to Dec 2017': a dated job in a sentence. Added when the company is not already one of the jobs."""
    for line in lines:
        text = clean(line.text)
        match = PAST_STATEMENT_RE.search(text)
        if not match:
            continue
        start, end, months, now = dates_of_text(text)
        if not start or now:
            continue
        company = employer_name(match.group("company").strip(" ,;:").rstrip("."), "sentence")
        if not company:
            continue
        key = norm_key(company)
        if any(entry.get("company") and (norm_key(entry["company"]) == key or key in norm_key(entry["company"]) or norm_key(entry["company"]) in key) for entry in experience):
            continue
        experience.append(new_job(clean(match.group("title") or "") or None, company, "sentence", start, end, months, False))
    return experience


GENERIC_COMPANY_END = {"development", "automation", "management", "testing", "engineering", "design", "scripts", "analysis", "deployment", "integration",
                       "monitoring", "migration", "implementation", "administration", "infrastructure", "applications", "application", "services"}


def current_company_ok(piece):
    """A piece of a line near a 'Present' date that may be the company: not a place ('New Delhi'), not a sentence or a bullet that starts with a verb
    ('Engineered infrastructure automation scripts'), not a skill or an activity ('End-to-End Android Development'). A company suffix (Pvt, Ltd,
    Technologies ...) always passes."""
    words = piece.split()
    if not words or len(words) > 6 or COMPANY_SUFFIX_RE.search(piece):
        return bool(words) and len(words) <= 8
    word_lists = lookups()
    return not (is_place_text(piece) or not (words[0][:1].isupper() or (words[0][:1].isdigit() and len(words) == 1)) or words[0].lower().endswith(("ed", "ing"))
                or words[-1].lower().strip(".,") in GENERIC_COMPANY_END or find_terms(piece, word_lists.skills))


CURRENT_END_RE = re.compile(r"^(?:present|current|up\s?to|till|now|ongoing|date|today)", re.IGNORECASE)


OPEN_END_RE = re.compile(r"(?P<start>(?:" + _DATE_PART + r"))\s*(?:[-–—]+|\bto\b|\buntil\b|\btill\b|\bupto\b|\bup to\b)\s*(?P<end>[^\d@]{0,40})$", re.IGNORECASE)


def current_date_span(text):
    """(start (year, month), (from, to) positions of the date text) when the text holds a START date and no real END date: the end is written as a word
    ('Present', 'Currently', 'Continuing', 'Onwards', 'Till date', 'Ongoing' ...), or is missing ('Jan 2023 -'), or the text says 'Since Jan 2023'. A closed
    range with two dates, or an end that is a place, a title or a company, is not current. Returns None otherwise."""
    match = DATE_RANGE_RE.search(text)
    if match:
        if CURRENT_END_RE.match(match.group("end").strip()):
            start = parse_date(match.group("start"))
            return (start, match.span()) if start else None
        return None                                                 # two real dates: a finished job
    mo = OPEN_END_RE.search(text.strip())
    if mo:
        end = clean(mo.group("end")).strip(" |,;.()[]")
        words = end.split()
        low = {word.lower().strip(".,") for word in words}
        if not end or (len(words) <= 3 and not re.search(r"\b(?:" + MONTHS + r")\b", end, re.IGNORECASE) and not (low & INDIAN_PLACES)
                       and not TITLE_HEAD.search(end + " ") and not COMPANY_SUFFIX_RE.search(end)):
            start = parse_date(mo.group("start"))
            if start:
                s = text.strip()
                off = text.find(s)
                return start, (off + mo.start(), off + mo.end())
    sm = SINCE_RE.search(text) if len(text.split()) <= 10 else None
    if sm and parse_date(sm.group(1)):
        return parse_date(sm.group(1)), sm.span()
    return None


def apply_current_dates(experience, lines, experience_lines=None, person=None):
    """A START date with no real END date (the end is 'Present', 'Currently', 'Continuing', 'Onwards' ... any word, or nothing) makes that job the current
    one, wherever the date is written: on the same line as the company ('Junior Software Engineer | [Royals Webtech Pvt Ltd Nagpur] | [Jan,2025] - Present'),
    above or below it, in a column to the left or right (the PDF then puts the date lines far from the company). The nearest line within two lines that names
    a company is the employer; when none does, the first company of the experience section. The nearest title line is its title. Only used when no
    current job is known yet."""
    if any(entry.get("current") and entry.get("company") and employer_name(entry["company"], entry.get("company_source", "")) and not bad_employer(entry["company"], person) for entry in experience):
        return experience                                           # a current job with a real company is already known: this rule only rescues a resume that has none
    texts = prepare_date_lines([line.text for line in lines])
    word_lists = lookups()
    in_experience = {id(line) for line in (experience_lines or [])}
    for i, text in enumerate(texts):
        found = current_date_span(text)
        if not found:
            continue
        start, span = found
        near = " ".join(texts[max(0, i - 1):i + 2])
        if SKIP_LABEL_RE.match(text) or (DEGREE_RE.search(near) and INSTITUTION_RE.search(near)):
            continue                                                # a project of a client, or a study period: not a job
        company = title = None
        forced = False
        for doc in (0, 1, -1, 2, -2):                                 # the date line first, then the lines next to it, nearest first
            j = i + doc
            if not 0 <= j < len(texts) or not texts[j].strip():
                continue
            segment = texts[j][:span[0]] + " | " + texts[j][span[1]:] if doc == 0 else texts[j]
            if doc and (current_date_span(segment) or DATE_RANGE_RE.search(segment) or len(segment.split()) > 12 or PROSE_START_RE.match(segment) or heading_section(segment)):
                continue
            for piece in re.split(r"\s*[|\[\]]\s*|\s+,\s+|\s{2,}", segment):
                piece = clean(piece).strip(" .,:;-–—()")
                if not piece or len(piece) < 2:
                    continue
                key = norm_key(piece)
                if title is None and (key in word_lists.titles or (TITLE_HEAD.search(piece.split()[-1] + " ") and len(piece.split()) <= 6
                                                          and not re.search(r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation)\b", piece, re.IGNORECASE))):
                    title = piece
                elif company is None and not bad_employer(piece, person) and current_company_ok(piece) and (employer_name(piece, "suffix") or fallback_employer(piece, "spacy", True)):
                    company = employer_name(piece, "suffix") or piece
            if company and title:
                break
        if not company:                                             # last resorts, nearest line first: a bracketed name above the date, then the client the job is for
            for doc in (-1, 1, -2, 2, 0):
                j = i + doc
                if not 0 <= j < len(texts):
                    continue
                mp = re.search(r"\(([A-Z0-9][A-Za-z0-9 .&'-]{2,30})\)\s*$", texts[j]) if doc < 0 else None             # 'E-commerce (8Farmer)' / '04/2025 - Present'
                company_match = re.search(r"\bclient(?:\s*name)?\s*:\s*([A-Z0-9][^|\n]{2,40}?)\s*(?:\||$)", texts[j], re.IGNORECASE)    # 'Client: State Farm USA | Project: ...'
                guess = clean((mp or company_match).group(1)).strip(" .,:;") if (mp or company_match) else ""
                if guess and (current_company_ok(guess) or (company_match and not mp and len(guess.split()) <= 6 and guess[:1].isupper())):
                    company, forced = guess, True                   # a name written right after 'Client:' needs no further test
                    break
        if not company and experience_lines and (id(lines[i]) in in_experience or i < 3):      # the date is in a column away from the company (or read first by the PDF): the only company of the section
            names = []
            for line in experience_lines:
                for piece in re.split(r"\s*[|\[\]]\s*|\s+,\s+|\s{2,}", clean(line.text)):
                    piece = clean(piece).strip(" .,:;-–—()")
                    if piece and COMPANY_SUFFIX_RE.search(piece) and current_company_ok(piece) and employer_name(piece, "suffix"):
                        n = employer_name(piece, "suffix")
                        if norm_key(n) not in {norm_key(item) for item in names}:
                            names.append(n)
            if len(names) == 1:
                company = names[0]                                  # only one company in the section: the date cannot belong to another
        if not company or not (forced or title or COMPANY_SUFFIX_RE.search(company)):
            continue                                                # no company, or a bare name with no title and no company suffix beside the date
        ckey = norm_key(company)
        for entry in experience:
            if entry.get("current") and not (entry.get("company") and employer_name(entry["company"], entry.get("company_source", "")) and not bad_employer(entry["company"], person)):
                entry["current"] = False                                # a 'current' job whose company is not a company ('Scripting & OS: Linux') is not current
        job = next((entry for entry in experience if entry.get("company") and (norm_key(entry["company"]) == ckey or ckey in norm_key(entry["company"]) or norm_key(entry["company"]) in ckey)), None)
        if job is None:
            job = new_job(title, company, "structure", "%04d-%02d" % start, "", 0, True)
            experience.insert(0, job)
        else:
            job["end"] = job.get("end") or ""
        job["current"] = True
        job["start"] = job.get("start") or "%04d-%02d" % start
        job["title"] = job.get("title") or title
        break                                                       # one current employer: the first (top-most) match
    return experience


def apply_current_statements(experience, lines):
    """A sentence that says the candidate works somewhere NOW names the current employer: 'Currently working with M.K. INTERIOR, from Jan 2018.'. A job
    of the same company becomes the current one; a job that only has the same start date was given the wrong company, so it gets this one; else the
    job is added."""
    for line in lines:
        match = CURRENT_LABEL_RE.match(clean(line.text)) or CURRENT_STATEMENT_RE.search(clean(line.text)) or TILL_DATE_WITH_RE.search(clean(line.text)) or CURRENT_NO_YEAR_RE.search(clean(line.text))
        if not match:
            continue
        company = match.group("company").strip(" ,;:")
        while len(company.split()) > 1 and company.split()[-1].lower().strip(",.") in INDIAN_PLACES:
            company = company.rsplit(None, 1)[0]                   # 'Perigon Technologies Hyderabad': the place is not part of the name
        if "title" not in match.groupdict():
            company = re.sub(r"^(?:(?i:currently|presently)\s+)?(?:(?i:working|employed)\s+)?(?:(?i:with|at|in|for)\s+)?", "", company)
        if company.endswith(".") and not re.search(r"\b(?:ltd|pvt|inc|co|corp|llc|l\.l\.c)\.$", company, re.IGNORECASE):
            company = company.rstrip(".")
        company = employer_name(company, "sentence")
        if not company:
            continue
        start, end, months, now = dates_of_text(clean(line.text))
        title = clean(match.groupdict().get("title") or "") or None
        key = norm_key(company)
        same = [entry for entry in experience if entry.get("company") and (norm_key(entry["company"]) == key or key in norm_key(entry["company"]) or norm_key(entry["company"]) in key)]
        if same:
            job = same[0]
        else:
            job = next((entry for entry in experience if start and entry.get("start") == start), None)
            if job:
                job["company"], job["company_source"] = company, "sentence"
            else:
                job = new_job(title, company, "sentence", start, end, months, True)
                experience.insert(0, job)
        job["current"] = True
        job["start"] = job.get("start") or start
        job["end"] = end or job.get("end") or ""
        job["title"] = job.get("title") or title
        for other in experience:
            if other is not job and other.get("current") and other.get("company") and norm_key(other["company"]) != key and start and other.get("start") == start:
                other["current"] = False
    return experience


PRIVATE_HEADING_RE = re.compile(r"personal|declaration|reference|referee|signature|particulars|\bplace\b|\bdate\b|\bbio.?data\b", re.IGNORECASE)
PRIVATE_DETAIL_RE = re.compile(r"\b(?:nationality|passport|marital|date of birth|dob|gender|sex|driving licen[sc]e|visa|aadhaar|aadhar|pan\b|religion|father|mother|permanent address|place\s*:|date\s*:)", re.IGNORECASE)


def close_other_section(*, buf, current, out):
    """Stores the lines collected for the section that is being read."""
    if current and buf and not PRIVATE_HEADING_RE.search(current):
        text = " ".join(part_text for part_text in buf if part_text).strip()
        text = PRIVATE_DETAIL_RE.split(text)[0].strip(" ,;:-")             # 'Languages: English, Hindi  Nationality: Indian  Passport ...': personal details are never kept
        entry = {"heading": current, "text": text}
        if text and entry not in out:
            out.append(entry)


def named_other_sections(lines):
    """Every part of the resume that is not one of the main sections (hobbies, awards, references, volunteering, declaration, an unknown heading) under the
    heading the candidate wrote: [{'heading': 'Hobbies', 'text': 'Reading, Cricket'}]. A section that has its own field (skills, experience, education,
    projects, certifications ...) is not repeated here. Nothing is cut by length or by number of lines: the part runs to the next heading."""
    sizes = collections.Counter(line.size for line in lines if line.size)
    body_size = sizes.most_common(1)[0][0] if sizes else 0.0
    out, current, buf = [], None, []

    for line in lines:
        text = clean(line.text)
        found = heading_section(text, line.bold, line.size, body_size) or vocab_heading(text)
        if found and len(text.split()) <= 8:
            close_other_section(buf=buf, current=current, out=out)
            buf = []
            current = text.rstrip(":-–— ").strip() if found == "other" else None
            continue
        if ":" in text and current is None:
            head, rest = text.split(":", 1)
            if heading_section(head, True) == "other" and rest.strip() and not PRIVATE_HEADING_RE.search(head):
                entry = {"heading": head.strip(), "text": PRIVATE_DETAIL_RE.split(rest.strip())[0].strip(" ,;:-")}
                if entry["text"] and entry not in out:
                    out.append(entry)
                continue
        if current:
            buf.append(text)
    close_other_section(buf=buf, current=current, out=out)
    return out


COLUMN_COMPANY_RE = re.compile(r"^\W*(?:name\s+of\s+(?:the\s+)?(?:company|organi[sz]ation|firm|employer)|company\s+name|organi[sz]ation\s+name|employer\s+name)\W*$", re.IGNORECASE)
COLUMN_HEADER_RE = re.compile(r"^\W*(?:sr\.?\s*no|s\.?\s*no|serial|designation|job profile|period|duration|location|role|project|experience|education|skills?)\b", re.IGNORECASE)


def column_companies(lines):
    """Jobs written as a table with a 'Name of the Company' header. Each job is a row: the company, the title, the description and the period, and a
    row ends with its number ('1', '2' ...). The company is the first short line of the row, and the row is current when its period says till date /
    present. Returns [(company, current)] in the order of the table."""
    texts = [clean(line.text) for line in lines]
    out = []
    for i, text in enumerate(texts):
        if not COLUMN_COMPANY_RE.match(text):
            continue
        row = []
        for u in texts[i + 1:i + 150]:
            if not u or COLUMN_HEADER_RE.match(u):
                continue
            if re.fullmatch(r"\d{1,2}", u):                         # the number that closes a row
                company = next((employer_name(item, "label") for item in row[:2] if employer_name(item, "label")), None)
                if company:
                    current = any(re.search(r"\b(?:present|till date|till now|to date)\b", item, re.IGNORECASE) for item in row)
                    if norm_key(company) not in {norm_key(c) for c, _ in out}:
                        out.append((company, current))
                row = []
                continue
            if COLUMN_COMPANY_RE.match(u) or heading_section(u) or PROSE_START_RE.match(u) or len(u.split()) > 14:
                break
            row.append(u)
        break
    return out


ACADEMIC_TABLE_HEADER_RE = re.compile(r"^\W*academic\s+records?\W*$", re.IGNORECASE)
DEGREE_TOKEN_RE = re.compile(r"^[A-Z][A-Za-z.]{1,12}(?:\s[A-Z][A-Za-z.]{1,12})?$")


def academic_table_rows(lines):
    """The 'ACADEMIC RECORD' table written one cell per line: 'Course' / 'Institution' / 'University/Board' ... then, row by row, the course (a short
    code such as 'B.Arch' or 'PGDUPDL', or a degree name) and the institution a few lines below it ('PDA College of Engineering,Gulbarga(Autonom' ...
    'IGNOU,DELHI'). Each row is one degree with its institution."""
    texts = [clean(line.text) for line in lines]
    rows = []
    for i, text in enumerate(texts):
        if not ACADEMIC_TABLE_HEADER_RE.match(text):
            continue
        for j in range(i + 1, min(len(texts), i + 60)):
            u = texts[j]
            if re.match(r"^\W*(?:professional|technical|personal|achievements?|skills|projects?|experience|declaration|references?|hobbies)\b", u, re.IGNORECASE):
                break
            if not (DEGREE_RE.search(u) or DEGREE_TOKEN_RE.match(u)) or len(u.split()) > 2 or INSTITUTION_RE.search(u):
                continue                                            # a course is a short code or a 1-2 word degree; a longer phrase is a description
            if re.search(r"university|board|technological", texts[j + 1] if j + 1 < len(texts) else "", re.IGNORECASE) or re.match(r"post\s*graduation$", u, re.IGNORECASE):
                continue                                            # 'Visvesvaraya' over 'Technological University' is the university column; 'Post graduation' describes the course above
            inst = None
            for k in range(j + 1, min(len(texts), j + 7)):
                if DEGREE_TOKEN_RE.match(texts[k]) and not INSTITUTION_RE.search(texts[k]):
                    break                                           # the next course starts here
                if re.search(r"\b(?:college|university|institute|institution|ignou|school|board)\b", texts[k], re.IGNORECASE) and not DEGREE_RE.match(texts[k]):
                    inst = texts[k]
                    break
            if not inst:
                continue
            inst = re.split(r"\s*\(|\s+(?:visvesvaraya|technological)\b", inst, flags=re.IGNORECASE)[0]
            inst = re.sub(r",(?=\S)", ", ", inst).strip(" ,.-")
            if real_institution(inst) or re.search(r"college|university|institute|ignou", inst, re.IGNORECASE):
                rows.append({"degree": u.strip(" .:"), "institution": inst, "year": "", "start_year": None, "end_year": None, "score": None, "field": ""})
    out, seen = [], set()
    for r in rows:
        key = norm_key(r["degree"])
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


PROJECT_JUNK_RE = re.compile(r"^\W*(?:name of|designation|job profile|experience|total experience|period|duration|from\s|sr\.?\s*no|serial|tenure|team size|"
                             r"client|role|description|responsibilit|technolog|\d)", re.IGNORECASE)
LABELLED_PROJECT_RE = re.compile(r"^\W*project\s*(?:title|name)\s*[:\-–]+\s*(?P<n>[A-Za-z0-9][^|]{2,120})$", re.IGNORECASE)


def labelled_projects(lines):
    """Projects written with a label, wherever they sit in the resume (the section may have no 'Projects' heading): 'Project title: Kubernetes Deployment on
    Amazon EKS', 'Project Name: ARISg'. The text after the label is the project; nothing is cut by the number of projects."""
    out = []
    for line in lines:
        match = LABELLED_PROJECT_RE.match(clean(line.text))
        if not match:
            continue
        name = clean(match.group("n")).strip(" .,;:-")
        if name and len(name.split()) <= 14 and not any(norm_key(name) == norm_key(part["name"]) for part in out):
            out.append({"name": name, "skills": [], "description": "", "link": ""})
    return out


SEMESTER_PROJECT_RE = re.compile(r"^\W*\d{1,2}(?:st|nd|rd|th)\s+semester\b\W*$", re.IGNORECASE)


def close_semester_project(*, buf, out):
    """Stores the lines collected for the project that is being read."""
    if buf:
        name = clean(" ".join(buf)).strip(" .,;:-")
        if name and not any(norm_key(name) == norm_key(part["name"]) for part in out):
            out.append({"name": name[:200], "skills": [], "description": "", "link": ""})


def semester_projects(lines):
    """Academic projects written as a list by semester: '1st semester' / 'Graphics, Concepts of Visual Arts & Basic' / 'design.' / '2nd semester' /
    'Residence, Detail study of functional aspects' / 'of a residence.'. The text under each semester heading (however many lines) is one project."""
    out, buf = [], None

    for line in lines:
        text = clean(line.text)
        if SEMESTER_PROJECT_RE.match(text):
            close_semester_project(buf=buf, out=out)
            buf = []
            continue
        if buf is not None and text:
            buf.append(text)
    close_semester_project(buf=buf, out=out)
    return out


def end_then_start(entry):
    """Sort key of a job: the end date, then the start date (an empty date sorts first)."""
    return (entry.get("end") or "", entry.get("start") or "")


def letters_only(text):
    """Lower case letters only, to compare two titles that differ by spaces or punctuation."""
    return re.sub(r"[^a-z]", "", text.lower())


def entry_end(entry):
    """Sort key: the end date of an entry."""
    return entry["end"]


def parse_lines(lines, filename, extra_links=()):
    """Everything above, in order. Returns the JSON-ready dict in the production format."""
    lines = join_split_headings(merge_label_cells(unglue_labels(lines)))
    sections = split_sections(lines)
    email, phone, linkedin = find_email(lines), find_phone(lines), find_linkedin(lines)
    name, name_confidence = find_name(lines, email, filename)
    if not name and linkedin:
        name = name_from_linkedin(linkedin)                      # a saved LinkedIn page has no readable name line, but its link has one
        name_confidence = 0.4 if name else 0.0
    skills = find_skills(sections)
    name_key = norm_key(name or "")
    skills = {k: v for k, v in skills.items() if not (name_key and norm_key(k) and (norm_key(k) == name_key or (len(norm_key(k)) >= 6 and norm_key(k) in name_key)))}   # the candidate's own name is not a skill
    titles = find_titles(lines, sections)
    industries = find_industries(lines, sections)
    education = enrich_from_header(find_education(sections, lines), sections.get("top", []))
    education = [entry for entry in education if not SEMESTER_PROJECT_RE.match(entry.get("degree") or "")]        # '10th semester' is a project list, not a degree
    have_degrees = {norm_key(entry.get("degree") or "") for entry in education}
    education += [r for r in academic_table_rows(lines) if norm_key(r["degree"]) not in have_degrees]
    experience_lines = sections.get("experience") or [line for name_ in ("top", "summary", "other", "skills") for line in sections.get(name_, [])]
    experience = [entry for entry in find_entries(experience_lines, undated=bool(sections.get("experience"))) if not entry.get("_education")]
    known = {norm_key(entry.get("company") or "") for entry in experience}
    for company, current in column_companies(lines):
        if norm_key(company) not in known:
            experience.append(new_job(None, company, "label", "", "", 0, current))
    others = []                                                      # dated jobs listed above the first heading or under another heading ('(1) ADGM | Sep 2025 – Now' / 'Lead | UAE'): the work history is read from where it is
    for name_ in ("top", "summary", "skills", "certifications", "other"):
        found_jobs = [entry for entry in find_entries(sections.get(name_, []), undated=False)
                      if entry.get("start") and entry.get("company") and entry.get("company_source") in ("suffix", "structure", "label", "sentence") and not entry.get("_education")]
        if len(found_jobs) < 2:
            found_jobs = find_timeline_entries(sections.get(name_, [])) or found_jobs
        if len(found_jobs) >= 2 or (found_jobs and not any(entry.get("start") for entry in experience)):
            others += found_jobs
    if others:
        have = {(norm_key(entry.get("company") or ""), entry.get("start")) for entry in experience}
        have_names = {norm_key(entry.get("company") or "") for entry in experience if entry.get("start")}
        experience = [entry for entry in experience if entry.get("start")] + [entry for entry in others if (norm_key(entry.get("company") or ""), entry.get("start")) not in have and norm_key(entry.get("company") or "") not in have_names]
        experience.sort(key=end_then_start, reverse=True)
    if not any(entry.get("start") or entry.get("company") for entry in experience):
        exp_texts = [clean(line.text) for line in sections.get("experience", [])]
        starts_ = [k for k, text in enumerate(exp_texts) if re.match(r"^\W*project(?:\s*(?:name|title))?\s*[:\-–]", text, re.IGNORECASE)]
        client_jobs = []
        for n_, k_ in enumerate(starts_):
            blk = exp_texts[k_:(starts_[n_ + 1] if n_ + 1 < len(starts_) else len(exp_texts))]
            cl = next((re.sub(r"^(?:company|employer|organi[sz]ation)(?:\s*name)?\s*[:\-–]\s*", "", text, flags=re.IGNORECASE).strip(" .") for text in blk
                       if re.match(r"(?:company|employer|organi[sz]ation)(?:\s*name)?\s*[:\-–]", text, re.IGNORECASE)), None)   # the employer, not the 'Client:'
            ro = next((re.sub(r"^role\s*[:\-–]\s*", "", text, flags=re.IGNORECASE).strip(" .") for text in blk if re.match(r"role\s*[:\-–]", text, re.IGNORECASE)), None)
            if cl or ro:
                client_jobs.append(new_job(re.sub(r"\s*\([^)]*\)\s*$", "", ro) if ro else None, cl, "label", description=shorten(" ".join(blk))))
        if client_jobs:
            experience = client_jobs                                     # the work history is written as projects with a client and a role
    internships = find_entries(sections.get("internships", []))
    internships += [entry for entry in experience if entry["title"] and re.search(r"\bintern(?:ship)?s?\b", entry["title"], re.IGNORECASE)]   # written inside experience
    intern_jobs = [entry for entry in experience if re.match(r"^\W*intern(?:ship)?\b", entry.get("company") or "", re.IGNORECASE)]
    internships += [dict(entry, title=entry["company"], company=None) for entry in intern_jobs]                 # 'Intern Software Development Engineer' read as the company of a job
    experience = [entry for entry in experience if entry not in internships and entry not in intern_jobs]
    experience = apply_past_statements(apply_current_statements(apply_current_dates(experience, lines, sections.get("experience", []), name), lines), lines)
    years, years_source = total_experience_years(sections, experience)
    designations = []                                                # the titles the candidate really held, as written ('Associate Data Engineer')
    for entry in experience:
        text = clean(entry.get("title") or "")
        if text and len(text.split()) <= 6 and (any(TITLE_HEAD.search(word + " ") for word in text.split()) or entry.get("company_source") == "label") \
                and not re.search(r"\bintern(?:ship)?s?\b", text, re.IGNORECASE):
            designations.append(text)
    if designations:
        headline = find_titles(lines, {"top": sections.get("top", [])[:6], "summary": sections.get("summary", [])[:4]})
        titles = list(dict.fromkeys(designations))
        titles += [heading for heading in headline if not any(letters_only(heading) in letters_only(doc) or letters_only(doc) in letters_only(heading) for doc in titles)]
        titles = titles
    around = " | ".join(item for entry in experience + internships for item in (entry.get("company"), entry.get("title")) if item)
    industries = find_industries([], {"experience": [Line(around)]}) if around else []
    if not industries:                                               # the employers say nothing: the career summary / objective of the resume
        industries = find_industries([], {"summary": sections.get("summary", [])[:12]})[:2]
    if not industries and re.search(r"\b(?:recruit\w*|staffing|talent acquisition|head-?hunt\w*|manpower)\b", around + " " + " ".join(clean(line.text) for line in sections.get("experience", [])[:40]), re.IGNORECASE):
        industries = ["Staffing & Recruitment"]                      # a recruiter's clients are many sectors: the sector of the work itself
    if not industries and len(re.findall(r"\b(?:java|python|sql|html|css|javascript|react|angular|node|aws|azure|linux|oracle|sap|\.net|c\+\+|spring|django|php|android|devops|selenium|etl|power bi|tableau|kubernetes|docker|jenkins|ansible|shell scripting|git|mysql|mongodb|typescript|kotlin|spark|hadoop|databricks)\b",
                                         " ".join(skills) if isinstance(skills, list) else " ".join(skills), re.IGNORECASE)) >= 3:
        industries = ["IT Services"]                                 # a technical skill set
    projects = find_projects(sections)
    have_projects = {norm_key(part["name"]) for part in projects}
    projects += [part for part in semester_projects(sections.get("projects", [])) if norm_key(part["name"]) not in have_projects]
    for part in projects:
        part["name"] = re.sub(r"^\s*\d{1,2}\s*[.)]\s+", "", part["name"])                                  # '1. Schwab' is the project Schwab, numbered
        link = re.search(r"https?://\S+", part["name"])
        if link:
            part["link"] = part.get("link") or link.group(0).strip("() ")
            part["name"] = re.sub(r"\(?\s*https?://\S*\s*\)?", "", part["name"]).strip(" .,:;-(")             # the link is kept in the project's link, not in its name
        part["name"] = re.sub(r"[\s(\[]+$", "", part["name"])                                              # a bracket left open at the end of the name
    have_projects = {norm_key(part["name"]) for part in projects}
    projects += [part for part in labelled_projects(lines) if norm_key(part["name"]) not in have_projects]     # 'Project title: X' / 'Project name: X' written anywhere
    projects = [part for part in projects if not PROJECT_JUNK_RE.match(part["name"]) and not re.search(r"\b(?:till date|present|till now|to date)\b", part["name"], re.IGNORECASE)
                and not re.search(r"\b(?:duration|description|modules?|tenure|role|client)\s*:", part["name"], re.IGNORECASE)
                and not re.fullmatch(r"\W*project\s*(?:description|duration|details?|name|title|period|role)?\W*", part["name"], re.IGNORECASE)]
    name_words = {word for word in re.split(r"[^a-z]+", (name or "").lower()) if len(word) >= 3}
    certifications = [c for c in find_certifications(sections, lines)
                      if norm_key(c) != norm_key(name or "-") and not re.search(r"year of passing|percentage|cgpa|marks", c, re.IGNORECASE)
                      and not (INSTITUTION_RE.search(c) and DEGREE_RE.search(c) and not CERT_WORDS.search(c) and not CERT_VENDORS.search(c))
                      and not (name_words and len(c.split()) <= 4 and len(name_words & {word for word in re.split(r"[^a-z]+", c.lower()) if len(word) >= 3}) >= 1
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
    education = [entry for entry in education if entry["degree"]]
    if education and not any(entry["institution"] for entry in education):
        issues.append("education found but no institution")
    for entry in experience + internships:
        if not (entry["title"] or entry["company"]):
            issues.append("an experience entry has neither title nor company")
            break
    confidence = {
        "name": round(name_confidence, 2),
        "contact": round(sum(bool(item) for item in (email, phone, linkedin)) / 3, 2),
        "skills": round(sum(skills.values()) / len(skills), 2) if skills else 0,
        "education": round(sum(bool(entry["institution"]) + bool(entry["year"]) + 1 for entry in education) / (3 * len(education)), 2) if education else 0,
        "experience": round(sum(bool(entry["title"]) + bool(entry["company"]) + 1 for entry in experience) / (3 * len(experience)), 2) if experience else 0,
    }

    # ---- the production format ----
    named = [entry for entry in experience if entry["title"] or entry["company"]]        # an entry with dates only still counts for the years, not as a job
    jobs = [build_job(entry) for entry in named]
    current = next((entry for entry in named if entry["current"]), None) or (max(named, key=entry_end) if named else None)
    education_out = build_education(education)
    whole_years = int(years) if years is not None else 0
    place = find_location(lines, sections, name)
    return {
        "name": name, "email": email, "phoneNumber": phone, "alternatePhoneNumbers": find_phones(lines, phone),
        "languagesKnown": find_languages(sections, lines),
        "otherSections": named_other_sections(lines),
        "relocationDetails": [],
        "currentCompanyDetails": {"companyName": (current or {}).get("company") or "", "designation": (current or {}).get("title") or ""},
        "employmentHistory": jobs,
        "ctcDetails": build_ctc(salary),
        "totalExperience": {"years": whole_years, "months": int(round(((years or 0) - whole_years) * 12))},
        "skills": [s.lower() for s in skills],
        "analysedSkills": rank_skills(skills, lines, jobs),
        "educationalBackground": education_out,
        "personalProjects": [{"projectName": part["name"], "projectDescription": part["description"], "projectLink": part["link"],
                              "skillsUsedInProject": [item.lower() for item in part["skills"]]} for part in projects],
        "certifications": build_certifications(certifications),
        "socialProfiles": find_social_profiles(lines, linkedin),
        "redFlags": build_red_flags(jobs, education_out, lines, issues),
        "currentNoticePeriod": find_notice_period(lines),
        # extras the production record does not have
        "titles": titles, "industries": industries, "internships": [build_job(entry) for entry in internships],
        "totalExperienceSource": years_source, "confidence": confidence, "issues": issues,
        "sections_found": [item for item in ("summary", "skills", "experience", "education", "projects", "internships", "certifications") if sections.get(item)],
        # the keys a candidate-list person has as well (one format for every JSON)
        "location": place["text"], "locationDetails": place, "links": find_links(lines, extra_links),
        "educationText": "; ".join(education_text(entry) for entry in education_out),
        "totalExperienceText": experience_text(whole_years, int(round(((years or 0) - whole_years) * 12))),
    }


def split_bundle(lines):
    """A file that holds several resumes (a recruiter's compiled PDF) -> a list of line lists, one per candidate.
    Heuristic: every different e-mail address that first appears far from the previous one starts a new resume; the
    short lines right above it (the name / contact block) go with it. A piece without at least two real sections
    (skills, experience, education ...) is not a resume of its own and is joined to the piece before it.
    Returns [lines] (no split) when fewer than two resumes are found."""
    firsts, seen = [], set()
    for i, line in enumerate(lines):
        for entry in EMAIL_RE.findall(line.text):
            if entry.lower() not in seen:
                seen.add(entry.lower())
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
        if sum(bool(found.get(item)) for item in ("skills", "experience", "education", "projects")) >= 2:
            kept.append(piece)
        else:
            kept[-1] = kept[-1] + piece
    return kept


# =====================================================================================================
# ONE FORMAT FOR EVERY JSON  +  THE SLIM SEARCH DOCUMENT
#   A resume and a person of a CSV / Excel list get exactly the same keys. What a spreadsheet row does not have stays empty.
#   search_document() turns a full JSON into the slim JSON that goes to MeiliSearch (Output\Valid).
# =====================================================================================================

EDUCATION_LABEL = {"bachelors": "Bachelor's", "masters": "Master's", "phd": "PhD", "diploma": "Diploma", "intermediate": "Intermediate",
                   "high_school": "High school", "associates": "Associate's"}


def education_text(entry):
    """{'educationType': 'bachelors', 'fieldOfStudy': 'Mechanical Engineering', 'institutionName': 'KVG College', 'endDate': '2026'}
    -> "Bachelor's in Mechanical Engineering, KVG College, 2026" """
    label = EDUCATION_LABEL.get(entry.get("educationType", ""), (entry.get("educationType") or "").replace("_", " ").capitalize())
    head = f"{label} in {entry['fieldOfStudy']}" if label and entry.get("fieldOfStudy") else (label or entry.get("fieldOfStudy") or "")
    return ", ".join(item for item in (head, entry.get("institutionName"), entry.get("endDate")) if item)


def job_text(j):
    where = " at ".join(item for item in (j.get("designation"), j.get("companyName")) if item)
    start, end = (j.get("startDate") or "")[:7], (j.get("endDate") or "")[:7]
    dates = f"{start} to {end}" if start else ""
    return f"{where} ({dates})" if where and dates else (where or dates)


def experience_text(years, months):
    """(5, 3) -> '5y 3mo'"""
    return " ".join(([f"{years}y"] if years else []) + ([f"{months}mo"] if months else []))


def parse_experience_text(text):
    """'5y 3mo' / '9y' / '5mo' (the spreadsheet's experience column) -> (years, months)"""
    part_text = str(text or "").lower()
    y, mo = re.search(r"(\d+(?:\.\d+)?)\s*y", part_text), re.search(r"(\d+)\s*mo", part_text)
    years = float(y.group(1)) if y else 0.0
    whole = int(years)
    months = (int(mo.group(1)) if mo else 0) + int(round((years - whole) * 12))
    return whole + months // 12, months % 12


def sheet_money(text):
    """'410000' / '900000.000' (the spreadsheet's CTC columns) -> 410000 / 900000. The number is copied as the sheet gives it. A value
    under 1000 ('100') is a placeholder, not a salary: it becomes 0 (unknown) instead of a guess."""
    try:
        value = float(str(text).replace(",", ""))
    except ValueError:
        return 0
    return int(value) if value >= 1000 else 0


def not_assessed(note):
    """The production 'redFlags' block when there is no resume text to judge."""
    flags = {}
    for name in ("containsMistakes", "frequentJobSwitcher", "overstatedAchievements", "employmentGap", "incompleteEducation", "lackOfRoleProgression"):
        flags[name] = False
        flags[name + "Description"] = note
    return flags


def education_levels_from_text(text):
    """A spreadsheet's education cell mixes degrees, schools and colleges without saying which belongs to which, so only the LEVELS
    are read ('Bachelor of Technology | RGMCET' -> bachelors). Institution and field stay empty: guessing them would be wrong."""
    out, seen = [], set()
    for segment in re.split(r"\s*[|;]\s*|,\s+", text or ""):
        match = DEGREE_RE.search(segment)
        if not match or (match.group(0).rstrip(".") in AMBIGUOUS_DEGREE and not AMBIGUOUS_OK.match(segment[match.end():])):
            continue
        kind = education_type(segment[match.start():])
        if kind and kind not in seen:
            seen.add(kind)
            out.append({"educationType": kind, "institutionName": "", "startDate": "", "endDate": "", "fieldOfStudy": "", "scoreAchieved": ""})
    return out


def list_person_record(c, source_file, folder, number, status):
    """One person of a candidate list (parse_table's dict) -> the same JSON format as a resume."""
    company = c.get("currentCompanyDetails") or {}
    cname, title = company.get("companyName") or "", company.get("designation") or ""
    years, months = parse_experience_text(c.get("totalExperienceText"))
    education_raw = c.get("educationText") or ""
    past = c.get("pastCompanyDetails") or {}
    jobs = [{"companyName": cname, "designation": title, "startDate": "", "endDate": "current",
             "workExperienceDescription": "", "skillsUsedInRole": []}] if (cname or title) else []
    if past.get("companyName") or past.get("designation"):           # 'Previous company name' / 'Previous designation' columns
        jobs.append({"companyName": past.get("companyName") or "", "designation": past.get("designation") or "", "startDate": "", "endDate": "",
                     "workExperienceDescription": "", "skillsUsedInRole": []})
    note = "Not assessed: this person comes from a spreadsheet row, there is no resume text to judge."
    return {
        "status": status, "reason": None, "category": None, "source_file": source_file, "source_folder": folder, "parsed": True,
        "sourceType": "candidate_list", "candidateNumber": number,
        "name": tidy_name(c.get("name")) or None, "email": c.get("email"), "phoneNumber": c.get("phoneNumber"),
        "languagesKnown": c.get("languages") or [], "relocationDetails": [],
        "currentCompanyDetails": {"companyName": cname, "designation": title},
        "employmentHistory": jobs,
        "ctcDetails": build_ctc({"current": {"per_year_inr": sheet_money(c.get("currentCTC"))},
                                 "expected": {"per_year_inr": sheet_money(c.get("expectedCTC"))}}),
        "totalExperience": {"years": years, "months": months},
        "skills": list(dict.fromkeys(item.lower() for item in (c.get("skills") or []))),
        "analysedSkills": {"advanced": [], "medium": [], "basic": []},
        "educationalBackground": education_levels_from_text(education_raw),
        "personalProjects": [], "certifications": [],
        "socialProfiles": c.get("socialProfiles") or [],
        "redFlags": not_assessed(note), "currentNoticePeriod": 0,
        "titles": c.get("titles") or [], "industries": c.get("industries") or [], "internships": [],
        "totalExperienceSource": "spreadsheet column" if (years or months) else None,
        "confidence": {}, "issues": [], "sections_found": [],
        "location": (c.get("locationDetails") or {}).get("text", ""), "locationDetails": c.get("locationDetails") or {}, "links": c.get("links") or [],
        "educationText": education_raw, "totalExperienceText": c.get("totalExperienceText") or "",
        "educationItems": c.get("educationItems") or [],
        "listDegrees": c.get("listDegrees") or [], "listFields": c.get("listFields") or [],
    }


DEGREE_CANON = {"btech": "B.Tech", "mtech": "M.Tech", "be": "B.E", "me": "M.E", "bsc": "B.Sc", "msc": "M.Sc", "bcom": "B.Com", "mcom": "M.Com",
                "bpharm": "B.Pharm", "mpharm": "M.Pharm", "barch": "B.Arch", "bba": "BBA", "mba": "MBA", "bca": "BCA", "mca": "MCA", "llb": "LLB",
                "llm": "LLM", "mbbs": "MBBS", "pgdm": "PGDM", "pgdca": "PGDCA", "pgdba": "PGDBA", "pgdbm": "PGDBM", "pgpm": "PGPM", "phd": "PhD", "doctorate": "PhD", "ba": "B.A", "ma": "M.A",
                "bed": "B.Ed", "msw": "MSW", "bsw": "BSW", "cma": "CMA", "diploma": "Diploma", "iti": "ITI"}
BACHELOR_OF = [("technolog", "B.Tech"), ("engineering", "B.E"), ("science", "B.Sc"), ("commerce", "B.Com"), ("arts", "B.A"), ("business admin", "BBA"),
               ("computer application", "BCA"), ("pharmac", "B.Pharm"), ("architect", "B.Arch"), ("law", "LLB"), ("education", "B.Ed")]
MASTER_OF = [("technolog", "M.Tech"), ("engineering", "M.E"), ("science", "M.Sc"), ("commerce", "M.Com"), ("arts", "M.A"), ("business admin", "MBA"),
             ("computer application", "MCA"), ("pharmac", "M.Pharm"), ("architect", "M.Arch"), ("law", "LLM"), ("social work", "MSW")]
FIELD_ABBR = {"cse": "Computer Science and Engineering", "cs": "Computer Science", "it": "Information Technology", "ece": "Electronics and Communication Engineering",
              "eee": "Electrical and Electronics Engineering", "eie": "Electronics and Instrumentation Engineering", "ete": "Electronics and Telecommunication Engineering",
              "mech": "Mechanical Engineering", "civil": "Civil Engineering", "ai": "Artificial Intelligence", "aiml": "Artificial Intelligence and Machine Learning",
              "etc": "Electronics and Telecommunication Engineering", "entc": "Electronics and Telecommunication Engineering", "extc": "Electronics and Telecommunication Engineering"}
FIELD_WORDS = re.compile(
    r"\b(computer|computers|science|sciences|engineering|technology|electronics|electrical|mechanical|civil|information|chemical|instrumentation|production|"
    r"industrial|automobile|automotive|aeronautical|aerospace|telecommunication|telecommunications|communication|communications|software|data|artificial|"
    r"intelligence|machine|learning|biotechnology|biomedical|mining|metallurgical|metallurgy|textile|agricultural|agriculture|marine|business|commerce|"
    r"accounting|accountancy|finance|financial|marketing|management|human|resource|resources|economics|english|hindi|history|political|sociology|psychology|"
    r"mathematics|physics|chemistry|biology|botany|zoology|arts|pharmacy|nursing|law|journalism|mass|fashion|design|architecture|hospitality|hotel|tourism|"
    r"education|statistics|geography|philosophy|applications|power|systems|embedded|vlsi|cyber|security|networks|networking|operations|international|banking|"
    r"insurance|taxation|accounts)\b", re.IGNORECASE)
_KNOWN_COMPANIES = None


def known_companies():
    """Company names written in the industry clue list ('Infosys', 'TCS', 'Axis Bank'): every word starts with a capital letter."""
    global _KNOWN_COMPANIES
    if _KNOWN_COMPANIES is None:
        _KNOWN_COMPANIES = {norm_key(word) for word in lookups().industry_of if len(word) >= 3 and all(part[:1].isupper() or part[:1].isdigit() for part in word.split())}
    return _KNOWN_COMPANIES


def canon_degree(text):
    """'B.Tech in CSE' / 'Bachelor of Technology' -> 'B.Tech'; only a degree that is written clearly. School levels (10th, 12th) and 'graduate' give ''."""
    text = clean(text or "")
    if re.search(r"\b(?:post|pg|p\.g)[\s.\-]*(?:graduate\s+)?diploma\b", text, re.IGNORECASE):
        return "PG Diploma"
    match = DEGREE_RE.search(text)
    if not match:
        return ""
    token = re.sub(r"[.\-\s'’]", "", match.group(0).lower())
    if token in ("be", "me", "ba", "ma") and not AMBIGUOUS_OK.match(text[match.end():]):
        return ""
    if token == "diploma" and re.search(r"\b(?:post|pg|p\.g)\b[\s.\-]*(?:graduate)?\s*$", text[max(0, match.start() - 22):match.start()], re.IGNORECASE):
        return "PG Diploma"
    if token.startswith(("bachelor", "master")):
        after = text[match.end():match.end() + 40].lower()
        bachelor = token.startswith("bachelor")
        for key, name in (BACHELOR_OF if bachelor else MASTER_OF):
            if after.lstrip().startswith("of") and key in after:
                return name
        return "Bachelor's" if bachelor else "Master's"
    return DEGREE_CANON.get(token, "")


FIELD_CUT = re.compile(r"\b(?:upper|first|second|third|class|honou?rs|distinction|with|as|from|at|aggregate|cgpa|gpa|grade|passed|pass|merit|dual|major|specialization|specialisation)\b|[-–—:]", re.IGNORECASE)
SINGLE_FIELD_OK = {"mechanical", "civil", "electrical", "electronics", "chemical", "finance", "marketing", "economics", "english", "hindi", "history",
                   "mathematics", "physics", "chemistry", "biology", "botany", "zoology", "nursing", "journalism", "accountancy", "accounting", "statistics",
                   "geography", "philosophy", "psychology", "sociology", "biotechnology", "biomedical", "agriculture", "textile", "mining", "metallurgy",
                   "aeronautical", "aerospace", "automobile", "instrumentation", "production", "hospitality", "tourism", "taxation", "banking", "insurance"}
FIELD_GENERIC = {"technology", "engineering", "science", "sciences", "arts", "commerce", "business administration", "computer applications", "computer application",
                 "pharmacy", "education", "law", "laws", "architecture", "management", "applications", "social work"}


def canon_field(text):
    """'Mechanical Engineering' / 'CSE' / 'Electronics & Communication' -> a clean field name. The words that are only the name of the degree
    ('Technology' of B.Tech, 'Business Administration' of MBA) are not a field of study."""
    part_text = clean(text or "").strip(" ()-,.:")
    part_text = re.sub(r"\s*&\s*", " and ", part_text)
    part_text = FIELD_CUT.split(part_text)[0].strip(" ()-,.:")
    part_text = re.sub(r"\s+(?:and|of|in|for|the)$", "", part_text, flags=re.IGNORECASE)
    if not part_text or re.search(r"[\d@]", part_text) or len(part_text.split()) > 7 or INSTITUTION_RE.search(part_text) or DEGREE_RE.search(part_text):
        return ""
    key = re.sub(r"[^a-z]", "", part_text.lower())
    if key in FIELD_ABBR:
        return FIELD_ABBR[key]
    vocab = FIELD_WORDS.pattern[3:-3].split("|")                    # a word with a typo ('Compute Engineering') is read as the nearest subject word
    fixed = []
    for word in part_text.split():
        wl = re.sub(r"[^a-z]", "", word.lower())
        if len(wl) >= 6 and wl not in vocab and wl not in ("administration",):
            near = process.extractOne(wl, vocab, scorer=fuzz.ratio, score_cutoff=88)
            word = near[0] if near else word
        fixed.append(word)
    part_text = " ".join(fixed)
    if len(part_text.split()) == 1 and part_text.lower() not in SINGLE_FIELD_OK:
        return ""
    part_text = re.sub(r"^(?:in|of|and|for|the)\s+", "", part_text, flags=re.IGNORECASE)
    part_text = re.sub(r"\bEngg\b\.?", "Engineering", part_text, flags=re.IGNORECASE)
    tokens = [word for word in re.findall(r"[a-z]+", part_text.lower()) if word not in ("and", "of", "in", "for", "the")]
    if not tokens or any(not (FIELD_WORDS.fullmatch(word) or word == "administration") for word in tokens):
        return ""                                                   # a word that is not a subject ('State', 'Council', 'Spl', 'HR') means the text is not a clean field
    if part_text.lower() in FIELD_GENERIC or not FIELD_WORDS.search(part_text) or SENTENCE_START_RE.match(part_text) or not re.fullmatch(r"[A-Za-z][A-Za-z ,/]*", part_text):
        return ""
    if len(tokens) == 1 and tokens[0] not in SINGLE_FIELD_OK:
        return ""
    small = {"and", "of", "in", "for", "the"}
    return " ".join(word if (word.lower() in small and k) else (word if (word.isupper() and len(word) <= 3) else word.capitalize()) for k, word in enumerate(part_text.split()))


CORE_INSTITUTION = re.compile(r"\b(universit(?:y|ies)|colleges?|collage|institutes?|institue|polytechnic|academy|iit|nit|iiit|iim|vidyapeeth|mahavidyalaya|campus|school of|bits|symbiosis|amity|manipal)\b", re.IGNORECASE)
NOT_INSTITUTE = re.compile(r"\b(board|cbse|icse|high school|public school|junior|jr|vidyalaya|matriculation|secondary|sslc|hsc|ssc|pu|puc|pre[- ]?university|competition|conducted|champion|winner|award|awarded|certified|certificate|course|training|exam|examination)\b", re.IGNORECASE)


def institute_name(text):
    """A college / university / institute: the name up to the city or the first comma. School boards, schools and fragments give ''."""
    part_text = re.sub(r"^(?:college|institute|institution|university|school)\s*(?:name)?\s*[:\-–]\s*", "", clean(text or ""), flags=re.IGNORECASE)
    part_text = re.split(r"\s*[,;|(\[]|\s[-–—]\s|[-–—]\s+|\s(?:passed|with|from|during|under|affiliated|in\s+(?:19|20)\d\d)\b|\s(?:in|at)\s+[A-Z]", part_text)[0].strip(" .,:;|-–—()/")
    while part_text.split() and part_text.split()[-1].lower() in ("on", "at", "in", "for", "the", "and", "of", "under", "with"):
        part_text = " ".join(part_text.split()[:-1])
    if re.search(r"\b(?:degree|pg|ug)\b", part_text, re.IGNORECASE):
        return ""
    core = list(CORE_INSTITUTION.finditer(part_text))
    if core:
        tail = part_text[core[-1].end():].split()
        if tail and tail[0].lower() not in ("of", "for") and not all(word.lower().strip(".,") in {"engineering", "technology", "science", "sciences", "management", "tech", "and", "&", "research", "studies"} for word in tail):
            part_text = part_text[:core[-1].end()]                                   # 'Visveswaraih technical University Belgaum' -> without the city
    if not part_text or part_text[:1].islower() or "\ufffd" in part_text or len(part_text.split()) < 2 or len(part_text.split()) > 8 or re.search(r"\d", part_text):
        return ""
    if not CORE_INSTITUTION.search(part_text) or NOT_INSTITUTE.search(part_text) or not plausible_item(part_text, 10) or heading_section(part_text):
        return ""
    return part_text


SUFFIX_WORDS = {"pvt", "private", "ltd", "limited", "llp", "inc", "llc", "corp", "corporation", "co", "company", "technologies", "technology", "solutions",
                "systems", "services", "industries", "labs", "lab", "infotech", "software", "consultancy", "consulting", "enterprises", "group", "bank",
                "motors", "global", "international", "associates", "partners", "studio", "networks", "digital", "and", "of", "the", "&"}
GENERIC_SECTOR = {"products", "product", "operating", "computer", "computers", "power", "innovation", "innovations", "network", "networks", "system",
                  "solution", "web", "mobile", "application", "applications", "research", "development", "design", "process", "quality", "support", "operations",
                  "healthcare", "information", "data", "cloud", "business", "financial", "banking", "education", "manufacturing", "retail", "telecom",
                  "automotive", "pharma", "it", "bpo", "ites", "marketing", "sales", "engineering", "management", "india", "indian"}
STOP_FIRST = {"about", "full", "company", "our", "the", "job", "work", "details", "profile", "software", "development", "project", "projects", "client",
              "core", "key", "basic", "good", "strong", "total", "current", "present", "previous", "last", "first", "role", "roles", "team"}
NOT_EMPLOYER = {"mechanical", "processes", "process", "electrical", "civil", "operations", "production", "quality", "maintenance", "designation", "position", "role", "company", "organization", "organisation", "employer", "client", "department", "confidential", "na", "n/a", "fresher", "self", "self employed", "freelance", "freelancer", "freelancing", "none", "nil", "not applicable",
                "various", "multiple", "others", "other"}


def employer_name(name, source=""):
    """A company is trusted only when it has a company word (Pvt, Ltd, Technologies ...), or was written as 'Company: X' / 'working as ... at X',
    or is a name from the industry clue list. The text is cut after the last company word ('Acme Pvt Ltd, Pune' -> 'Acme Pvt Ltd'), and a
    sentence, a number of months, or a sector word on its own is not a company."""
    if re.match(r"^\W*(?:client|customer|project|product|end client)\b", str(name or ""), re.IGNORECASE):
        return ""                                                   # 'Client - Wipro Ltd' is the client of a project, not the employer
    trusted = source == "sheet"                                     # a column the sheet itself calls company / previous company
    text = re.sub(r"^[^A-Za-z0-9]+", "", clean(name or ""))
    text = re.sub(r"\s*\([^()]{8,}\)\s*$", "", text)                      # 'Sterling BackCheck (USA No.2 Background Verifications company)': the note in brackets is not the name
    text = re.sub(r"^(?:name of (?:the )?)?(?:company|organi[sz]ation|employer|client)(?:\s+name)?\s*[:\-–]\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^(?:(?:previously|currently|presently|earlier|formerly)\s+)?(?:(?:working|worked|employed)\s+(?:in|at|with|for)\s+)?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\W+", "", text, flags=re.IGNORECASE)
    if re.search(r"\b(?:an|a|the)\b", text) and not re.match(r"(?:the)\b", text, re.IGNORECASE):
        return ""
    if re.search(r"\b(?:pvt|private|ltd|limited|llp|inc|llc)\b", text, re.IGNORECASE):
        text = re.split(r"\b(?:in|from|at|with|for)\b\s+", text, flags=re.IGNORECASE)[-1]
        text = text.split(":")[-1].strip()
        pieces = re.split(r"\s+under\s+|\s[-–—]\s*|[-–—]\s+", text)
        with_suffix = [part for part in pieces if re.search(r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation)\b", part, re.IGNORECASE)]
        text = (with_suffix[0] if with_suffix else pieces[-1]) if len(pieces) > 1 else text             # 'Accenture Solution Pvt Ltd - Pune, INDIA': the piece with the company suffix, not the place after the dash
    text = text.strip(" .,:;|-–—")
    if not text or re.match(r"^(?:" + ("" if trusted else r"\d|") + r"deputed|worked|working|currently|presently|trainee|intern\b|as\b|at\b|in\b|from\b|since\b)", text, re.IGNORECASE):
        return ""
    last = None
    for last in re.finditer(r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation)\b\.?", text, re.IGNORECASE):
        pass
    if last:
        text = text[:last.end()]
    else:
        text = re.split(r"\s*[,:\[|]|\s*\((?![A-Za-z][A-Za-z .&-]{0,13}\))|\s[-–—]\s*|\s(?:from|since|as|at|in|for|with|working)\b", text)[0]
    text = text.strip(" .,:;|-–—()")
    if not text or len(text) > 60 or len(text.split()) > 7 or not plausible_item(text, 8) or re.search(r"\d{3,}|[\[\]]", text) or text.count("(") != text.count(")") or text.lower() in NOT_EMPLOYER:
        return ""
    word_lists = lookups()
    key = norm_key(text)
    suffix = COMPANY_SUFFIX_RE.search(text) is not None
    if key in word_lists.skills or key in word_lists.titles or heading_section(text) or DEGREE_RE.search(text):
        return ""
    if (TITLE_HEAD.search(text.split()[-1] + " ") or INSTITUTION_RE.search(text) or PROSE_START_RE.match(text)) and not suffix and not trusted:
        return ""
    if {word.lower().strip(",.") for word in text.split()} <= (INDIAN_PLACES | NAME_BAD_WORDS):
        return ""
    if re.search(r"\b(?:project|client|role|responsibilit\w*|description|duration|period|environment|courses?|months?|years?)\b", text, re.IGNORECASE):
        return ""
    distinctive = [word for word in re.findall(r"[A-Za-z][A-Za-z.&']+", text) if word.lower().strip(".") not in (SUFFIX_WORDS | GENERIC_SECTOR)]
    if not distinctive:
        return ""
    words = text.split()
    if not trusted and (TITLE_HEAD.search(words[0] + " ") or words[0].lower() in {"analyst", "engineer", "manager", "executive", "developer", "consultant", "associate"}):
        return ""                                                   # 'ANALYST INFRASTRUCTURE SERVICE DXC TECHNOLOGY': the job title is glued to the company
    legal = re.search(r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation)\b", text, re.IGNORECASE) is not None
    if CORE_INSTITUTION.search(text) and not legal:
        return ""                                                   # a college or university written where a company was expected
    capitalised = all(word[:1].isupper() or word[:1].isdigit() or word.lower() in ("and", "of", "&", "for") for word in words)
    verbish = any(re.search(r"(?:ing|ed)$", word.lower()) for word in words if word.lower() not in ("holding", "holdings", "trading", "consulting", "marketing", "engineering", "manufacturing"))
    weak_ok = suffix and words[-1].lower().strip(".,") in SUFFIX_WORDS and 2 <= len(words) <= 5 and words[0].lower() not in STOP_FIRST and capitalised and not verbish
    if words[0].lower().endswith("ing") and not legal and not trusted and words[0].lower() not in ("holding", "holdings", "trading", "consulting", "marketing", "engineering", "manufacturing", "sterling", "kingfisher"):
        return ""                                                   # 'Deploying solutions in Azure ...' is a sentence, not a company
    return text if (legal or weak_ok or trusted or source in ("label", "sentence", "structure", "suffix") or key in known_companies()) else ""


MONTH_ONLY_RE = re.compile(r"^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?(?:\s*['’]?\s*\d{0,4})?$", re.IGNORECASE)


def fallback_employer(raw, source="", dated=False):
    """A company name that employer_name() did not accept but that the job shows: kept only when it really reads like a company. A month ('Aug'), a
    place ('Maharashtra'), an education word ('Pursuing'), a title, a number ('45%'), a college, a skill or a project name is not one; a one-word name needs a
    company suffix or a known company; a name that spaCy or the layout guessed (not a label or a sentence) needs a suffix or a known company too."""
    raw = clean(raw).strip(" .,:;|-–—")
    words = raw.split()
    if not raw or len(raw) > 60 or len(words) > 6 or PROSE_START_RE.match(raw) or DATE_RANGE_RE.search(raw) or raw.lower() in NOT_EMPLOYER:
        return False
    if re.match(r"^\W*(?:client|customer|project|product|end client)\b", raw, re.IGNORECASE):
        return False                                                # a client of a project is not the employer
    low = {word.lower().strip(",.") for word in words}
    word_lists = lookups()
    key = norm_key(raw)
    if (re.search(r"\d", raw) and len(re.findall(r"[A-Za-z]", raw)) < 4) or re.search(r"%|@|[\[\]{}]", raw) or MONTH_ONLY_RE.match(raw) or low <= (INDIAN_PLACES | NAME_BAD_WORDS | GENERIC_SECTOR | EXTRA_NOT_NAME) \
            or re.match(r"^\W*(?:pursuing|present|current|till|since|from|to|completed|passed|expected|worked|working)\b", raw, re.IGNORECASE) \
            or TITLE_HEAD.search(words[-1] + " ") or INSTITUTION_RE.search(raw) or CORE_INSTITUTION.search(raw) or DEGREE_RE.search(raw) or heading_section(raw) \
            or key in word_lists.skills or key in word_lists.titles or raw.startswith(":"):
        return False
    known = key in known_companies()
    suffix = COMPANY_SUFFIX_RE.search(raw) is not None
    if known or suffix:
        return True
    capitalised = all(word[:1].isupper() or (word[:1].isdigit() and len(words) == 1) or word.lower() in ("and", "of", "&", "the", "for") for word in words)
    if dated and source in ("spacy", "structure", "caps") and capitalised and len(raw) >= 3:
        return True                                                 # 'IBM' / 'WIPRO' next to a date range: spaCy found an organisation and a period, nothing above rejects it
    if len(words) == 1:
        return False
    return source in ("label", "sentence") and capitalised


def bad_employer(name, person):
    """A value that cannot be an employer: the candidate's own name ('Vyankatesh Raut'), a company suffix on its own ('Limited'), a piece of a date
    ('May/'), or a job title that ends in a number ('Software Developer 2')."""
    name = clean(name or "")
    words = {word for word in re.findall(r"[a-z]+", name.lower()) if len(word) >= 3}
    person_words = {word for word in re.findall(r"[a-z]+", (person or "").lower()) if len(word) >= 3}
    if person_words and words and len(words) >= 2 and words <= person_words:
        return True
    if norm_key(name) and norm_key(name) == norm_key(person or "-"):
        return True
    if name.lower().strip(" .,") in SUFFIX_WORDS:
        return True
    if re.fullmatch(r"\W*(?:" + MONTHS + r")\W*(?:(?:19|20)\d{2})?\W*", name, re.IGNORECASE):
        return True
    bare = re.sub(r"\s*\([^)]*\)?\s*$", "", name)                    # 'Software Developer 2 (Flutter Developer)' -> 'Software Developer 2'
    return bool(re.search(r"\d+$", bare) and TITLE_HEAD.search(re.sub(r"\s*\d+$", "", bare) + " ")
                and not re.search(r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation)\b", name, re.IGNORECASE))


EMP_LABEL_JUNK_RE = re.compile(r"^(?:company|employer|organi[sz]ation|client)\s*(?:profile|url|name|details?|description|overview)?\s*(?:n/?a)?$|^still working$|^(?:software|technical|key|core|it)\s+(?:expertise|skills?|pattern)$|^company\s*:\s*position$|^roles?.*responsibilit|^responsibilit", re.IGNORECASE)
EMP_TITLE_WORDS = re.compile(r"\b(engineer|developer|devloper|manager|executive|officer|consultant|analyst|lead|head|specialist|associate|assistant|architect|designer|supervisor|coordinator|administrator|programmer|tester|technician|accountant|director|scientist|recruiter|representative|intern|trainee|counsel+or|counselor|admin)s?\b", re.IGNORECASE)
EMP_TITLE_START_RE = re.compile(r"^(?:(?:senior|junior|sr|jr|lead|principal|associate|assistant|software|web|mobile|ios|android|java|python|full[- ]?stack|front[- ]?end|back[- ]?end|devops|cloud|qa|sdet|data|network|system|technical|it|hr|sales|business|project|product|operation|operations)\b\.?\s*)+", re.IGNORECASE)
COMPANY_SUFFIX_RE = COMPANY_SUFFIX_RE


EMP_LEGAL_RE = r"\b(?:pvt|private|ltd|limited|llp|inc|llc|corp|corporation)\b"
EMP_MONTH_TAIL_RE = re.compile(r"\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s*(?:19|20)\d{2}\s*[-–—]*\s*$", re.IGNORECASE)
EMP_LABEL_JUNK_RE2 = re.compile(r"^(?:latest|current|previous|last|present)\s+(?:project|company|employer|organi[sz]ation)s?$|^(?:ind|india|present|roles?|current|previous)$|^software engineering experience$|^(?:sr\.?|senior)\s+(?:sdet|hr-data operations)$|^software engniee?r$", re.IGNORECASE)


def final_employer(name):
    """Returns (cleaned, reason). cleaned '' means the value is dropped. Only a value that is wholly a job title or a label is dropped; a title glued
    in front of a confirmed company ('Software Engineer in Wipro') is cut; anything unsure stays as it is."""
    text = clean(name or "").strip(" ,:;|-–—")
    if not text:
        return "", ""
    orig = text
    text = re.sub(r"^\W*\d+\W*\s*(?:company|employer|organi[sz]ation)\s*(?:name)?\s*[:\-–]+\s*", "", text, flags=re.I)
    text = re.sub(r"^\(\d+\)\s*(?:company|employer|organi[sz]ation)\s*(?:name)?\s*[:\-–]+\s*", "", text, flags=re.I)
    text = re.sub(r"^(?:company|employer|organi[sz]ation)\s*\d+\s*[:\-–]+\s*", "", text, flags=re.I)
    text = re.sub(r"^(?:\w{0,3}ly\s+)?(?:currently\s+)?working\s+(?:(?:in|at|with|for)\s+)?", "", text, flags=re.I)
    text = re.sub(r"^\d+(?:\.\d+)?\+?\s*years?\s*(?:of\s+)?(?:experience\s+)?(?:in|at|with)\s+", "", text, flags=re.I)
    text = re.sub(r"^(?:current|present|previous)\s+(?:company|employer|organi[sz]ation)\s*[:\-–]?\s*(?=\S)", "", text, flags=re.I)
    text = re.sub(r"[“”\"]", "", text)
    text = re.sub(r"\s*\((?:working|current|present|currently working)\)\s*$", "", text, flags=re.I)
    text = re.sub(r"\s*[:\-–]+\s*\d+\s*(?:years?|months?|yrs?)\b.*$", "", text, flags=re.I)
    text = EMP_MONTH_TAIL_RE.sub("", text)
    text = text.strip(" ,:;|-–—")
    match = re.match(r"^(?P<title>.+?)\s*(?P<sep>,|\bin\b|\bat\b|\bfrom\b|\bwith\b|\bfor\b|[-–—|/@])\s+(?P<co>.+)$|^(?P<title2>[^(]+?)\s*(?P<sep2>\()\s*(?P<co2>[A-Za-z][^)]*)$", text, flags=re.I)
    if match and match.group("title2"):
        match = re.match(r"^(?P<title>[^(]+?)\s*(?P<sep>\()\s*(?P<co>[A-Za-z][^)]*)$", text)
    if match and EMP_TITLE_WORDS.search(match.group("title")) and len(match.group("title").split()) <= 6 and not re.search(EMP_LEGAL_RE, match.group("title"), re.I):
        co = re.sub(r"\s+(?:from|since)$", "", match.group("co").strip(" .,:;|-–—("), flags=re.I)
        co = re.sub(r"^in\s+", "", co, flags=re.I)
        cw = co.split()
        dash = match.group("sep") in ("-", "–", "—", "|", "/")
        ok = False
        if co and len(cw) <= 6:
            if norm_key(co) in known_companies():
                ok = True
            elif re.search(EMP_LEGAL_RE, co, re.I) and len(cw) >= 2:
                ok = True
            elif not dash and len(cw) >= 2 and COMPANY_SUFFIX_RE.search(cw[-1] + " "):
                ok = True
            elif not dash and len(cw) <= 3 and all(word[:1].isupper() and len(word) >= 3 for word in cw) and not EMP_TITLE_WORDS.search(co) and not any(word.lower() in ("test", "testing", "production", "projects", "operations", "team", "support") for word in cw):
                ok = not (heading_section(co) or norm_key(co) in lookups().skills or norm_key(co) in lookups().titles or {word.lower().strip(",.") for word in cw} <= INDIAN_PLACES)
        if ok:
            text = co
    if not text or EMP_LABEL_JUNK_RE.match(text) or EMP_LABEL_JUNK_RE2.match(text):
        return "", "label"
    legal = re.search(EMP_LEGAL_RE, text, re.I) is not None
    words = re.findall(r"[A-Za-z]+", text)
    last = words[-1] if words else ""
    ends_company = COMPANY_SUFFIX_RE.search(last + " ") is not None
    if not legal and not ends_company and words:
        level = EMP_TITLE_START_RE.match(text) is not None and EMP_TITLE_WORDS.search(text) is not None
        singular_title_end = EMP_TITLE_WORDS.search(last + " ") is not None and not last.lower().endswith("s")
        if (level or singular_title_end) and norm_key(text) not in known_companies() and "," not in text:
            return "", "title"
    return text, ("cleaned" if text != orig else "")


def employers_of(doc, is_list):
    """(current employers, past employers) of a record. Current only when the resume says the job is current (Present, Till date ...)."""
    current, past = [], []
    if is_list:
        c = doc.get("currentCompanyDetails") or {}
        name = final_employer(employer_name(c.get("companyName"), "sheet"))[0]
        old = [final_employer(employer_name(j.get("companyName"), "sheet"))[0] for j in doc.get("employmentHistory") or [] if j.get("endDate") != "current"]
        old = [item for item in dict.fromkeys(o for o in old if o) if not name or norm_key(item) != norm_key(name)]
        return ([name] if name else []), old
    for j in (doc.get("employmentHistory") or []) + (doc.get("internships") or []):
        name = employer_name(j.get("companyName"), j.get("companySource", ""))
        raw = clean(j.get("companyName") or "")
        if not name and raw and fallback_employer(raw, j.get("companySource", ""), bool(j.get("startDate"))):
            name = raw                                              # the job shows this company, so the employer field shows it too
        if name and bad_employer(name, doc.get("name")):
            name = ""
        if name:
            name = final_employer(name)[0]                          # a job title or a label is not an employer; a title glued to the company is cut
        if name:
            (current if j.get("endDate") == "current" else past).append(name)
    current = list(dict.fromkeys(current))
    keys = [norm_key(item) for item in current]
    past = [item for item in dict.fromkeys(past) if norm_key(item) not in keys]
    return current, past


def unique_ignoring_case(items):
    """The items without repeats ('Python' and 'python' are one); the last spelling wins, the order of first appearance stays."""
    return list({item.lower(): item for item in items}.values())


def education_facts(doc, is_list):
    """(degrees, fields of study, institutes) of a record: each value is a canonical degree / a field with a subject word / a real institution."""
    if is_list:
        rows = [{"degreeText": i if (DEGREE_RE.search(i) and not CORE_INSTITUTION.search(i)) else "", "fieldOfStudy": "", "institutionName": i if real_institution(i) else ""}
                for i in doc.get("educationItems") or []]
    else:
        rows = doc.get("educationalBackground") or []
    degrees, fields, institutes = [], [], []
    for entry in rows:
        text = entry.get("degreeText") or ""
        degree = canon_degree(text)
        shown_degree = degree or (clean(text) if (not is_list and text) else "")        # a school level ('MATRICULATION', 'HIGHER SECONDARY') is a value of the section as well
        if shown_degree:
            degrees.append(shown_degree)
        field = canon_field(entry.get("fieldOfStudy") or (field_of_study(text) if (degree and DEGREE_RE.match(text.strip())) else ""))
        if not field and entry.get("fieldOfStudy") and len(entry["fieldOfStudy"].split()) <= 8 and plausible_item(entry["fieldOfStudy"], 8):
            field = clean(entry["fieldOfStudy"])                        # a field the dictionary does not know, written in the resume ('Operations Research & Customer Relationship Management')
        if field and (degree or not is_list):
            fields.append(field)
        inst = institute_name(entry.get("institutionName")) or (clean(entry.get("institutionName") or "") if (not is_list and shown_degree) else "")
        if inst and (shown_degree or is_list):
            institutes.append(inst)
    if is_list and doc.get("listDegrees"):
        degrees += doc["listDegrees"]
        fields += doc.get("listFields") or []
    return unique_ignoring_case(degrees), unique_ignoring_case(fields), unique_ignoring_case(institutes)


def make_id(text):
    """A MeiliSearch id: letters and digits only. The same input always gives the same id."""
    return hashlib.md5(text.encode("utf-8")).hexdigest()


def job_line_text(j):
    j = dict(j)
    for field in ("designation", "companyName"):
        if j.get(field) and not plausible_item(j[field], 12):
            j[field] = ""                                    # a sentence, number or date is not a title or a company
    return job_text(j)


def search_document(doc):
    """A full JSON (resume or list person) -> the slim search document, or None when it cannot be searched (a list row with no
    e-mail, phone or LinkedIn). The text fields are shown and searched; the *_list fields, source_type and source_file are
    stored but hidden: the *_list ones exist for filters, the other two tell where the record came from."""
    is_list = doc.get("sourceType") == "candidate_list"
    if is_list:
        key = (doc.get("email") or "").lower() or re.sub(r"\D", "", doc.get("phoneNumber") or "")[-10:] or \
            next((part.get("url") for part in doc.get("socialProfiles") or [] if part.get("platform") == "linkedin"), "")
        if not key:
            if not doc.get("name"):
                return None
            key = f"row:{doc.get('source_file') or ''}#{doc.get('candidateNumber') or ''}:{(doc.get('name') or '').lower()}"      # a row with only a name is still a person: one JSON per row, the other fields stay empty
        doc_id = make_id("candidate:" + key)                    # one document per person, even if he is in several lists
    else:
        doc_id = make_id((doc.get("source_file") or doc.get("name") or "") + (f"#{doc['candidateNumber']}" if doc.get("candidateNumber") else ""))
    skills = [item for item in dict.fromkeys(doc.get("skills") or []) if skill_ok(item)]
    titles = [item for item in dict.fromkeys(doc.get("titles") or []) if plausible_item(item, 8)]
    industries = [item for item in dict.fromkeys(doc.get("industries") or []) if plausible_item(item, 8)]
    languages = doc.get("languagesKnown") or []
    other_sections = [f"{o['heading']}: {o['text']}" for o in doc.get("otherSections") or [] if o.get("text")]
    education = [education_text(entry) for entry in doc.get("educationalBackground") or []]
    if is_list:
        education = doc.get("educationItems") or []               # the degrees and institutions of the cell, without the sheet's own words
    education = [item for item in dict.fromkeys(education) if listed_item(item)]
    certifications = [item for item in dict.fromkeys(c.get("name") for c in doc.get("certifications") or [] if c.get("name")) if listed_item(item)]


    current_employers, past_employers = employers_of(doc, is_list)
    degrees, fields_of_study, institutes = education_facts(doc, is_list)
    experience = [text for text in (job_line_text(j) for j in doc.get("employmentHistory") or []) if text]
    if is_list and doc.get("totalExperienceText"):
        experience.append("Experience: " + doc["totalExperienceText"])
    loc = doc.get("locationDetails") or {}
    return {
        "id": doc_id,
        "name": doc.get("name"), "email": doc.get("email"),
        "linkedin": next((part.get("url") for part in doc.get("socialProfiles") or [] if part.get("platform") == "linkedin"), None),
        "phone": doc.get("phoneNumber"),
        "phones": [item for item in dict.fromkeys([doc.get("phoneNumber")] + list(doc.get("alternatePhoneNumbers") or [])) if item],
        "skills": ", ".join(skills), "titles": ", ".join(titles), "industries": ", ".join(industries), "languages": ", ".join(languages),
        "educations": "; ".join(education), "certifications": "; ".join(certifications),
        "projects": [item for item in dict.fromkeys(part.get("projectName") for part in doc.get("personalProjects") or [] if part.get("projectName")) if listed_item(item)],
        "internships": [text for text in (job_line_text(j) for j in doc.get("internships") or []) if text],
        "experience": experience,
        "current_employers": "; ".join(current_employers), "past_employers": "; ".join(past_employers),
        "degrees": "; ".join(degrees), "fields_of_study": "; ".join(fields_of_study), "institutes": "; ".join(institutes),
        "location": loc.get("text", ""), "city": loc.get("city", ""), "state": loc.get("state", ""), "country": loc.get("country", ""),
        "links": [u for u in dict.fromkeys(doc.get("links") or [])][:30],
        "other_sections": other_sections,
        # hidden: for filters and for tracing
        "cities_list": loc.get("cities", []), "states_list": loc.get("states", []), "countries_list": loc.get("countries", []),
        "current_employers_list": current_employers, "past_employers_list": past_employers,
        "degrees_list": degrees, "fields_of_study_list": fields_of_study, "institutes_list": institutes,
        "skills_list": skills, "titles_list": titles, "industries_list": industries, "languages_list": languages,
        "educations_list": education, "certifications_list": certifications,
        "source_type": "candidate_list" if is_list else "resume_file", "source_file": doc.get("source_file"),
    }


def identity_keys(doc):
    """What says that two JSONs are the same person: the e-mail, the 10 digits of the phone, the LinkedIn link."""
    keys = []
    if doc.get("email"):
        keys.append("e:" + doc["email"].lower())
    digits = re.sub(r"\D", "", doc.get("phone") or "")[-10:]
    if len(digits) == 10:
        keys.append("p:" + digits)
    handle = (doc.get("linkedin") or "").lower().rstrip("/").rsplit("/", 1)[-1]
    if len(handle) >= 8:                                                # '/in/ajay' names nobody in particular
        keys.append("l:" + doc["linkedin"].lower().rstrip("/"))
    return keys


def richness(doc):
    """How much a slim JSON holds: of two copies of one person the richer one stays in Valid (a resume beats a spreadsheet row)."""
    fields = ("skills_list", "titles_list", "industries_list", "educations_list", "certifications_list", "projects", "internships", "experience")
    return sum(len(doc.get(k) or []) for k in fields) + 3 * bool(doc.get("name")) + 2 * (doc.get("source_type") == "resume_file")


def career_facts(doc):
    """The employers and institutes of a slim JSON, as lower-case letters only, to tell two people with the same contact details apart."""
    items = (doc.get("current_employers_list") or []) + (doc.get("past_employers_list") or []) + (doc.get("institutes_list") or [])
    return {re.sub(r"[^a-z]", "", str(item).lower()) for item in items if len(re.sub(r"[^a-z]", "", str(item).lower())) >= 4}


def career_conflict(facts_a, facts_b):
    """Two resumes with the same e-mail / phone but no employer and no institute in common are two people (a shared office phone, a family
    e-mail, a recruiter's number). Needs at least 2 facts on each side; a candidate-list row or a thin resume never separates people."""
    if len(facts_a) < 2 or len(facts_b) < 2:
        return False
    return not any(a == b or (min(len(a), len(b)) >= 8 and fuzz.partial_ratio(a, b) >= 90) for a in facts_a for b in facts_b)


def same_person(name_a, name_b):
    """Two JSONs with the same e-mail / phone are one person unless both have a name and the names have nothing in common."""
    a = {word for word in re.split(r"[^a-z]+", (name_a or "").lower()) if len(word) >= 3}
    b = {word for word in re.split(r"[^a-z]+", (name_b or "").lower()) if len(word) >= 3}
    if not a or not b:
        return True
    common = a & b
    if len(common) >= 2 or (common and min(len(a), len(b)) == 1):
        return True                                                     # two name parts in common, or a one-word name that is in the other
    if fuzz.token_sort_ratio(" ".join(sorted(a)), " ".join(sorted(b))) >= 88:       # a typo, not just a shared first name ('Ajay Bansod' / 'Ajay Kumar')
        return True
    glued_a, glued_b = re.sub(r"[^a-z]", "", (name_a or "").lower()), re.sub(r"[^a-z]", "", (name_b or "").lower())
    return fuzz.ratio(glued_a, glued_b) >= 88 or (min(len(glued_a), len(glued_b)) >= 8 and (glued_a in glued_b or glued_b in glued_a))       # 'GopalLalKumawat' / 'Gopal Lal Kumawat'


def run_job(job):
    """Worker: one file -> (job, JSON dict, parts). Never raises: a file that cannot be read still gets a JSON with the reason.
    parts = [(candidate number, JSON dict)] with one entry per candidate when --split-candidates found several, else []."""
    parts = []
    data = {"status": job["status"], "reason": job["reason"], "category": job["category"],
            "source_file": job["name"], "source_folder": job["folder"], "parsed": False,
            "sourceType": "candidate_list" if job["mode"] == "table" else "resume", "candidateNumber": None}
    try:
        if job["mode"] == "table":
            data.update({"type": "candidate_list", **parse_table(read_table(long_path(job["read"])))})
            data["parsed"] = True
            if len(data["candidates"]) >= 1:
                parts = [(n, list_person_record(c, job["name"], job["folder"], n, job["status"])) for n, c in enumerate(data["candidates"], 1)]
        elif job["mode"] == "parse":
            lines = read_lines(long_path(job["read"]))
            if not lines:
                raise ValueError("no text could be read from the file")
            hyperlinks = read_hyperlinks(long_path(job["read"]))
            pieces = split_bundle(lines) if job.get("split_bundles") else [lines]
            if len(pieces) > 1:
                parts = [(n, {"status": job["status"], "reason": job["reason"], "category": job["category"], "source_file": job["name"],
                              "source_folder": job["folder"], "parsed": True, "sourceType": "resume", "candidateNumber": n,
                              **parse_lines(piece, job["name"], hyperlinks)}) for n, piece in enumerate(pieces, 1)]
                data["type"] = "bundle"
            else:
                data.update(parse_lines(lines, job["name"], hyperlinks))
                data["text_fingerprint"] = text_fingerprint(lines)
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


def clean_before_lists(out_list, out_dup):
    """A lists run makes the list JSON again: Output/List is emptied, and the old JSON of list people in Duplicates is deleted.
    Returns the names of the resume JSON in Duplicates (needed to put back a resume afterwards)."""
    for name in os.listdir(out_list):
        if name.lower().endswith(".json"):
            os.remove(os.path.join(out_list, name))
    resume_dups, removed = set(), 0
    for name in os.listdir(out_dup):
        if not name.lower().endswith(".json"):
            continue
        path = os.path.join(out_dup, name)
        try:
            with open(path, encoding="utf-8") as f:
                kind = json.load(f).get("source_type")
        except Exception:
            continue
        if kind == "candidate_list":
            os.remove(path)
            removed += 1
        else:
            resume_dups.add(name)
    print(f"List folder emptied; {removed} old list-person JSON removed from Duplicates.")
    return resume_dups


def text_fingerprint(lines):
    """Two resumes are the same resume when every word of them is the same: all the words of the whole document, each counted as often as it
    appears, spacing and reading order ignored (the same page read column by column or line by line gives the same words)."""
    words = sorted(word for line in lines if line.text for word in line.text.split())
    return hashlib.sha256(" ".join(words).encode("utf-8")).hexdigest()


def content_fingerprint(doc):
    """What makes two slim JSONs identical: every field except the ones that only say where it came from."""
    body = {k: v for k, v in doc.items() if k not in ("id", "source_file", "duplicate_of", "source_folder", "candidateNumber")}
    return hashlib.sha256(json.dumps(body, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()


def load_json_or_none(path):
    """The JSON content of a file, or None when it cannot be read."""
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def repoint_duplicates(out_slim, out_dup):
    """Makes sure every Duplicates JSON names, in 'duplicate_of', the identical JSON that is in Valid (same content). Returns (repointed,
    unresolved); a duplicate whose twin is not in Valid any more is unresolved."""
    twin, in_valid = {}, set(os.listdir(out_slim))
    for name in in_valid:
        doc = load_json_or_none(os.path.join(out_slim, name))
        if doc:
            twin.setdefault(content_fingerprint(doc), name)
    fixed = unresolved = 0
    for name in os.listdir(out_dup):
        path = os.path.join(out_dup, name)
        doc = load_json_or_none(path)
        if not doc or doc.get("duplicate_of") in in_valid:
            continue
        target = twin.get(content_fingerprint(doc))
        if target is None:
            unresolved += 1
        elif doc.get("duplicate_of") != target:
            doc["duplicate_of"] = target
            with open(path, "w", encoding="utf-8") as f:
                json.dump(doc, f, ensure_ascii=False, indent=1)
            fixed += 1
    return fixed, unresolved


def drop_lists_from_valid(out_slim, out_dup, resume_dups):
    """After a lists run: the list people are in Output/List now, so their JSON is deleted from Valid. A resume that sits in Duplicates only because a
    list copy of the same person had taken its place in Valid is moved back to Valid."""
    removed, resume_keys = set(), set()
    for name in os.listdir(out_slim):
        if not name.lower().endswith(".json"):
            continue
        path = os.path.join(out_slim, name)
        try:
            with open(path, encoding="utf-8") as f:
                doc = json.load(f)
        except Exception:
            continue
        if doc.get("source_type") == "candidate_list":
            os.remove(path)
            removed.add(name)
        else:
            resume_keys.update(identity_keys(doc))
    restored = 0
    for name in resume_dups:
        path = os.path.join(out_dup, name)
        try:
            with open(path, encoding="utf-8") as f:
                doc = json.load(f)
        except Exception:
            continue
        if doc.get("duplicate_of") in removed and not (set(identity_keys(doc)) & resume_keys):
            doc.pop("duplicate_of", None)
            with open(os.path.join(out_slim, name), "w", encoding="utf-8") as f:
                json.dump(doc, f, ensure_ascii=False, indent=1)
            os.remove(path)
            resume_keys.update(identity_keys(doc))
            restored += 1
    print(f"{len(removed)} list-person JSON deleted from Valid; {restored} resume JSON put back from Duplicates.")


def save_json(folder, file_name, doc):
    """Writes one JSON file."""
    with open(os.path.join(folder, file_name), "w", encoding="utf-8") as handle:
        json.dump(doc, handle, ensure_ascii=False, indent=1)


class ValidWriter:
    """Writes the slim JSON of every parsed record to Valid, or to Duplicates when everything in it is the same as a record already in Valid."""

    def __init__(self, out_slim, out_list, out_dup, kind):
        self.out_slim = out_slim
        self.out_list = out_list
        self.out_dup = out_dup
        self.kind = kind
        self.seen_fingerprints = {}                                     # content fingerprint -> file that stays in Valid
        self.seen_ids = {}                                              # document id -> fingerprint
        self.stats = collections.Counter()

    def write(self, file_name, record):
        """A JSON goes to Duplicates only when EVERYTHING in it is the same as one already in Valid (name, contact, skills, jobs, education ...).
        Same e-mail / phone / name with anything different is another record and stays in Valid. Every file that classification put in Valid
        is parsed and written here, even when no contact data was found in it."""
        doc = search_document(record) if (record.get("parsed") and record.get("status") == "valid" and "name" in record) else None
        if not doc:
            return
        is_list = doc.get("source_type") == "candidate_list"
        home = self.out_list if (is_list and self.kind == "lists") else self.out_slim
        fingerprint = record.get("text_fingerprint") or content_fingerprint(doc)       # a resume: its whole text; a person of a list: all the fields of the row
        if fingerprint in self.seen_fingerprints:
            doc["duplicate_of"] = self.seen_fingerprints[fingerprint]
            save_json(self.out_dup, file_name, doc)
            self.stats["duplicate -> Duplicates"] += 1
            return
        if self.seen_ids.get(doc["id"], fingerprint) != fingerprint:   # same person key, different content: its own document in the search index
            doc["id"] = make_id(doc["id"] + ":" + fingerprint[:12])
        self.seen_ids[doc["id"]] = fingerprint
        self.seen_fingerprints[fingerprint] = file_name
        save_json(home, file_name, doc)
        self.stats["unique -> List" if home == self.out_list else "unique -> Valid"] += 1


def main():
    arg_parser = argparse.ArgumentParser(description="Parse every classified file into its own JSON (parsing only).")
    arg_parser.add_argument("--classification", default=DEFAULT_CLASSIFICATION, help="folder with Valid and Invalid (default: ..\\classification)")
    arg_parser.add_argument("--output", default=DEFAULT_OUTPUT, help="folder that gets Valid and Invalid sub-folders of JSON files (default: ..\\Output)")
    arg_parser.add_argument("--only", choices=["both", "valid", "invalid"], default="both")
    arg_parser.add_argument("--workers", type=int, default=min(os.cpu_count() or 1, 3))
    arg_parser.add_argument("--split-candidates", action="store_true",
                    help="kept so old commands still work: a candidate list (CSV/Excel) is ALWAYS written as one JSON per person ("
                         "'<name> - 001.json', ...). A resume file always gives ONE JSON.")
    arg_parser.add_argument("--split-bundles", action="store_true",
                    help="also cut a PDF/DOCX that holds several resumes into one JSON per resume (off by default: it wrongly cut "
                         "28 of 31 single resumes in two, because a second e-mail address is not a second candidate)")
    arg_parser.add_argument("--limit", type=int, default=0, help="only the first N files of each folder (to try it out)")
    arg_parser.add_argument("--kind", choices=["all", "resumes", "lists"], default="all",
                    help="all (default): resumes and candidate lists, everything goes to Output\\Valid; resumes: everything except the lists; "
                         "lists: only the candidate lists, into Output\\List (and their old JSON is deleted from Valid)")
    args = arg_parser.parse_args()

    jobs = collect_jobs(args.classification, args.only, args.limit)
    if args.kind == "lists":
        jobs = [j for j in jobs if j["status"] == "valid" and j["mode"] == "table"]
    elif args.kind == "resumes":
        jobs = [j for j in jobs if not (j["status"] == "valid" and j["mode"] == "table")]
    for j in jobs:
        j["split"] = args.split_candidates
        j["split_bundles"] = args.split_bundles
    out_slim = os.path.join(args.output, "Valid")                      # the slim search JSON of every valid resume / list person: what is uploaded
    out_invalid = os.path.join(args.output, "Invalid")                 # a JSON with the reason for every invalid file (and for a valid file that could not be read)
    out_dup = os.path.join(args.output, "Duplicates")                  # the second, third ... copy of a person who is already in Valid
    out_list = os.path.join(args.output, "List")                       # the people of the candidate lists (CSV / Excel), one JSON per person
    for folder in (out_slim, out_invalid, out_dup) + ((out_list,) if args.kind == "lists" else ()):
        os.makedirs(folder, exist_ok=True)
    resume_dups = clean_before_lists(out_list, out_dup) if args.kind == "lists" else set()
    if args.kind != "lists" and os.path.isdir(out_list):                  # the list people are in Valid again: the old List folder goes
        for name in os.listdir(out_list):
            if name.lower().endswith(".json"):
                os.remove(os.path.join(out_list, name))
        try:
            os.rmdir(out_list)
        except OSError:
            pass
    writer = ValidWriter(out_slim, out_list, out_dup, args.kind)

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
                if job["status"] == "valid" and data.get("parsed"):
                    stem = json_name(out_list if (parts and args.kind == "lists") else out_slim, job["name"], used)[:-5]
                    if parts:                                              # a candidate list: one slim JSON per person, in Output\List
                        for k, part in parts:
                            writer.write(f"{stem} - {k:03d}.json", part)
                        done[("valid", "list split into one JSON per person")] += 1
                    else:
                        writer.write(stem + ".json", data)
                        done[("valid", "parsed")] += 1
                else:                                                      # invalid, or a valid file that could not be read: the reason goes to Invalid
                    with open(os.path.join(out_invalid, json_name(out_invalid, job["name"], used)), "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=1)
                    done[(job["status"], "parsed" if data["parsed"] else "reason only / not readable")] += 1
                if n % 250 == 0:
                    print(f"  {n}/{len(jobs)} done ({n / (time.time() - started):.1f} per second)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if args.kind == "lists":
        drop_lists_from_valid(out_slim, out_dup, resume_dups)
    fixed, unresolved = repoint_duplicates(out_slim, out_dup)
    print(f"{fixed} 'duplicate_of' pointers updated to the copy that is now in Valid; {unresolved} could not be resolved.")
    print(f"name_words.json rebuilt from the names in Valid: {build_name_words(out_slim)} words (used to cut letter-spaced names on the next parse).")
    print(f"\nDONE in {round(time.time() - started)} seconds.")
    for (status, kind), count in sorted(done.items()):
        print(f"  {status:8} {kind:30} {count}")
    for kind, count in sorted(writer.stats.items()):
        print(f"  {kind:36} {count}")
    print(f"  Valid      -> {out_slim}   (one JSON per person: resumes and the people of the candidate lists)\n"
          + (f"  List       -> {out_list}   (--kind lists: the people of the candidate lists)\n" if args.kind == "lists" else "")
          + f"  Duplicates -> {out_dup}   (a copy of a person who is already in Valid)\n  Invalid    -> {out_invalid}")


if __name__ == "__main__":
    main()
