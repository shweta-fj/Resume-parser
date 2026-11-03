"""
classify_resumes.py  --  CLASSIFICATION ONLY: is each file a real resume, or not (and why)?

It does not parse resumes into JSON and it never touches MeiliSearch. It looks at every file in the
Input folder, whatever its type, and sorts it:

    classification/Valid/      the real resumes (the files themselves)
    classification/Invalid/    everything else, in a sub-folder named after the reason

and writes the reason for every file next to them:

    classification/Valid/valid_resumes.csv
    classification/Invalid/invalid_reasons.csv
    classification/classification_report.csv      (every file, one row each)
    classification/classification_summary.json

There is no list of allowed file types. Every file is opened and read as far as it can be:
  PDF (text layer, then OCR of the pages, then OCR of the pictures inside)   DOCX
  DOC / RTF / ODT / HTML (turned into DOCX with LibreOffice)                 images (OCR)
  plain text, markdown, code, CSV                                            XLSX / XLS / ODS spreadsheets
  zip files (opened, every file inside is classified, nested zips too)
  files saved inside a web-upload body ("------WebKitFormBoundary...") (the real file is cut out)
  anything else (videos, programs, ...) is classified too, as "not a document", with the reason.

A file is a valid resume unless it is clearly something else: a certificate / invoice / statement,
 a job description (no contact details, advert wording), or a file with no readable text. There is
no requirement for a name, a section, contact details or a number of sections. A spreadsheet that lists
candidates (one per row, with name, contact and profile columns) is valid too and is marked as a candidate
list. Two files with the same content: the first
is classified, the others are Invalid ("duplicate of ...").

HOW TO RUN (from this folder, in PowerShell):
    python classify_resumes.py
    python classify_resumes.py --input "..\\Input" --output "..\\classification" --workers 3
Files are put into Valid / Invalid as hard links by default: no extra disk space, the Input folder is left
exactly as it is. Use --copy for real copies or --move to move them out of the Input folder.
"""

import os
import re
import io
import sys
import csv
import json
import time
import shutil
import hashlib
import tempfile
import argparse
import subprocess
import collections
import xml.etree.ElementTree as ET
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import resume_engine as E          # the reading engine (same rules as the parser)

ROOT = os.path.dirname(HERE)
DEFAULT_INPUT = os.path.join(ROOT, "Input")
DEFAULT_OUTPUT = os.path.join(ROOT, "classification")

CODE_EXTENSIONS = {".js", ".css", ".json", ".xml", ".svg", ".map", ".ts", ".py", ".java", ".sql", ".yml", ".yaml",
                   ".ini", ".log", ".bat", ".sh", ".php", ".c", ".cpp", ".h"}
VIDEO_AUDIO_SIGNS = ((b"\x1a\x45\xdf\xa3", "video (webm/mkv)"), (b"OggS", "audio/video (ogg)"), (b"ID3", "audio (mp3)"),
                     (b"RIFF", "audio/video (wav/avi)"), (b"fLaC", "audio (flac)"))
SPREADSHEET_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".ods"}

# short reason -> folder name inside Invalid/
CATEGORIES = [
    ("duplicate of", "duplicate"),
    ("empty file", "empty_file"),
    ("certificate, statement", "certificate_invoice_statement"),
    ("job description", "job_description"),
    ("no readable text", "no_readable_text"),
    ("spreadsheet or csv that is not", "spreadsheet_not_candidates"),
    ("not a document", "not_a_document"),
    ("web or code file", "web_or_code_file"),
    ("could not", "could_not_read"),
]


def category_of(reason):
    low = reason.lower()
    for key, folder in CATEGORIES:
        if key in low:
            return folder
    return "other"


def long_path(path):
    path = os.path.abspath(path)
    return "\\\\?\\" + path if os.name == "nt" and not path.startswith("\\\\?\\") else path


def pick_temp_folder(wanted):
    if wanted:
        os.makedirs(wanted, exist_ok=True)
        return wanted
    best, space = None, -1
    for drive in ("D:\\", "E:\\", "C:\\"):
        if os.path.exists(drive):
            free = shutil.disk_usage(drive).free
            if free > space:
                best, space = drive, free
    folder = os.path.join(best or tempfile.gettempdir(), "classify_tmp")
    os.makedirs(folder, exist_ok=True)
    return folder


# ----------------------------------------------------------------------------------------------------
#  looking at a file
# ----------------------------------------------------------------------------------------------------

def looks_like_text(head):
    """True if these first bytes are readable text (UTF-8, UTF-16 or a Latin-1 style code page)."""
    if not head:
        return False
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return True
    if b"\x00" in head:
        return False
    printable = sum(1 for b in head if b in (9, 10, 13) or 32 <= b < 127 or b >= 160)
    return printable / len(head) > 0.9


def repair_zip_tail(data, target):
    """Some saved Word/Excel files have a web page or other junk written after the zip. Cut at the end of the
    zip (its end-of-directory record) and keep the part that is the real document. Returns the new path or None."""
    import struct
    i = data.rfind(b"PK\x05\x06")
    if i < 0 or i + 22 > len(data):
        return None
    end = i + 22 + struct.unpack("<H", data[i + 20:i + 22])[0]
    try:
        with open(target, "wb") as handle:
            handle.write(data[:end])
        with zipfile.ZipFile(target) as z:
            return target if z.namelist() else None
    except (zipfile.BadZipFile, OSError):
        return None


def is_xlsx(path):
    try:
        with zipfile.ZipFile(path) as z:
            return "xl/workbook.xml" in z.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def read_xlsx_rows(path, max_rows=80):
    """The first rows of the first sheet of an .xlsx file, without any extra library."""
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rows = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(ns + "si"):
                shared.append("".join(t.text or "" for t in si.iter(ns + "t")))
        sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
        if not sheets:
            return rows
        for _, row in ET.iterparse(io.BytesIO(z.read(sheets[0])), events=("end",)):
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
                value = c.find(ns + "v")
                text = ""
                if c.get("t") == "s" and value is not None and value.text and value.text.isdigit():
                    index = int(value.text)
                    text = shared[index] if index < len(shared) else ""
                elif c.get("t") == "inlineStr":
                    text = "".join(t.text or "" for t in c.iter(ns + "t"))
                elif value is not None:
                    text = value.text or ""
                cells[col] = text
            if cells:
                rows.append([cells.get(i, "") for i in range(1, max(cells) + 1)])
            row.clear()
            if len(rows) >= max_rows:
                break
    return rows


def convert_spreadsheets_to_csv(paths, temp_folder):
    """Old Excel / OpenDocument spreadsheets -> CSV text with LibreOffice. {original: csv path}"""
    out = {}
    office = E.find_libreoffice()
    if not paths or not office:
        return out
    os.makedirs(temp_folder, exist_ok=True)
    copies = {}
    for number, path in enumerate(paths):
        copy = os.path.join(temp_folder, f"{number:06d}{os.path.splitext(path)[1].lower() or '.xls'}")
        try:
            shutil.copy2(path, copy)
        except OSError:
            continue
        copies[path] = copy
    for start in range(0, len(copies), 25):
        group = list(copies.values())[start:start + 25]
        try:
            subprocess.run([office, "--headless", "--norestore", "--convert-to", "csv", "--outdir", temp_folder] + group,
                           capture_output=True, timeout=60 * len(group), check=False)
        except Exception:
            pass
    for path, copy in copies.items():
        csv_path = os.path.splitext(copy)[0] + ".csv"
        if os.path.exists(csv_path):
            out[path] = csv_path
    return out


_NAME_COLUMN = re.compile(r"candidate\s*name|name\s+of\s+(the\s+)?candidate|full\s*name|first\s*name|^name$|^name\b", re.IGNORECASE)
_NOT_A_PERSON_NAME = re.compile(r"executive|recruiter|vendor|service\s+provider|client|company|account|dealer|req\b|requirement|manager|lead", re.IGNORECASE)
_CONTACT_COLUMN = re.compile(r"mob|phone|contact\s*n|e-?mail", re.IGNORECASE)
_PROFILE_COLUMN = re.compile(r"skill|experience|qualif|designation|education|ctc|notice|current\s*company|location|resume|\bcv\b|degree|college", re.IGNORECASE)


def candidate_list_header(header):
    """True if these column names are a list of candidates: a person's name column, a contact column
    (mobile, phone, email) and at least one profile column (skills, experience, qualification, designation,
    CTC, location ...). This covers ATS exports, job-portal exports and recruiter submission trackers.
    A list of companies, executives to call, statistics or bank lines has no such combination."""
    names = [str(h).strip() for h in header if str(h).strip()]
    if E.detect_candidate_csv((",".join(names)).encode("utf-8")):
        return True
    has_name = any(_NAME_COLUMN.search(h) and not _NOT_A_PERSON_NAME.search(h) for h in names)
    has_contact = any(_CONTACT_COLUMN.search(h) for h in names)
    has_profile = any(_PROFILE_COLUMN.search(h) for h in names)
    return has_name and has_contact and has_profile


def describe_columns(header, limit=8):
    names = [str(h).strip() for h in header if str(h).strip()][:limit]
    return ", ".join(names) if names else "no header"


def classify_table(rows):
    """('valid' | 'invalid', reason, note) for a table (CSV or spreadsheet) given as a list of rows."""
    if not rows:
        return "invalid", "could not read any rows from this spreadsheet or CSV", ""
    header = [str(c).strip().lower() for c in rows[0]]
    if len(header) <= 2 and header and "\t" in header[0]:
        header = header[0].split("\t")
        rows = [r[0].split("\t") if r and "\t" in r[0] else r for r in rows]
    if candidate_list_header(header):
        return "valid", "", f"candidate list: {max(0, len(rows) - 1)}+ rows, one person per row (columns: {describe_columns(header)})"
    return "invalid", f"spreadsheet or CSV that is not a list of candidates (columns: {describe_columns(header)})", ""


# ----------------------------------------------------------------------------------------------------
#  placing the files
# ----------------------------------------------------------------------------------------------------

def safe_name(text):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text).strip(" .") or "file"


def place(source, folder, wanted_name, mode, used):
    """Puts the file into the folder as a link, copy or move. Returns the name it got there."""
    os.makedirs(folder, exist_ok=True)
    base, extension = os.path.splitext(safe_name(wanted_name))
    base = base[:120]
    name = base + extension
    counter = 1
    while (folder, name.lower()) in used or os.path.exists(os.path.join(folder, name)):
        counter += 1
        name = f"{base} ({counter}){extension}"
    used.add((folder, name.lower()))
    target = long_path(os.path.join(folder, name))
    src = long_path(source)
    if mode == "move":
        shutil.move(src, target)
    elif mode == "copy":
        shutil.copy2(src, target)
    else:
        try:
            os.link(src, target)
        except OSError:
            shutil.copy2(src, target)     # another disk, or a file system without links
    return name


# ----------------------------------------------------------------------------------------------------
#  the run
# ----------------------------------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Classify every file in the Input folder as a valid resume or not.")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="folder with the files (default: ..\\Input)")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="classification folder (default: ..\\classification)")
    parser.add_argument("--workers", type=int, default=min(os.cpu_count() or 1, 3), help="files read at once (default 3)")
    parser.add_argument("--temp", default=None, help="folder for temporary files (default: the drive with most free space)")
    how = parser.add_mutually_exclusive_group()
    how.add_argument("--copy", action="store_true", help="put real copies in Valid/Invalid (needs the disk space)")
    how.add_argument("--move", action="store_true", help="move the files out of the Input folder")
    how.add_argument("--report-only", action="store_true", help="only write the reports, put no files in Valid/Invalid")
    parser.add_argument("--clean", action="store_true", help="empty Valid and Invalid first (needed to run again)")
    args = parser.parse_args()
    mode = "move" if args.move else "copy" if args.copy else "link"

    if not os.path.isdir(args.input):
        print(f"Input folder not found: {args.input}")
        return
    valid_folder = os.path.join(args.output, "Valid")
    invalid_folder = os.path.join(args.output, "Invalid")
    for folder in (valid_folder, invalid_folder):
        os.makedirs(folder, exist_ok=True)
        if os.listdir(folder):
            if not args.clean:
                print(f"{folder} is not empty. Run again with --clean to empty Valid and Invalid first.")
                return
    if args.clean:
        for folder in (valid_folder, invalid_folder):
            for entry in os.listdir(folder):
                path = os.path.join(folder, entry)
                if os.path.isdir(path):
                    shutil.rmtree(long_path(path), ignore_errors=True)
                else:
                    os.remove(path)

    temp_root = pick_temp_folder(args.temp)
    temp = tempfile.mkdtemp(prefix="run_", dir=temp_root)
    zip_temp = os.path.join(temp, "zip")
    work_temp = os.path.join(temp, "work")
    os.makedirs(zip_temp), os.makedirs(work_temp)
    started = time.time()
    try:
        # 1. every file, zips opened
        files, zip_problems = E.find_all_files(args.input, zip_temp, [], True)
        files = E.unwrap_uploaded_files(files, zip_temp)
        files.sort(key=lambda f: (".zip/" in f[0].lower(), f[0].lower()))   # a loose file wins over its copy inside a zip
        print(f"Found {len(files)} files in {args.input} (zip files opened).")

        records = {}      # id -> record
        order = []

        def new_record(inside, real):
            rid = E.make_file_id(inside)
            rec = {"id": rid, "file": inside, "real_path": real, "result": "", "reason": "", "note": "",
                   "category": "", "name": "", "email": "", "phone": "", "linkedin": "", "sections": "",
                   "section_count": "", "kind": ""}
            records[rid] = rec
            order.append(rid)
            return rec

        def mark(rec, result, reason="", note=""):
            rec["result"], rec["reason"], rec["note"] = result, reason, note
            if result == "Invalid":
                rec["category"] = category_of(reason)

        for inside, real in files:
            new_record(inside, real)
        for inside, reason in zip_problems:
            rec = new_record(inside, "")
            mark(rec, "Invalid", f"could not open: {reason}")

        # 2. look at each file: empty, duplicate, what it is
        seen_content = {}
        to_read = []              # (path_to_read, inside, mime, id) for the reading engine
        old_formats, sheets = [], []
        by_kind = collections.Counter()
        for rid in list(order):
            rec = records[rid]
            real, inside = rec["real_path"], rec["file"]
            if not real or rec["result"]:
                continue
            ext = os.path.splitext(inside)[1].lower()
            try:
                size = os.path.getsize(real)
                with open(real, "rb") as handle:
                    data = handle.read()
            except OSError as problem:
                mark(rec, "Invalid", f"could not read the file: {problem}")
                continue
            if size == 0:
                mark(rec, "Invalid", "empty file (0 bytes)")
                continue
            digest = hashlib.sha256(data).hexdigest()
            if digest in seen_content:
                mark(rec, "Invalid", f"duplicate of {seen_content[digest]}")
                continue
            seen_content[digest] = inside
            head = data[:4096]
            if head[:4] == b"PK\x03\x04" and not zipfile.is_zipfile(real):
                fixed = repair_zip_tail(data, os.path.join(work_temp, f"repaired_{len(seen_content):06d}{ext or '.zip'}"))
                if fixed:
                    rec["real_path"] = real = fixed
                    head = open(real, "rb").read(4096)
            mime = E.detect_mime_type(real)
            if ext in SPREADSHEET_EXTENSIONS or (mime is None and head[:2] == b"PK" and is_xlsx(real)):
                rec["kind"] = "spreadsheet"
                if head[:2] == b"PK" and is_xlsx(real):
                    try:
                        verdict, reason, note = classify_table(read_xlsx_rows(real))
                    except Exception as problem:
                        verdict, reason, note = "invalid", f"could not read this spreadsheet: {problem}", ""
                    mark(rec, "Valid" if verdict == "valid" else "Invalid", reason, note)
                else:
                    sheets.append(rid)
                by_kind["spreadsheet"] += 1
                continue
            if mime == E.CSV_MIME:
                rec["kind"] = "candidate list (CSV)"
                try:
                    header, rows = E.read_csv_rows(real)
                    mark(rec, "Valid", "", f"candidate list: {len(rows)} rows, one person per row (columns: {describe_columns(header)})")
                except Exception as problem:
                    mark(rec, "Invalid", f"could not read the CSV: {problem}")
                by_kind["candidate list (CSV)"] += 1
                continue
            if mime in (E.PDF_MIME, E.DOCX_MIME) or mime in E.IMAGE_MIMES:
                rec["kind"] = {E.PDF_MIME: "PDF", E.DOCX_MIME: "DOCX"}.get(mime, "image")
                to_read.append((real, inside, mime, rid))
            elif mime in E.CONVERT_MIMES:
                rec["kind"] = {E.DOC_MIME: "DOC", E.RTF_MIME: "RTF", E.ODT_MIME: "ODT", E.HTML_MIME: "HTML"}[mime]
                old_formats.append(rid)
            else:
                signature = next((label for sign, label in VIDEO_AUDIO_SIGNS if head.startswith(sign)), None)
                if signature is None and len(head) > 8 and head[4:8] == b"ftyp":
                    signature = "video (mp4/mov)"
                if signature is None and head[:2] == b"MZ":
                    signature = "program (exe/dll)"
                if signature:
                    rec["kind"] = signature
                    mark(rec, "Invalid", f"not a document: {signature}, it has no resume text")
                elif looks_like_text(head):
                    if ext in (".csv", ".tsv"):
                        rec["kind"] = "CSV"
                        try:
                            header, rows = E.read_csv_rows(real)
                            verdict, reason, note = classify_table([header] + rows[:50])
                        except Exception as problem:
                            verdict, reason, note = "invalid", f"could not read the CSV: {problem}", ""
                        mark(rec, "Valid" if verdict == "valid" else "Invalid", reason, note)
                    else:                                       # no file type is refused by its extension: a text file is read like any other, whatever it is called
                        rec["kind"] = "text"
                        to_read.append((real, inside, E.TEXT_MIME, rid))
                else:
                    rec["kind"] = "binary"
                    mark(rec, "Invalid", f"not a document: unknown binary file ({ext or 'no extension'})")
            by_kind[rec["kind"]] += 1

        # 2b. spreadsheets that need LibreOffice
        if sheets:
            converted = convert_spreadsheets_to_csv([records[r]["real_path"] for r in sheets], os.path.join(work_temp, "sheets"))
            for rid in sheets:
                rec = records[rid]
                csv_path = converted.get(rec["real_path"])
                if not csv_path:
                    mark(rec, "Invalid", "could not read this spreadsheet (LibreOffice could not convert it)")
                    continue
                try:
                    header, rows = E.read_csv_rows(csv_path)
                    verdict, reason, note = classify_table([header] + rows[:50])
                except Exception as problem:
                    verdict, reason, note = "invalid", f"could not read this spreadsheet: {problem}", ""
                mark(rec, "Valid" if verdict == "valid" else "Invalid", reason, note)

        # 2c. DOC / RTF / ODT / HTML -> DOCX
        if old_formats:
            converted = E.convert_old_formats([records[r]["real_path"] for r in old_formats], os.path.join(work_temp, "convert"))
            print(f"Old formats and web pages (DOC/RTF/ODT/HTML): {len(converted)} of {len(old_formats)} converted for reading.")
            for rid in old_formats:
                rec = records[rid]
                docx = converted.get(rec["real_path"])
                if docx:
                    to_read.append((docx, rec["file"], E.DOCX_MIME, rid))
                else:
                    mark(rec, "Invalid", f"could not convert this {rec['kind']} file to read it (LibreOffice failed)")

        # 3. read every file and decide
        print(f"{len(to_read)} files to read.")

        def when_done(kind, data):
            rec = records.get(data.get("id"))
            if rec is None:
                return
            for key in ("name", "email", "phone", "linkedin"):
                if data.get(key):
                    rec[key] = data[key]
            if kind == "real":
                rec["sections"] = ", ".join(data.get("sections_found", []))
                rec["section_count"] = data.get("section_count", "")
                mark(rec, "Valid")
            elif kind == "invalid":
                mark(rec, "Invalid", data.get("reason", "not a resume"))
            else:
                error = str(data.get("error", "unknown error"))
                low = error.lower()
                if "encrypted" in low or "password" in low or "document closed" in low:
                    mark(rec, "Invalid", "could not read: password-protected or encrypted file (it cannot be opened without the password)")
                elif "trailer" in low or "page count" in low or "xref" in low:
                    mark(rec, "Invalid", "could not read: damaged PDF (broken structure, and OCR found no text)")
                elif "index out of range" in low or "no text" in low or "empty" in low:
                    mark(rec, "Invalid", "no readable text in the file (an empty document)")
                else:
                    mark(rec, "Invalid", f"could not read the file: {error[:150]}")

        if to_read:
            E.parse_all_files(to_read, args.workers, os.path.join(HERE, "lookups"), when_done)
        for rid in order:
            if not records[rid]["result"]:
                mark(records[rid], "Invalid", "could not read the file (no result)")

        # 4. put the files in Valid / Invalid and write the reports
        used = set()
        if not args.report_only:
            for rid in order:
                rec = records[rid]
                if not rec["real_path"]:
                    continue
                folder = valid_folder if rec["result"] == "Valid" else os.path.join(invalid_folder, rec["category"])
                try:
                    rec["placed_as"] = place(rec["real_path"], folder, os.path.basename(rec["file"]), mode, used)
                    rec["placed_in"] = os.path.relpath(folder, args.output)
                except OSError as problem:
                    rec["placed_as"], rec["placed_in"] = "", f"not placed: {problem}"
        columns = ["file", "result", "reason", "category", "kind", "note", "name", "email", "phone", "linkedin",
                   "sections", "section_count", "placed_in", "placed_as"]

        def write(path, rows, cols):
            with open(path, "w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle, quoting=csv.QUOTE_ALL)
                writer.writerow(cols)
                for r in rows:
                    writer.writerow([r.get(c, "") for c in cols])

        valid_rows = [records[r] for r in order if records[r]["result"] == "Valid"]
        invalid_rows = [records[r] for r in order if records[r]["result"] == "Invalid"]
        write(os.path.join(args.output, "classification_report.csv"), [records[r] for r in order], columns)
        write(os.path.join(valid_folder, "valid_resumes.csv"), valid_rows,
              ["file", "kind", "name", "email", "phone", "linkedin", "sections", "section_count", "note", "placed_as"])
        write(os.path.join(invalid_folder, "invalid_reasons.csv"), invalid_rows,
              ["file", "reason", "category", "kind", "name", "email", "phone", "placed_in", "placed_as"])
        summary = {
            "input": args.input, "files": len(order), "valid": len(valid_rows), "invalid": len(invalid_rows),
            "invalid_by_reason": dict(collections.Counter(r["category"] for r in invalid_rows).most_common()),
            "valid_by_kind": dict(collections.Counter(r["kind"] for r in valid_rows).most_common()),
            "all_by_kind": dict(collections.Counter(records[r]["kind"] or "?" for r in order).most_common()),
            "placed_by": "move" if mode == "move" else "copy" if mode == "copy" else "hard link (copy if not possible)",
            "seconds": round(time.time() - started),
        }
        with open(os.path.join(args.output, "classification_summary.json"), "w", encoding="utf-8") as handle:
            json.dump(summary, handle, indent=1, ensure_ascii=False)
    finally:
        shutil.rmtree(long_path(temp), ignore_errors=True)

    print()
    print(f"DONE in {summary['seconds']} seconds.  Files looked at: {summary['files']}")
    print(f"  Valid   : {summary['valid']}   -> {valid_folder}")
    print(f"  Invalid : {summary['invalid']}   -> {invalid_folder}  (a sub-folder per reason)")
    for reason, count in summary["invalid_by_reason"].items():
        print(f"      {count:6}  {reason}")
    print(f"  Reasons for every file: {os.path.join(invalid_folder, 'invalid_reasons.csv')}")
    print(f"  Full report           : {os.path.join(args.output, 'classification_report.csv')}")


if __name__ == "__main__":
    main()
