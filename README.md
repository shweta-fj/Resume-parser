# Resume parsing: what each script does, who calls whom, and how to read the hard parts

Three scripts do the work. They never import each other. They hand work over through **folders on disk**. One extra
script (`run_pipeline.py`) starts the three in order. One library file (`resume_engine.py`) is imported by the classifier.

```
Input\  (any files: PDF, DOCX, DOC, RTF, images, CSV/Excel lists, zips ...)
   |
   |  1. main\classify_resumes.py            imports  main\resume_engine.py    word lists: main\lookups\
   v
classification\Valid\                        the real resumes (hard links, no extra disk space)
classification\Invalid\<reason>\             everything else, one folder per reason
classification\classification_report.csv     one row per file     (+ Valid\valid_resumes.csv, Invalid\invalid_reasons.csv)
   |
   |  2. parser\parse_resumes_simple.py      self-contained         word lists: parser\lookups\
   v
Output\Valid\<same name>.json                one JSON per valid file (production format)
Output\Invalid\<same name>.json              one JSON per invalid file, with the reason inside
   |
   |  3. parser\upload_to_meilisearch.py     needs: pip install meilisearch
   v
MeiliSearch index  new_resumes               old documents deleted first, then the new ones uploaded
```

| File | Role |
|---|---|
| `main\classify_resumes.py` | **Script 1.** Decides, for every file, "real resume or not, and why". |
| `main\resume_engine.py` | **Library**, imported by script 1 (never run by hand). Reads files and applies the validity rules. |
| `parser\parse_resumes_simple.py` | **Script 2.** Turns every classified file into a JSON file. |
| `parser\upload_to_meilisearch.py` | **Script 3.** Replaces the documents of the MeiliSearch index with the valid JSONs. |
| `run_pipeline.py` | **Runner.** Starts 1, 2, 3 one after the other with `subprocess` (works from `Resume parsing\`, `main\` or `parser\`). |

## 1. Run it

One command for everything (stops at the first step that fails):

```powershell
cd "C:\Users\ankur\Downloads\Resume parsing"
python run_pipeline.py                       # classify -> parse -> upload
python run_pipeline.py --skip-classify       # classification already done
python run_pipeline.py --skip-upload         # stop after the JSON files
python run_pipeline.py --only upload         # one step only: classify | parse | upload
python run_pipeline.py --limit 200 --skip-upload      # try on a small batch
python run_pipeline.py --clean               # DELETES the JSON in Output, and (if the classify step runs) empties classification\Valid and \Invalid first
```

Or each step by hand:

```powershell
cd "C:\Users\ankur\Downloads\Resume parsing\main"
python classify_resumes.py                   # --clean to run it a second time, --copy / --move / --report-only for other modes

cd "..\parser"
python parse_resumes_simple.py --split-candidates          # --limit 200 to try it, --only valid|invalid

python upload_to_meilisearch.py --dry-run                  # shows what would be uploaded, touches nothing
python upload_to_meilisearch.py --include-candidate-lists  # asks you to type YES, deletes the old documents, uploads
```

Never use `--clean` while a parse or an upload is running on the same folders.

## 2. Who calls whom

```
run_pipeline.py
  |-- subprocess -> main\classify_resumes.py  (cwd = main\)
  |                    |-- import resume_engine as E
  |                    |       E.find_all_files, E.unwrap_uploaded_files, E.detect_mime_type, E.detect_candidate_csv,
  |                    |       E.read_csv_rows, E.convert_old_formats, E.find_libreoffice, E.make_file_id,
  |                    |       E.parse_all_files  ->  parse_one_file -> parse_resume -> check_if_real_resume -> validity_problem
  |                    '-- own functions: place, category_of, classify_table, read_xlsx_rows, convert_spreadsheets_to_csv ...
  |-- subprocess -> parser\parse_resumes_simple.py   (reads classification\ , writes Output\)
  |                    main -> collect_jobs -> convert_old_files -> ProcessPoolExecutor -> run_job -> read_lines / parse_lines / parse_table
  '-- subprocess -> parser\upload_to_meilisearch.py  (reads Output\Valid)
                       main -> resume_document / candidate_documents -> meilisearch client
```

What the steps pass to each other (the only "interface"):

| From -> to | Passed through | What the next step relies on |
|---|---|---|
| 1 -> 2 | `classification\Valid\*` and `classification\Invalid\<category>\*` | `Invalid\invalid_reasons.csv` has the columns `category`, `placed_as`, `reason`; script 2 uses them to write the reason into each invalid JSON. |
| 2 -> 3 | `Output\Valid\*.json` | Each JSON has `status`, `parsed`, `source_file` and the production fields (`name`, `skills`, `employmentHistory` ...). |

## 3. Script 1: `main\classify_resumes.py` (+ `resume_engine.py`)

**Goal:** sort every file in `Input` into Valid or Invalid and write the reason. It never writes resume JSON.

Flow of `main()`:
1. **Find files.** `E.find_all_files` lists everything, opening zip files (up to 3 levels). `E.unwrap_uploaded_files` cuts the real file out of web-upload bodies (`------WebKit...`).
2. **Look at each file** (cheap checks first): empty (0 bytes) -> Invalid; same SHA-256 as an earlier file -> Invalid "duplicate of ..."; broken zip tail -> `repair_zip_tail`; then the type is found with `E.detect_mime_type` (by the file's bytes, not its extension).
   - spreadsheet (xlsx/xls/ods) -> `classify_table`: a list of candidates is Valid (marked "candidate list"), any other spreadsheet is Invalid. `.xls/.ods` are first turned into CSV with LibreOffice.
   - candidate CSV (`E.detect_candidate_csv`) -> Valid; any other CSV -> `classify_table` decides.
   - video / audio / program / unknown binary -> Invalid "not a document" (found by the first bytes, see `VIDEO_AUDIO_SIGNS`).
   - code / web files (`CODE_EXTENSIONS`) -> Invalid.
   - PDF, DOCX, images, text -> queued for reading. DOC/RTF/ODT/HTML -> `E.convert_old_formats` (LibreOffice) -> DOCX -> queued.
3. **Read and judge** the queue: `E.parse_all_files` runs worker processes; each runs `parse_one_file`, which runs `parse_resume` (text, sections, contact details) and then `validity_problem`. If a worker crashes (usually out of memory) the engine retries with half as many workers.
4. **Place the files** (`place`): hard link by default (`--copy` / `--move` / `--report-only`). Invalid files go to `Invalid\<category>`.
5. **Reports:** `classification_report.csv`, `Valid\valid_resumes.csv`, `Invalid\invalid_reasons.csv`, `classification_summary.json`.

The validity rules (`validity_problem` in the engine, plus the checks above). A file is **Invalid** if it is:
empty, a duplicate, unreadable (damaged or password-protected), not a document, code/web, a spreadsheet that is not a candidate list,
a certificate/invoice/statement (2+ different "paper" words in the first 600 characters, or a certificate title in the first 300, and 3 or fewer sections),
a job description (no email/phone/LinkedIn, and advert wording in the first 400 characters or 2 different phrases),
or has fewer than 20 readable characters. Everything else is Valid, **including cover letters**. A name, a section or contact details are *not* required.

`resume_engine.py` is long (about 3,000 lines) because it started as an older full parser. The classifier uses only the parts listed in "Who calls whom"; the many `extract_*` functions (education, projects, titles, skills ...) are there for `parse_resume`, which the classifier needs only to count sections and find contact details.

## 4. Script 2: `parser\parse_resumes_simple.py`

**Goal:** one JSON per input file, with the same name (`Rahul CV.pdf` -> `Rahul CV.json`), in the production (LLM-parsed) format.

Flow of `main()`:
1. `collect_jobs` lists the files of `classification\Valid` and `classification\Invalid\*` and decides a **mode** for each: `parse` (read as a resume), `table` (CSV/Excel candidate list), or `skip` (invalid files that cannot be parsed: JSON with the reason only).
2. `convert_old_files` turns DOC/RTF/ODT/HTML into DOCX (and XLS/ODS into CSV) with LibreOffice, in one batch, before the workers start.
3. A `ProcessPoolExecutor` runs `run_job` per file. `run_job` never raises: a failure becomes a JSON with `parse_error`.
   - mode `parse`: `read_lines` -> list of `Line` -> (`split_bundle` if `--split-candidates`) -> `parse_lines`.
   - mode `table`: `read_table` -> `parse_table` (columns are matched by name, e.g. "mob.no", "e-mail id").
4. The main process writes the JSON (`json_name` keeps the input's name; only characters Windows forbids are replaced; a clash becomes `name (pdf).json`).

`parse_lines(lines, filename)` is the heart. In order:

| Step | Function | Idea |
|---|---|---|
| sections | `split_sections` (uses `heading_section`) | a line that is short, matches a heading word list (+ fuzzy for typos and lost spaces) starts a section: `summary`, `skills`, `experience`, `education`, `projects`, `internships`, `certifications`, `personal`, `other`; everything before the first heading is `top`. |
| contact | `find_email`, `find_phone`, `find_linkedin` | email: regex + `email-validator`. Phone: the `phonenumbers` library, default country India, returned as E.164 (`+919876543210`). Looks in the first 40 lines first. |
| name | `find_name` | see "Names" below. |
| skills | `find_skills` | dictionary n-gram matching (`find_terms`) over `skills.json` with a confidence score; `rank_skills` splits them into advanced / medium / basic. |
| titles, industries | `find_titles`, `find_industries` | dictionary and clue-word matching. |
| education | `find_education` -> `build_education` | degree + institution + year; `education_type` maps to the production enum (`bachelors`, `masters`, `intermediate`, `diploma`, `high_school`, `associates`, `phd`). |
| jobs | `find_entries` (+ `find_undated_entries`) -> `build_job` | see "Employment" below. |
| other | `find_projects`, `find_certifications` -> `build_certifications`, `find_salary` -> `build_ctc`, `find_languages`, `find_notice_period`, `find_social_profiles` | |
| years | `total_experience_years` | "6+ years of experience" if the text says so, otherwise the length of the dated jobs with overlaps counted once (`totalExperienceSource` says which). |
| flags | `build_red_flags` | rule-based version of the production `redFlags` (gap > 1 month, 3+ jobs of a year or less, same title 3 times, degree without institution, glued words / no email). |

Extras in the JSON that the production record does not have: `titles`, `industries`, `internships`, `totalExperienceSource`, `confidence`, `issues`, `sections_found`.

**`--split-candidates`** (on in `run_pipeline.py`): a CSV/Excel list becomes one JSON per row, and a file that holds several resumes (`split_bundle`) becomes one JSON per resume: `<name> - 001.json`, `- 002.json`. The JSON with the input's own name stays as an **index** (`"split": true`, `candidateCount`, `candidateFiles`).

## 5. Script 3: `parser\upload_to_meilisearch.py`

1. Reads every JSON in `Output\Valid`. It skips: invalid status, split **index** files (`"split": true`, their parts are uploaded instead), and candidate-list people unless `--include-candidate-lists`.
2. `resume_document(data)` turns a resume JSON into one flat search document (`skills` as text for searching + `skills_list` for filtering, `total_experience_years`, red-flag booleans ...). `candidate_documents(data)` does the same for a person of a candidate list (`source_type = "candidate_list"`, one document per email / phone so nobody is added twice).
3. Asks you to type **YES**, deletes all documents of the index (the index itself stays), sets filterable / displayed / searchable fields (`FILTERABLE`, `SHOWN`, `SEARCHED`) and uploads in batches of 500.
4. Address and key: `--url` / `--api-key`, or `MEILI_URL` / `MEILI_API_KEY`, or a `.env` file next to the script.

## 6. Hard-to-read names, explained

### Shared by the scripts

| Name | Meaning |
|---|---|
| `long_path(path)` | Adds the `\\?\` prefix on Windows so paths over 260 characters work. It looks odd in output but is the same path. |
| `Line` (parser) | One line of the resume: `.text`, `.size` (font size, 0 if unknown), `.bold`. Name and project detection use size and bold. |
| `clean(text)` | Collapses whitespace and removes bullet and private-use characters that PDFs leave behind. |
| `norm_key(text)` | Comparison key: lower case, only letters, digits, `+`, `#`. `MS-Excel`, `ms excel`, `MSExcel` all become `msexcel`. A text starting with `.` keeps a leading `.` so `.NET` differs from `net`. |
| `E.make_file_id` | md5 of the path inside `Input`: the same file always gets the same id. Not a security hash, only an id. |
| `E.detect_mime_type` | Finds the file type from its first bytes. `CSV_MIME` (`text/x-candidate-csv`) is a made-up type for "a CSV that is a candidate list". |

### Dates and employment (`parse_resumes_simple.py`)

| Name | Meaning |
|---|---|
| `MONTHS` vs `MONTHS_STRICT` | `MONTHS` accepts any word starting with a month ("decided" matches "dec" + "ided"; kept loose for the older patterns). `MONTHS_STRICT` needs a real month name plus a word boundary. |
| `_DATE_PART` | One date: `Jan 2020`, `2018 December`, `03/2021`, `Jan'20`, `2019`. Built as one big regex string, so it is hard to read. |
| `DATE_RANGE_RE` | `start` and `end` groups around a dash / `to` / `till` / `until`; `end` may be `present`, `current`, `till date` ... (`_END_WORDS`). The dash can be a long dash or U+FFFD (a dash lost by the PDF). |
| `SINCE_RE` | `Since Jan 2020`, `Joined in 2019`, `w.e.f 14th February 2011`: one date, still working there. |
| `ONE_DATE_RE` | A line that holds only one date. `prepare_date_lines` joins two such lines (`October 2018` / `January 2020`) into one range. |
| `prepare_date_lines` | Cleans the text before `find_entries`: puts a dash back, drops `[ ]`, merges `Company` / `- X Ltd` label lines into `Company: X Ltd`, merges split dates. |
| `find_entries(lines, undated)` | Every date range starts one entry. `header` = the date line without the date + the lines just above/below (`header_rows`); the title comes from a title word or the titles list, the company from a suffix (`Pvt Ltd`, `Technologies` ...), a `Company:` label, a sentence ("Working as X in Y"), a short capitalised line, or spaCy ORG as the last resort. `begin` / `stop` mark the lines of the job description. |
| `find_undated_entries` | Used only if the experience section exists but has no usable dates: a title line next to a company line. Entries have empty `start` / `end`; they count as jobs but not for the years. |
| entries with only dates | Kept for the total years, not written as jobs. |
| `PROSE_START_RE` | "Does this line start like a sentence?" (`developed`, `worked`, `the`, `a`, `to`, `process` ...). Such a line is never a title, company, project name or certificate. |
| `COMPANY_SUFFIX_RE`, `TITLE_HEAD` | Words that make a text look like a company (`Pvt`, `Ltd`, `Technologies`, `Bank` ...) or a job title (`engineer`, `manager` ...). |

### Names

`find_name` order: (1) `name_by_scoring` (`name_candidates`): looks at the first 25 lines, drops emails, phones, URLs, headings, job titles, companies, places and skills, and scores the rest: earlier is better, biggest font, bold, all-caps or title case, 2-3 words, a `Name:` label (+1.2), a word that also appears in the email or file name (+0.7). A score of 1.0 or more is accepted. (2) spaCy `PERSON`, only if the email or file name agrees. (3) The text name if the email or file name agrees. (4) A name built from the file name or email.
`looks_like_person` is the guard: a normal 2-4 word name made of ordinary words is **never** replaced by the file name. `name_from_file` reads `Rutvik CV - Rutwik Dave.docx` -> "Rutwik Dave" (the part after ` - `) or `Naukri_RashmiJori[11y_0m].pdf` -> "Rashmi Jori" (CamelCase split, junk removed). `better_with_file` upgrades only a **one-word** name when the file name shares a word with it.

### Education

| Name | Meaning |
|---|---|
| `DEGREE_RE` | Degree words. The short ambiguous ones (`BE`, `ME`, `BA`, `MA`, `CA`; `AMBIGUOUS_DEGREE`) count only if what follows looks like a branch or a delimiter (`AMBIGUOUS_OK`), otherwise "ME Thailand and Power Line ..." would be a degree. |
| `real_institution` / `GENERIC_INSTITUTION` | An institution line needs a school word (`INSTITUTION_RE`) and must not be only a table header such as `College/University` or `Board`. |
| `find_education` internals | `degree_text` = the degree words up to a delimiter or a year. `key_at` = position of an institution word inside the degree text ("B.Sc. ACHARYA NAGARJUNA UNIVERSITY"): the institution is cut out of it. `tail` = the text after the degree with dates and brackets removed; `segments` = tail split at `from`, `-`, `|`, `:`, `/`, `]`. Then a `from X in 2018` fallback, an all-caps acronym (RGPV) and, for tables read column by column, a `pool` of unused institution lines handed out in order. |
| `field_of_study` | "B.Tech in Mechanical Engineering" -> "Mechanical Engineering". |

### Skills and certificates

| Name | Meaning |
|---|---|
| `find_terms(text, table, max_words)` | Looks up every 1..4-word window of the text in a dictionary (`skills.json` 40k, `titles.json` 25k). |
| `whole_item_only`, `known_tech`, `is_strong_skill` | A generic word ("load", "output", "cluster") counts as a skill only when it is the whole list item. In running text only multi-word, well-known or tech-looking skills (`AWS`, `C++`, `Node.js`) count. |
| `CERT_WORDS`, `CERT_VENDORS` | Words that make a line a certificate (`certified`, `certification`) or name a vendor/platform (`aws`, `coursera`, `scrum` ...). |
| `CERT_NOISE` | Lines that are never a certificate: marks, table headers, `DOB`, `Gender`, verbs like "Worked ...", `Mob: ...`. |
| wrapped lines in `find_certifications` | A name the PDF broke over two lines ("... Training and" / "Investing Course") is joined before it is judged. |

### Classification and upload

| Name | Meaning |
|---|---|
| `CATEGORIES`, `category_of(reason)` | Maps a reason text to the folder name under `Invalid\`. |
| `repair_zip_tail`, `_WRAPPED_START` | A file saved as a web-upload body or with a damaged zip end is repaired or cut out so it can be read. |
| `BrokenProcessPool` handling | If a worker dies (out of memory) the engine retries the unfinished files with half the workers, and finally records the one crashing file as an error. |
| `FILTERABLE`, `SHOWN`, `SEARCHED` | The MeiliSearch settings: fields you can filter on, fields returned, fields searched. |
| `make_id` | md5 of the source file name (+ `#<candidateNumber>` for split parts) so re-uploading the same file gives the same document id, and split parts do not overwrite each other. |

## 7. Good to know

- Re-running classification needs `--clean` (it refuses to write into non-empty Valid / Invalid). Re-running the parse overwrites JSON files of the same name but does not delete old ones: empty `Output` first for a clean run.
- Everything runs on your machine; nothing is sent anywhere except the upload to the MeiliSearch address you give.
- Accuracy is checked with format checks and by re-reading samples, not against a hand-labelled answer set. The LLM-parsed production JSON is a reference, not the truth.
