# parse_resumes_simple.py — what changed from the old script to the new one
The parser is still rule-based. It uses no LLM. The libraries are the same: pymupdf, python-docx, LibreOffice, Tesseract, phonenumbers, rapidfuzz, dateparser, spaCy and email-validator.

The rules the new script follows:

1. A value comes only from its own section of the resume.
2. A blank means "not in the resume".
3. A wrong value is worse than a blank.
4. There is no fixed limit on the number of jobs, projects, skills and so on.
5. Every section is captured.

---

## 1. Code structure (this refactor — behaviour unchanged)

| | Before | After |
|---|---|---|
| Lambdas | 24 | **0**, each is a named function |
| Functions defined inside functions | 24 | **0**, each is a module-level function |
| Functions with a leading underscore | 5 | 0, they have plain names |
| Variables named `L`, `l`, `t`, `m`, `p`, `w`, `e`, `d` ... | everywhere | renamed (`word_lists`, `line`, `text`, `match`, `part`, `word`, `entry`, `doc` ...) |
| Writing results to Valid / Duplicates | two functions inside `main()` using shared variables | a class `ValidWriter` with a `write()` method and a `save_json()` function |

- **Moved-out helpers:** a helper that needed variables from its parent function now takes them as keyword-only parameters. For example `is_title_text(text, *, word_lists)`.
- **Renamed functions:**
  - `_find_name` → `find_name_and_score`
  - `_find_education` → `find_education_with_labels`
  - `_find_education_lines` → `find_education_in_lines`
  - `_dates_of` → `dates_of_text`
  - `_find_entries_dated` → `find_dated_entries`
- **New named functions replacing lambdas:**
  - sort keys: `first_of_pair`, `position_in_text`, `negative_score`, `job_start_date`, `match_start`, `end_then_start`, `entry_end`
  - helpers: `heading_style`, `capitalise_name_word`, `is_school_name`, `full_month_name`, `is_month_word`, `full_year`, `letters_only`, `unique_ignoring_case`, `drop_trailing_id`, `decode_word_escape`, `decode_control_character`, `load_lookup_file`
- **New named functions replacing nested helpers:**
  - `add_docx_paragraph`, `block_starts_with_heading`, `block_position`
  - `name_is_backed`, `add_skill_score`
  - `job_line_head`, `is_title_text`, `is_company_text`, `is_company_like`
  - `usable_header_text`, `tidy_header_text`, `is_sentence_line`
  - `close_labelled_job`, `tidy_project_name`, `keep_certification`
  - `table_cell`, `first_number_in_row`, `ctc_block`, `split_link_parts`
  - `close_other_section`, `close_semester_project`, `job_line_text`, `load_json_or_none`

**How it was checked:** 300 resumes were parsed before the refactor and after it, and every JSON field was compared. All 300 are identical (240 parsed; the other 60 are old-format `.doc` files that the pipeline converts first, and they fail the same way in both runs). A script also found no undefined names and no wrong function calls.

---

## 2. Behaviour changes (new rules, in the order of the resume)

### Reading the file
- Word encoding escapes (`_x0092_`), zero-width characters, ligatures and bullets are cleaned (`clean`, `decode_word_escape`, `decode_control_character`).
- PDFs with a sidebar or two columns are put in reading order (`column_order`). The header is read first, then the left column of all pages, then the right.
- A DOCX is read with its hyperlinks (`read_hyperlinks`, `read_docx_raw`). Labels glued to values are split (`unglue_labels`, `merge_label_cells`, `join_split_headings`). Typos in headings are fixed (`tidy_line_text`, `HEADING_TYPOS`).

### Sections
- Headings are recognised from a vocabulary and from the resume's own heading style (`vocab_heading`, `heading_style`, `reclassify_sections`).
- Section names that the script has no field for are kept as `other_sections` (`named_other_sections`).
- The uploader now shows `other_sections` too.

### Name, contact, location
- The name is read only from the resume text. It is checked against the e-mail and file name, and against `name_words.json`, which is rebuilt from the Valid names after every parse. If the name is unsure the field stays blank (`find_name_and_score`, `genuine_name`, `plausible_name`, `trim_name`, `cut_glued_name`, `unspace_name`).
- Phones, links and location are read with new functions (`find_phones`, `find_links`, `links_from_cell`, `find_location`, `location_from_pin`, `compose_location`).
- The e-mail is validated, and a domain glued to text is cut (`cut_domain_end`).

### Experience and employers
- **Current employer by date.** A job is current when it has a start date and no real end date. The end may be any word (Present, Current, Continuing, Till date, Ongoing) or missing. The company may be above, below or beside the date (`current_date_span`, `OPEN_END_RE`, `apply_current_dates`, `apply_current_statements`, `apply_past_statements`, `current_company_ok`).
- **Several layouts for jobs** (all new): dated entries, undated entries, labelled entries ("Employer:", "Designation:", "Duration:", "Status:"), numbered entries, a career-timeline table read column by column, a growth table, and column-by-column company lists (`find_dated_entries`, `find_labelled_entries`, `find_numbered_entries`, `find_table_entries`, `find_timeline_entries`, `find_growth_entries`, `column_companies`).
- **Employer checks:**
  - `employer_name`, `fallback_employer`, `bad_employer` and `final_employer` reject places, months, skills, sentences, the candidate's own name and a client of a project.
  - **`final_employer`** (the last change) also drops a job title or a label ("Sr. Software Engineer", "Company Profile") and cuts a title glued to a company ("Software Engineer in Wipro" → "Wipro").
  - On the full set it brought wrong-looking employer values from 378 down to about 41, which is about 8 wrong values once real companies starting with "Sri" or "SR" are not counted.
- `employers_of` builds the current and past employer lists, for resumes and for candidate lists.
- **Fixed crash:** a bug in the employment-gap check (`int('')`) sent 52 real resumes to `Invalid`.
- **No limits:** 23 caps on the number of entries and the 200-line limits were removed.

### Education
- Degrees and fields are normalised (`canon_degree`, `canon_field`, `education_facts`).
- Academic-record tables and labelled education are read (`academic_table_rows`, `find_labelled_education`, `find_education_with_labels`).
- A school name that contains "Secondary" is no longer taken as a degree (`is_school_name`).
- "Passed" is no longer taken as a degree. Institutes are checked (`institute_name`).

### Skills, titles, projects, certifications, internships
- Skills: junk is filtered out (`skill_ok`, `SKILL_NOT_A_SKILL`, `listed_tool`, `LISTED_NOT_SKILL`). A tool that is listed in the skills section but unknown to the dictionary is kept.
- Titles are read from experience, headline and labelled text (`titles_from_text`, `title_like`, `sentence_titles`).
- Projects: semester projects, labelled projects ("Project Name:", "Project Title:") and project names without numbers or links (`semester_projects`, `labelled_projects`, `PROJECT_JUNK_RE`).
- Certifications: a stray bullet letter "e" is removed.
- Internships: an internship is recognised from the company (intern-by-company rule).

### Duplicates and output
- Two records are duplicates only when everything in them is the same (`text_fingerprint`, `content_fingerprint`, `repoint_duplicates`).
- `duplicate_of` always names the file that really is in Valid.
- Candidate lists give one JSON per person (`drop_lists_from_valid`, `clean_before_lists`).

---

## 3. Results

- **Agreement with an LLM** on the same 500 resumes: 59.7% at the start of this work, 62.1% now, with 0 crashes.
  - Strong fields: e-mail 97.2%, phone 97.8%, name 94.6%, current employers 71.6%.
  - Weak fields: skills 14.8% (the LLM lists many more), degrees 31.0% (partly naming), past employers 47.4%.
  - The LLM is not ground truth; several times the script was right.
- **Full set** (18,450 Valid records, 6,514 from resumes and 11,890 from candidate lists):
  - name 97.6%, e-mail 95.3%, phone 96.0%
  - experience 90.8%, current employer about 54% (after the employer fix)
  - Invalid holds only real non-resumes: duplicates, spreadsheets that are not candidate lists, job descriptions, certificates, empty or password-protected files.

## 4. Files

Run order (`run_pipeline.py`): `classify_resumes.py` → `parse_resumes_simple.py` → `upload_to_meilisearch.py`.
