# Resume Parser

This tool reads a folder full of resumes and pulls out useful information
from each one — name, email, phone number, LinkedIn link, skills, job
titles, industries, languages, education, certifications, projects, and
internships. Everything is written straight into a
[MeiliSearch](https://www.meilisearch.com/) index — nothing is saved as
local JSON files. Every input file gets a matching document there — a real
result, a duplicate note, an unsupported-type note, or an error — so the
index alone always tells the full story for every file you pointed the
tool at, and it's immediately searchable with no separate loading step.

It works on PDF files, Word files (both `.docx` and the older `.doc`
format), and image files.

**One script, self-contained.** `parse_resumes_structured.py` is the only
script in this project — no other file it depends on. It's meant to keep
being the one place you run new real resumes through as you get more of
them.

## Quick start

```bash
# 1. Install the Python packages
pip install -r requirements.txt

# 2. Start MeiliSearch (see "Making the results searchable with
#    MeiliSearch" below if you don't have it running yet)
docker run -d --name meilisearch -p 7700:7700 -v meili_data:/meili_data getmeili/meilisearch:latest

# 3. Put your resumes in input/resumes/ (the default folder it reads from —
#    or point it at any other folder as the first argument, see below)

# 4. Make sure skills.json, titles.json, industries.json, languages.json,
#    and normalized_category_mapping.json are in lookups/

# 5. Run it
python parse_resumes_structured.py
```

That's the whole thing for a basic run — results land in the
`resumes_structured` MeiliSearch index, one document per resume.

Needs **Python 3.9 or newer**. No specific operating system requirement —
developed and tested on Windows, but nothing in it is Windows-specific
except two small install-path fallbacks for Tesseract/LibreOffice (see
Setup below), which simply won't trigger on Mac/Linux since those systems
put programs on `PATH` more consistently.

## How the project folder is organized

```
resumes/                            (project root)
├── parse_resumes_structured.py     the one script
├── requirements.txt
├── Dockerfile                      runs the parser in a container (see its comments)
├── .env                            MeiliSearch address and key (see "How to run it")
├── input/
│   └── resumes/                    resumes to read go here (not committed)
├── lookups/                        the reference lists
│   ├── skills.json
│   ├── titles.json
│   ├── industries.json
│   ├── languages.json
│   ├── normalized_category_mapping.json
│   ├── skills_extra.json           extra skills, short forms, and blocked words
│   └── industry_keywords.json      sector name -> clue words and employers
└── hash_store_structured/          the tool's own bookkeeping file, seen_hashes.json
```

Every result — real, duplicate, skipped, or error — is written as a
document into MeiliSearch, not into this folder. `hash_store_structured/`
is the only thing the tool keeps on local disk, since it's just a small
bookkeeping file, not resume content. None of the paths are hardcoded —
every one can be pointed somewhere else with a command-line flag, covered
below in "How to run it".

## What file types work

The tool figures out what a file actually is by peeking at its raw content
(the first few bytes — sometimes called a file's "magic bytes"), not by
trusting its filename extension. A file's extension is just a naming
convention — nothing stops someone renaming an HTML page to `Resume.doc`.
Checking the content directly means a mislabeled file gets caught
immediately as unsupported, instead of being attempted and failing
partway through.

| File type (by content, not extension) | How it's read |
|---|---|
| PDF | Text is pulled out directly. If the PDF is just a scanned picture with no real text in it, the tool automatically falls back to OCR for that one file. |
| DOCX | Text is pulled out directly. |
| DOC (the old Word format) | Converted to `.docx` first using LibreOffice (running invisibly in the background), then read the same way as any other `.docx`. Needs LibreOffice installed — see Setup below. |
| Pictures (JPEG, PNG, BMP, TIFF — including `.jfif`) | Read using OCR, since a picture has no real text in it, only pixels. Section headings are less accurate for these, since there's no font information to work with. |

Everything else is skipped on purpose: `.csv`, `.xlsx`, `.htm`/`.html`,
`.txt`, videos, programs, and so on — and anything whose *content* doesn't
match one of the types above, regardless of what its extension claims.
Each one still gets a document in MeiliSearch explaining why it was
skipped, so nothing silently disappears.

If the exact same file shows up twice under different names, the second
copy isn't processed — it gets a short document saying `"The file is a
duplicate of the <original filename> file."` The tool compares file
*content*, not names, so it won't catch a resume that was re-saved or
slightly edited, only an exact copy.

## What a result document looks like

```json
{
  "id": "d41d8cd98f00b204e9800998ecf8427e",
  "name": "Jane Doe",
  "email": "jane@example.com",
  "linkedin": "linkedin.com/in/janedoe",
  "phone": "9876543210",
  "skills": "Python, AutoCAD, SolidWorks",
  "titles": "Software Engineer, Design Engineer",
  "industries": "Manufacturing",
  "languages": "English, Hindi, Kannada",
  "skills_list": ["Python", "AutoCAD", "SolidWorks"],
  "titles_list": ["Software Engineer", "Design Engineer"],
  "industries_list": ["Manufacturing"],
  "languages_list": ["English", "Hindi", "Kannada"],
  "educations": "B.Tech Mechanical Engineering, XYZ College of Engineering, Pune; PUC, ABC PU College",
  "educations_list": ["B.Tech Mechanical Engineering, XYZ College of Engineering, Pune", "PUC, ABC PU College"],
  "certifications": "AWS Certified Cloud Practitioner",
  "certifications_list": ["AWS Certified Cloud Practitioner"],
  "projects": ["Design and Fabrication of Electric Micro Car", "..."],
  "internships": ["Summer Intern, ABC Pvt Ltd", "..."],
  "unmatched": "some other skill, another title, ..."
}
```

Each list-like field comes twice: a **plain string** that's easy to read in
MeiliSearch's document browser, and a real **array** (`*_list`) with one
item per element, which is what filters should use. `skills_list = "Java"`
matches that one skill exactly, where a filter on the comma-separated
string could only match the whole list. Education and certification entries
contain commas themselves ("B.Tech, XYZ College, Pune"), so their strings
separate entries with `"; "`. `projects`/`internships` are arrays only. `id`
is a hash MeiliSearch requires as the primary key; there's no `source_file`
field, by design.

Where each field comes from:
- **Sections.** Every heading is sorted into a category (education,
  projects, experience, skills, ...). A known wording ("Educational
  Profile", "Academic Credentials", "Work Experience (Total 15 Years)")
  counts even in plain body text. A looser keyword match also needs a style
  cue (bold, bigger, ALL CAPS, trailing colon), so a skills-list line like
  "Project Management" is never mistaken for a heading.
- **Skills.** Everything in a skills section counts: known terms, plus list
  items that merely look like skills ("Tally ERP 9", "Desktop Support").
  So do "Tech Stack:"/"Environment:" lines in jobs and projects. In job and
  project prose, only known skills written the way the list writes them
  count ("Java", not "software"). `lookups/skills_extra.json` adds common
  skills `skills.json` lacks: it stores many under long names such as
  "Python (Programming Language)", and their short forms are matched too.
  It also blocks generic words ("Birth", "Status", "Email") and lists
  ambiguous ones ("Excel", "Go", "C") that only count as whole list items.
- **Titles.** Read from the headline under the name, the summary, and
  job/internship entries: "Designation:" labels, "worked as a ..." phrases
  and short title lines. A title must end in an occupation word
  ("Engineer", "Manager", "Head", "Trainee", ...). That drops the generic
  words `titles.json` also holds ("Management", "Academic", "High School")
  and keeps real titles it doesn't know ("Procurement Head").
- **Industries.** Clues in `lookups/industry_keywords.json`: a sector word
  ("thermal power plant", "banking domain") or a well-known employer
  ("Infosys", "Tata Motors"), mapped to short sector names. Explicit
  "Industry:"/"Domain:" lines count too, including the longer
  `industries.json` names.

`unmatched` is one combined, deduplicated, length-capped string of
phrases the tool found but couldn't match to anything in your reference
lists — useful for spotting real skills/titles/industries missing from
those lists. It's built to exclude anything already counted elsewhere
(company names, education content, certification mentions, personal
details like marital status/religion/gender), so it stays a list of
genuinely uncategorized phrases, not noise.

`educations` has one entry per qualification: the degree/12th/10th line
plus the institution and place lines around it, which table layouts put on
separate lines. Grades (CGPA/SGPA/%) and years are stripped.
`certifications` keeps only certificate-title-looking lines. It drops bare
issuer-and-year lines (`"Coursera - 2023"`), participation and appreciation
notes, and validity/ID lines. Without a Certifications section, it takes
lines elsewhere that name a certification outright ("Six Sigma Green Belt",
"CCNA"). `projects` holds project titles only: explicit "Project Title:"
labels anywhere, otherwise title-looking lines in the Projects section, not
their descriptions.

Duplicate, unsupported-type and error files get a short note document (an
`id`, a `message`, and `duplicate_of` for duplicates). These notes go to a
**separate index**, `<index>_notes` by default (`--meili-notes-index` to
change it), so the resume index holds only real candidates. Re-runs check
both indexes, so nothing is processed twice.

## Setting it up

### 1. Install the Python packages

```bash
pip install -r requirements.txt
```

### 2. Install three extra programs (only needed for OCR and old `.doc` files)

These aren't Python packages — they're separate programs the tool calls
out to. If you're only working with normal PDFs and `.docx` Word files,
you can skip this step.

| Program | What it's for | How to install on Windows |
|---|---|---|
| **Tesseract OCR** | Reads text out of pictures, and out of scanned PDFs | `winget install --id UB-Mannheim.TesseractOCR -e` |
| **Poppler** | Turns a PDF page into a picture so Tesseract can read it | `winget install --id oschwartz10612.Poppler -e` |
| **LibreOffice** | Converts old `.doc` files to `.docx` so they can be read | `winget install --id TheDocumentFoundation.LibreOffice -e` |

On a Mac: `brew install tesseract poppler libreoffice`.
On Linux: `apt install tesseract-ocr poppler-utils libreoffice` (or
whatever your Linux version uses).

**Important on Windows:** if you install any of these while a terminal
window is already open, that terminal won't see them until you close it
and open a new one.

### 3. Keep the reference lists nearby

`skills.json`, `titles.json`, `industries.json`, `languages.json`, and
`normalized_category_mapping.json` should sit in the `lookups/` folder
next to the script. If one is missing, the tool still runs — it just
won't be able to match anything from that list — and it prints a warning.

- `skills.json`/`titles.json`/`industries.json` are flat lists of terms.
- `languages.json` maps a canonical language name to its matchable
  spellings/scripts: `{"English": ["english", "english language"], ...}`.
- `normalized_category_mapping.json` maps skill-related *categories*
  (technical, soft, management, etc.) to the many different heading
  phrasings a resume might use for a skills section (`"Areas of
  Expertise"`, `"Skill Profile"`, `"Key Competencies"`, ...) — this is
  what lets the tool recognize far more skill-section headings than the
  handful of common ones it knows by default.
- `skills_extra.json` and `industry_keywords.json` are covered under
  "Where each field comes from" above.

### 4. Start MeiliSearch

See "Making the results searchable with MeiliSearch" below.

## How to run it

```bash
python parse_resumes_structured.py
```

With no arguments, it reads from `input/resumes/` and writes into the
`resumes_structured` index at `http://localhost:7700`. You can point it
anywhere instead:

```bash
python parse_resumes_structured.py path/to/your/resumes --meili-index resumes_structured --lookups-dir path/to/lookups
```

| Flag | What it changes | Default |
|---|---|---|
| (first argument) | which folder of resumes to read | `input/resumes` |
| `--workers` | how many resumes to process at once | number of CPU cores on your machine |
| `--lookups-dir` | where the reference lists live | `lookups` |
| `--hash-dir` | where the content-hash checkpoint file is stored | `hash_store_structured` |
| `--meili-url` | the MeiliSearch server's address | `http://localhost:7700` |
| `--meili-index` | the index name results are written into | `resumes_structured` |
| `--meili-api-key` | your MeiliSearch master key, if you set one | read from `.env`'s `MEILI_API_KEY` |
| `--meili-notes-index` | the index that gets duplicate/unsupported/error notes | `<meili-index>_notes` |

The server address and key come from a `.env` file next to the script:

```
MEILI_URL=http://localhost:7700
# MEILI_URL=https://<your-tunnel>.ngrok-free.dev
MEILI_API_KEY=your-master-key
```

Leave exactly one `MEILI_URL` line uncommented: local Docker, or the ngrok
tunnel to the shared server. `MEILI_API_KEY` must be the key of whichever
server you pick. A real environment variable beats `.env`
(`MEILI_URL=... python parse_resumes_structured.py` for a one-off run), and
a command-line flag beats both. With no `MEILI_URL` anywhere, it uses
`http://localhost:7700`.

## How it stays safe to re-run

- **Checkpointed against MeiliSearch itself.** Before processing anything,
  the tool fetches every document ID already in the index in one pass. A
  file whose document is already there is skipped without even being
  opened — so if the tool crashes or you stop it partway through, just run
  the same command again and it picks up where it left off.
- **Content-hash duplicate detection**, saved to
  `hash_store_structured/seen_hashes.json` between runs, so re-hashing an
  already-checked file's content is never repeated.
- **A file that can't be opened, or a `.doc` that fails to convert**, gets
  a document with a clear error message instead of stopping the batch.
- **If a worker process crashes outright**, that step stops early instead
  of the whole program crashing — whatever wasn't finished yet just has no
  document yet, so re-running picks it up.
- **Each document's ID is computed from the whole original filename,
  extension included** — two files that share a name but differ only in
  extension (`Resume.pdf` vs `Resume.docx`) get different IDs, so one
  can't silently overwrite the other's result.

## How names are found (and why it's more involved than it sounds)

The name is usually the largest/boldest text on the resume — but resumes
are messy, so a long chain of checks rules out things that aren't
actually the name before accepting one:

- Section headings and form labels (`"Name of the College"`, `"Name:
  John Doe"` — the label is stripped, keeping just `"John Doe"`)
- Anything with a colon anywhere (a heading or a `"Label: value"` line)
- A sentence (too many words to plausibly be a name)
- A phone number or email
- Any resume-jargon word anywhere in the line (`"Skills"`, `"Contact"`,
  `"Address"`, ...)
- A line that's mostly a matched job title (`"Purchase & Sales
  Engineer"`) or a company name (`"Mohan Exim India Pvt Ltd"`)
- A place name (`"Karnataka"`) or a date/duration (`"December 2021 to
  Present"`)

Letter-spaced text (`"J O H N   D O E"`) is rejoined where possible, and a
trailing credentials segment (`"- MBA IT | PMP"`) is cut off. If nothing
on the page passes all of this, the tool falls back to a name derived from
the email address as a last resort.

## Data-quality filters

A few categories of noise get filtered out wherever they'd otherwise leak
into a field that isn't the right home for them:
- Numeric-only entries, bare domain fragments (`.com`, `www`), and
  boilerplate words (`"Declaration"`, `"Summary"`) never count as a
  skill/title/industry, matched or unmatched — a few had crept into the
  reference lists themselves and have been removed.
- Company names, personal-details content (marital status, religion,
  nationality, gender), and dates/places are excluded from candidate
  generation and from `unmatched`.
- A language mention (`"English"`, `"Hindi"`) is only ever counted once,
  in `languages` — it's excluded from `skills`/`titles`/`industries` and
  their unmatched leftovers.

This filtering isn't exhaustive — a badly laid-out resume (a scrambled
PDF table, for instance) can still produce some odd output — but it
removes the great majority of the noise these reference lists and
heuristics would otherwise let through.

## Things it can't do yet

- Old-style `.doc` files need LibreOffice installed to convert them.
  Without it, `.doc` files fail one at a time with a clear error.
- OCR needs Tesseract and Poppler installed. Without them, scanned/image
  files fail one at a time with a clear error.
- A resume that had to go through OCR loses bold/font-size information,
  which the tool normally uses to spot section headings and the name —
  OCR resumes are read a bit less accurately than normal ones.
- A resume whose source PDF extracted text in a scrambled order (common
  with table-based layouts) can still produce jumbled `educations`/
  `unmatched` content — reordering already-identified pieces can't fix
  text that was extracted out of order to begin with.
- If a file's name suggests it's not really a resume, the tool has no way
  to tell from content alone, so it still tries to parse it as one.

## Making the results searchable with MeiliSearch

The tool writes results straight into MeiliSearch as it parses — no local
files, no separate loading step.

### Running MeiliSearch

```bash
docker run -d --name meilisearch -p 7700:7700 -v meili_data:/meili_data -e MEILI_MASTER_KEY="pick-a-long-random-secret" getmeili/meilisearch:latest
```

The `-v meili_data:/meili_data` part stores its data in a named Docker
volume, so your indexed resumes survive even if you remove and recreate
the container. Setting `MEILI_MASTER_KEY` is worth doing even for local
use — without it, anyone who can reach that port can read and write
everything in it.

### Searching

MeiliSearch talks over a simple web API. If you've set a master key, pass
it as a bearer token:

```bash
curl -X POST 'http://localhost:7700/indexes/resumes_structured/search' \
  -H 'Content-Type: application/json' \
  -H 'Authorization: Bearer YOUR_MASTER_KEY' \
  -d '{"q": "AutoCAD"}'
```

**Filtering:** filter on the array fields (`skills_list`, `titles_list`,
`industries_list`, `languages_list`, `educations_list`,
`certifications_list`), e.g. `"filter": "skills_list = \"AutoCAD\""`. An
array filter matches one element exactly, so it's a precise "resumes with
skill X" lookup. The tool makes these filterable on every run. It only
*adds* to the index's filterable attributes and never replaces them, so
settings added by hand survive re-runs. Free-text search (the `q`
parameter above) still finds a term anywhere in a resume.

MeiliSearch also ships a small built-in web UI for poking around — open
`http://localhost:7700` in a browser. Note that this built-in UI doesn't
request search-result highlighting, and its generic data-viewer treats
comma-separated strings as plain text, not a structured list — for a
properly highlighted, structured search experience you'd want a small
custom page calling the API directly with `attributesToHighlight`.
