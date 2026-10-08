"""
upload_to_meilisearch.py  --  put the VALID parsed resumes into a MeiliSearch index (default: new_resumes).

WHAT IT DOES
  1. Reads every JSON file in  ..\\Output\\Valid   (made by parse_resumes_simple.py, production format).
  2. Turns each one into a search document (flat fields that are easy to search and filter).
  3. DELETES ALL DOCUMENTS that are in the index now (the index itself stays).
  4. Sets up the index (filterable, shown and searched fields) and uploads the new documents in batches of 500.

Files in Output\\Invalid are never uploaded.
A valid CSV / Excel candidate list (one person per row) is NOT uploaded unless you add --include-candidate-lists;
then every person in those lists becomes one document (source_type = "candidate_list", one document per email /
phone, so the same person is not added twice).

HOW TO RUN (PowerShell, from this folder)
    python upload_to_meilisearch.py --dry-run                           # shows what would be uploaded; touches nothing
    python upload_to_meilisearch.py --include-candidate-lists           # asks you to type YES, deletes the old data, uploads
    python upload_to_meilisearch.py --include-candidate-lists --yes     # same, without the question
    python upload_to_meilisearch.py --index new_resumes --url http://localhost:7700 --api-key YOUR_KEY
The address and key can also be put in a .env file next to this script (MEILI_URL, MEILI_API_KEY).
"""

import os
import re
import sys
import json
import hashlib
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DEFAULT_INPUT = os.path.join(ROOT, "Output", "Valid")
BATCH_SIZE = 500
TABLE_EXT = (".csv", ".tsv", ".xlsx", ".xlsm", ".xls", ".ods")      # a candidate list: one person per row
CORE_SECTIONS = ("skills", "titles", "industries", "educations", "certifications", "projects", "internships")

# Fields that can be used in filters (for example:  skills_list = "python"  or  total_experience_years >= 5).
FILTERABLE = ["skills_list", "titles_list", "industries_list", "languages_list", "advanced_skills_list", "medium_skills_list",
              "basic_skills_list", "skills", "titles", "industries", "languages", "name", "email", "phone", "linkedin", "github",
              "educations", "certifications", "educations_list", "certifications_list", "education_types_list", "degrees_list",
              "institutions_list", "companies_list", "current_company", "current_designation", "location", "source_folder",
              "sections_found", "section_count", "has_all_sections", "name_source", "parser_version", "source_type",
              "total_experience_years", "salary_current", "salary_expected", "notice_period_days",
              "employment_gap", "frequent_job_switcher", "incomplete_education", "lack_of_role_progression", "contains_mistakes"]
# Fields returned in search results (the *_list helper fields are stored but not shown).
SHOWN = ["id", "name", "email", "phone", "linkedin", "github", "skills", "titles", "industries", "languages", "educations", "certifications",
         "projects", "internships", "experience", "current_company", "current_designation", "total_experience_years", "salary_current",
         "salary_expected", "notice_period_days", "location", "social_profiles", "source_file", "source_folder", "source_type",
         "sections_found", "section_count", "has_all_sections", "employment_gap", "frequent_job_switcher", "incomplete_education",
         "lack_of_role_progression", "contains_mistakes", "analysed_skills"]
# Fields the search box looks inside.
SEARCHED = ["name", "email", "phone", "linkedin", "skills", "titles", "industries", "languages", "skills_list", "titles_list",
            "industries_list", "educations", "certifications", "projects", "internships", "experience", "companies_list",
            "current_company", "current_designation", "location", "source_file"]


def read_env_file():
    path = os.path.join(HERE, ".env")
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def make_id(text):
    """A MeiliSearch id: letters, digits, - and _ only. The same file always gets the same id."""
    return hashlib.md5(text.encode("utf-8")).hexdigest()


def education_text(e):
    head = " ".join(x for x in (e.get("educationType", "").replace("_", " "), e.get("fieldOfStudy")) if x)
    return ", ".join(x for x in (head, e.get("institutionName"), e.get("endDate")) if x)


def job_text(j):
    where = " at ".join(x for x in (j.get("designation"), j.get("companyName")) if x)
    dates = f"{(j.get('startDate') or '')[:7]} to {(j.get('endDate') or '')[:7]}".strip()
    return f"{where} ({dates})" if where else dates


def profile_url(profiles, platform):
    return next((p.get("url") for p in profiles if p.get("platform") == platform), None)


def resume_document(data):
    """The parsed JSON of one resume (production format) -> one MeiliSearch document."""
    skills = data.get("skills") or []
    jobs = data.get("employmentHistory") or []
    internships = data.get("internships") or []
    education = data.get("educationalBackground") or []
    certifications = data.get("certifications") or []
    projects = [p.get("projectName") for p in (data.get("personalProjects") or []) if p.get("projectName")]
    profiles = data.get("socialProfiles") or []
    current = data.get("currentCompanyDetails") or {}
    titles = data.get("titles") or list(dict.fromkeys(j["designation"] for j in jobs if j.get("designation")))
    industries = data.get("industries") or []
    languages = data.get("languagesKnown") or []
    experience = data.get("totalExperience") or {}
    ctc = data.get("ctcDetails") or {}
    flags = data.get("redFlags") or {}
    levels = data.get("analysedSkills") or {}
    cert_names = [c.get("name") for c in certifications if c.get("name")]
    doc = {
        "id": make_id((data.get("source_file") or data.get("name") or json.dumps(data)[:200]) + (f"#{data['candidateNumber']}" if data.get("candidateNumber") else "")),
        "name": data.get("name"), "email": data.get("email"), "phone": data.get("phoneNumber"),
        "linkedin": profile_url(profiles, "linkedin"), "github": profile_url(profiles, "github"), "social_profiles": profiles,
        "skills": ", ".join(skills), "skills_list": skills,
        "advanced_skills_list": levels.get("advanced") or [], "medium_skills_list": levels.get("medium") or [],
        "basic_skills_list": levels.get("basic") or [], "analysed_skills": levels,
        "titles": ", ".join(titles), "titles_list": titles,
        "industries": ", ".join(industries), "industries_list": industries,
        "languages": ", ".join(languages), "languages_list": languages,
        "educations": "; ".join(education_text(e) for e in education), "educations_list": [education_text(e) for e in education],
        "education_types_list": list(dict.fromkeys(e["educationType"] for e in education if e.get("educationType"))),
        "degrees_list": [e["fieldOfStudy"] for e in education if e.get("fieldOfStudy")],
        "institutions_list": [e["institutionName"] for e in education if e.get("institutionName")],
        "certifications": "; ".join(cert_names), "certifications_list": cert_names,
        "projects": projects, "internships": [job_text(j) for j in internships],
        "experience": [job_text(j) for j in jobs],
        "companies_list": [j["companyName"] for j in jobs if j.get("companyName")],
        "current_company": current.get("companyName") or None, "current_designation": current.get("designation") or None,
        "total_experience_years": round((experience.get("years") or 0) + (experience.get("months") or 0) / 12, 1),
        "salary_current": (ctc.get("currentCTC") or {}).get("min") or None,
        "salary_expected": (ctc.get("expectedCTC") or {}).get("min") or None,
        "notice_period_days": data.get("currentNoticePeriod"),
        "employment_gap": flags.get("employmentGap"), "frequent_job_switcher": flags.get("frequentJobSwitcher"),
        "incomplete_education": flags.get("incompleteEducation"), "lack_of_role_progression": flags.get("lackOfRoleProgression"),
        "contains_mistakes": flags.get("containsMistakes"),
        "source_file": data.get("source_file"), "source_folder": "Valid", "source_type": "resume_file",
        "parser_version": "simple-2", "name_source": None,
    }
    present = {"skills": skills, "titles": titles, "industries": industries, "educations": education,
               "certifications": certifications, "projects": projects, "internships": internships}
    doc["sections_found"] = [k for k in CORE_SECTIONS if present[k]]
    doc["section_count"] = len(doc["sections_found"])
    doc["has_all_sections"] = doc["section_count"] == len(CORE_SECTIONS)
    return doc


def candidate_documents(data):
    """A candidate list (CSV / Excel): one document per person."""
    for person in data.get("candidates") or []:
        skills = person.get("skills") or []
        key = (person.get("email") or "").lower() or re.sub(r"\D", "", person.get("phoneNumber") or "")[-10:] or \
            (profile_url(person.get("socialProfiles") or [], "linkedin") or "")
        if not key:
            continue
        company = person.get("currentCompanyDetails") or {}
        title = company.get("designation") or ""
        doc = {
            "id": make_id("candidate:" + key),
            "name": person.get("name"), "email": person.get("email"), "phone": person.get("phoneNumber"),
            "linkedin": profile_url(person.get("socialProfiles") or [], "linkedin"),
            "skills": ", ".join(skills), "skills_list": skills,
            "titles": title, "titles_list": [title] if title else [],
            "industries": "", "industries_list": [], "languages": "", "languages_list": [],
            "educations": person.get("educationText") or "", "educations_list": [person["educationText"]] if person.get("educationText") else [],
            "certifications": "", "certifications_list": [], "projects": [], "internships": [],
            "experience": [person["totalExperienceText"]] if person.get("totalExperienceText") else [],
            "companies_list": [company["companyName"]] if company.get("companyName") else [],
            "current_company": company.get("companyName") or None, "current_designation": title or None,
            "location": person.get("location"),
            "source_file": data.get("source_file"), "source_folder": "Valid", "source_type": "candidate_list",
            "parser_version": "simple-2", "name_source": "csv",
        }
        present = [k for k, v in (("skills", skills), ("titles", doc["titles_list"]), ("educations", doc["educations_list"])) if v]
        doc["sections_found"], doc["section_count"], doc["has_all_sections"] = present, len(present), False
        yield doc


def main():
    read_env_file()
    ap = argparse.ArgumentParser(description="Replace the documents of a MeiliSearch index with the valid parsed resumes.")
    ap.add_argument("--input", default=DEFAULT_INPUT, help="folder with the valid parsed JSON files (default: ..\\Output\\Valid)")
    ap.add_argument("--index", default="new_resumes", help="index to refill (default: new_resumes)")
    ap.add_argument("--url", default=os.environ.get("MEILI_URL", "http://localhost:7700"))
    ap.add_argument("--api-key", default=os.environ.get("MEILI_API_KEY"))
    ap.add_argument("--include-candidate-lists", action="store_true", help="also upload the people of the CSV / Excel candidate lists")
    ap.add_argument("--dry-run", action="store_true", help="only read the files and show what would be uploaded")
    ap.add_argument("--yes", action="store_true", help="do not ask before deleting the old documents")
    args = ap.parse_args()

    if not os.path.isdir(args.input):
        sys.exit(f"Folder not found: {args.input}")
    documents, ids, lists_skipped, unreadable = [], set(), 0, 0
    for name in sorted(os.listdir(args.input)):
        if not name.lower().endswith(".json"):
            continue
        try:
            data = json.load(open(os.path.join(args.input, name), encoding="utf-8"))
        except (OSError, ValueError):
            unreadable += 1
            continue
        if data.get("status") not in (None, "valid"):
            continue                                              # never upload anything that is not a valid resume
        if data.get("split"):
            continue                                              # the index of a split file: its parts ("name - 001.json") are uploaded instead
        if data.get("candidateNumber") and str(data.get("source_file", "")).lower().endswith(TABLE_EXT):
            if not args.include_candidate_lists:                  # one person of a CSV / Excel list
                lists_skipped += 1
                continue
            new_docs = list(candidate_documents({"candidates": [data], "source_file": data.get("source_file")}))
        elif data.get("type") == "candidate_list":
            if not args.include_candidate_lists:
                lists_skipped += 1
                continue
            new_docs = list(candidate_documents(data))
        elif data.get("parsed", True):
            new_docs = [resume_document(data)]
        else:
            continue
        for doc in new_docs:
            if doc["id"] not in ids:
                ids.add(doc["id"])
                documents.append(doc)
    by_type = {t: sum(1 for d in documents if d["source_type"] == t) for t in ("resume_file", "candidate_list")}
    print(f"Read {len(documents)} documents from {args.input}: {by_type['resume_file']} resumes, {by_type['candidate_list']} candidate-list people"
          f"  (candidate-list files left out: {lists_skipped}; unreadable JSON: {unreadable})")
    if not documents:
        sys.exit("Nothing to upload.")
    if args.dry_run:
        print("DRY RUN - nothing was sent. First document:")
        print(json.dumps(documents[0], ensure_ascii=False, indent=1)[:1800])
        return

    import meilisearch
    client = meilisearch.Client(args.url, args.api_key)
    try:
        client.health()
    except Exception as problem:
        sys.exit(f"Could not reach MeiliSearch at {args.url} ({problem}). Is it running?  (docker start meilisearch)")
    index = client.index(args.index)
    try:
        before = index.get_stats().number_of_documents
    except Exception:
        before = 0
    print(f"Index '{args.index}' holds {before} documents now. ALL of them will be deleted and replaced by {len(documents)}.")
    if not args.yes and input("Type YES to continue: ").strip() != "YES":
        sys.exit("Stopped. Nothing was changed.")

    wait = lambda task: client.wait_for_task(task.task_uid, timeout_in_ms=1_800_000)
    if before:
        wait(index.delete_all_documents())
        print("  old documents deleted")
    wait(index.update_filterable_attributes(FILTERABLE))
    wait(index.update_displayed_attributes(SHOWN))
    wait(index.update_searchable_attributes(SEARCHED))
    for start in range(0, len(documents), BATCH_SIZE):
        wait(index.add_documents(documents[start:start + BATCH_SIZE], primary_key="id"))
        print(f"  uploaded {min(start + BATCH_SIZE, len(documents))} of {len(documents)}")
    print(f"Done. Index '{args.index}' now holds {index.get_stats().number_of_documents} documents.")


if __name__ == "__main__":
    main()
