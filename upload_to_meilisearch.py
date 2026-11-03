"""
upload_to_meilisearch.py  --  put the SLIM search JSONs into a MeiliSearch index (default: new_resumes).

WHAT IT DOES
  1. Reads every JSON file in  ..\\Output\\Valid   (the slim search JSONs, made by parse_resumes_simple.py: one slim JSON per valid resume and per
     person of a CSV / Excel candidate list). What is in those files is exactly what goes into the index.
  2. DELETES ALL DOCUMENTS that are in the index now (the index itself stays).
  3. Sets up the index and uploads the documents in batches of 500.

THE DOCUMENT
  shown and searched (texts):   name, email, linkedin, phone, skills, titles, industries, languages, educations, certifications,
                                current_employers, past_employers, degrees, fields_of_study, institutes
  shown and searched (lists):   projects, internships, experience
  hidden (stored, not shown):   skills_list, titles_list, industries_list, languages_list, educations_list, certifications_list,
                                current_employers_list, past_employers_list, degrees_list, fields_of_study_list, institutes_list
                                (the same sections as lists: use them in filters, e.g.  skills_list = "python"),
                                source_type ("resume_file" / "candidate_list") and source_file (which file it came from)

People of the candidate lists are NOT uploaded unless you add --include-candidate-lists. One document per person: the same
e-mail / phone in several lists gives one document (the first file in alphabetical order wins).
Files in Output\\Invalid are never uploaded.

HOW TO RUN (PowerShell, from this folder)
    python upload_to_meilisearch.py --dry-run                           # shows what would be uploaded; touches nothing
    python upload_to_meilisearch.py --include-candidate-lists           # asks you to type YES, deletes the old data, uploads
    python upload_to_meilisearch.py --include-candidate-lists --yes     # same, without the question
    python upload_to_meilisearch.py --index new_resumes --url http://localhost:7700 --api-key YOUR_KEY
The address and key can also be put in a .env file next to this script (MEILI_URL, MEILI_API_KEY).
"""

import os
import sys
import json
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


DEFAULT_INPUT = os.path.join(ROOT, "Output", "Valid")
BATCH_SIZE = 500

SHOWN = ["id", "name", "email", "linkedin", "phone", "skills", "titles", "industries", "languages", "educations", "certifications",
         "projects", "internships", "experience", "current_employers", "past_employers", "degrees", "fields_of_study", "institutes",
         "other_sections"]                                           # what the search result shows
SEARCHED = [f for f in SHOWN if f != "id"]                           # what the search box looks inside
LIST_FIELDS = ["skills_list", "titles_list", "industries_list", "languages_list", "educations_list", "certifications_list",
               "current_employers_list", "past_employers_list", "degrees_list", "fields_of_study_list", "institutes_list"]
FILTERABLE = LIST_FIELDS + ["source_type"]                           # hidden fields that filters can use


def read_env_file():
    path = os.path.join(HERE, ".env")
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def main():
    read_env_file()
    ap = argparse.ArgumentParser(description="Replace the documents of a MeiliSearch index with the slim search JSONs.")
    ap.add_argument("--input", default=DEFAULT_INPUT, help="folder with the slim JSON files (default: ..\\Output\\Valid)")
    ap.add_argument("--index", default="new_resumes", help="index to refill (default: new_resumes)")
    ap.add_argument("--url", default=os.environ.get("MEILI_URL", "http://localhost:7700"))
    ap.add_argument("--api-key", default=os.environ.get("MEILI_API_KEY"))
    ap.add_argument("--include-candidate-lists", action="store_true", help="also upload the people of the CSV / Excel candidate lists")
    ap.add_argument("--dry-run", action="store_true", help="only read the files and show what would be uploaded")
    ap.add_argument("--yes", action="store_true", help="do not ask before deleting the old documents")
    args = ap.parse_args()

    if not os.path.isdir(args.input):
        sys.exit(f"Folder not found: {args.input}   (run parse_resumes_simple.py first: it writes the slim JSON files)")
    documents, ids, lists_skipped, duplicates, unreadable = [], set(), 0, 0, 0
    for name in sorted(os.listdir(args.input)):
        if not name.lower().endswith(".json"):
            continue
        try:
            doc = json.load(open(os.path.join(args.input, name), encoding="utf-8"))
        except (OSError, ValueError):
            unreadable += 1
            continue
        if not doc.get("id"):
            continue
        if doc.get("source_type") == "candidate_list" and not args.include_candidate_lists:
            lists_skipped += 1
            continue
        if doc["id"] in ids:
            duplicates += 1                                           # the same person again (another list, or the same e-mail)
            continue
        ids.add(doc["id"])
        documents.append(doc)
    people = sum(1 for d in documents if d.get("source_type") == "candidate_list")
    print(f"Read {len(documents)} documents from {args.input}: {len(documents) - people} resumes, {people} candidate-list people"
          f"  (same person again, left out: {duplicates}; list files left out: {lists_skipped}; unreadable JSON: {unreadable})")
    if not documents:
        sys.exit("Nothing to upload.")
    if args.dry_run:
        print("DRY RUN - nothing was sent. First document:")
        print(json.dumps(documents[0], ensure_ascii=False, indent=1)[:2200])
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
    except Exception as problem:
        text = str(problem).lower()
        if "invalid_api_key" in text or "missing_authorization" in text or "403" in text or "401" in text:
            sys.exit("MeiliSearch refused the API key (or none was given). Nothing was changed.\n"
                     "Give the master key with  --api-key YOUR_KEY  or  $env:MEILI_API_KEY = \"YOUR_KEY\"  or a .env file next to this script.")
        before = 0                                                    # the index does not exist yet
    print(f"Index '{args.index}' holds {before} documents now. ALL of them will be deleted and replaced by {len(documents)}.")
    if not args.yes and input("Type YES to continue: ").strip() != "YES":
        sys.exit("Stopped. Nothing was changed.")

    wait = lambda task: client.wait_for_task(task.task_uid, timeout_in_ms=1_800_000)
    if before:
        wait(index.delete_all_documents())
        print("  old documents deleted")
    wait(index.update_filterable_attributes(FILTERABLE))
    wait(index.update_displayed_attributes(SHOWN))                    # the *_list fields, source_type and source_file stay hidden
    wait(index.update_searchable_attributes(SEARCHED))
    for start in range(0, len(documents), BATCH_SIZE):
        wait(index.add_documents(documents[start:start + BATCH_SIZE], primary_key="id"))
        print(f"  uploaded {min(start + BATCH_SIZE, len(documents))} of {len(documents)}")
    print(f"Done. Index '{args.index}' now holds {index.get_stats().number_of_documents} documents.")


if __name__ == "__main__":
    main()
