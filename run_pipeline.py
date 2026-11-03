"""
run_pipeline.py  --  runs the three steps one after the other.  Each step is the existing script, started exactly as you would start it by hand.

  1. classification   main\\classify_resumes.py      Input\\           ->  classification\\Valid, classification\\Invalid
                      (it imports main\\resume_engine.py itself)
  2. parsing          main\\parse_resumes_simple.py classification\\  ->  Output\\Valid, Output\\Duplicates, Output\\Invalid
  3. uploading        main\\upload_to_meilisearch.py Output\\Valid     ->  MeiliSearch index new_resumes
                      (deletes the old documents of the index first and asks you to type YES, unless --yes)

  python run_pipeline.py                              # all three steps
  python run_pipeline.py --skip-classify              # classification is already done
  python run_pipeline.py --skip-upload                # stop after the JSON files
  python run_pipeline.py --only upload                # one step only: classify | parse | upload
  python run_pipeline.py --clean                      # empty classification\\Valid/Invalid and the JSON in Output first (needed for a second run)
  python run_pipeline.py --limit 200 --skip-upload    # try it on a small batch

The pipeline stops at the first step that fails, so a broken step never feeds the next one.
"""
import os
import sys
import glob
import time
import argparse
import subprocess

ROOT = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(ROOT).lower() in ("main", "parser"):            # works from "Resume parsing\\", "Resume parsing\main\\" or "Resume parsing\parser\\"
    ROOT = os.path.dirname(ROOT)
PY = sys.executable
CLASSIFY = os.path.join(ROOT, "main", "classify_resumes.py")
PARSE = os.path.join(ROOT, "main", "parse_resumes_simple.py")
UPLOAD = os.path.join(ROOT, "main", "upload_to_meilisearch.py")


def run(title, command):
    print("\n" + "=" * 100 + f"\n{title}\n  " + " ".join(f'"{c}"' if " " in c else c for c in command) + "\n" + "=" * 100, flush=True)
    started = time.time()
    code = subprocess.call(command, cwd=os.path.dirname(command[1]))
    print(f"\n-> {title}: {'finished' if code == 0 else 'FAILED (exit code %d)' % code} in {round(time.time() - started)} s", flush=True)
    return code


def main():
    ap = argparse.ArgumentParser(description="classification -> parsing -> MeiliSearch upload")
    ap.add_argument("--only", choices=["classify", "parse", "upload"], help="run just this step")
    ap.add_argument("--skip-classify", action="store_true")
    ap.add_argument("--skip-parse", action="store_true")
    ap.add_argument("--skip-upload", action="store_true")
    ap.add_argument("--clean", action="store_true", help="empty the classification folders and the JSON in Output before running")
    ap.add_argument("--mode", choices=["copy", "move", "report-only"], help="how classification puts the files in Valid / Invalid (default: the script's own default)")
    ap.add_argument("--limit", type=int, default=0, help="parsing: only the first N files of each folder (to try it out)")
    ap.add_argument("--workers", type=int, default=0, help="number of files read at once (classification and parsing)")
    ap.add_argument("--split-candidates", dest="split", action="store_true", default=True, help="parsing: one JSON per candidate for lists and bundles (default on)")
    ap.add_argument("--no-split-candidates", dest="split", action="store_false")
    ap.add_argument("--include-candidate-lists", action="store_true", help="upload: also upload the people of the CSV / Excel lists")
    ap.add_argument("--url", help="MeiliSearch address (default: MEILI_URL or http://localhost:7700)")
    ap.add_argument("--api-key", help="MeiliSearch key (default: MEILI_API_KEY)")
    ap.add_argument("--yes", action="store_true", help="upload: do not ask before deleting the old documents")
    ap.add_argument("--dry-run", action="store_true", help="upload: only show what would be uploaded")
    args = ap.parse_args()

    steps = {"classify": not args.skip_classify, "parse": not args.skip_parse, "upload": not args.skip_upload}
    if args.only:
        steps = {k: k == args.only for k in steps}

    if args.clean and steps["parse"]:
        removed = 0
        for folder in ("Valid", "Duplicates", "Invalid"):
            for path in glob.glob(os.path.join(ROOT, "Output", folder, "*.json")):
                os.remove(path)
                removed += 1
        print(f"--clean: {removed} JSON files removed from Output (Valid, Duplicates and Invalid)")

    if steps["classify"]:
        cmd = [PY, CLASSIFY]
        if args.clean:
            cmd.append("--clean")
        if args.mode:
            cmd.append("--" + args.mode)
        if args.workers:
            cmd += ["--workers", str(args.workers)]
        if run("STEP 1 - classification (Valid / Invalid)", cmd):
            sys.exit(1)

    if steps["parse"]:
        cmd = [PY, PARSE]
        if args.split:
            cmd.append("--split-candidates")
        if args.limit:
            cmd += ["--limit", str(args.limit)]
        if args.workers:
            cmd += ["--workers", str(args.workers)]
        if run("STEP 2 - parsing (JSON per file)", cmd):
            sys.exit(1)

    if steps["upload"]:
        cmd = [PY, UPLOAD]
        if args.include_candidate_lists:
            cmd.append("--include-candidate-lists")
        if args.url:
            cmd += ["--url", args.url]
        if args.api_key:
            cmd += ["--api-key", args.api_key]
        if args.yes:
            cmd.append("--yes")
        if args.dry_run:
            cmd.append("--dry-run")
        if run("STEP 3 - upload to MeiliSearch (index new_resumes)", cmd):
            sys.exit(1)

    print("\nAll selected steps finished.")


if __name__ == "__main__":
    main()
