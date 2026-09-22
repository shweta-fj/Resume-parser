# Containerizes the PARSER itself (Python + Tesseract + Poppler +
# LibreOffice + the script) -- so anyone can run parse_resumes_structured.py
# without installing any of that on their own machine.
#
# This is separate from the resumes_meilisearch_docker package, which is
# just a preloaded MeiliSearch DATA snapshot with no parsing code in it.
# This image needs a MeiliSearch server to talk to (run separately, e.g.
# via that other package, or any MeiliSearch instance reachable from here).
#
# Build:
#   docker build -t resumes-parser .
#
# Run (Windows/Mac, MeiliSearch running in another local container):
#   docker run --rm ^
#     -v "%cd%\input\resumes:/app/input/resumes" ^
#     -v "%cd%\hash_store_structured:/app/hash_store_structured" ^
#     -e MEILI_URL=http://host.docker.internal:7700 ^
#     -e MEILI_API_KEY=your-key ^
#     resumes-parser
#
# `host.docker.internal` is Docker's special address for "the machine this
# container is running on" -- use it instead of localhost, since inside a
# container, localhost means the container itself, not your computer.

FROM python:3.11-slim

# Tesseract (OCR), Poppler (PDF -> image for OCR), and LibreOffice (.doc ->
# .docx conversion) -- the three non-Python programs the parser calls out to.
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    poppler-utils \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY parse_resumes_structured.py .
COPY lookups/ lookups/

ENTRYPOINT ["python", "parse_resumes_structured.py"]
