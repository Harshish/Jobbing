# Jobbing

A local, reactive dashboard for a more intentional job search: **discover → match → decide**.
Built with Streamlit. The CLI and web interface share the same scraping and matching functions.

## Start the web interface

Python 3.10+ required. Run these commands from the project root:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# First setup only; do not overwrite an existing .env:
cp -n .env.example .env
# Set OPENAI_API_KEY in .env.
streamlit run app.py
```

Open **http://127.0.0.1:8501**. The server binds to localhost; no account or database is needed.
Do not expose this app publicly: it has no authentication and saved files are shared across sessions.

### 1. Discover

Enter keywords, location, result limit, and posting age, then click **Find jobs**.
The app scrapes LinkedIn with full descriptions and shows the resulting jobs in a preview.
You can instead select an existing `jobs*.csv` from `data/` or a previous web run.
A new search clears the active shortlist selection; historical results remain available.
Scraping may be slow, rate-limited, or return fewer jobs than requested. Follow the source site's terms.

### 2. Match

Use `data/resume.docx` or upload a different DOCX in the sidebar. Set the model,
minimum resume-fit score, and jobs per batch, then click **Match & rank jobs**.
A progress bar updates after every completed LLM batch.

Running sends the full resume text (including contact details) and job data to the
configured provider and incurs API charges. Remove unwanted personal details from
a copy of the resume first. Uploaded resumes are saved locally in the run directory.
The API key stays server-side and is never displayed.

### 3. Decide

Browse the shortlist as ranked role cards or an interactive comparison table:

- Company, title, location, disclosed salary, fit score, and preference tiers.
- Expandable ranking evidence, qualification gaps, and full descriptions.
- Search by role/company/location; filter by fit score or remote-only roles.
- Open the original posting or download the full ranked CSV.
- Switch between saved shortlists without making new API calls.

Display filters do not rescore jobs or change the saved CSV. Downloads contain the
full shortlist. Network operations run only when you explicitly click a scrape,
match, or document-generation button.

### 4. Create a tailored application

Each role card has a **Create resume** button. Enable **Include a motivation letter**
above the cards to generate both documents in one request. Use the original resume
selected in the sidebar. This works
with previously saved shortlists; you do not need to scrape or match again.

The generator uses the original DOCX text plus that job's description, not a prior
generated resume or the matching LLM's opinions. It asks the LLM to preserve facts,
use relevant supported keywords naturally, and flag requirements not evidenced in
your resume. It exports a text-based **PDF** plus an editable DOCX, with real
paragraphs and bullets rather than images or hidden keywords.

The resume PDF uses a clean, recruiter-friendly single-column design:

- A prominent name, distinct professional subtitle, and compact contact details.
- Embedded sans-serif fonts, 10.5-point body text, and comfortable line spacing.
- Dark green section headings, subtle dividers, and consistent 0.7-inch side margins.
- Bold experience headers, hanging bullets, and spacing between roles.
- Headings kept with the following content, widow/orphan control, and discreet page numbers.

Name/contact text and clickable hyperlinks are preserved; links omitted by the LLM
are collected in a final Links section instead of being lost. The editable DOCX keeps
the original template styling. Both resume exports normalize long dashes to ordinary
hyphens, curly quotes to straight quotes, and ellipses to three periods. This applies
to visible text only; hyperlink targets, accents, and factual AI/LLM skills stay intact.
The writing prompt also avoids generic filler and authorship notes. The original resume
and saved versions are never modified.

PDF export uses ReportLab locally; Microsoft Word and LibreOffice are not required.
It embeds installed Calibri fonts, falls back to Arial, then bundled Vera, without
network font downloads. PDF pagination can differ from the editable DOCX; content is
not shrunk or dropped to force a page count. Motivation letters retain their own simpler
paragraph layout. This is not a pixel-exact Word layout engine: floating images, text
boxes, columns, and Word headers/footers are not reproduced.

**ATS-friendly is not a measured or guaranteed ATS score.** These are drafts. Review
all claims, dates, contact details, and wording before submitting; prompting cannot
eliminate LLM hallucinations. No application is submitted automatically.

Every explicit generation makes a paid API request and creates a new version:

```text
data/applications/<company-role-job-fingerprint>/<timestamp-id>/
  resume.pdf                # Primary export, selectable text and clickable links
  resume.docx               # Editable copy
  motivation_letter.pdf     # Only when requested
  motivation_letter.docx    # Only when requested
  review_notes.txt           # Keywords, missing requirements, checks to make
  application.json           # Job snapshot, model, source resume hash and provenance
```

Use **Saved applications** on a card to select versions and download PDFs or DOCX files.
Older versions remain unchanged; create a new version to use the redesigned PDF layout.
Downloads and browsing never regenerate documents. Versions remain available after
restart and never overwrite the original resume or earlier applications. Rendering
is staged so failures do not publish half-finished applications. Delete unwanted
application folders manually. These files contain personal data and are ignored by Git.

## Project layout

```text
app.py                       # Streamlit entry point and presentation
jobbing/
  __init__.py                # Project/data paths
  job_scrape.py              # Reusable scraper and CLI
  filter_jobs.py             # LLM assessment, ranking, CSV IO, and CLI
  generate_application.py    # Per-job LLM drafts and versioned artifacts
  documents.py               # Original-template DOCX styling and PDF export
assets/style.css             # Dashboard styling
.streamlit/config.toml       # Localhost binding, theme, upload limit, no telemetry
tests/                       # Offline pipeline and web interaction tests
data/                        # Local files; ignored by Git
  resume.docx                # Default resume
  jobs.csv                   # Existing scraped jobs
  matched_jobs.csv           # Existing ranked matches
  runs/<timestamp-id>/       # Separate outputs from scraping and matching
  applications/              # Versioned application documents, grouped by job
.env                         # API credentials; ignored by Git
.env.example                 # Configuration template
requirements.txt
```

Existing root-level resume/jobs/matches were moved into `data/` without altering
their contents. The old root scripts now live in `jobbing/`; use the module commands below.
Web runs never overwrite prior results. Files in `data/runs/` persist after restart;
delete unwanted run directories manually. Empty or failed operations may leave a run
folder without a shortlist. Browser selections reset on a new session, but files remain.

## CLI

```bash
# These examples use new output names so existing CSVs remain intact.
python -m jobbing.job_scrape --search "software engineer" \
  --location "Amsterdam, Netherlands" --results 50 --hours-old 24 \
  --output data/jobs_new.csv

python -m jobbing.filter_jobs --jobs data/jobs_new.csv \
  --min-score 75 --batch-size 5 --output data/matched_jobs_new.csv
```

Defaults are `data/resume.docx`, `data/jobs.csv`, and `data/matched_jobs.csv`, resolved
relative to the project, not the working directory. Explicit paths are relative to
your working directory. Existing output paths are rejected before network calls;
choose a new path for each scraping/matching CLI run. Their output directory must already exist.

Generate documents for the first row of a matched CSV (1-based, original CSV order):

```bash
python -m jobbing.generate_application --matches data/matched_jobs.csv \
  --job-number 1 --with-letter
```

Omit `--with-letter` for resume only. Optional flags: `--resume`, `--model`, and
`--output-dir` (default `data/applications`). Application output folders are created
automatically, with a fresh version on every successful run.

## Configuration and ranking

- `.env` is always loaded from the project root. Existing environment variables take
  precedence. Restart the web server after changing `.env`.
- `OPENAI_MODEL` defaults to `gpt-4.1-mini`.
- `OPENAI_REASONING_EFFORT` is optional (`low`, `medium`, `high`) and is passed
  straight to OpenAI chat completions for supported reasoning models/endpoints;
  no LangChain layer is needed.
- `OPENAI_BASE_URL` can point to an OpenAI-compatible endpoint supporting
  structured JSON-schema outputs.
- Default fit threshold: **75/100**. Explicit hard blockers are excluded.
- Matches sort by **company tier → pay tier → benefits tier → resume fit**, descending.
  Earlier criteria always take precedence. Benefits cannot outweigh company/pay tiers.
- Preference tiers run from 0 to 3. Company tiers favor large (1000+ employees),
  established businesses using supplied company data, not name recognition.
- Pay tiers judge disclosed compensation against demonstrated experience and location,
  considering currency and pay period. Undisclosed or scraper-estimated salaries are
  unknown; "competitive salary" is not evidence of decent pay.
- Benefits use only the job description. Multiple substantial, concrete benefits
  outrank one substantial benefit, then basic/vague perks.
- Missing evidence is a blank/unknown tier, ranked after known tiers within the same
  higher-priority group. This is conservative prioritization, not proof an employer is worse.
- Fit scores are **not callback probabilities**. Stability is an unverified proxy;
  pay estimates are not verified market benchmarks. Review evidence and gaps yourself.
- All original columns are preserved, followed by `llm_match_score`, `llm_reason`,
  `llm_gaps`, `llm_hard_blockers`, `llm_company_score`, `llm_pay_score`,
  `llm_benefits_score`, and `llm_preference_reason`.

## Limits

DOCX body paragraphs and tables are supported, not images, headers, or footers.
Input CSVs need `title` and `description` columns and use the scraper's backslash
escaping and UTF-8 encoding (optional BOM). No job URLs are fetched during matching.
All job text is sent without truncation; reduce batch size for context-limit errors.

This is a local, single-user tool, not a background job service. Keep the page open
until a step completes; stopping/restarting does not resume a partial run. Matching
uses sequential batches with SDK retries. API/validation failures happen before a
new shortlist is written; rerunning reevaluates all jobs. A completed run with no
matches saves a header-only CSV and shows an empty state, not old results.

## Offline tests

```bash
pip install -r tests/requirements.txt
python -m unittest discover -s tests -v
```

Tests mock JobSpy and the LLM client. They cover ranking, CSV preservation, overwrite
protection, progress, failure handling, DOCX/PDF generation, extractable PDF text,
typography, page-break handling, hyperlink preservation, application versioning, and Streamlit interactions.
The test-only `pypdf` dependency inspects PDF content and link annotations.
No network, API key, or paid calls are needed for tests.
