"""Scrape LinkedIn jobs with descriptions, for the CLI or local web app."""

import argparse
import csv
from pathlib import Path

from jobspy import scrape_jobs

from jobbing import DATA


def run_scrape(output_path, *, search_term="software engineer",
               location="Amsterdam, Netherlands", results_wanted=50, hours_old=24):
    output_path = Path(output_path)
    if not search_term.strip() or not location.strip():
        raise ValueError("Search term and location are required.")
    if not 1 <= results_wanted <= 500 or not 1 <= hours_old <= 720:
        raise ValueError("Use 1-500 results and a posting age of 1-720 hours.")
    if output_path.exists():
        raise ValueError(f"Output already exists: {output_path}. Choose a new --output path.")
    if not output_path.parent.is_dir():
        raise ValueError(f"Output directory does not exist: {output_path.parent}")

    jobs = scrape_jobs(
        site_name=["linkedin"], search_term=search_term.strip(), location=location.strip(),
        results_wanted=results_wanted, hours_old=hours_old, fetch_description=True,
    )
    # JobSpy can return an empty frame without columns. Save a valid empty CSV,
    # never silently reuse jobs from an earlier search.
    if jobs.empty:
        jobs = jobs.reindex(columns=list(dict.fromkeys([*jobs.columns, "title", "description"])))
    elif not {"title", "description"}.issubset(jobs.columns):
        raise ValueError("Scraper returned an unexpected schema; no CSV was written.")
    with output_path.open("x", newline="", encoding="utf-8-sig") as handle:
        jobs.to_csv(handle, quoting=csv.QUOTE_NONNUMERIC, escapechar="\\", index=False)
    return len(jobs)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--search", default="software engineer")
    parser.add_argument("--location", default="Amsterdam, Netherlands")
    parser.add_argument("--results", type=int, default=50)
    parser.add_argument("--hours-old", type=int, default=24)
    parser.add_argument("--output", type=Path, default=DATA / "jobs.csv")
    args = parser.parse_args()
    try:
        count = run_scrape(args.output, search_term=args.search, location=args.location,
                           results_wanted=args.results, hours_old=args.hours_old)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    print(f"Saved {count} jobs to {args.output}")


if __name__ == "__main__":
    main()
