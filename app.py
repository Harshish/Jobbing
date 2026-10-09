"""Local, step-by-step job search dashboard. Run: streamlit run app.py"""

from datetime import datetime
from html import escape
import math
import os
from pathlib import Path
from urllib.parse import urlsplit
from uuid import uuid4

from dotenv import load_dotenv
import pandas as pd
import streamlit as st

from jobbing import DATA, ROOT
from jobbing.filter_jobs import read_jobs, run_filter
from jobbing.job_scrape import run_scrape
from jobbing.generate_application import generate_application, job_folder_name


def saved_files(pattern):
    return sorted(
        [*DATA.glob(pattern), *DATA.glob(f"runs/*/{pattern}")],
        key=lambda path: path.stat().st_mtime, reverse=True,
    )


def new_run():
    path = DATA / "runs" / f"{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:8]}"
    path.mkdir(parents=True)
    return path


def file_label(path):
    return str(Path(path).relative_to(DATA)) if path else "No file selected"


def safe_job_url(value):
    """Only expose ordinary web links from untrusted scraped content."""
    try:
        parts = urlsplit(value or "")
        if parts.scheme in {"https", "http"} and parts.hostname and not parts.username:
            return value
    except ValueError:
        pass
    return None


def salary_label(job):
    low, high = job.get("min_amount", ""), job.get("max_amount", "")
    if not low and not high:
        return "Salary not disclosed"
    amount = f"{low} – {high}" if low and high else f"From {low}" if low else f"Up to {high}"
    return " · ".join(filter(None, [amount, job.get("currency"), job.get("interval"),
                                   job.get("salary_source")]))


def tier(value):
    return f"{value}/3" if value not in (None, "") else "Unknown"


def render_application_actions(job, widget_key, resume_source, model, enabled, include_letter):
    label = "Create resume + letter" if include_letter else "Create resume"
    if st.button(label, key=f"generate:{widget_key}", disabled=not enabled):
        with st.spinner("Tailoring your application from the original resume…"):
            try:
                destination = generate_application(
                    resume_source, job, DATA / "applications", model=model,
                    include_letter=include_letter,
                )
                st.session_state[f"version:{widget_key}"] = str(destination)
                st.success("Drafts created. Review the facts and notes before submitting.")
            except Exception as exc:
                st.error(f"Could not create the application: {exc}. Previous versions are unchanged.")

    group = DATA / "applications" / job_folder_name(job)
    versions = sorted((p for p in group.glob("*") if p.is_dir() and not p.name.startswith(".")
                       and (p / "resume.docx").is_file()), key=lambda p: p.stat().st_mtime, reverse=True)
    if not versions:
        return
    with st.expander(f"Saved applications · {len(versions)} version(s)"):
        options = list(map(str, versions))
        key = f"version:{widget_key}"
        if st.session_state.get(key) not in options:
            st.session_state[key] = options[0]
        version = Path(st.selectbox("Application version", options,
                                   format_func=lambda value: Path(value).name, key=key))
        st.caption(f"Saved in {file_label(version)} · These drafts may use a different resume than the current sidebar selection.")
        for filename, label in (("resume.pdf", "Download resume · PDF"),
                                ("motivation_letter.pdf", "Download motivation letter · PDF"),
                                ("resume.docx", "Editable resume · DOCX"),
                                ("motivation_letter.docx", "Editable motivation letter · DOCX")):
            file = version / filename
            if file.is_file():
                mime = "application/pdf" if file.suffix == ".pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                st.download_button(label, file.read_bytes(), file_name=filename, mime=mime,
                                   key=f"download:{widget_key}:{filename}")
        if not (version / "resume.pdf").exists():
            st.caption("This older version has no PDF. Create a new version to use the redesigned PDF layout.")
        notes = version / "review_notes.txt"
        if notes.is_file():
            st.text(notes.read_text(encoding="utf-8"))


def render_results(resume_source, model, generation_enabled):
    st.header("03 / Your shortlist")
    st.caption("Company → pay → benefits → resume fit. Unknown does not mean bad.")
    paths = [None, *map(str, saved_files("matched_jobs*.csv"))]
    if "match_path" not in st.session_state:
        st.session_state.match_path = paths[1] if len(paths) > 1 else None
    if st.session_state.match_path not in paths:
        st.session_state.match_path = None
    chosen = st.selectbox("Saved shortlist", paths, format_func=file_label, key="match_path")
    if not chosen:
        st.info("Run step 2 or select a saved shortlist. Previous runs remain available here.")
        return
    path = Path(chosen)
    try:
        columns, rows = read_jobs(path, allow_assessed=True)
        if "llm_match_score" not in columns:
            raise ValueError("This file has no resume-fit scores. Run matching first.")
        for row in rows:
            score = int(row["llm_match_score"])
            if not 0 <= score <= 100:
                raise ValueError("Invalid resume-fit score in the saved CSV.")
    except (OSError, ValueError, TypeError) as exc:
        st.error(f"Cannot load shortlist: {exc}")
        return

    st.caption(f"Viewing saved results: {file_label(path)}. Changing settings does not rescore this file.")
    stats = st.columns(3)
    stats[0].metric("Shortlisted roles", len(rows))
    stats[1].metric("Companies", len({row.get("company") for row in rows if row.get("company")}))
    stats[2].metric("Best resume fit", f"{max(int(r['llm_match_score']) for r in rows)} / 100" if rows else "—")
    st.download_button("↓ Download full shortlist", path.read_bytes(), file_name="matched_jobs.csv",
                       mime="text/csv", key="download_matches")
    if not rows:
        st.info("No jobs met the fit threshold without hard blockers. Try a broader search or a lower threshold.")
        return

    st.markdown("**Tailored application drafts**")
    include_letter = st.checkbox("Include a motivation letter when creating a resume", key="include_letter")
    st.caption("Uses the original resume selected in the sidebar and the current LLM model. "
               "Enable data-sharing consent in step 2 to create drafts. Each click incurs an API call. "
               "ATS-friendly formatting is not a guaranteed ATS score.")

    search_col, score_col, remote_col = st.columns([3, 2, 1])
    query = search_col.text_input("Search roles, companies or locations", key="search_results").strip().casefold()
    floor = score_col.slider("Display minimum fit", 0, 100, 0, key="display_score")
    remote_only = remote_col.checkbox("Remote only", key="remote_only")
    visible = [(rank, job) for rank, job in enumerate(rows, 1)
               if int(job["llm_match_score"]) >= floor
               and (not query or query in " ".join(job.get(k, "") for k in
                    ("title", "company", "location")).casefold())
               and (not remote_only or job.get("is_remote", "").lower() in {"true", "1"})]
    st.caption(f"Showing {len(visible)} of {len(rows)} roles · Original preference rank preserved")
    if not visible:
        st.info("No roles match these display filters. Clear the search or lower the display threshold.")
        return

    cards, table = st.tabs(["✦ Role cards", "▦ Compare in a table"])
    with cards:
        page_count = math.ceil(len(visible) / 10)
        page = st.number_input("Page", min_value=1, max_value=page_count, value=1, step=1,
                               key=f"page:{chosen}:{query}:{floor}:{remote_only}")
        for rank, job in visible[(page - 1) * 10:page * 10]:
            with st.container(border=True):
                left, right = st.columns([5, 1])
                with left:
                    st.markdown(
                        f'<div class="eyebrow">#{rank:02d} · {escape(job.get("company") or "Company undisclosed")}</div>'
                        f'<h3 class="role-title">{escape(job.get("title") or "Untitled role")}</h3>'
                        f'<div class="role-meta">{escape(job.get("location") or "Location unspecified")}'
                        f' · {escape(job.get("job_type") or "Type unspecified")}</div>',
                        unsafe_allow_html=True,
                    )
                right.metric("Resume fit", f"{job['llm_match_score']}/100")
                st.caption(salary_label(job))
                st.markdown(
                    '<div class="tier-row">'
                    + "".join(f'<span>{label} <b>{escape(tier(job.get(column)))}</b></span>' for label, column in [
                        ("Company", "llm_company_score"), ("Pay", "llm_pay_score"),
                        ("Benefits", "llm_benefits_score")]) + '</div>',
                    unsafe_allow_html=True,
                )
                st.text(job.get("llm_reason") or "No match explanation recorded.")
                with st.expander("Why this rank · gaps · full description"):
                    st.markdown("**Preference evidence**")
                    st.text(job.get("llm_preference_reason") or "Not recorded in this CSV.")
                    st.markdown("**Gaps to check**")
                    st.text(job.get("llm_gaps") or "No gaps recorded. Verify the requirements yourself.")
                    if job.get("llm_hard_blockers"):
                        st.warning(job["llm_hard_blockers"])
                    st.markdown("**Job description**")
                    st.text(job.get("description") or "No description provided.")
                url = safe_job_url(job.get("job_url_direct")) or safe_job_url(job.get("job_url"))
                if url:
                    st.link_button("View job ↗", url)
                render_application_actions(job, f"{chosen}:{rank}", resume_source, model,
                                           generation_enabled, include_letter)
    with table:
        table_rows = []
        for rank, job in visible:
            table_rows.append({
                "Rank": rank, "Role": job.get("title", ""), "Company": job.get("company", ""),
                "Location": job.get("location", ""), "Fit": int(job["llm_match_score"]),
                "Company tier": tier(job.get("llm_company_score")), "Pay tier": tier(job.get("llm_pay_score")),
                "Benefits tier": tier(job.get("llm_benefits_score")), "Salary": salary_label(job),
                "Why it fits": job.get("llm_reason", ""), "Gaps": job.get("llm_gaps", ""),
                "Link": safe_job_url(job.get("job_url_direct")) or safe_job_url(job.get("job_url")),
            })
        st.caption("Use Role cards to create and download application documents for each job.")
        st.dataframe(pd.DataFrame(table_rows), hide_index=True, width="stretch",
                     column_config={"Fit": st.column_config.ProgressColumn("Resume fit", min_value=0, max_value=100),
                                    "Link": st.column_config.LinkColumn("Job link", display_text="Open ↗")})
    st.caption("Fit scores are not callback probabilities. Company stability and pay suitability are unverified estimates.")


def main():
    st.set_page_config(page_title="Jobbing · Your next chapter", page_icon="✦", layout="wide")
    load_dotenv(ROOT / ".env")
    st.markdown(f"<style>{(ROOT / 'assets' / 'style.css').read_text()}</style>", unsafe_allow_html=True)
    with st.sidebar:
        st.markdown("## ✦ Jobbing")
        st.caption("A more intentional job search.")
        st.divider()
        st.subheader("Your resume")
        upload = st.file_uploader("Use another resume", type=["docx"], key="resume_upload")
        resume_path = DATA / "resume.docx"
        if upload:
            st.caption("This uploaded original is used for matching and application drafts. Matching saves a local copy.")
        elif resume_path.exists():
            st.success("data/resume.docx ready")
        else:
            st.warning("Upload a DOCX resume to enable matching.")
        st.subheader("LLM settings")
        model = st.text_input("LLM model", value=os.getenv("OPENAI_MODEL") or "gpt-4.1-mini")
        threshold = st.slider("Minimum resume fit", 0, 100, 75)
        batch_size = st.number_input("Jobs per LLM batch", min_value=1, max_value=20, value=5)
        key_ready = bool(os.getenv("OPENAI_API_KEY", "").strip())
        if key_ready:
            st.caption("✓ API key configured on the server")
        else:
            st.warning("Set OPENAI_API_KEY in the project .env, then restart the app.")
        st.divider()
        st.markdown("**Your priorities**\n\n01 · Large, established companies\n\n02 · Decent pay for your experience\n\n03 · Better benefits")
        st.caption("Local, single-user workspace. Your API key is never displayed.")

    st.markdown('''<section class="hero"><div class="eyebrow">LESS NOISE. BETTER OPPORTUNITIES.</div>
        <h1>Your next chapter,<br>with a little more clarity.</h1>
        <p>Discover roles. Find your fit. Focus on the companies worth your time.</p>
        <div class="steps"><span>01 &nbsp; Discover</span><span>02 &nbsp; Match</span><span>03 &nbsp; Decide</span></div>
        </section>''', unsafe_allow_html=True)

    with st.expander("01 / Discover jobs", expanded=True):
        st.caption("Search LinkedIn for recent roles, or reuse a saved jobs CSV below.")
        with st.form("scrape_form"):
            left, right = st.columns(2)
            search_term = left.text_input("Role or keywords", "software engineer")
            location = right.text_input("Location", "Amsterdam, Netherlands")
            count = left.number_input("Maximum results", min_value=1, max_value=500, value=50)
            hours = right.number_input("Posted within (hours)", min_value=1, max_value=720, value=24)
            scrape_clicked = st.form_submit_button("Find jobs", type="primary")
        if scrape_clicked:
            with st.status("Searching LinkedIn and fetching descriptions…", expanded=True) as status:
                try:
                    output = new_run() / "jobs.csv"
                    found = run_scrape(output, search_term=search_term, location=location,
                                       results_wanted=count, hours_old=hours)
                    st.session_state.job_path = str(output)
                    st.session_state.match_path = None
                    status.update(label=f"Search complete · {found} jobs saved", state="complete")
                    if not found:
                        st.warning("No jobs returned. Broaden the keywords, location or posting window.")
                except Exception as exc:
                    status.update(label="Search failed", state="error")
                    st.error(f"Could not scrape jobs: {exc}. Previous files were not changed.")

        job_paths = [None, *map(str, saved_files("jobs*.csv"))]
        if "job_path" not in st.session_state:
            st.session_state.job_path = job_paths[1] if len(job_paths) > 1 else None
        if st.session_state.job_path not in job_paths:
            st.session_state.job_path = None
        chosen_jobs = st.selectbox("Jobs to match", job_paths, format_func=file_label, key="job_path",
                                  on_change=lambda: st.session_state.update(match_path=None))
        jobs = []
        if chosen_jobs:
            try:
                _, jobs = read_jobs(Path(chosen_jobs))
                st.caption(f"{len(jobs)} jobs loaded · {file_label(chosen_jobs)}")
                if jobs:
                    preview = pd.DataFrame(jobs)
                    st.dataframe(preview[[c for c in ("title", "company", "location", "date_posted") if c in preview]],
                                 hide_index=True, width="stretch", height=200)
            except (OSError, ValueError) as exc:
                st.error(f"Cannot load jobs: {exc}")

    with st.container(border=True):
        st.header("02 / Find your fit")
        st.write("Compare each role with your resume, then rank the matches around your priorities.")
        consent = st.checkbox("I agree to send my resume text and job listings to the configured LLM provider. API charges apply.",
                              key="llm_consent")
        ready = bool(jobs and key_ready and model.strip() and (upload or resume_path.exists()) and consent)
        if st.button("Match & rank jobs", type="primary", disabled=not ready, key="run_filter"):
            st.session_state.match_path = None
            with st.status("Reading your resume…", expanded=True) as status:
                progress = st.progress(0, text="Preparing the first batch")
                try:
                    run_dir = new_run()
                    if upload:
                        resume_path = run_dir / "resume.docx"
                        resume_path.write_bytes(upload.getvalue())
                    output = run_dir / "matched_jobs.csv"
                    selected = run_filter(
                        resume_path, Path(chosen_jobs), output, model=model.strip(), min_score=threshold,
                        batch_size=batch_size,
                        on_progress=lambda done, total: progress.progress(
                            done / total if total else 1.0, text=f"Assessed {done} of {total} jobs"),
                    )
                    st.session_state.match_path = str(output)
                    status.update(label=f"Matching complete · {len(selected)} roles shortlisted", state="complete")
                except Exception as exc:
                    status.update(label="Matching failed", state="error")
                    st.error(f"Could not finish matching: {exc}. No new shortlist is selected; previous files are preserved.")
        st.caption("Runs only when you click. Display filters and page changes never trigger API calls.")

    generation_enabled = bool(key_ready and model.strip() and (upload or resume_path.exists()) and consent)
    render_results(upload.getvalue() if upload else resume_path, model.strip(), generation_enabled)


if __name__ == "__main__":
    main()
