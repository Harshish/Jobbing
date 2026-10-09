import csv
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from docx import Document

import pandas as pd
from streamlit.testing.v1 import AppTest

import jobbing
from jobbing import filter_jobs, job_scrape
from tests.test_filter_jobs import fake_client, result


APP = jobbing.ROOT / "app.py"


def write_resume(path):
    document = Document()
    document.styles["Normal"].font.name = "Arial"
    document.add_paragraph("Python engineer with SQL experience")
    document.save(path)


def sample_jobs():
    return pd.DataFrame([
        {"title": "Backend Engineer", "company": "Example One", "location": "Amsterdam",
         "description": "Python and SQL", "is_remote": True, "job_url": "https://example.com/1"},
        {"title": "Platform Engineer", "company": "Example Two", "location": "Amsterdam",
         "description": "Python APIs", "is_remote": False, "job_url": "https://example.com/2"},
    ])


class PipelineTests(unittest.TestCase):
    def test_scrape_parameters_and_empty_results(self):
        with TemporaryDirectory() as folder, patch.object(job_scrape, "scrape_jobs") as scraper:
            output = Path(folder) / "jobs.csv"
            scraper.return_value = pd.DataFrame()
            self.assertEqual(job_scrape.run_scrape(output, search_term="Python", location="Utrecht",
                                                   results_wanted=12, hours_old=48), 0)
            self.assertEqual(filter_jobs.read_jobs(output)[1], [])
            self.assertEqual(scraper.call_args.kwargs["search_term"], "Python")
            self.assertEqual(scraper.call_args.kwargs["hours_old"], 48)
            self.assertTrue(scraper.call_args.kwargs["fetch_description"])
            with self.assertRaises(ValueError):
                job_scrape.run_scrape(output)
            self.assertEqual(scraper.call_count, 1)

    def test_progress_and_no_output_on_failed_batch(self):
        with TemporaryDirectory() as folder, patch.object(filter_jobs, "OpenAI"), \
                patch.object(filter_jobs, "assess_batch") as assess, \
                patch.dict("os.environ", {"OPENAI_API_KEY": "offline-test"}):
            root = Path(folder)
            resume, jobs, output = (root / name for name in ("resume.docx", "jobs.csv", "matched.csv"))
            write_resume(resume)
            sample_jobs().to_csv(jobs, index=False, escapechar="\\", quoting=csv.QUOTE_ALL)
            updates = []
            assess.side_effect = [[result(0, 90)], ValueError("Invalid model response")]
            with self.assertRaisesRegex(ValueError, "Invalid model response"):
                filter_jobs.run_filter(resume, jobs, output, batch_size=1,
                                       on_progress=lambda done, total: updates.append((done, total)))
            self.assertEqual(updates, [(0, 2), (1, 2)])
            self.assertFalse(output.exists())
            assess.side_effect = [[result(0, 90)], [result(1, 80)]]
            filter_jobs.run_filter(resume, jobs, output, batch_size=1,
                                   on_progress=lambda done, total: updates.append((done, total)))
            self.assertEqual(updates[-1], (2, 2))
            self.assertEqual(len(filter_jobs.read_jobs(output, allow_assessed=True)[1]), 2)


class WebTests(unittest.TestCase):
    def test_untrusted_job_links(self):
        from app import safe_job_url

        self.assertEqual(safe_job_url("https://example.com/job"), "https://example.com/job")
        for value in (None, "", "javascript:alert(1)", "data:text/html,test", 
                      "file:///etc/passwd", "//example.com", "https://user:secret@example.com"):
            self.assertIsNone(safe_job_url(value))

    def test_scrape_match_browse_without_network(self):
        with TemporaryDirectory() as folder, patch.object(jobbing, "DATA", Path(folder)), \
                patch.object(job_scrape, "scrape_jobs", return_value=sample_jobs()) as scraper, \
                patch.object(filter_jobs, "OpenAI") as llm, \
                patch.dict("os.environ", {"OPENAI_API_KEY": "offline-test"}):
            write_resume(Path(folder) / "resume.docx")
            llm.return_value.__enter__.return_value = fake_client([
                result(0, 85, company=3, pay=2, benefits=3), result(1, 95, company=2, pay=3, benefits=2),
            ])
            at = AppTest.from_file(str(APP), default_timeout=20).run()
            self.assertFalse(at.exception)
            self.assertTrue(at.button(key="run_filter").disabled)
            scraper.assert_not_called()
            llm.assert_not_called()

            next(button for button in at.button if button.label == "Find jobs").click().run()
            self.assertFalse(at.exception)
            self.assertEqual(scraper.call_count, 1)
            self.assertFalse(at.button(key="run_filter").disabled)
            at.button(key="run_filter").click().run()
            self.assertFalse(at.exception)
            self.assertFalse(at.error)
            self.assertEqual(at.metric[0].value, "2")
            self.assertEqual(len(list(Path(folder).glob("runs/*/matched_jobs.csv"))), 1)
            self.assertEqual(llm.call_count, 1)

            at.text_input(key="search_results").set_value("Example One").run()
            self.assertFalse(at.exception)
            self.assertEqual(len(at.dataframe[-1].value), 1)
            self.assertEqual(at.dataframe[-1].value.iloc[0]["Company"], "Example One")
            at.text_input(key="search_results").set_value("no such company").run()
            self.assertTrue(any("No roles match" in notice.value for notice in at.info))
            self.assertEqual(llm.call_count, 1)  # Reactive browsing never reruns paid matching.
            self.assertEqual(scraper.call_count, 1)

            # A new empty scrape must not leave old jobs or results active.
            scraper.return_value = pd.DataFrame()
            next(button for button in at.button if button.label == "Find jobs").click().run()
            self.assertFalse(at.exception)
            self.assertTrue(at.button(key="run_filter").disabled)
            self.assertIsNone(at.session_state["match_path"])
            self.assertEqual(len(list(Path(folder).glob("runs/*/matched_jobs.csv"))), 1)

    def test_failure_states_and_empty_shortlist(self):
        with TemporaryDirectory() as folder, patch.object(jobbing, "DATA", Path(folder)), \
                patch.object(job_scrape, "scrape_jobs", side_effect=RuntimeError("Scraper offline")), \
                patch.object(filter_jobs, "OpenAI") as llm, \
                patch.dict("os.environ", {"OPENAI_API_KEY": "offline-test"}):
            root = Path(folder)
            write_resume(root / "resume.docx")
            sample_jobs().to_csv(root / "jobs.csv", index=False, escapechar="\\", quoting=csv.QUOTE_ALL)
            at = AppTest.from_file(str(APP), default_timeout=20).run()
            next(button for button in at.button if button.label == "Find jobs").click().run()
            self.assertTrue(any("Scraper offline" in error.value for error in at.error))
            self.assertFalse(at.exception)

            llm.return_value.__enter__.side_effect = RuntimeError("Provider offline")
            at.button(key="run_filter").click().run()
            self.assertFalse(at.exception)
            self.assertTrue(any("Provider offline" in error.value for error in at.error))
            self.assertFalse(list(root.glob("runs/*/matched_jobs.csv")))

            llm.return_value.__enter__.side_effect = None
            llm.return_value.__enter__.return_value = fake_client([result(0, 20), result(1, 30)])
            at.button(key="run_filter").click().run()
            self.assertFalse(at.exception)
            self.assertEqual(at.metric[0].value, "0")
            self.assertTrue(any("No jobs met" in notice.value for notice in at.info))


if __name__ == "__main__":
    unittest.main()
