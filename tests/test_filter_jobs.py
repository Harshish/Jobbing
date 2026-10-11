import csv
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
from zipfile import ZipFile

from jobbing import filter_jobs as app


def result(index, score, blockers=None, company=None, pay=None, benefits=None):
    return app.Assessment(row_index=index, match_score=score, reason="Relevant Python work",
                          gaps=["Location unknown"], hard_blockers=blockers or [],
                          company_score=company, pay_score=pay, benefits_score=benefits,
                          preference_reason="Company, pay and benefits evidence or unknowns")


def fake_client(assessments):
    message = SimpleNamespace(refusal=None, parsed=app.BatchResult(assessments=assessments))
    response = SimpleNamespace(choices=[SimpleNamespace(message=message)])
    return SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(
        parse=lambda **kwargs: response)))


class FilterJobsTest(unittest.TestCase):
    def test_threshold_blockers_and_sorting(self):
        jobs = [{"title": str(i)} for i in range(4)]
        matches = app.shortlist(jobs, [result(0, 75), result(1, 74),
                                     result(2, 95, ["Mandatory language mismatch"]),
                                     result(3, 90)], 75)
        self.assertEqual([m["title"] for m in matches], ["3", "0"])
        self.assertEqual(matches[0]["llm_gaps"], "Location unknown")

    def test_preference_order_and_unknowns(self):
        assessments = [
            result(0, 99, company=2, pay=3, benefits=3),
            result(1, 95, company=3, pay=1, benefits=3),
            result(2, 90, company=3, pay=2, benefits=1),
            result(3, 80, company=3, pay=2, benefits=3),
            result(4, 85, company=3, pay=2, benefits=3),
            result(5, 99, company=3, pay=None, benefits=3),
            result(6, 99, company=None, pay=3, benefits=3),
            result(7, 99, company=3, pay=2, benefits=None),
            result(8, 74, company=3, pay=3, benefits=3),
            result(9, 99, blockers=["Mandatory mismatch"], company=3, pay=3, benefits=3),
        ]
        jobs = [{"title": str(i)} for i in range(len(assessments))]
        matches = app.shortlist(jobs, assessments, 75)
        self.assertEqual([m["title"] for m in matches], ["4", "3", "2", "7", "1", "5", "0", "6"])
        self.assertEqual(matches[0]["llm_company_score"], 3)
        self.assertIsNone(matches[-1]["llm_company_score"])
        self.assertTrue(matches[0]["llm_preference_reason"])

    def test_preference_tier_validation(self):
        for field in ("company", "pay", "benefits"):
            for invalid in (-1, 4, "3", True):
                with self.subTest(field=field, invalid=invalid), self.assertRaises(ValueError):
                    result(0, 80, **{field: invalid})
        self.assertEqual(app.shortlist([], [], 75), [])

    def test_visa_sponsorship_signal(self):
        self.assertEqual(app.visa_sponsorship_status({"description": "We offer visa sponsorship for this role."}), "yes")
        self.assertEqual(app.visa_sponsorship_status({"description": "Candidates must already have work authorization; we cannot sponsor visas."}), "no")
        self.assertEqual(app.visa_sponsorship_status({"description": "Python and SQL"}), "unknown")

        assessments = [
            result(0, 90, company=3, pay=2, benefits=2),
            result(1, 95, company=3, pay=2, benefits=2),
            result(2, 99, company=3, pay=2, benefits=2),
        ]
        jobs = [
            {"title": "sponsored", "description": "Visa sponsorship available."},
            {"title": "no", "description": "Must already have work authorization; no visa sponsorship."},
            {"title": "unknown", "description": "Distributed systems role."},
        ]
        self.assertEqual([m["title"] for m in app.shortlist(jobs, assessments, 75)],
                         ["sponsored", "unknown", "no"])

    def test_full_time_signal_and_sorting(self):
        self.assertEqual(app.employment_type_status({"job_type": "Full-time"}), "full_time")
        self.assertEqual(app.employment_type_status({"employment_type": "Permanent"}), "full_time")
        self.assertEqual(app.employment_type_status({"job_type": "Contract"}), "other")
        self.assertEqual(app.employment_type_status({}), "unknown")

        assessments = [
            result(0, 99, company=3, pay=2, benefits=2),
            result(1, 80, company=3, pay=2, benefits=2),
            result(2, 90, company=3, pay=2, benefits=2),
        ]
        jobs = [
            {"title": "contract", "job_type": "Contract"},
            {"title": "full", "job_type": "Full-time"},
            {"title": "unknown"},
        ]
        self.assertEqual([m["title"] for m in app.shortlist(jobs, assessments, 75)],
                         ["full", "contract", "unknown"])

    def test_batch_indexes_and_score_validation(self):
        batch = [(0, {"title": "Engineer"}), (1, {"title": "Developer"})]
        valid = [result(1, 80), result(0, 90)]
        self.assertEqual(app.assess_batch(fake_client(valid), "fake", "resume", batch), valid)
        for invalid in ([result(0, 80)], [result(0, 80), result(0, 90)],
                        [result(0, 80), result(2, 90)]):
            with self.assertRaises(ValueError):
                app.assess_batch(fake_client(invalid), "fake", "resume", batch)
        with self.assertRaises(ValueError):
            result(0, 101)

    def test_reasoning_effort_forwarded(self):
        assessment = result(0, 85)
        message = SimpleNamespace(refusal=None, parsed=app.BatchResult(assessments=[assessment]))
        response = SimpleNamespace(choices=[SimpleNamespace(message=message)])
        parse = Mock(return_value=response)
        client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(parse=parse)))

        self.assertEqual(app.assess_batch(client, "fake", "resume", [(0, {"title": "Engineer"})],
                                          reasoning_effort="low"), [assessment])
        self.assertEqual(parse.call_args.kwargs["reasoning_effort"], "low")

    def test_offline_pipeline_and_overwrite_protection(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            resume, jobs, output = (root / name for name in
                                    ("resume.docx", "jobs.csv", "matched.csv"))
            with ZipFile(resume, "w") as archive:
                archive.writestr("word/document.xml", '''<w:document
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:body><w:p><w:r><w:t>Python experience</w:t></w:r></w:p>
                    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>SQL</w:t></w:r></w:p>
                    </w:tc></w:tr></w:tbl></w:body></w:document>''')
            self.assertEqual(app.read_resume(resume), "Python experience\nSQL")
            original = {"title": "Engineer", "description": 'Python, SQL\n"Backend" \\ APIs'}
            with jobs.open("w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=list(original), escapechar="\\",
                                        quoting=csv.QUOTE_ALL)
                writer.writeheader()
                writer.writerow(original)
            self.assertEqual(app.read_jobs(jobs)[1], [original])
            argv = ["filter_jobs.py", "--resume", str(resume), "--jobs", str(jobs),
                    "--output", str(output)]
            with patch("sys.argv", argv), patch.dict("os.environ", {"OPENAI_API_KEY": "fake"}), \
                    patch.object(app, "OpenAI") as constructor:
                constructor.return_value.__enter__.return_value = fake_client([result(0, 85)])
                app.main()
                with output.open(newline="", encoding="utf-8-sig") as handle:
                    matches = list(csv.DictReader(handle, escapechar="\\"))
                self.assertEqual(len(matches), 1)
                self.assertEqual(matches[0]["description"], original["description"])
                self.assertEqual(matches[0]["llm_company_score"], "")
                self.assertEqual(matches[0]["llm_pay_score"], "")
                self.assertEqual(matches[0]["llm_benefits_score"], "")
                self.assertTrue(matches[0]["llm_preference_reason"])
                before = output.read_bytes()
                with self.assertRaises(SystemExit):
                    app.main()
                self.assertEqual(output.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
