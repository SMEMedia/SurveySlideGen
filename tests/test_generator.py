import io
import zipfile
from pathlib import Path
import unittest

from openpyxl import Workbook

from generator import PURCHASE_PROFILES, discover_metric_options, generate_presentation, metric_value, read_rows


def fixture_workbook_bytes():
    wb = Workbook()
    ws = wb.active
    rows = [
        ["Question 10 (Single Choice)"],
        ["What is your role in purchasing decisions? (Please select one)"],
        ["Number", "Choice", None, "Total", "Total Magazine Audience", "Mfg Weekly Newsletter", "SME Event Attendees", "AM.org USers", "Video or Podcast"],
        ["(Net)", "purchase influence net", "Frequency", .815, .86, .79, .90, .93, .83],
        ["Question 12 (Multiple Choice)"],
        ["Which types of products/services do you influence or purchase? (Please select all that apply)"],
        ["Number", "Choice", None, "Total", "Total Magazine Audience", "Mfg Weekly Newsletter", "SME Event Attendees", "AM.org USers", "Video or Podcast"],
    ]
    product_values = {
        "Machine Tools": [.533, .57, .53, .55, .70, .51],
        "Software": [.518, .54, .55, .59, .63, .74],
        "Automation / Robotics": [.477, .50, .53, .37, .57, .60],
        "Additive Equipment": [.387] * 6,
        "Consulting": [.296] * 6,
        "Training": [.427] * 6,
        "Other": [.151] * 6,
        "MT & Automation": [.683, .71, .68, .62, .93, .69],
    }
    for number, (label, values) in enumerate(product_values.items(), start=1):
        rows.append([number, label, "Frequency", *values])
    rows += [
        ["Question 13 (Multiple Choice)"],
        ["In which industry or markets is your organization primarily involved? (Please select all that apply)"],
        ["Number", "Choice", None, "Total", "Total Magazine Audience"],
        ["(Net)", "SSAB", "Frequency", .69, .75],
        ["Question 14 (Multiple Choice)"],
        ["Which of the following best describes your job title or primary function? (Please choose the option that best applies to you)"],
        ["Number", "Choice", None, "Total", "Total Magazine Audience"],
        ["(Net)", "leadership + purchasing", "Frequency", .43, .49],
    ]
    for row in rows:
        ws.append(row)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


class GeneratorTests(unittest.TestCase):
    def test_discovers_fixed_decision_metric(self):
        rows = read_rows(fixture_workbook_bytes())
        options = discover_metric_options(rows, "What is your role in purchasing decisions?")
        self.assertEqual(metric_value(options, "purchase influence net", "Total"), .815)

    def test_generates_downloadable_pptx(self):
        rows = read_rows(fixture_workbook_bytes())
        template = Path("assets/sme_audience_template.pptx").read_bytes()
        result = generate_presentation(
            template_bytes=template,
            rows=rows,
            company_name="Example Co",
            report_year=2026,
            qualified_value=.43,
            qualified_description="Hold manufacturing leadership and engineering roles",
            qualified_denominator=5,
            industry_value=.75,
            industry_description="Work across target manufacturing industries",
            purchase_profile_name="Machine Tools",
        )
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            slide1 = archive.read("ppt/slides/slide1.xml").decode("utf-8")
            self.assertIn("Example Co", slide1)
            self.assertIn(">X<", slide1)
            self.assertIn("ppt/charts/chart1.xml", archive.namelist())

    def test_all_purchase_profiles_generate(self):
        rows = read_rows(fixture_workbook_bytes())
        template = Path("assets/sme_audience_template.pptx").read_bytes()
        for profile in PURCHASE_PROFILES:
            with self.subTest(profile=profile):
                result = generate_presentation(
                    template_bytes=template,
                    rows=rows,
                    company_name="Example Co",
                    report_year=2026,
                    qualified_value=.43,
                    qualified_description="Hold manufacturing leadership and engineering roles",
                    qualified_denominator=5,
                    industry_value=.75,
                    industry_description="Work across target manufacturing industries",
                    purchase_profile_name=profile,
                )
                self.assertGreater(len(result), 100_000)


if __name__ == "__main__":
    unittest.main()
