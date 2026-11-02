r"""Opt-in acceptance tests against real desktop Excel and the supplied files.

Run: $env:MTC_EXCEL_TESTS='1'; .venv\Scripts\python -m unittest discover -s tests -v
Optional: MTC_FIXTURE_DIR overrides the default Downloads fixture directory.
"""
from __future__ import annotations

import hashlib
import math
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from mtc_tool.generator import generate
from mtc_tool.workbook import RAW_HEADERS, ValidationError, WorkbookError, discover_sheets
from test_workbook import sample_row, save_raw


@unittest.skipUnless(os.environ.get("MTC_EXCEL_TESTS") == "1", "Set MTC_EXCEL_TESTS=1 to run desktop Excel acceptance tests")
class ExcelAcceptanceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.fixtures = Path(os.environ.get("MTC_FIXTURE_DIR", Path.home() / "Downloads"))

    def tearDown(self):
        self.temp.cleanup()

    def test_supplied_raw_counts_formulas_and_original_unchanged(self):
        path = self.fixtures / "MTC Customer Open Order as of 10.31.2026 10.32 AM.xlsx"
        before = hashlib.sha256(path.read_bytes()).digest()
        result = generate(path, discover_sheets(path)[0], self.folder)
        self.assertEqual((result.input_rows, result.output_rows, result.excluded_rows), (173, 164, 9))
        self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), before)
        wb = load_workbook(result.path)
        self.assertEqual(wb["RAW"]["S2"].value, "0017")
        self.assertEqual(wb["OPEN TEMPLATE"]["O2"].value, "=RAW!X3")
        self.assertEqual(wb["Sheet1"].sheet_state, "hidden")
        wb.close()

    def test_entire_sample_report_values_match(self):
        expected = load_workbook(self.fixtures / "MTC OUTPUT.xlsx", data_only=True)
        try:
            raw = Workbook()
            raw.active.title = "Sample Raw"
            raw.active.append(list(RAW_HEADERS))
            for row in expected["RAW"].iter_rows(min_row=2, max_col=36, values_only=True):
                values = row[1:36]
                if values[0] is not None:
                    raw.active.append(values)
            path = self.folder / "sample-raw.xlsx"
            raw.save(path)
            raw.close()
            result = generate(path, "Sample Raw", self.folder)
            self.assertEqual((result.input_rows, result.output_rows, result.excluded_rows), (215, 206, 9))
            actual = load_workbook(result.path, data_only=True)
            try:
                for row in expected["OPEN TEMPLATE"].iter_rows(min_row=1, max_row=207, max_col=42):
                    for cell in row:
                        a = actual["OPEN TEMPLATE"][cell.coordinate].value
                        b = cell.value
                        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                            self.assertTrue(math.isclose(a, b, rel_tol=1e-10, abs_tol=1e-6),
                                            f"{cell.coordinate}: {a!r} != {b!r}")
                        else:
                            self.assertEqual(a, b, cell.coordinate)
            finally:
                actual.close()
        finally:
            expected.close()

    def test_original_raw_can_be_regenerated_directly_without_report_copy(self):
        from mtc_tool.reconcile import compare_reports
        path = self.fixtures / "MTC OUTPUT.xlsx"
        result = generate(path, "RAW", self.folder, path)
        comparison = compare_reports(path, result.path)
        self.assertTrue(comparison.matches)
        self.assertTrue(comparison.order_matches)
        self.assertEqual(comparison.duplicate_candidate, [])
        self.assertEqual(result.input_rows, 215)
        self.assertEqual(result.output_rows, 206)
        self.assertEqual(comparison.totals(comparison.actual)["PO QTY"], 893)
        self.assertAlmostEqual(comparison.totals(comparison.actual)["NET AMOUNT"], 2196884.22)

    def test_zero_served_and_partial_served_rules(self):
        path = self.folder / "raw.xlsx"
        for served, count in ((0, 0), (1, 1)):
            row = sample_row()
            row[26] = served
            save_raw(path, [row])
            result = generate(path, "Orders", self.folder)
            self.assertEqual(result.output_rows, count)
            wb = load_workbook(result.path, data_only=True)
            try:
                if count:
                    ws = wb["OPEN TEMPLATE"]
                    self.assertEqual(ws["V2"].value, 2)
                    self.assertEqual(ws["W2"].value, 1)
                    self.assertEqual(ws["Y2"].value, 2)
                    self.assertEqual(ws["AI2"].value, "10/22/2026")
                    self.assertEqual(ws["AE2"].value, 222)
                    self.assertEqual(ws["AF2"].value, 222)
                else:
                    self.assertEqual(wb["OPEN TEMPLATE"].max_row, 1)
            finally:
                wb.close()

    def test_failed_validation_and_calculation_publish_nothing(self):
        path = self.folder / "raw.xlsx"
        output = self.folder / "outputs"
        row = sample_row()
        row[0] = "UNKNOWN"
        save_raw(path, [row])
        with self.assertRaises(ValidationError):
            generate(path, "Orders", output)
        self.assertFalse(output.exists())
        save_raw(path, [sample_row()])
        with patch("mtc_tool.generator.recalculate_with_excel", side_effect=WorkbookError("Excel unavailable")):
            with self.assertRaises(WorkbookError):
                generate(path, "Orders", output)
        self.assertEqual(list(output.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
