from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from mtc_tool.reconcile import compare_reports, save_comparison
from mtc_tool.workbook import WorkbookError


class ReconciliationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.original = self.folder / "original.xlsx"
        self.candidate = self.folder / "candidate.xlsx"

    def tearDown(self):
        self.temp.cleanup()

    def save(self, path, entries):
        wb = Workbook()
        ws = wb.active
        ws.title = "OPEN TEMPLATE"
        headers = [f"Column {i}" for i in range(42)]
        headers[13], headers[4], headers[17] = "PO REF #", "UNIQUE CODE", "MATERIAL CODE"
        ws.append(headers)
        for po, branch, material, qty, amount in entries:
            row = [None] * 42
            row[13], row[4], row[17], row[21], row[30] = po, branch, material, qty, amount
            ws.append(row)
        wb.save(path)
        wb.close()

    def test_key_matching_order_and_all_column_differences(self):
        self.save(self.original, [("PO1", "B1", 123, 2, 20), ("PO2", "B2", 124, 3, 30)])
        self.save(self.candidate, [("PO2", "B2", 124, 3, 30), ("PO1", "B1", "123", 2, 20)])
        result = compare_reports(self.original, self.candidate)
        self.assertEqual(result.missing, [])
        self.assertEqual(result.extra, [])
        self.assertFalse(result.order_matches)
        # Key types normalize for matching; different cell types remain visible.
        self.assertEqual(len(result.differences), 1)
        self.assertEqual(result.differences[0][2], "R")

    def test_missing_extra_duplicate_multiplicity_and_totals(self):
        self.save(self.original, [("PO1", "B1", 123, 2, 20)] * 2)
        self.save(self.candidate, [("PO1", "B1", 123, 2, 20), ("PO2", "B2", 124, 3, 30)])
        result = compare_reports(self.original, self.candidate)
        self.assertEqual(len(result.missing), 1)
        self.assertEqual(len(result.extra), 1)
        self.assertEqual(len(result.duplicate_original), 1)
        self.assertEqual(result.totals(result.actual)["NET AMOUNT"], 50)
        self.assertFalse(result.matches)
        audit = self.folder / "audit.xlsx"
        save_comparison(result, audit)
        wb = load_workbook(audit)
        self.assertEqual(wb["Missing records"].max_row, 2)
        self.assertEqual(wb["Extra records"].max_row, 2)
        wb.close()
        with self.assertRaises(WorkbookError):
            save_comparison(result, self.original)

    def test_exact_duplicate_pairs_matched_before_changed_pairs(self):
        self.save(self.original, [("PO1", "B1", 123, 1, 20), ("PO1", "B1", 123, 2, 20)])
        self.save(self.candidate, [("PO1", "B1", 123, 2, 20), ("PO1", "B1", 123, 3, 20)])
        result = compare_reports(self.original, self.candidate)
        self.assertEqual(len(result.differences), 1)
        self.assertEqual(result.differences[0][3:], ("Column 21", 1, 3))

    def test_numeric_tolerance_and_no_false_match_for_missing_duplicate(self):
        self.save(self.original, [("PO1", "B1", 123, 2, 20)])
        self.save(self.candidate, [("PO1", "B1", 123, 2, 20.000000001)])
        self.assertTrue(compare_reports(self.original, self.candidate).matches)
        self.save(self.candidate, [("PO1", "B1", 123, 2, 20)] * 2)
        self.assertEqual(len(compare_reports(self.original, self.candidate).extra), 1)


if __name__ == "__main__":
    unittest.main()
