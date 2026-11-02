from __future__ import annotations

import csv
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from mtc_tool.paths import bundled_template
from mtc_tool.settings import Settings
from mtc_tool.workbook import (RAW_HEADERS, Issue, ValidationError, WorkbookError,
                               discover_sheets, export_issues, first_lookup,
                               lookup_key, prepare_data, write_prepared)


def sample_row():
    return [
        "MT01 DELCO PASIG", "10/20/2026", "POMTDL261020015", "CREATED",
        "PO414226101620445", "10/21/2026", "10/22/2026", "TAXABLE", 12,
        "EXCLUSIVE", "SBIC", "MT", "SM GROUP", "SA001", "240762",
        "SANFORD MARKETING CORPORATION", "SA001-N005-SU", "0017", "SMCO DC NCR",
        "LIPTON", "LI01LI09CAX150", "8720608020102",
        "69979478 LIPT EXS CHAMLINDEN 1.3GX25X6 [CASE x 150]", "CASE", 150,
        2, 1, 1, 1061.38, 1061.38, 111, 222, "mccortez",
        "10/20/2026  02:57:06 PM", "VMACOSTA",
    ]


def save_raw(path, rows, headers=RAW_HEADERS):
    wb = Workbook()
    wb.active.title = "Orders"
    wb.active.append(list(headers))
    for row in rows:
        wb.active.append(row)
    wb.save(path)
    wb.close()


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.raw = self.folder / "raw.xlsx"

    def tearDown(self):
        self.temp.cleanup()

    def test_filter_preserves_order_negative_served_and_source_rows(self):
        rows = []
        for served in (0, 1, -2, 0, 3):
            row = sample_row()
            row[26] = served
            rows.append(row)
        save_raw(self.raw, rows)
        data = prepare_data(self.raw, "Orders")
        self.assertEqual(data.retained, (1, 2, 4))
        self.assertEqual(data.excluded, 2)
        target = self.folder / "output.xlsx"
        write_prepared(data, target)
        wb = load_workbook(target)
        try:
            report = wb["OPEN TEMPLATE"]
            self.assertEqual(report.max_row, 4)
            self.assertEqual(report["O2"].value, "=RAW!X3")
            self.assertEqual(report["O4"].value, "=RAW!X6")
            self.assertEqual(report["Y2"].value, "=V2")
            self.assertEqual(report["AI2"].value, "=RAW!H3")
            self.assertEqual(report["AE2"].value, "=RAW!AG3")
            self.assertEqual(report["AF2"].value, "=RAW!AG3")
            self.assertEqual(report["U2"].value, "CASE")
            self.assertIsNone(report["J2"].value)
            self.assertIsNone(report["AP2"].value)
            self.assertEqual(wb["RAW"]["S2"].value, "0017")
            self.assertEqual(wb["RAW"]["S2"].data_type, "s")
            self.assertEqual(wb["Sheet1"].sheet_state, "hidden")
            self.assertEqual(report.auto_filter.ref, "A1:AP4")
            self.assertFalse(report.sheet_properties.filterMode)
            self.assertNotIn("A1341.", wb.defined_names)
        finally:
            wb.close()

    def test_reordered_headers_and_multiple_matching_sheets(self):
        save_raw(self.raw, [list(reversed(sample_row()))], list(reversed(RAW_HEADERS)))
        wb = load_workbook(self.raw)
        wb.copy_worksheet(wb.active).title = "Second"
        wb.save(self.raw)
        wb.close()
        self.assertEqual(discover_sheets(self.raw), ["Orders", "Second"])
        self.assertEqual(prepare_data(self.raw, "Second").records[0].get("WAREHOUSE"), "MT01 DELCO PASIG")

    def test_missing_lookup_and_bad_data_report_source_rows(self):
        row = sample_row()
        row[0] = "UNKNOWN WAREHOUSE"
        row[16] = "UNKNOWN BRANCH"
        row[22] = "UNKNOWN SKU"
        row[25] = "not numeric"
        row[6] = "02/30/2026"
        save_raw(self.raw, [row])
        with self.assertRaises(ValidationError) as caught:
            prepare_data(self.raw, "Orders")
        issues = caught.exception.issues
        self.assertEqual({i.field for i in issues},
                         {"WAREHOUSE", "BRANCH CODE", "SKU DESCRIPTION", "PO QTY", "CANCEL DATE"})
        self.assertTrue(all(i.row == 2 for i in issues))

    def test_zero_served_does_not_need_customer_or_sku_lookup(self):
        row = sample_row()
        row[26], row[16], row[22] = 0, "NEW BRANCH", "NEW SKU"
        save_raw(self.raw, [row])
        data = prepare_data(self.raw, "Orders")
        self.assertEqual(data.retained, ())
        write_prepared(data, self.folder / "empty-report.xlsx")
        wb = load_workbook(self.folder / "empty-report.xlsx")
        self.assertEqual(wb["OPEN TEMPLATE"].max_row, 1)
        self.assertEqual(wb["OPEN TEMPLATE"].auto_filter.ref, "A1:AP1")
        wb.close()

    def test_actual_dates_and_numeric_text(self):
        row = sample_row()
        row[1] = datetime(2026, 10, 20)
        row[26] = "2"
        save_raw(self.raw, [row])
        data = prepare_data(self.raw, "Orders")
        self.assertEqual(data.records[0].get("PO DATE"), "10/20/2026")
        self.assertEqual(data.records[0].get("SERVED QTY"), 2)

    def test_rejects_formulas_missing_headers_empty_and_invalid_files(self):
        row = sample_row()
        row[2] = '=HYPERLINK("https://example.com")'
        save_raw(self.raw, [row])
        with self.assertRaises(ValidationError):
            prepare_data(self.raw, "Orders")
        save_raw(self.raw, [sample_row()], RAW_HEADERS[:-1])
        self.assertEqual(discover_sheets(self.raw), [])
        with self.assertRaises(WorkbookError):
            prepare_data(self.raw, "Orders")
        save_raw(self.raw, [])
        with self.assertRaises(WorkbookError):
            prepare_data(self.raw, "Orders")
        self.raw.write_text("not an Excel workbook")
        with self.assertRaises(WorkbookError):
            discover_sheets(self.raw)

    def test_first_match_case_insensitive_lookup_keeps_key_types(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Same", 10])
        ws.append(["SAME", 99])
        lookup = first_lookup(ws, 1, 1)
        self.assertEqual(lookup[lookup_key("same")][1], 10)
        self.assertNotEqual(lookup_key("0017"), lookup_key(17))
        wb.close()

    def test_custom_reference_updates_and_invalid_layout(self):
        save_raw(self.raw, [sample_row()])
        ref = self.folder / "refs.xlsx"
        wb = load_workbook(bundled_template())
        wb["SKU"]["N3"] = "UPDATED WAREHOUSE"
        wb.save(ref)
        wb.close()
        data = prepare_data(self.raw, "Orders", ref)
        target = self.folder / "custom.xlsx"
        write_prepared(data, target, ref)
        wb = load_workbook(target)
        self.assertEqual(wb["SKU"]["N3"].value, "UPDATED WAREHOUSE")
        wb.close()
        wb = Workbook()
        wb.save(ref)
        wb.close()
        with self.assertRaises(WorkbookError):
            prepare_data(self.raw, "Orders", ref)

    def test_large_then_small_generation_has_no_stale_rows(self):
        target = self.folder / "output.xlsx"
        for count in (650, 1):
            save_raw(self.raw, [sample_row() for _ in range(count)])
            write_prepared(prepare_data(self.raw, "Orders"), target)
            wb = load_workbook(target)
            self.assertEqual(wb["RAW"].max_row, count + 1)
            self.assertEqual(wb["OPEN TEMPLATE"].max_row, count + 1)
            self.assertEqual(wb["OPEN TEMPLATE"].auto_filter.ref, f"A1:AP{count + 1}")
            self.assertEqual(wb["OPEN TEMPLATE"][f"O{count+1}"].value, f"=RAW!X{count+1}")
            wb.close()

    def test_settings_and_issue_export(self):
        path = self.folder / "settings.json"
        Settings("reference.xlsx", "output").save(path)
        self.assertEqual(Settings.load(path), Settings("reference.xlsx", "output"))
        path.write_text("[bad settings")
        self.assertEqual(Settings.load(path), Settings())
        issues = [Issue("Orders", 2, "SKU DESCRIPTION", "=BAD()", "Missing lookup")]
        target = self.folder / "issues.csv"
        export_issues(issues, target)
        with target.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        self.assertEqual(rows[1][3], "'=BAD()")


if __name__ == "__main__":
    unittest.main()
