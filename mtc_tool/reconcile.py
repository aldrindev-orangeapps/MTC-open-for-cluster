"""Deterministic report comparison; never modifies either input workbook.

python -m mtc_tool.reconcile ORIGINAL.xlsx CANDIDATE.xlsx --output audit.xlsx
Add --source EXPORT.xlsx to trace missing/extra keys to the selected raw export.
"""
from __future__ import annotations

import argparse
import math
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel

from .workbook import WorkbookError, lookup_key, open_xlsx, source_columns

KEY_COLUMNS = (13, 4, 17)  # PO REF #, UNIQUE CODE, MATERIAL CODE (zero-based).
METRICS = {"PO QTY": 21, "SRV QTY": 22, "UNSERVED QTY": 23, "FINAL QTY": 24,
           "LPBT CASE": 25, "VAT EX LPBT": 26, "LPAT": 27, "VAT IN": 28,
           "NET PRICE": 29, "NET AMOUNT": 30, "NET AMOUNT VAT EX": 31}


def same_value(a, b) -> bool:
    if isinstance(a, (date, datetime)):
        a = to_excel(a)
    if isinstance(b, (date, datetime)):
        b = to_excel(b)
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=1e-10, abs_tol=1e-6)
    return a == b


def key_text(value) -> str:
    # Treat numeric 801842 and its Excel text representation as the same
    # material identifier, while preserving leading zeros in text identifiers.
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return format(value, ".15g")
    return str(value or "").casefold()


def record_key(row) -> tuple:
    return tuple(key_text(row[c]) for c in KEY_COLUMNS)


@dataclass(frozen=True)
class ReportRow:
    row: int
    values: tuple

    @property
    def key(self):
        return record_key(self.values)


@dataclass
class Comparison:
    original: Path
    candidate: Path
    headers: tuple
    expected: list[ReportRow]
    actual: list[ReportRow]
    missing: list[ReportRow]
    extra: list[ReportRow]
    differences: list[tuple]
    matched: int
    duplicate_original: list[tuple]
    duplicate_candidate: list[tuple]
    candidate_errors: list[tuple]
    order_matches: bool

    @property
    def matches(self) -> bool:
        return not (self.missing or self.extra or self.differences or self.candidate_errors)

    def totals(self, records) -> dict[str, float]:
        return {name: math.fsum(float(r.values[c] or 0) for r in records) for name, c in METRICS.items()}


def read_report(wb):
    if "OPEN TEMPLATE" not in wb.sheetnames:
        raise WorkbookError("Comparison requires an OPEN TEMPLATE worksheet.")
    ws = wb["OPEN TEMPLATE"]
    headers = tuple(c.value for c in next(ws.iter_rows(min_row=1, max_row=1, max_col=42)))
    for col, expected in ((13, "PO REF #"), (4, "UNIQUE CODE"), (17, "MATERIAL CODE")):
        if str(headers[col] or "").strip() != expected:
            raise WorkbookError(f"Unexpected report header in column {get_column_letter(col + 1)}.")
    records = []
    errors = []
    for row_number, cells in enumerate(ws.iter_rows(min_row=2, max_col=42), 2):
        values = tuple(c.value for c in cells)
        if all(v is None for v in values):
            continue
        records.append(ReportRow(row_number, values))
        errors.extend(("OPEN TEMPLATE", c.coordinate, c.value) for c in cells if c.data_type == "e")
    return headers, records, errors


def compare_reports(original: Path, candidate: Path) -> Comparison:
    a = open_xlsx(original, data_only=True, read_only=True)
    b = open_xlsx(candidate, data_only=True, read_only=True)
    try:
        headers, expected, _ = read_report(a)
        actual_headers, actual, errors = read_report(b)
        differences = [(1, 1, get_column_letter(i + 1), headers[i], x, y)
                       for i, (x, y) in enumerate(zip(headers, actual_headers)) if not same_value(x, y)]
        emap, amap = defaultdict(list), defaultdict(list)
        for record in expected:
            emap[record.key].append(record)
        for record in actual:
            amap[record.key].append(record)
        missing, extra, matched = [], [], 0
        # Keep multiplicities. A dict overwrite or set comparison could hide a
        # duplicate or wrongly reuse one actual row for two expected rows.
        for key in dict.fromkeys([*emap, *amap]):
            erows, arows = list(emap[key]), list(amap[key])
            remaining = []
            for erow in erows:
                exact = next((i for i, row in enumerate(arows)
                              if all(same_value(x, y) for x, y in zip(erow.values, row.values))), None)
                if exact is not None:
                    arows.pop(exact)
                    matched += 1
                else:
                    remaining.append(erow)
            for erow in remaining:
                if arows:
                    arow = arows.pop(0)
                    matched += 1
                    differences.extend((erow.row, arow.row, get_column_letter(i + 1), headers[i], x, y)
                                       for i, (x, y) in enumerate(zip(erow.values, arow.values)) if not same_value(x, y))
                else:
                    missing.append(erow)
            extra.extend(arows)
        return Comparison(original, candidate, headers, expected, actual, missing, extra, differences,
                          matched, [(key, len(rows)) for key, rows in emap.items() if len(rows) > 1],
                          [(key, len(rows)) for key, rows in amap.items() if len(rows) > 1], errors,
                          [r.key for r in expected] == [r.key for r in actual])
    finally:
        a.close()
        b.close()


def trace_source(path: Path, sheet: str | None = None, reference: Path | None = None):
    """Return raw source row evidence by the report's three business key fields."""
    wb = open_xlsx(path, read_only=True)
    ref = open_xlsx(reference or path, data_only=True, read_only=True)
    try:
        if sheet is None:
            sheet = "RAW" if "RAW" in wb.sheetnames else wb.sheetnames[0]
        columns = source_columns(wb[sheet])
        if "SKU" not in ref.sheetnames:
            raise WorkbookError("Source tracing requires a reference workbook with a SKU sheet.")
        sku = {}
        for row in ref["SKU"].iter_rows(min_row=4, max_col=3, values_only=True):
            if row[1] is not None:
                sku.setdefault(lookup_key(row[1]), row[2])
        evidence = defaultdict(list)
        for number, row in enumerate(wb[sheet].iter_rows(min_row=2, values_only=True), 2):
            def get(name):
                index = columns[name] - 1
                return row[index] if index < len(row) else None
            if get("WAREHOUSE") is None:
                continue
            key = (key_text(get("PO REF #")), key_text(get("BRANCH CODE")),
                   key_text(sku.get(lookup_key(get("SKU DESCRIPTION")))))
            evidence[key].append((number, get("SERVED QTY"), get("PO STATUS")))
        return evidence
    finally:
        wb.close()
        ref.close()


def save_comparison(comparison: Comparison, destination: Path, source: Path | None = None,
                    source_sheet: str | None = None) -> None:
    protected = {comparison.original.resolve(), comparison.candidate.resolve()}
    if source:
        protected.add(source.resolve())
    if destination.resolve() in protected:
        raise WorkbookError("The audit path must not overwrite an input workbook.")
    original_trace = trace_source(comparison.original) if source else {}
    selected_trace = trace_source(source, source_sheet, comparison.original) if source else {}
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    summary.append(["Check", "Original", "Candidate", "Result"])
    summary.append(["Workbook", str(comparison.original), str(comparison.candidate)])
    summary.append(["Data rows", len(comparison.expected), len(comparison.actual)])
    summary.append(["Missing records", 0, len(comparison.missing)])
    summary.append(["Extra records", 0, len(comparison.extra)])
    summary.append(["Changed cells on shared records", 0, len(comparison.differences)])
    summary.append(["Duplicate keys", len(comparison.duplicate_original), len(comparison.duplicate_candidate)])
    summary.append(["Candidate report errors", 0, len(comparison.candidate_errors)])
    summary.append(["Record order matches", "", "", comparison.order_matches])
    summary.append(["Overall match", "", "", comparison.matches])
    original_totals = comparison.totals(comparison.expected)
    candidate_totals = comparison.totals(comparison.actual)
    for name, value in original_totals.items():
        summary.append(["SUM " + name, value, candidate_totals[name], candidate_totals[name] - value])
    for title, records in (("Missing records", comparison.missing), ("Extra records", comparison.extra)):
        ws = wb.create_sheet(title)
        ws.append(["Report row", "PO REF #", "UNIQUE CODE", "MATERIAL CODE", "MATERIAL DESCRIPTION",
                   "PO QTY", "SRV QTY", "UNSERVED QTY", "NET AMOUNT", "Original RAW rows",
                   "Selected source rows", "Selected source status and served qty", "Explanation"])
        for record in records:
            r = record.values
            original_rows = original_trace.get(record.key, [])
            selected_rows = selected_trace.get(record.key, [])
            explanation = ""
            if source:
                if title == "Missing records":
                    explanation = ("Absent from selected source; present in original RAW." if not selected_rows
                                   else "Present in selected source; inspect filtering or values.")
                else:
                    explanation = ("Present in selected source; absent from original RAW." if not original_rows
                                   else "Present in original RAW; inspect values or multiplicity.")
            ws.append([record.row, r[13], r[4], r[17], r[18], r[21], r[22], r[23], r[30],
                       ", ".join(str(n) for n, _, _ in original_rows),
                       ", ".join(str(n) for n, _, _ in selected_rows), str(selected_rows) if selected_rows else "", explanation])
    changed = wb.create_sheet("Changed cells")
    changed.append(["Original row", "Candidate row", "Column", "Field", "Original value", "Candidate value"])
    for difference in comparison.differences:
        changed.append(list(difference))
    duplicates = wb.create_sheet("Duplicate keys")
    duplicates.append(["Workbook", "PO REF #", "UNIQUE CODE", "MATERIAL CODE", "Count"])
    for label, values in (("Original", comparison.duplicate_original), ("Candidate", comparison.duplicate_candidate)):
        for key, count in values:
            duplicates.append([label, *key, count])
    errors = wb.create_sheet("Report errors")
    errors.append(["Sheet", "Cell", "Error"])
    for error in comparison.candidate_errors:
        errors.append(list(error))
    for ws in wb:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="242E48")
        for row in ws:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
        for column in ws.columns:
            width = min(65, max(14, max(len(str(c.value or "")) for c in column) + 2))
            ws.column_dimensions[column[0].column_letter].width = width
    try:
        wb.save(destination)
    finally:
        wb.close()


def main():
    parser = argparse.ArgumentParser(description="Compare MTC reports by PO REF # + UNIQUE CODE + MATERIAL CODE.")
    parser.add_argument("original", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--source", type=Path)
    parser.add_argument("--source-sheet")
    args = parser.parse_args()
    result = compare_reports(args.original, args.candidate)
    save_comparison(result, args.output, args.source, args.source_sheet)
    print(f"Match: {result.matches}; missing: {len(result.missing)}; extra: {len(result.extra)}; "
          f"changed cells: {len(result.differences)}; audit: {args.output}")


if __name__ == "__main__":
    main()
