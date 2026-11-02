from __future__ import annotations

import csv
import math
import re
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.filters import AutoFilter

from .paths import bundled_template

RAW_HEADERS = (
    "WAREHOUSE", "PO DATE", "PO REF #", "PO STATUS", "SOURCE PO #", "DEL DATE",
    "CANCEL DATE", "VAT TYPE", "VAT RATE", "VAT PRICE", "BUSINESS UNIT",
    "TRADE GROUP", "CUSTOMER GROUP", "CUSTOMER CODE", "ACCT CUSTOMER CODE",
    "CUSTOMER NAME", "BRANCH CODE", "ACCT BRANCH CODE", "BRANCH NAME", "BRAND",
    "SKU CODE", "BARCODE", "SKU DESCRIPTION", "PACKAGING", "PCS IN PKG", "PO QTY",
    "SERVED QTY", "UNSERVED QTY", "UNIT PRICE", "NET PRICE", "NET AMOUNT",
    "VAT EX AMOUNT", "CREATED BY", "CREATED DATE", "SALES REP",
)
REFERENCE_SHEETS = ("CENSUS", "SKU", "SKU MASTER LIST")
NUMERIC_FIELDS = ("VAT RATE", "PCS IN PKG", "PO QTY", "SERVED QTY", "UNSERVED QTY",
                  "UNIT PRICE", "NET PRICE", "NET AMOUNT", "VAT EX AMOUNT")
DATE_FIELDS = ("PO DATE", "DEL DATE", "CANCEL DATE")
Progress = Callable[[int, str], None]


class WorkbookError(Exception):
    """An actionable error that can be shown directly to the user."""


@dataclass(frozen=True)
class Issue:
    sheet: str
    row: int
    field: str
    value: str
    message: str


class ValidationError(WorkbookError):
    def __init__(self, issues: list[Issue]):
        self.issues = issues
        super().__init__(f"Found {len(issues)} issue(s). Update the raw or reference workbook, then try again.")


@dataclass(frozen=True)
class Record:
    source_row: int
    values: tuple

    def get(self, header: str):
        return self.values[RAW_HEADERS.index(header)]


@dataclass(frozen=True)
class PreparedData:
    sheet: str
    records: tuple[Record, ...]
    retained: tuple[int, ...]  # Zero-based indices into records, preserving source order.

    @property
    def excluded(self) -> int:
        return len(self.records) - len(self.retained)


def open_xlsx(path: Path, *, data_only: bool = False, read_only: bool = False):
    if path.suffix.lower() != ".xlsx":
        raise WorkbookError("Please select an .xlsx Excel workbook.")
    try:
        return load_workbook(path, data_only=data_only, read_only=read_only, keep_links=False)
    except Exception as exc:
        raise WorkbookError(f"Cannot read {path.name}. Check that it is a valid, unencrypted .xlsx file.\n{exc}") from exc


def normalized_header(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).upper()


def header_columns(ws) -> dict[str, list[int]]:
    columns: dict[str, list[int]] = {}
    for cell in next(ws.iter_rows(min_row=1, max_row=1), ()):
        name = normalized_header(cell.value)
        if name:
            columns.setdefault(name, []).append(cell.column)
    return columns


def source_columns(ws) -> dict[str, int]:
    """Resolve an export, or the original workbook's RAW B:AJ source block.

    In the latter layout column A is a calculated warehouse label, while B
    contains the original warehouse key. Only this complete, recognized layout
    may disambiguate the duplicate WAREHOUSE header; never guess for other files.
    """
    columns = header_columns(ws)
    missing = [h for h in RAW_HEADERS if h not in columns]
    if missing:
        raise WorkbookError("Missing raw columns: " + ", ".join(missing))
    headers = [normalized_header(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))]
    embedded_layout = headers[:36] == ["WAREHOUSE", *RAW_HEADERS]
    duplicates = [h for h in RAW_HEADERS if len(columns[h]) > 1
                  and not (h == "WAREHOUSE" and embedded_layout and columns[h] == [1, 2])]
    if duplicates:
        raise WorkbookError("Ambiguous duplicate raw columns: " + ", ".join(duplicates))
    return {h: (2 if h == "WAREHOUSE" and embedded_layout else columns[h][0]) for h in RAW_HEADERS}


def discover_sheets(path: Path) -> list[str]:
    wb = open_xlsx(path, read_only=True)
    try:
        result = []
        for ws in wb:
            columns = header_columns(ws)
            if all(h in columns for h in RAW_HEADERS):
                result.append(ws.title)
        return result
    finally:
        wb.close()


def numeric(value) -> float:
    if isinstance(value, bool) or value is None:
        raise ValueError("A number is required")
    result = float(value)
    if not math.isfinite(result):
        raise ValueError("A finite number is required")
    return result


def date_text(value) -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%m/%d/%Y")
    if not isinstance(value, str):
        raise ValueError("Use an Excel date or MM/DD/YYYY")
    return datetime.strptime(value.strip(), "%m/%d/%Y").strftime("%m/%d/%Y")


def lookup_key(value):
    # Excel distinguishes numeric keys from text keys, including leading zeros.
    if isinstance(value, str):
        return ("text", value.casefold())
    if value is None:
        return ("blank", "")
    return ("number", value)


def first_lookup(ws, key_column: int, start_row: int) -> dict:
    result = {}
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        key = row[key_column - 1] if len(row) >= key_column else None
        if key is not None:
            result.setdefault(lookup_key(key), row)
    return result


def validate_reference(wb) -> None:
    missing = [name for name in REFERENCE_SHEETS if name not in wb.sheetnames]
    if missing:
        raise WorkbookError("Reference workbook is missing sheet(s): " + ", ".join(missing))
    checks = {"CENSUS": (3, {2: "IMS BRANCH CODE", 7: "TRADE GROUP", 14: "HUB"}),
              "SKU": (3, {2: "GENERATE NAME", 3: "SKU", 5: "LPBT", 13: "MT01 DELCO PASIG"}),
              "SKU MASTER LIST": (2, {1: "BRAND", 2: "SKU"})}
    for name, (row, fields) in checks.items():
        for col, expected in fields.items():
            # Warehouse mappings start on row 3, but their first item may change.
            if name == "SKU" and col == 13:
                continue
            if normalized_header(wb[name].cell(row, col).value) != expected:
                raise WorkbookError(f"Reference layout changed: {name}!{get_column_letter(col)}{row} must be {expected}.")


def prepare_data(raw_path: Path, sheet: str, reference_path: Path | None = None) -> PreparedData:
    raw_wb = open_xlsx(raw_path, read_only=True)
    ref_wb = None
    try:
        if sheet not in raw_wb.sheetnames:
            raise WorkbookError(f"Worksheet not found: {sheet}")
        ws = raw_wb[sheet]
        columns = source_columns(ws)
        indices = [columns[h] - 1 for h in RAW_HEADERS]
        ref_wb = open_xlsx(reference_path or bundled_template(), data_only=True, read_only=True)
        validate_reference(ref_wb)
        warehouses = first_lookup(ref_wb["SKU"], 13, 3)
        branches = first_lookup(ref_wb["CENSUS"], 2, 4)
        skus = first_lookup(ref_wb["SKU"], 2, 4)
        prices = first_lookup(ref_wb["SKU"], 3, 4)
        records, retained, issues = [], [], []
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            values = [row[i] if i < len(row) else None for i in indices]
            if all(v is None or v == "" for v in values):
                continue
            for h in NUMERIC_FIELDS:
                i = RAW_HEADERS.index(h)
                try:
                    values[i] = numeric(values[i])
                except (ValueError, TypeError, OverflowError):
                    issues.append(Issue(sheet, row_num, h, str(values[i]), "Enter a finite number."))
            for h in DATE_FIELDS:
                i = RAW_HEADERS.index(h)
                try:
                    values[i] = date_text(values[i])
                except (ValueError, TypeError):
                    issues.append(Issue(sheet, row_num, h, str(values[i]), "Use MM/DD/YYYY or an Excel date."))
            # Reject formula input rather than trusting potentially stale cached values.
            for h, value in zip(RAW_HEADERS, values):
                if isinstance(value, str) and value.startswith("="):
                    issues.append(Issue(sheet, row_num, h, value, "Raw fields must contain values, not formulas."))
            record = Record(row_num, tuple(values))
            if lookup_key(record.get("WAREHOUSE")) not in warehouses:
                issues.append(Issue(sheet, row_num, "WAREHOUSE", str(record.get("WAREHOUSE")), "Missing in SKU columns M:N."))
            try:
                keep = numeric(record.get("SERVED QTY")) != 0
            except (ValueError, TypeError, OverflowError):
                keep = False
            if keep:
                for field, lookup, location in (("BRANCH CODE", branches, "CENSUS column B"),
                                                 ("SKU DESCRIPTION", skus, "SKU column B")):
                    if lookup_key(record.get(field)) not in lookup:
                        issues.append(Issue(sheet, row_num, field, str(record.get(field)), f"Missing in {location}."))
                sku_row = skus.get(lookup_key(record.get("SKU DESCRIPTION")))
                if sku_row:
                    price_row = prices.get(lookup_key(sku_row[2]))
                    if price_row is None:
                        issues.append(Issue(sheet, row_num, "MATERIAL CODE", str(sku_row[2]), "Missing in SKU column C."))
                    elif price_row[4] is not None:
                        try:
                            numeric(price_row[4])
                        except (ValueError, TypeError, OverflowError):
                            issues.append(Issue(sheet, row_num, "LPBT", str(price_row[4]), "Reference LPBT must be numeric or blank."))
                retained.append(len(records))
            records.append(record)
        if not records:
            raise WorkbookError("The selected worksheet has no raw data records.")
        if issues:
            raise ValidationError(issues)
        return PreparedData(sheet, tuple(records), tuple(retained))
    finally:
        raw_wb.close()
        if ref_wb:
            ref_wb.close()


def copy_reference_sheets(source, destination) -> None:
    """Copy cells and styles across workbooks; never copy report rows from the sample."""
    validate_reference(source)
    for name in REFERENCE_SHEETS:
        old = destination[name]
        index = destination.sheetnames.index(name)
        destination.remove(old)
        target = destination.create_sheet(name, index)
        ws = source[name]
        for row in ws:
            for cell in row:
                other = target.cell(cell.row, cell.column, cell.value)
                if cell.has_style:
                    other.font = copy(cell.font)
                    other.fill = copy(cell.fill)
                    other.border = copy(cell.border)
                    other.alignment = copy(cell.alignment)
                    other.number_format = cell.number_format
                    other.protection = copy(cell.protection)
                if cell.hyperlink:
                    other.hyperlink = copy(cell.hyperlink)
                if cell.comment:
                    other.comment = copy(cell.comment)
        for key, dimension in ws.column_dimensions.items():
            target.column_dimensions[key] = copy(dimension)
        for key, dimension in ws.row_dimensions.items():
            target.row_dimensions[key] = copy(dimension)
        for merged in ws.merged_cells.ranges:
            target.merge_cells(str(merged))
        target.sheet_format = copy(ws.sheet_format)
        target.sheet_properties = copy(ws.sheet_properties)
        target.sheet_state = ws.sheet_state
        target.freeze_panes = ws.freeze_panes
        target.auto_filter = copy(ws.auto_filter)
        target.sheet_view.showGridLines = ws.sheet_view.showGridLines


def remove_obsolete_external_names(wb) -> None:
    # The supplied template contains A1341., referring to a missing external
    # workbook ('[1]RAW MAROSE'!#REF!). Retaining that name after removing links
    # makes Excel reject the entire file, even though no report formula uses it.
    for name, definition in list(wb.defined_names.items()):
        if re.search(r"\[\d+\]", definition.attr_text or ""):
            del wb.defined_names[name]


def report_formulas(output_row: int, raw_row: int) -> dict[str, object]:
    r, s = output_row, raw_row
    return {
        "A": f'=IFERROR(VLOOKUP($E{r},CENSUS!$B:$H,6,FALSE),0)',
        "B": f'=RAW!L{s}', "C": f'=RAW!A{s}',
        "D": f'=IFERROR(VLOOKUP($E{r},CENSUS!$B:$E,4,FALSE),0)',
        "E": f'=RAW!R{s}',
        "F": f'=IFERROR(VLOOKUP($E{r},CENSUS!$B:$E,3,FALSE),0)',
        "G": f'=RAW!N{s}', "H": f'=RAW!E{s}', "I": f'=RAW!AH{s}',
        "M": f'=RAW!C{s}', "N": f'=RAW!D{s}', "O": f'=RAW!X{s}',
        "P": f'=IFERROR(VLOOKUP($O{r},SKU!$B:$J,8,FALSE),0)',
        "Q": f'=IFERROR(VLOOKUP($O{r},SKU!$B:$J,9,FALSE),0)',
        "R": f'=IFERROR(VLOOKUP($O{r},SKU!$B:$J,2,FALSE),0)',
        "S": f'=IFERROR(VLOOKUP($O{r},SKU!$B:$J,3,FALSE),0)',
        "U": "CASE", "V": f'=RAW!AA{s}', "W": f'=RAW!AB{s}', "X": f'=RAW!AC{s}',
        "Y": f'=V{r}', "Z": f'=IFERROR(VLOOKUP($R{r},SKU!$C:$F,3,FALSE),0)',
        "AA": f'=Z{r}*Y{r}', "AB": f'=ROUND(Z{r}*1.12,2)', "AC": f'=AA{r}*1.12',
        "AD": f'=RAW!AE{s}', "AE": f'=RAW!AG{s}', "AF": f'=RAW!AG{s}',
        # DATEVALUE depends on the user's Windows/Excel date locale. The raw
        # dates are normalized to MM/DD/YYYY, so parse their components explicitly.
        "AG": f'=RAW!C{s}', "AH": f'=DATE(RIGHT(AG{r},4),LEFT(AG{r},2),MID(AG{r},4,2))',
        "AI": f'=RAW!H{s}', "AJ": f'=DATE(RIGHT(AI{r},4),LEFT(AI{r},2),MID(AI{r},4,2))', "AK": f'=RAW!F{s}',
        "AL": f'=IFERROR(VLOOKUP($E{r},CENSUS!$B:$N,11,FALSE),0)',
        "AM": f'=IFERROR(VLOOKUP($E{r},CENSUS!$B:$N,12,FALSE),0)',
        "AN": f'=IFERROR(VLOOKUP($E{r},CENSUS!$B:$N,13,FALSE),0)',
        "AO": f'=RAW!AJ{s}',
    }


def write_prepared(data: PreparedData, destination: Path, reference_path: Path | None = None) -> None:
    wb = open_xlsx(bundled_template())
    reference = None
    try:
        remove_obsolete_external_names(wb)
        if reference_path:
            reference = open_xlsx(reference_path)
            copy_reference_sheets(reference, wb)
        raw, report = wb["RAW"], wb["OPEN TEMPLATE"]
        # Capture the styled template row, then remove the entire old data region.
        raw_styles = [copy(raw.cell(2, c)._style) for c in range(1, 39)]
        report_styles = [copy(report.cell(2, c)._style) for c in range(1, 43)]
        raw_height, report_height = raw.row_dimensions[2].height, report.row_dimensions[2].height
        for ws in (raw, report):
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
            for key in list(ws.row_dimensions):
                if key > 1:
                    del ws.row_dimensions[key]
            ws.auto_filter = AutoFilter()
            ws.sheet_properties.filterMode = False
        for r, record in enumerate(data.records, 2):
            for col, style in enumerate(raw_styles, 1):
                raw.cell(r, col)._style = copy(style)
            raw.row_dimensions[r].height = raw_height
            raw.cell(r, 1, f'=IFERROR(VLOOKUP($B{r},SKU!M:N,2,FALSE),0)')
            for c, value in enumerate(record.values, 2):
                cell = raw.cell(r, c, value)
                if isinstance(value, str):
                    cell.data_type = "s"  # Preserve codes and never execute source text.
            raw.cell(r, 2).number_format = "@"
        for r, index in enumerate(data.retained, 2):
            for c, style in enumerate(report_styles, 1):
                report.cell(r, c)._style = copy(style)
            report.row_dimensions[r].height = report_height
            for column, value in report_formulas(r, index + 2).items():
                report[f'{column}{r}'] = value
        raw.auto_filter.ref = f'A1:AL{max(1, len(data.records) + 1)}'
        report.auto_filter.ref = f'A1:AP{max(1, len(data.retained) + 1)}'
        report.print_area = f'A1:AP{max(1, len(data.retained) + 1)}'
        wb.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True)
        wb.active = wb.sheetnames.index("OPEN TEMPLATE")
        wb.save(destination)
    finally:
        wb.close()
        if reference:
            reference.close()


def export_issues(issues: list[Issue], path: Path) -> None:
    # Prevent spreadsheet formula execution when a user opens the CSV in Excel.
    def safe(value):
        text = str(value)
        return "'" + text if text.lstrip().startswith(("=", "+", "-", "@")) else text
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(("Sheet", "Source row", "Field", "Value", "Issue"))
        for issue in issues:
            writer.writerow([safe(v) for v in (issue.sheet, issue.row, issue.field, issue.value, issue.message)])
