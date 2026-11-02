from __future__ import annotations

import gc
import os
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .workbook import (PreparedData, Progress, WorkbookError, open_xlsx,
                       prepare_data, report_formulas, write_prepared)


@dataclass(frozen=True)
class GenerationResult:
    path: Path
    sheet: str
    input_rows: int
    output_rows: int
    excluded_rows: int


def recalculate_with_excel(path: Path) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise WorkbookError("Excel automation is unavailable. Install the app dependencies (pywin32).") from exc
    excel = workbook = None
    pythoncom.CoInitialize()
    try:
        try:
            # DispatchEx creates our own instance; never attach to the user's Excel.
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            raise WorkbookError("Cannot start Microsoft Excel. Install or repair desktop Excel, then try again.") from exc
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # Disable macros in files opened by this instance.
        try:
            workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
            if workbook.ReadOnly:
                raise WorkbookError("Excel opened the generated workbook read-only. Check folder permissions.")
            excel.Calculation = -4105  # Automatic
            workbook.ForceFullCalculation = False
            excel.CalculateFullRebuild()
            deadline = time.monotonic() + 60
            # Excel may retain xlPending (2) for the legacy hidden sheet's
            # broken references. Wait for active calculation, then verify every
            # generated formula cache, independently of that legacy sheet.
            while excel.CalculationState == 1:
                if time.monotonic() >= deadline:
                    raise WorkbookError("Excel did not finish calculating within 60 seconds. Please try again.")
                pythoncom.PumpWaitingMessages()
                time.sleep(0.1)
            workbook.Save()
        except WorkbookError:
            raise
        except Exception as exc:
            raise WorkbookError(f"Excel could not recalculate or save the output.\n{exc}") from exc
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            workbook = None
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                excel = None
                gc.collect()
                pythoncom.CoUninitialize()


def check_calculated_output(path: Path, data: PreparedData) -> None:
    wb = open_xlsx(path, data_only=True, read_only=True)
    try:
        errors = []
        for name, rows, columns in (("RAW", len(data.records), 38),
                                     ("OPEN TEMPLATE", len(data.retained), 42)):
            if rows == 0:
                continue
            for row in wb[name].iter_rows(min_row=2, max_row=rows + 1, max_col=columns):
                for cell in row:
                    if cell.data_type == "e":
                        errors.append(f"{name}!{cell.coordinate}: {cell.value}")
                        if len(errors) == 15:
                            break
                if len(errors) == 15:
                    break
        if errors:
            raise WorkbookError("Calculated output contains Excel errors:\n" + "\n".join(errors))
        if data.records and wb["RAW"]["A2"].value is None:
            raise WorkbookError("Excel did not save calculated warehouse values.")
        if data.retained and wb["OPEN TEMPLATE"]["AH2"].value is None:
            raise WorkbookError("Excel did not save calculated report values.")
        for row in wb["RAW"].iter_rows(min_row=2, max_row=len(data.records) + 1, max_col=1):
            if row[0].value is None:
                raise WorkbookError(f"Missing calculated value at RAW!{row[0].coordinate}.")
        if data.retained:
            formula_columns = [c for c, v in report_formulas(2, 2).items() if isinstance(v, str) and v.startswith("=")]
            # Stream through the cached report once; avoid repeated read-only
            # worksheet indexing, which reparses the XML for each cell.
            from openpyxl.utils import column_index_from_string
            indices = [column_index_from_string(c) - 1 for c in formula_columns]
            for row in wb["OPEN TEMPLATE"].iter_rows(min_row=2, max_row=len(data.retained) + 1, max_col=42):
                for index in indices:
                    if row[index].value is None:
                        raise WorkbookError(f"Missing calculated value at OPEN TEMPLATE!{row[index].coordinate}.")
    finally:
        wb.close()


def generate(raw_path: Path, sheet: str, output_folder: Path,
             reference_path: Path | None = None, progress: Progress | None = None) -> GenerationResult:
    notify = progress or (lambda value, message: None)
    notify(5, "Validating raw data and reference lookups…")
    data = prepare_data(raw_path, sheet, reference_path)
    notify(30, f"Validated {len(data.records):,} rows. Preparing workbook…")
    try:
        output_folder.mkdir(parents=True, exist_ok=True)
        filename = f"{raw_path.stem}_MTC_Output_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
        # Keep paths below Excel's practical filename limit even for long exports.
        if len(str(output_folder.resolve() / filename)) > 210:
            filename = f"MTC_Output_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
        target = output_folder / filename
        with tempfile.TemporaryDirectory(prefix=".mtc-", dir=output_folder) as staging:
            temporary = Path(staging) / "MTC Output.xlsx"
            write_prepared(data, temporary, reference_path)
            notify(60, "Calculating formulas in Microsoft Excel…")
            recalculate_with_excel(temporary)
            notify(90, "Checking calculated report…")
            check_calculated_output(temporary, data)
            # Windows rename refuses to overwrite an existing file.
            os.rename(temporary, target)
    except WorkbookError:
        raise
    except OSError as exc:
        raise WorkbookError(f"Cannot save the output. Check folder access and available disk space.\n{exc}") from exc
    notify(100, "Output ready.")
    return GenerationResult(target, sheet, len(data.records), len(data.retained), data.excluded)
