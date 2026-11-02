# MTC Open Order

Windows desktop tool for **Select Raw Excel → Generate Output → Open File**.
All processing stays on the computer. Microsoft Excel desktop must be installed
and working. The packaged EXE does not require Python.

## Use

1. Open `dist/MTC Open Order.exe`.
2. Click **Browse…** and select the customer open order `.xlsx` export, or a
   workbook containing its original `RAW` sheet.
3. If multiple worksheets match, select the one to process.
4. Click **Generate Output**, then **Open File** or **Open Folder**.

Reports go to `Documents\MTC Open Order\Outputs` by default. Each report has a
unique timestamp. Source files are never overwritten. The interface shows input,
report, and excluded row counts. You can open an output in Excel and edit RAW;
the report formulas still work. Row inclusion is determined at generation time;
generate again if records are added or served quantities change.

The selected source determines which transactions exist. An export with 173
records cannot reproduce an original whose RAW sheet contains 215 different
records. To reproduce an existing final report, select the workbook containing
its complete RAW source and choose `RAW`. Only RAW source fields are imported;
OPEN TEMPLATE report rows are never copied or merged into the selected source.
If that workbook has updated lookup data, also select it in reference Settings.

After generation, **Compare with Original…** compares all 42 report columns,
matches records using **PO REF # + UNIQUE CODE + MATERIAL CODE**, and saves an
Excel audit containing missing/extra records, changed cells, duplicates, totals,
and their original/selected raw source row evidence. A successful generation
means the selected source was converted; use this comparison to establish that
it also matches a particular original final report.

**Settings** lets you change the output folder or select an updated reference
workbook containing `CENSUS`, `SKU`, and `SKU MASTER LIST` in the original layout.
Leave the reference field blank to use the bundled references. Settings are saved
under `%LOCALAPPDATA%\MTCOpenOrder\settings.json`.

Missing warehouse, customer, or SKU lookups stop generation. The issue window
shows source row numbers, can export a CSV, and links to reference settings.
Update the selected workbook and retry. Keep the reference workbook saved before
processing; cached values in reference formulas must be current.

## Confirmed conversion rules

- Raw source headers are on row 1. All 35 fields are required; header order may
  change. Duplicate required headers are rejected except for the complete known
  embedded RAW layout: its A column is the calculated warehouse label and B:AJ
  holds the 35 original fields. In that layout B supplies the warehouse key;
  A's lookup formula is ignored. Source formulas in imported fields are rejected.
- All raw records go to RAW B:AJ. RAW A uses the warehouse lookup in SKU M:N.
- OPEN TEMPLATE retains records with **SERVED QTY != 0**, including negative
  quantities, in source order. No grouping or deduplication is applied.
- Customer lookups use BRANCH CODE against CENSUS B; SKU description uses SKU B.
  Exact VLOOKUP uses the first match, ignoring letter case. Identifiers retain
  their original text/number types, including leading zeros.
- **FINAL QTY = PO QTY**. Delivery date comes from **CANCEL DATE**. Both net
  amount columns come from **VAT EX AMOUNT**. UOM is `CASE`; KAM is SALES REP.
- Price uses the first matching material code in SKU C, returning LPBT from E.
  VAT calculations use 12%, with rounding only where the sample uses ROUND.
- Raw dates are normalized to MM/DD/YYYY. Date formulas use explicit month, day,
  and year components instead of locale-sensitive DATEVALUE, producing the same
  dates on computers whose regional settings use DD/MM/YYYY.
- Blank reference return fields retain Excel's zero/blank behavior. Missing
  lookup keys stop processing rather than silently becoming zero.
- Generated reports preserve template formatting and working formulas. Filters
  are reset and resized; no previous report records are carried over.
- The hidden `Sheet1` retains the template's pre-existing broken formulas; it is
  not a report source. Formula-error checks cover generated RAW and OPEN TEMPLATE.

The bundled workbook uses the original template layout and the three reference
sheets from `MTC OUTPUT.xlsx`. No sample transaction rows are bundled. An unused,
broken external defined name (`A1341.` referencing `[1]RAW MAROSE`) was removed:
Excel rejects openpyxl-saved files that keep that name after external links are
removed. Reference updates also discard external workbook links; use self-contained
reference data.

## Development

Tested on Windows with Python 3.14 and Microsoft Excel desktop.

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements-dev.txt
.\.venv\Scripts\python mtc_open_order.py
```

Build the standalone, windowed EXE:

```powershell
.\.venv\Scripts\python -m PyInstaller --noconfirm 'MTC Open Order.spec'
```

Core tests:

```powershell
.\.venv\Scripts\python -m unittest discover -s tests -v
```

Acceptance tests launch their own Excel instances. Put the supplied raw and
`MTC OUTPUT.xlsx` in Downloads, or set `MTC_FIXTURE_DIR` to their folder:

```powershell
$env:MTC_EXCEL_TESTS = '1'
.\.venv\Scripts\python -m unittest discover -s tests -v
```

Acceptance includes 173 → 164 supplied raw records and all 42 columns of the
215 → 206 sample regression, with numeric tolerance for Excel floating-point
values. Generated files are recalculated in a private Excel instance, checked
for report formula errors, and published only after those checks pass.

You can also save a reconciliation without the GUI:

```powershell
.\.venv\Scripts\python -m mtc_tool.reconcile ORIGINAL.xlsx GENERATED.xlsx --output audit.xlsx --source RAW_EXPORT.xlsx
```

`--source` is optional; it adds evidence showing whether missing/extra keys exist
in the selected export and in the original workbook's RAW sheet. With source
tracing, the original must include its RAW and SKU sheets. Inputs are read-only.
Duplicate records are compared with their multiplicities rather than silently
deduplicated. Monetary differences use a 0.000001 absolute / 1e-10 relative
tolerance; the report retains the original calculation and rounding rules.
