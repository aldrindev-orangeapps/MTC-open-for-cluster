# MTC accuracy investigation ? 2 November 2026

## Root cause and evidence

The selected 173-record export and the original final report use different transaction datasets. The original workbook contains 215 source records in RAW B:AJ. All 44 report records missing from the generated report are absent from the selected export, and all 44 exist in the original RAW. The two additional generated records exist in the selected export and do not exist in the original RAW. This is a source dataset mismatch; the files alone cannot establish why the upstream SPIC export differs.

The previous confirmation of correctness applied only to converting the selected export. It did not establish equality with the original final report. The original is the acceptance authority in this investigation.

| Check | Selected export / previous report | Original / regenerated report |
|---|---:|---:|
| Input records | 173 | 215 |
| Zero-served records excluded | 9 | 9 |
| Report records | 164 | 206 |
| PO QTY | 711 | 893 |
| SRV QTY | 711 | 893 |
| UNSERVED QTY | 0 | 0 |
| FINAL QTY | 711 | 893 |
| NET AMOUNT | 2,317,175.63 | 2,196,884.22 |

The 162 shared records have zero changed cells across all 42 columns. Both original and generated reports have zero duplicate business keys. The original has 11 PO references absent from the prior generated report; the prior generated report has one PO reference absent from the original.

Quantity reconciliation: 711 - 80 (two extra records) + 262 (44 missing records) = 893.
Amount reconciliation: 2,317,175.63 - 603,238.75 + 482,947.34 = 2,196,884.22.

## Data flow and code changes

1. `mtc_tool/app.py::start_generation` passes the user-selected file and worksheet to `generator.generate`. The application does not retrieve SPIC data, connect to a database, execute SQL, paginate, or call an AI service. There is no upstream query or join layer in this repository to repair.
2. `mtc_tool/workbook.py::prepare_data` reads all selected source rows, preserves source order, validates numeric/date fields and warehouse lookups, and retains report records where SERVED QTY != 0. Both source datasets have the same nine zero-served exclusions. All 173 selected-export records and all 215 original-RAW records have status CREATED. No PO status predicate is applied; adding an OPEN-only predicate would contradict the original.
3. Customer/branch mapping uses BRANCH CODE against CENSUS B. Warehouse mapping uses SKU M:N. SKU DESCRIPTION maps against SKU B; material pricing uses the first matching code in SKU C and LPBT in E. Missing keys stop generation; lookups do not silently drop transaction rows. No aggregation, deduplication, row limit, or AI summarization occurs.
4. `workbook.write_prepared` transfers all selected source records to a fresh template and writes formulas for retained records. `generator.recalculate_with_excel` calculates them in a private Excel instance; `check_calculated_output` validates caches and report errors before publication.
5. Added `workbook.source_columns` to recognize the complete original RAW layout. A contains the calculated warehouse label; B:AJ holds the actual 35 source fields. The prior reader rejected the two WAREHOUSE headers. The new reader uses B for this recognized layout and continues rejecting ambiguous duplicates elsewhere. This removes an ingestion limitation; it does not manufacture records missing from an export.
6. Added `mtc_tool/reconcile.py::compare_reports`, `trace_source`, and `save_comparison`, plus the GUI Compare with Original button. Comparison uses PO REF # + UNIQUE CODE + MATERIAL CODE, preserves duplicate multiplicities, checks all 42 columns and all requested totals, and exports missing/extra source evidence.
7. Existing filtering, formulas, first-match mapping, rounding, report formatting, and ordinary raw-export support are preserved. No transaction identifiers or row counts are hardcoded into application logic.

## Calculation validation

All 11 requested quantity/price/amount fields match the original for every shared record before the fix and for every one of the 206 regenerated records. LPBT and CASE VAT EX correspond to the template headers LPBT CASE (Z) and VAT EX LPBT (AA).

| Field | Rule retained from original |
|---|---|
| PO QTY / SRV QTY / UNSERVED QTY | Corresponding source fields |
| FINAL QTY | PO QTY |
| LPBT CASE | First material-code match in SKU C, returning SKU E |
| VAT EX LPBT / CASE VAT EX | LPBT CASE ? FINAL QTY |
| LPAT | ROUND(LPBT CASE ? 1.12, 2) |
| VAT IN | VAT EX LPBT ? 1.12, without extra rounding |
| NET PRICE | Source NET PRICE |
| NET AMOUNT / NET AMOUNT VAT EX | Both use source VAT EX AMOUNT, matching the original formulas |

Numeric comparison tolerance is 0.000001 absolute / 1e-10 relative. No additional transaction rounding was introduced. Date formulas retain the original values while avoiding locale-sensitive DATEVALUE. Generated RAW and OPEN TEMPLATE have no Excel formula errors. The pre-existing hidden Sheet1 errors are not used for report generation.

## Exact previously missing records

Every record below is present in original RAW and absent from the selected export. The Before reconciliation.xlsx workbook includes its original RAW row number, the report row, description, quantities, amount, and source-presence explanation.

| Original report row | PO REF # | UNIQUE CODE | MATERIAL CODE | PO QTY | NET AMOUNT |
|---:|---|---|---|---:|---:|
| 132 | POMTCN261031004 | RO005-R018-SU | 12461722 | 3 | 1,338.06 |
| 133 | POMTCN261031004 | RO005-R018-SU | 12461725 | 3 | 1,338.06 |
| 134 | POMTCN261031004 | RO005-R018-SU | 12463271 | 3 | 1,605.87 |
| 135 | POMTCN261031005 | RO005-R019-SU | 69768551 | 27 | 56,137.86 |
| 136 | POMTCN261031005 | RO005-R019-SU | 69994913 | 2 | 9,760.22 |
| 137 | POMTCN261031005 | RO005-R019-SU | 801832 | 3 | 9,408.18 |
| 138 | POMTCN261031005 | RO005-R019-SU | 801855 | 1 | 2,846.06 |
| 139 | POMTCN261031006 | RO005-T031-SU | 69768551 | 12 | 24,950.16 |
| 140 | POMTCN261031006 | RO005-T031-SU | 69979478 | 28 | 41,312.60 |
| 141 | POMTCN261031006 | RO005-T031-SU | 69979481 | 4 | 5,901.80 |
| 142 | POMTCN261031006 | RO005-T031-SU | 69979488 | 3 | 4,426.35 |
| 143 | POMTCN261031006 | RO005-T031-SU | 69994913 | 7 | 34,160.77 |
| 144 | POMTCN261031006 | RO005-T031-SU | 801025 | 1 | 5,304.26 |
| 145 | POMTCN261031006 | RO005-T031-SU | 801026 | 1 | 5,304.26 |
| 146 | POMTCN261031006 | RO005-T031-SU | 801027 | 2 | 10,608.52 |
| 147 | POMTCN261031006 | RO005-T031-SU | 801832 | 6 | 18,816.36 |
| 148 | POMTCN261031006 | RO005-T031-SU | 801842 | 8 | 40,249.44 |
| 149 | POMTCN261031006 | RO005-T031-SU | 801855 | 6 | 17,076.36 |
| 150 | POMTCN261031007 | RO005-T031-SU | 12461705 | 13 | 5,798.26 |
| 151 | POMTCN261031007 | RO005-T031-SU | 12461722 | 32 | 14,272.64 |
| 152 | POMTCN261031007 | RO005-T031-SU | 12461725 | 30 | 13,380.60 |
| 153 | POMTCN261031007 | RO005-T031-SU | 12463266 | 16 | 7,136.32 |
| 154 | POMTCN261031008 | RO005-T031-SU | 69979478 | 3 | 4,426.35 |
| 155 | POMTCN261031008 | RO005-T031-SU | 69979481 | 4 | 5,901.80 |
| 156 | POMTCN261031008 | RO005-T031-SU | 69979488 | 3 | 4,426.35 |
| 157 | POMTCN261031008 | RO005-T031-SU | 69994913 | 3 | 14,640.33 |
| 158 | POMTCN261031008 | RO005-T031-SU | 801025 | 1 | 5,304.26 |
| 159 | POMTCN261031008 | RO005-T031-SU | 801026 | 1 | 5,304.26 |
| 160 | POMTCN261031008 | RO005-T031-SU | 801027 | 1 | 5,304.26 |
| 161 | POMTCN261031008 | RO005-T031-SU | 801028 | 1 | 5,304.26 |
| 162 | POMTCN261031008 | RO005-T031-SU | 801832 | 2 | 6,272.12 |
| 163 | POMTCN261031008 | RO005-T031-SU | 801842 | 2 | 10,062.36 |
| 164 | POMTCN261031008 | RO005-T031-SU | 801855 | 2 | 5,692.12 |
| 197 | POMTPG261031003 | RO005-R003-SU | 801832 | 1 | 3,136.06 |
| 198 | POMTPG261031004 | RO005-R003-SU | 801832 | 1 | 3,136.06 |
| 199 | POMTPG261031005 | RO005-R003-SU | 801832 | 1 | 3,136.06 |
| 200 | POMTPG261031006 | RO005-R003-SU | 801832 | 1 | 3,136.06 |
| 201 | POMTPG261031007 | RO005-R004-SU | 69768551 | 2 | 4,158.36 |
| 202 | POMTPG261031007 | RO005-R004-SU | 69994913 | 2 | 9,760.22 |
| 203 | POMTPG261031007 | RO005-R004-SU | 801832 | 1 | 3,136.06 |
| 204 | POMTPG261031007 | RO005-R004-SU | 801855 | 1 | 2,846.06 |
| 205 | POMTPG261031008 | RO005-R004-SU | 69768551 | 7 | 14,554.26 |
| 206 | POMTPG261031008 | RO005-R004-SU | 801832 | 3 | 9,408.18 |
| 207 | POMTPG261031008 | RO005-R004-SU | 801855 | 8 | 22,768.48 |

## Exact previously extra records

Both records below come from the selected export, have positive served quantities, and are absent from original RAW. They were included legitimately for that export; they are not invented records.

| Prior report row | PO REF # | UNIQUE CODE | MATERIAL CODE | PO QTY | NET AMOUNT |
|---:|---|---|---|---:|---:|
| 36 | POMTDL261030096 | SO009-B001-DS | 69610488 | 30 | 351,679.82 |
| 37 | POMTDL261030096 | SO009-B001-DS | 801842 | 50 | 251,558.93 |

## Regeneration and acceptance result

Used `MTC OUTPUT.xlsx` **RAW source records**, with its CENSUS/SKU reference data. The final OPEN TEMPLATE rows were rebuilt from those source fields and recalculated in Excel; they were not copied from the original final report.

Regenerated workbook: `C:\Users\Administrator\Documents\MTC Open Order\Outputs\MTC OUTPUT_MTC_Output_20261102_124010_741157.xlsx`.

| Acceptance check | Before | After |
|---|---:|---:|
| Missing records | 44 | 0 |
| Extra records | 2 | 0 |
| Changed cells on matched records | 0 | 0 |
| Report records | 164 | 206 |
| PO QTY | 711 | 893 |
| NET AMOUNT | 2,317,175.63 | 2,196,884.22 |
| Duplicate business keys | 0 | 0 |

After reconciliation.xlsx confirms equality across all 42 columns / 8,652 data cells, matching row order and all requested totals. The regenerated report matches the original.

Validation completed: all 21 automated tests passed, including desktop Excel acceptance tests, direct original-RAW ingestion, ordinary export compatibility, partial/zero-served cases, ambiguous headers, first-match lookups, and duplicate-aware reconciliation. The rebuilt EXE also passed the actual file-picker → original RAW → Generate Output → Compare with Original workflow. Packaged comparison.xlsx records that separate successful EXE run.

For future runs, use the complete source export for the intended reporting period. Selecting the old 173-record export still correctly produces 164 records; there is no evidence-based way to turn that source alone into the original 206-record report. The new comparison audit makes a dataset mismatch visible instead of treating successful generation as proof of equality.
