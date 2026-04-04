# PDF Extraction Tool

This project extracts structured data from PDF documents and writes the results into Excel workbooks.

It supports:
- rule-based extraction for known table and figure patterns
- OCR fallback for scanned/image-based PDFs
- per-PDF extractor selection through `extract_requests.txt`
- lightweight history-based extractor recommendation
- human-in-the-loop feedback through reviewed Excel files

## Files

- [extract_production_process_table.py](d:\App\pdfscrapper\extract_production_process_table.py): main extraction script
- [extract_requests.txt](d:\App\pdfscrapper\extract_requests.txt): per-PDF extractor configuration
- [extraction_history.json](d:\App\pdfscrapper\extraction_history.json): learned historical pattern memory
- [feedback_history.json](d:\App\pdfscrapper\feedback_history.json): learned human feedback corrections

## Requirements

Install the main Python packages:

```bash
python -m pip install pdfplumber pypdfium2 openpyxl pytesseract
```

For scanned PDFs, install the Tesseract OCR engine as well.

Expected Windows install path:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

The script auto-detects that location.

## Supported Extractors

Current extractor names:

- `Production Process`
- `Table 1`
- `Table 3 ASTM`
- `Table 6`
- `CI-4 Hardness`
- `ISO VG 10`
- `Fig 1 Temperatures`
- `Product Tables`
- `Greases`
- `OXX`
- `ABB Greasing`

These names are the same names used in `extract_requests.txt` and as Excel sheet names.

## Basic Execution

Run with the default request file:

```bash
python extract_production_process_table.py
```

Run for a specific PDF:

```bash
python extract_production_process_table.py myfile.pdf
```

Run for a specific PDF and request file:

```bash
python extract_production_process_table.py myfile.pdf extract_requests.txt
```

Run while also passing reviewed Excel workbooks explicitly:

```bash
python extract_production_process_table.py myfile.pdf extract_requests.txt reviewed_output.xlsx
```

## Request File Format

The script reads extraction requests from `extract_requests.txt`.

You can define:

1. Global options for a single PDF:

```text
Greases, OXX
```

2. Per-PDF options for batch processing:

```text
testLubricant.pdf: Greases
prodTable.pdf: Product Tables, OXX
ABB.pdf: ABB Greasing
```

If no explicit request is available for a PDF, the history-based recommender may suggest extractors automatically.

## Output Files

For each PDF, the script creates one Excel workbook:

```text
<pdf_stem>_tables.xlsx
```

Examples:

- `testLubricant_tables.xlsx`
- `prodTable_tables.xlsx`
- `ABB_tables.xlsx`

If the workbook is open and locked, the script saves instead to:

```text
<pdf_stem>_tables_updated.xlsx
```

## Progress Output

The script prints progress to the console, including:

- selected extractor count
- start/completion message for each extractor
- row count per extractor
- total time taken

Example:

```text
PDF: D:\App\pdfscrapper\testLubricant.pdf
Selected extractors: 1
Starting Greases...
Completed Greases in 0.24s with 13 rows.
Extracted 13 rows across 1 sheets to: D:\App\pdfscrapper\testLubricant_tables.xlsx
Total time taken: 0.27s
```

## OCR Behavior

The script first tries native PDF text extraction where possible.

If a page is scanned or image-based, it falls back to OCR.

This is used for:

- scanned ABB regreasing cards
- image-heavy figure extraction
- scanned table lookups where text extraction is weak

## History-Based Recommendation

The script maintains a lightweight memory in `extraction_history.json`.

How it works:

1. Successful extractions store page snippets and extractor names.
2. When a new PDF has no explicit request entry, the script scores its pages against historical snippets.
3. The best-matching extractors are recommended and can be run automatically.

This is a lightweight ML-style similarity layer, not a trained neural network.

## Human Feedback Loop

Every output row now includes:

- `Review Field`
- `Human Evaluation`

### How to Review

After the workbook is generated:

1. Open the Excel file.
2. Review the extracted values.
3. Leave `Review Field` as-is unless you want to target a different column.
4. In `Human Evaluation`:
- write `Correct` if the extracted value is right
- write the corrected value if the extracted value is wrong

### Example

If a row looks like this:

| Production Process | Rt (mm) | Ra (mm) | Review Field | Human Evaluation |
|---|---|---|---|---|
| grinding | 2.00-6.0 | 0.400-0.8 | Ra (mm) | 0.750-3.5 |

then on the next run the script can learn that correction and apply it automatically.

### How Feedback Is Learned

The script reads reviewed workbooks from:

- explicitly passed `.xlsx` arguments
- any `*_tables.xlsx` files in the current folder
- any `*_tables_updated.xlsx` files in the current folder

It stores:

- confirmations in `feedback_history.json`
- corrections in `feedback_history.json`

Then, during future extraction runs, matching values are auto-corrected before writing the next workbook.

### Feedback Workflow

1. Run extraction:

```bash
python extract_production_process_table.py
```

2. Open the generated workbook and fill `Human Evaluation`.

3. Run the script again:

```bash
python extract_production_process_table.py
```

4. The script learns from the reviewed workbook and applies matching corrections.

### Notes

- `Correct` means the value is confirmed as-is.
- Any other non-empty `Human Evaluation` text is treated as the corrected value.
- Blank `Human Evaluation` means not reviewed yet.
- Corrections are safest when `Review Field` points to the exact column being reviewed.

## ABB Greasing Output

`ABB Greasing` writes one section per detected ABB regreasing card occurrence.

Each occurrence includes:

- bearings
- amount of grease
- greased in factory with
- the 4-column grease table

The sheet includes occurrence and page markers such as:

```text
Occurrence 1 | Page 1
Occurrence 2 | Page 2
```

## Troubleshooting

### `Skipping <name>: no details were found.`

This means the selected extractor did not find its expected pattern in that PDF.

### `No details were found.`

No selected extractor found usable content in the PDF.

### OCR is not working

Check:

- `pytesseract` is installed
- Tesseract OCR engine is installed
- `tesseract.exe` exists in a standard Windows path

### Workbook is locked

If the Excel file is open, the script writes to a fallback filename ending in `_updated.xlsx`.

## Recommended Usage Pattern

For repeated vendor PDFs:

1. Add the PDF name and extractor names to `extract_requests.txt`.
2. Run the script.
3. Review the workbook.
4. Fill `Human Evaluation`.
5. Re-run the script so it learns from the review.
6. Let the history and feedback files improve future extraction quality.
