# Incident Report Parser

Condenses structured, template-based PDF reports into a single, easy-to-read
spreadsheet — one row per report page — pulling out only the cells that matter.
Point it at one PDF or a whole folder of them. Output is a print-ready Excel
`.xlsx` by default (CSV or JSON optional).

Extraction is **anchor-based**: each field is located by a fixed label on the page
(not by hard-coded coordinates), so it keeps working as long as the template's
labels are stable, and the set of extracted columns is a one-line change.

## Quick start

The easy way (no setup — [`uv`](https://docs.astral.sh/uv/) installs the
dependencies automatically the first time):

```bash
# a single PDF (which may contain many pages) -> incidents.xlsx
uv run parse_incidents.py report.pdf

# a whole folder of PDFs -> one combined workbook
uv run parse_incidents.py /path/to/pdfs -o incidents.xlsx

# other formats
uv run parse_incidents.py report.pdf -f csv -o out.csv
```

If the script is executable you can also run `./parse_incidents.py report.pdf`.
Prefer plain `pip`? `pip install -r requirements.txt` then
`python parse_incidents.py report.pdf`.

There's also a small web front-end (`web/app.py`, FastAPI) that accepts an upload,
parses it in memory, and returns the `.xlsx` — see the container/deployment notes.

## Privacy

Everything runs locally; the PDFs and the output never leave your machine. The
only network access is the one-off dependency download the first time `uv` runs.
If you need to share a parsing problem without sharing data, use `--inspect`: it
reports page structure and field *shapes* (lengths / character classes) with every
value masked, so the output contains no page content. `--diagnose` reports text-
layer health and prints no page content either.

### Sharing a full-page example safely — `sanitize_pdf.py`

When the structure itself needs to be seen (a new layout, a heading that isn't
being picked up), `sanitize_pdf.py` rebuilds a PDF with the **layout preserved but
the data masked**:

```
./sanitize_pdf.py real_report.pdf                # -> real_report.sanitized.pdf
./sanitize_pdf.py real_report.pdf --aggressive   # mask the labels too (wireframe)
```

It keeps word positions, table lines and the heading bars and preserves the fixed
template labels, but masks every value to its shape (`letters→x`, `digits→0`,
punctuation / `€` / spacing kept). Pages are rebuilt from scratch, so there's **no
hidden text layer underneath** — open the result, check it, then share it.

## How it works

- **Anchor-based fields.** Each extracted column is declared in a `FIELDS` registry
  near the bottom of `parse_incidents.py` as `(key, header, locator)`. A locator
  finds a fixed label and reads the cell relative to it. Adding a column is one
  line; no other code changes.
- **Page detection.** Cover/continuation pages are detected and skipped (listed so
  you can verify them); if a file matches none, every page is extracted instead, so
  detection can never cause silent data loss.
- **Nested grouping.** If the report groups rows under banner headings, those are
  read once and rendered as nested coloured dividers (and filled down as hidden
  group columns in CSV/JSON).
- **Print-first `.xlsx`.** The workbook has five sheets — Summary, the main rows,
  Legend, Discrepancies, and Skipped Pages — styled for printing/review and
  **ink-light**: solid fills mark only the noteworthy cells (flags, gaps,
  "yes" indicators), while headers and the nested dividers use coloured text and
  border rules on white, so printing spends toner only where something needs
  attention. Amber marks unreadable cells vs grey italics for recorded "no data";
  per-page hyperlinks lead back to the source PDF. CSV/JSON stay flat and keep the
  full column set.

### Printing

Each sheet is **print-ready**: open it in Excel/LibreOffice and *File → Print* (or
*Export as PDF*) — no separate export step. The main sheet is preset to A4,
landscape, fit-to-width with the header repeated per page.

## Options

```
uv run parse_incidents.py INPUT [INPUT ...] [options]

  -o, --output PATH     output file (default: incidents.xlsx / .csv / .json)
  -f, --format FMT      xlsx (default) | csv | json
  --debug               print every extracted field for each page
  --engine ENGINE       PDF text back-end: pdfplumber (default, pure-Python) or
                        pdfium (pypdfium2 — several times faster at ~the same
                        memory). Both feed identical extraction logic. Also
                        settable via the PARSER_ENGINE environment variable.
  --inspect             privacy-safe structural report (every value masked)
  --diagnose            check whether the PDF's text is actually extractable
  --self-test           parse the bundled template.pdf and check every *** cell
  --region-map FILE     station -> region code/name map, applied to Sec 19
                        incidents only (auto-detected next to the script if
                        present, encrypted preferred over plaintext; see
                        parse_region_map() for the file format). A station
                        with no match is flagged for review, not skipped.
  --encrypt-region-map [FILE]
                        encrypt FILE (default: region_stations.md) to
                        FILE.enc with a passphrase, then exit
```

## Adding more fields

1. Open `parse_incidents.py`, find the `FIELDS` list near the bottom.
2. Append one entry: `Field("key", "Column Header", locator_function)`.
3. Most fields are a one-line locator built from the existing helpers
   (`_table_cell`, the roles helpers, or `PageView.region` / `PageView.find`).

The new column flows through to xlsx/csv/json automatically.

## Notes & assumptions

- **Input must be text-based PDFs** (as exported from the source system), not scans
  or screenshots saved as PDF — the parser reads the text layer, not pixels. Check
  a suspect file with `--diagnose`; an image/scan must be OCR'd first.
- One row per report page; overflow/continuation pages are recognised and skipped
  rather than turned into blank rows.
- The parser never crashes on a bad page or missing cell — it logs a warning and
  leaves the field blank, so a large batch always produces output.
- **Validate on real data:** verified against the bundled template and rendered
  pages with realistic values, but real reports may have edge cases. Run `--debug`
  on a handful of real pages first and confirm the values look right.

## Security

Values from an uploaded PDF are treated as untrusted: they are neutralised against
spreadsheet/CSV formula injection before being written. The optional web service is
stateless (per-request temp dir, nothing persisted), enforces an upload size cap and
a concurrency limit, binds its Print action to server-produced output, sets strict
response security headers, and can verify the identity proxy's signed assertion
itself. The region map can be encrypted at rest (AES-256-GCM). No report data is ever
committed. See [SECURITY.md](SECURITY.md) for the full model, configuration, and a
deployment hardening checklist.
