# Batch File Toolkit — Split & Consolidate

A polished, single-page web tool that:

- **Splits** a large `.xlsx`, `.xls` or `.csv` file into fixed-size batches
  (you choose the rows per file), downloaded together as one ZIP.
- **Consolidates** many `.xlsx` / `.xls` / `.csv` files into a single file,
  with two merge modes and a column map so you can spot problems before
  downloading.

Everything runs **entirely in the browser**. There is no backend, nothing is
uploaded, and the tool works offline once the folder is on your machine.

---

## Quick start

| Option | How |
| --- | --- |
| Local | Double-click `index.html` (or open it in any modern browser) |
| Shared drive | Copy this folder anywhere your team can reach and share the path |
| Hosted | Serve this folder from any static web server (see Hosting below) |

No installation, no build step, no internet connection required.

## Split

1. Drag & drop (or browse for) one file.
2. If the workbook has several sheets, pick the sheet to use.
3. Review the first-10-rows preview.
4. Choose **rows per file** (default 500), untick any **columns** you want to
   drop (all are kept by default), and pick **CSV or Excel** output.
5. Click **Split & download ZIP** — you get
   `{name}_split.zip` containing `{name}_split1.csv`, `{name}_split2.csv`, …
   each with the header row included.

## Combine files (consolidate)

Both tools live on the same page — Split at the top, Combine below it.

1. Drag & drop any number of files — picking more files **adds** to the list.
2. Choose the merge mode:
   - **Match by header name** (default). Columns are aligned by their header
     text, not their position. If `file1` has columns `A,B` and `file2` has
     `B,A`, the data still lands in the right output columns. Headers are
     trimmed and matched case-insensitively ("UBI " ≡ "ubi"). The output
     contains the union of all columns; cells a file doesn't have stay blank.
   - **Simple stack (append rows)**. Rows are appended positionally with no
     remapping — for files that already share the same column layout. The
     header row comes from the first file. Files with a different column
     count are flagged; short rows are padded, wider files keep their extra
     data under blank headers.
3. Check the **column map** (which files contribute which columns — a quick
   way to catch a typo'd header) and the merged preview.
4. Pick CSV or Excel output, name the file, click **Combine & download**.

## Fidelity notes

- CSV input is read as plain text, so values like `007` keep their leading
  zeros end-to-end.
- Excel input keeps numbers as numbers and dates as dates; dates are written
  to CSV as `YYYY-MM-DD` (or `YYYY-MM-DD HH:MM:SS` when they have a time).
- CSV output uses a UTF-8 BOM and CRLF line endings so Excel opens it
  correctly, including non-ASCII text.
- Blank or duplicate headers are auto-named (`(Column C)`, `name (2)`) with a
  visible warning.

## Hosting

The folder is a self-contained static site — host it like any static asset:

- **Internal web server / IIS / nginx**: copy the folder into the web root.
  No server-side code, so any static host works.
- **GitHub Pages**: push this folder to a repository, then Settings → Pages →
  deploy from branch. Note: the tool's **URL becomes public** (the data you
  process still never leaves the browser).
- **Quick local server for testing**:
  `python -m http.server 8000` in this folder → http://127.0.0.1:8000

Because the two libraries are bundled locally in `lib/`, the site has zero
external dependencies at runtime and keeps working with no internet at all.

## Folder contents

```
index.html          the app
css/styles.css      light dashboard theme
js/core.js          pure data logic (parsing, batching, merging, output)
js/ui.js            DOM wiring, downloads, toasts
js/tests.js         test suite for core.js
lib/                pinned local libraries (no CDN calls)
tests.html          core-logic test runner — open directly in a browser
ui_verify.html      end-to-end UI test harness — needs a local http server
README.md           this file
```

## Pinned libraries

| Library | Version | Source |
| --- | --- | --- |
| SheetJS (xlsx.full.min.js) | 0.20.3 | cdn.sheetjs.com |
| JSZip | 3.10.1 | cdnjs.cloudflare.com |

## Running the tests

- **Core logic** (37 checks: batching edges, swapped-column merges,
  case/whitespace header variants, leading-zero preservation, stack-mode
  padding, XLSX round-trips): open `tests.html` — it reports
  `TESTS PASSED n/n`.
- **UI end-to-end** (13 checks driving the real page in an iframe: upload →
  preview → split, consolidate in both modes, file removal): serve the folder
  (`python -m http.server 8000`) and open
  http://127.0.0.1:8000/ui_verify.html.

## Limits

- Processing happens in browser memory. Typical batch files (up to a few
  hundred thousand rows) are fine; multi-hundred-MB exports may be slow or
  hit memory limits, and a warning appears for files over 50 MB.
- Legacy `.xls` files can be read but output is always `.csv` or `.xlsx`.
