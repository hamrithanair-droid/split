/*
 * core.js — pure data logic for the Batch File Toolkit (Split & Consolidate).
 *
 * No DOM access here. Depends only on SheetJS (XLSX). Wrapped as UMD so the
 * same file runs in the browser (window.Core) and in Node for tests.
 *
 * Data model used throughout:
 *   table = { name: string, headers: string[], rows: any[][] }
 *   - headers are cleaned display names (trimmed, blanks named, dupes suffixed)
 *   - rows are arrays aligned to headers; cells hold typed values
 *     (string | number | boolean | Date | null)
 */
(function (root, factory) {
  if (typeof module === "object" && typeof module.exports === "object") {
    module.exports = factory(require("../lib/xlsx.full.min.js"));
  } else {
    root.Core = factory(root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX) {
  "use strict";

  var VERSION = "1.0.0";

  /* ------------------------------------------------------------------ *
   * Cell + header helpers
   * ------------------------------------------------------------------ */

  function pad2(n) {
    return n < 10 ? "0" + n : "" + n;
  }

  /* Render any cell value as display text. Numbers keep full precision
     (no scientific notation for normal ids), dates become ISO-ish text,
     strings pass through untouched (so "007" stays "007"). */
  function formatCell(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) {
      if (isNaN(value.getTime())) return "";
      var base =
        value.getFullYear() + "-" + pad2(value.getMonth() + 1) + "-" + pad2(value.getDate());
      if (value.getHours() === 0 && value.getMinutes() === 0 && value.getSeconds() === 0) {
        return base;
      }
      return (
        base + " " + pad2(value.getHours()) + ":" + pad2(value.getMinutes()) + ":" + pad2(value.getSeconds())
      );
    }
    if (typeof value === "number") {
      if (!isFinite(value)) return "";
      return String(value);
    }
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return String(value);
  }

  /* Display form of a header cell: rendered as text and trimmed. */
  function normalizeHeader(value) {
    return formatCell(value).trim();
  }

  /* Canonical key used to decide whether two headers are "the same
     column": trimmed, lower-cased, internal whitespace collapsed. */
  function canonKey(header) {
    return normalizeHeader(header).toLowerCase().replace(/\s+/g, " ");
  }

  function columnLetter(index) {
    return XLSX.utils.encode_col(index);
  }

  /* Clean a raw header row:
     - trim whitespace
     - blank headers become "(Column A)" style names
     - duplicates (case/space-insensitive) become "Name (2)", "Name (3)", ...
     `width` lets callers force extra unnamed columns when data rows are
     wider than the header row. */
  function cleanHeaders(rawHeaderRow, width) {
    var headers = [];
    var warnings = [];
    var used = new Set();
    var n = Math.max(width || 0, rawHeaderRow.length);
    for (var i = 0; i < n; i++) {
      var name = normalizeHeader(rawHeaderRow[i]);
      if (name === "") {
        name = "(Column " + columnLetter(i) + ")";
        warnings.push('A blank header in column ' + columnLetter(i) + ' was named "' + name + '".');
      }
      var key = canonKey(name);
      if (used.has(key)) {
        var counter = 2;
        var candidate, candidateKey;
        do {
          candidate = name + " (" + counter + ")";
          candidateKey = canonKey(candidate);
          counter++;
        } while (used.has(candidateKey));
        warnings.push('Duplicate header "' + name + '" was renamed to "' + candidate + '".');
        name = candidate;
        key = candidateKey;
      }
      used.add(key);
      headers.push(name);
    }
    return { headers: headers, warnings: warnings };
  }

  /* ------------------------------------------------------------------ *
   * Reading files
   * ------------------------------------------------------------------ */

  function extensionOf(filename) {
    var match = /\.([^.\\\/]+)$/.exec(String(filename || ""));
    return match ? match[1].toLowerCase() : "";
  }

  /* Parse raw file bytes (ArrayBuffer/Uint8Array) or a string into a
     SheetJS workbook.
     - CSV/TSV/TXT: parsed as plain text with raw:true so values keep
       their exact text ("007" is not turned into 7).
     - XLSX/XLS: parsed with typed cells; date cells become JS Dates. */
  function readWorkbook(data, filename) {
    var ext = extensionOf(filename);
    if (ext === "csv" || ext === "txt" || ext === "tsv") {
      var text;
      if (typeof data === "string") {
        text = data;
      } else {
        var bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
        text = new TextDecoder("utf-8").decode(bytes);
      }
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
      return XLSX.read(text, { type: "string", raw: true, dense: true });
    }
    var bin = data instanceof Uint8Array ? data : new Uint8Array(data);
    return XLSX.read(bin, { type: "array", cellDates: true, dense: true });
  }

  /* Extract a rectangular table from one sheet of a workbook.
     First non-blank row is the header row; blank data rows are skipped. */
  function tableFromSheet(workbook, sheetName) {
    var ws = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
    if (!ws) return { headers: [], rows: [], warnings: ['Sheet "' + sheetName + '" was not found.'] };
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
    if (!aoa.length) return { headers: [], rows: [], warnings: [] };

    var width = 0;
    for (var i = 0; i < aoa.length; i++) width = Math.max(width, aoa[i].length);

    var cleaned = cleanHeaders(aoa[0], width);
    var rows = new Array(aoa.length - 1);
    for (var r = 1; r < aoa.length; r++) {
      var src = aoa[r];
      var row = new Array(width).fill(null);
      for (var c = 0; c < src.length && c < width; c++) {
        row[c] = src[c] === undefined ? null : src[c];
      }
      rows[r - 1] = row;
    }
    return { headers: cleaned.headers, rows: rows, warnings: cleaned.warnings };
  }

  /* ------------------------------------------------------------------ *
   * Split
   * ------------------------------------------------------------------ */

  /* Plan the batches: returns [{start, count}, ...].
     Throws a friendly error when batchSize is not a whole number >= 1. */
  function makeBatches(totalRows, batchSize) {
    var size = Number(batchSize);
    if (!isFinite(size) || Math.floor(size) !== size || size < 1) {
      throw new Error("Rows per file must be a whole number of 1 or more.");
    }
    var batches = [];
    for (var start = 0; start < totalRows; start += size) {
      batches.push({ start: start, count: Math.min(size, totalRows - start) });
    }
    return batches;
  }

  /* Keep only the given column indices (in the given order). */
  function selectColumns(table, keepIndices) {
    var headers = keepIndices.map(function (i) {
      return table.headers[i];
    });
    var rows = table.rows.map(function (row) {
      return keepIndices.map(function (i) {
        return i < row.length ? row[i] : null;
      });
    });
    return { headers: headers, rows: rows };
  }

  /* ------------------------------------------------------------------ *
   * Consolidate
   * ------------------------------------------------------------------ */

  function listNames(names) {
    if (names.length <= 3) return names.join(", ");
    return names.slice(0, 3).join(", ") + " and " + (names.length - 3) + " more";
  }

  /* Mode 1 — match by header name.
     Output columns are the union of all headers (first-seen order and
     casing). Every row is remapped by header text, so column position in
     the source files never matters. Cells for columns a file does not
     have stay null (blank). */
  function mergeByHeader(tables) {
    var outHeaders = [];
    var keyToIndex = new Map();
    var variants = new Map(); // canon key -> Set of display spellings seen

    tables.forEach(function (t) {
      t.headers.forEach(function (h) {
        var disp = normalizeHeader(h);
        var key = canonKey(disp);
        if (!keyToIndex.has(key)) {
          keyToIndex.set(key, outHeaders.length);
          outHeaders.push(disp);
        }
        if (!variants.has(key)) variants.set(key, new Set());
        variants.get(key).add(disp);
      });
    });

    // Per table: output column index -> source column index (or -1).
    var maps = tables.map(function (t) {
      var m = new Array(outHeaders.length).fill(-1);
      t.headers.forEach(function (h, j) {
        var idx = keyToIndex.get(canonKey(h));
        if (m[idx] === -1) m[idx] = j;
      });
      return m;
    });

    var rows = [];
    var perFile = [];
    tables.forEach(function (t, ti) {
      var m = maps[ti];
      t.rows.forEach(function (src) {
        var row = new Array(outHeaders.length).fill(null);
        for (var c = 0; c < outHeaders.length; c++) {
          var j = m[c];
          if (j >= 0 && j < src.length) row[c] = src[j];
        }
        rows.push(row);
      });
      perFile.push({ name: t.name, rows: t.rows.length, columns: t.headers.length });
    });

    var columnMap = outHeaders.map(function (h, c) {
      return {
        header: h,
        presentIn: tables.map(function (_t, ti) {
          return maps[ti][c] >= 0;
        }),
      };
    });

    var warnings = [];
    variants.forEach(function (set) {
      if (set.size > 1) {
        warnings.push(
          "Headers " +
            Array.from(set)
              .map(function (s) {
                return '"' + s + '"';
              })
              .join(", ") +
            " were treated as the same column."
        );
      }
    });
    columnMap.forEach(function (cm) {
      var missing = [];
      cm.presentIn.forEach(function (present, ti) {
        if (!present) missing.push(tables[ti].name);
      });
      if (missing.length > 0 && missing.length < tables.length) {
        warnings.push(
          'Column "' + cm.header + '" is missing from ' + listNames(missing) + "; blank cells were filled in."
        );
      }
    });

    return { headers: outHeaders, rows: rows, warnings: warnings, perFile: perFile, columnMap: columnMap };
  }

  /* Mode 2 — simple stack (append by position).
     Header row comes from the first file; later files' data rows are
     appended as-is, position by position, with no remapping. Files with
     a different column count are flagged; short rows are padded with
     blanks, wider files keep their data under blank headers. */
  function appendByPosition(tables) {
    if (!tables.length) return { headers: [], rows: [], warnings: [], perFile: [], columnMap: [] };

    var first = tables[0];
    var width = 0;
    tables.forEach(function (t) {
      width = Math.max(width, t.headers.length);
    });

    var headers = new Array(width).fill("");
    for (var i = 0; i < first.headers.length; i++) headers[i] = first.headers[i];

    var warnings = [];
    tables.forEach(function (t, ti) {
      if (ti > 0 && t.headers.length !== first.headers.length) {
        warnings.push(
          '"' + t.name + '" has ' + t.headers.length + ' columns but "' + first.name + '" has ' +
            first.headers.length + ". " +
            (t.headers.length < first.headers.length
              ? "Its rows were padded with blank cells."
              : "Its extra columns were kept under blank headers.")
        );
      }
    });

    var rows = [];
    var perFile = [];
    tables.forEach(function (t) {
      t.rows.forEach(function (src) {
        var row = new Array(width).fill(null);
        for (var j = 0; j < src.length && j < width; j++) row[j] = src[j];
        rows.push(row);
      });
      perFile.push({ name: t.name, rows: t.rows.length, columns: t.headers.length });
    });

    var columnMap = headers.map(function (h, c) {
      return {
        header: h === "" ? "(blank)" : h,
        presentIn: tables.map(function (t) {
          return c < t.headers.length;
        }),
      };
    });

    return { headers: headers, rows: rows, warnings: warnings, perFile: perFile, columnMap: columnMap };
  }

  /* ------------------------------------------------------------------ *
   * Writing files
   * ------------------------------------------------------------------ */

  function csvEscape(text) {
    return /[",\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  }

  /* CSV text with a UTF-8 BOM (so Excel opens it with the right
     encoding) and CRLF line endings. */
  function toCSV(headers, rows) {
    var lines = [
      headers
        .map(function (h) {
          return csvEscape(formatCell(h));
        })
        .join(","),
    ];
    rows.forEach(function (row) {
      lines.push(
        row
          .map(function (cell) {
            return csvEscape(formatCell(cell));
          })
          .join(",")
      );
    });
    return "\uFEFF" + lines.join("\r\n") + "\r\n";
  }

  /* Excel sheet names: max 31 chars, no [ ] : * ? / \  */
  function sanitizeSheetName(name) {
    var s = String(name || "")
      .replace(/[\[\]:*?\/\\]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (s.charAt(0) === "'") s = s.slice(1);
    if (s.length > 31) s = s.slice(0, 31).trim();
    return s || "Data";
  }

  /* .xlsx file bytes (ArrayBuffer) for one table. Typed values (numbers,
     dates) stay typed; strings stay strings. */
  function toXLSX(headers, rows, sheetName) {
    var ws = XLSX.utils.aoa_to_sheet([headers].concat(rows), { cellDates: true });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheetName));
    return XLSX.write(wb, { type: "array", bookType: "xlsx", cellDates: true, compression: true });
  }

  /* Safe base name for output files (extension stripped, illegal
     filename characters replaced). */
  function sanitizeFileBase(name) {
    var base = String(name || "");
    var dot = base.lastIndexOf(".");
    if (dot > 0) base = base.slice(0, dot);
    base = base
      .replace(/[\\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/\.+$/, "");
    return base || "file";
  }

  /* ------------------------------------------------------------------ */

  return {
    VERSION: VERSION,
    formatCell: formatCell,
    normalizeHeader: normalizeHeader,
    canonKey: canonKey,
    columnLetter: columnLetter,
    cleanHeaders: cleanHeaders,
    extensionOf: extensionOf,
    readWorkbook: readWorkbook,
    tableFromSheet: tableFromSheet,
    makeBatches: makeBatches,
    selectColumns: selectColumns,
    mergeByHeader: mergeByHeader,
    appendByPosition: appendByPosition,
    csvEscape: csvEscape,
    toCSV: toCSV,
    sanitizeSheetName: sanitizeSheetName,
    toXLSX: toXLSX,
    sanitizeFileBase: sanitizeFileBase,
  };
});
