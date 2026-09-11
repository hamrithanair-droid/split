/*
 * ui.js — DOM wiring for the Batch File Toolkit.
 * All data logic lives in js/core.js; this file only handles state,
 * rendering and downloads. User-provided text is always inserted with
 * textContent (never innerHTML) to avoid injection.
 */
(function () {
  "use strict";

  /* ------------------------------ helpers ------------------------------ */

  function $(id) {
    return document.getElementById(id);
  }

  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined && text !== null) node.textContent = text;
    return node;
  }

  function show(node, visible) {
    node.classList.toggle("hidden", !visible);
  }

  function num(n) {
    return Number(n).toLocaleString("en-US");
  }

  function formatBytes(bytes) {
    if (!isFinite(bytes)) return "";
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }

  var ACCEPTED = { xlsx: true, xls: true, csv: true };

  function isAccepted(filename) {
    return !!ACCEPTED[Core.extensionOf(filename)];
  }

  /* ------------------------------- toasts ------------------------------ */

  var toastHost = $("toasts");

  function toast(message, type, sticky) {
    var box = el("div", "toast " + (type || ""));
    box.appendChild(el("span", null, message));
    var close = el("button", "t-close", "\u00D7");
    close.setAttribute("aria-label", "Dismiss notification");
    close.addEventListener("click", function () {
      box.remove();
    });
    box.appendChild(close);
    toastHost.appendChild(box);
    if (!sticky) {
      setTimeout(function () {
        box.remove();
      }, 6000);
    }
  }

  /* ---------------------------- busy overlay --------------------------- */

  var busyNode = $("busy");
  var busyText = $("busy-text");

  /* Run `fn` behind the busy overlay. The tiny delay lets the overlay
     paint before SheetJS blocks the main thread on big files. */
  function runBusy(text, fn) {
    busyText.textContent = text;
    show(busyNode, true);
    return new Promise(function (resolve) {
      setTimeout(resolve, 50);
    })
      .then(fn)
      .finally(function () {
        show(busyNode, false);
      });
  }

  /* ----------------------------- downloads ----------------------------- */

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 20000);
  }

  /* --------------------------- file -> table --------------------------- */

  function parseFileToEntry(file) {
    return file.arrayBuffer().then(function (buffer) {
      var workbook = Core.readWorkbook(buffer, file.name);
      var sheetNames = workbook.SheetNames.slice();
      var sheetName = sheetNames[0];
      var table = Core.tableFromSheet(workbook, sheetName);
      table.name = file.name;
      return {
        name: file.name,
        size: file.size,
        workbook: workbook,
        sheetNames: sheetNames,
        sheetName: sheetName,
        table: table,
      };
    });
  }

  function switchEntrySheet(entry, sheetName) {
    entry.sheetName = sheetName;
    entry.table = Core.tableFromSheet(entry.workbook, sheetName);
    entry.table.name = entry.name;
  }

  /* ---------------------------- preview table -------------------------- */

  var PREVIEW_ROWS = 10;
  var PREVIEW_COLS = 30;

  function renderPreviewTable(tableNode, headers, rows) {
    tableNode.textContent = "";
    var colCount = Math.min(headers.length, PREVIEW_COLS);
    var clippedCols = headers.length - colCount;

    var thead = document.createElement("thead");
    var headRow = document.createElement("tr");
    for (var c = 0; c < colCount; c++) {
      var th = el("th", null, headers[c] === "" ? "(blank)" : headers[c]);
      th.title = headers[c];
      headRow.appendChild(th);
    }
    if (clippedCols > 0) headRow.appendChild(el("th", null, "\u2026 +" + clippedCols + " more"));
    thead.appendChild(headRow);
    tableNode.appendChild(thead);

    var tbody = document.createElement("tbody");
    var shown = Math.min(rows.length, PREVIEW_ROWS);
    for (var r = 0; r < shown; r++) {
      var tr = document.createElement("tr");
      for (var c2 = 0; c2 < colCount; c2++) {
        var text = Core.formatCell(rows[r][c2]);
        var td = el("td", text === "" ? "blank-cell" : null, text);
        if (text.length > 40) td.title = text;
        tr.appendChild(td);
      }
      if (clippedCols > 0) tr.appendChild(el("td", "blank-cell", "\u2026"));
      tbody.appendChild(tr);
    }
    if (shown === 0) {
      var emptyRow = document.createElement("tr");
      var emptyCell = el("td", "blank-cell", "No data rows");
      emptyCell.colSpan = Math.max(colCount, 1);
      emptyRow.appendChild(emptyCell);
      tbody.appendChild(emptyRow);
    }
    tableNode.appendChild(tbody);
  }

  function renderWarnings(node, warnings) {
    node.textContent = "";
    if (!warnings.length) {
      show(node, false);
      return;
    }
    node.appendChild(el("strong", null, "Heads up:"));
    var list = document.createElement("ul");
    warnings.forEach(function (w) {
      list.appendChild(el("li", null, w));
    });
    node.appendChild(list);
    show(node, true);
  }

  function renderStats(node, pairs) {
    node.textContent = "";
    pairs.forEach(function (pair) {
      var wrap = document.createElement("div");
      wrap.appendChild(el("dt", null, pair[0]));
      var dd = el("dd", pair[2] ? "small-val" : null, pair[1]);
      wrap.appendChild(dd);
      node.appendChild(wrap);
    });
  }

  /* ---------------------------- drop zones ----------------------------- */

  function wireDropzone(zone, input, onFiles) {
    zone.addEventListener("click", function () {
      input.click();
    });
    zone.addEventListener("keydown", function (event) {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        input.click();
      }
    });
    zone.addEventListener("dragover", function (event) {
      event.preventDefault();
      zone.classList.add("dragover");
    });
    zone.addEventListener("dragleave", function () {
      zone.classList.remove("dragover");
    });
    zone.addEventListener("drop", function (event) {
      event.preventDefault();
      zone.classList.remove("dragover");
      if (event.dataTransfer && event.dataTransfer.files.length) {
        onFiles(Array.prototype.slice.call(event.dataTransfer.files));
      }
    });
    input.addEventListener("change", function () {
      if (input.files.length) onFiles(Array.prototype.slice.call(input.files));
      input.value = "";
    });
  }

  var BIG_FILE_BYTES = 50 * 1024 * 1024;

  function warnIfHuge(file) {
    if (file.size > BIG_FILE_BYTES) {
      toast('"' + file.name + '" is ' + formatBytes(file.size) + " — processing may take a moment.", "warn");
    }
  }

  /* ====================================================================== *
   *  SPLIT
   * ====================================================================== */

  var splitState = { entry: null };

  var splitPreviewCard = $("split-preview-card");
  var splitOptionsCard = $("split-options-card");
  var splitResultCard = $("split-result-card");

  wireDropzone($("split-drop"), $("split-input"), function (files) {
    var file = files[0];
    if (files.length > 1) toast("Split works on one file at a time — using the first one.", "warn");
    if (!isAccepted(file.name)) {
      toast('"' + file.name + '" is not a supported file type. Please use .xlsx, .xls or .csv.', "error");
      return;
    }
    warnIfHuge(file);
    runBusy('Reading "' + file.name + '"\u2026', function () {
      return parseFileToEntry(file).then(function (entry) {
        splitState.entry = entry;
        show(splitResultCard, false);
        renderSplit();
      });
    }).catch(function (err) {
      toast('Could not read "' + file.name + '":\n' + ((err && err.message) || err), "error", true);
    });
  });

  $("split-clear").addEventListener("click", resetSplit);
  $("split-again").addEventListener("click", resetSplit);

  function resetSplit() {
    splitState.entry = null;
    show($("split-fileinfo"), false);
    show($("split-warnings"), false);
    show(splitPreviewCard, false);
    show(splitOptionsCard, false);
    show(splitResultCard, false);
  }

  function renderSplit() {
    var entry = splitState.entry;
    if (!entry) return;
    var table = entry.table;

    $("split-filename").textContent = entry.name;
    $("split-filemeta").textContent =
      formatBytes(entry.size) + " \u00B7 " + num(table.rows.length) + " data rows \u00D7 " +
      num(table.headers.length) + " columns";
    show($("split-fileinfo"), true);

    var sheetWrap = $("split-sheet-wrap");
    var sheetSelect = $("split-sheet");
    if (entry.sheetNames.length > 1) {
      sheetSelect.textContent = "";
      entry.sheetNames.forEach(function (name) {
        var option = el("option", null, name);
        option.value = name;
        option.selected = name === entry.sheetName;
        sheetSelect.appendChild(option);
      });
      show(sheetWrap, true);
    } else {
      show(sheetWrap, false);
    }

    renderWarnings($("split-warnings"), table.warnings);
    renderPreviewTable($("split-preview"), table.headers, table.rows);
    show(splitPreviewCard, true);

    var grid = $("split-columns");
    grid.textContent = "";
    table.headers.forEach(function (header, index) {
      var chip = el("label", "chip");
      var input = document.createElement("input");
      input.type = "checkbox";
      input.checked = true;
      input.dataset.index = String(index);
      var span = el("span");
      span.appendChild(el("span", "chip-label", header));
      chip.appendChild(input);
      chip.appendChild(span);
      input.addEventListener("change", updateSplitPlan);
      grid.appendChild(chip);
    });

    show(splitOptionsCard, true);
    updateSplitPlan();
  }

  $("split-sheet").addEventListener("change", function () {
    var entry = splitState.entry;
    if (!entry) return;
    switchEntrySheet(entry, this.value);
    renderSplit();
  });

  $("split-cols-all").addEventListener("click", function () {
    setAllSplitColumns(true);
  });
  $("split-cols-none").addEventListener("click", function () {
    setAllSplitColumns(false);
  });

  function setAllSplitColumns(checked) {
    $("split-columns")
      .querySelectorAll("input[type=checkbox]")
      .forEach(function (input) {
        input.checked = checked;
      });
    updateSplitPlan();
  }

  function checkedColumnIndices() {
    var indices = [];
    $("split-columns")
      .querySelectorAll("input[type=checkbox]")
      .forEach(function (input) {
        if (input.checked) indices.push(Number(input.dataset.index));
      });
    return indices;
  }

  function splitFormat() {
    return document.querySelector('input[name="split-format"]:checked').value;
  }

  $("split-size").addEventListener("input", updateSplitPlan);
  document.querySelectorAll('input[name="split-format"]').forEach(function (radio) {
    radio.addEventListener("change", updateSplitPlan);
  });

  function updateSplitPlan() {
    var entry = splitState.entry;
    if (!entry) return;
    var plan = $("split-plan");
    var runButton = $("split-run");
    var kept = checkedColumnIndices();
    var totalRows = entry.table.rows.length;

    $("split-cols-count").textContent = kept.length + " of " + entry.table.headers.length + " selected";

    var problems = [];
    if (totalRows === 0) problems.push("This sheet has no data rows to split.");
    if (kept.length === 0) problems.push("Select at least one column.");

    var batches = [];
    if (!problems.length) {
      try {
        batches = Core.makeBatches(totalRows, $("split-size").value);
      } catch (err) {
        problems.push(err.message);
      }
    }

    if (problems.length) {
      plan.textContent = problems.join(" ");
      plan.style.background = "var(--danger-soft)";
      plan.style.borderColor = "#f2c3be";
      plan.style.color = "var(--danger)";
      runButton.disabled = true;
      return;
    }

    plan.style.background = "";
    plan.style.borderColor = "";
    plan.style.color = "";
    var last = batches[batches.length - 1];
    var sizeValue = Number($("split-size").value);
    var text =
      num(totalRows) + " rows \u2192 " + num(batches.length) +
      (batches.length === 1 ? " file" : " files") + " of up to " + num(sizeValue) + " rows";
    if (batches.length > 1 && last.count !== sizeValue) {
      text += " (last file " + num(last.count) + " rows)";
    }
    text += " \u00B7 " + kept.length + " of " + entry.table.headers.length + " columns \u00B7 " +
      (splitFormat() === "csv" ? "CSV" : "Excel") + " output";
    plan.textContent = text;
    runButton.disabled = false;
  }

  $("split-run").addEventListener("click", function () {
    var entry = splitState.entry;
    if (!entry) return;
    var format = splitFormat();
    var kept = checkedColumnIndices();
    var base = Core.sanitizeFileBase(entry.name);

    runBusy("Splitting and zipping\u2026", function () {
      var filtered = Core.selectColumns(entry.table, kept);
      var batches = Core.makeBatches(filtered.rows.length, $("split-size").value);
      var zip = new JSZip();
      var folder = zip.folder(base + "_split");

      batches.forEach(function (batch, index) {
        var rows = filtered.rows.slice(batch.start, batch.start + batch.count);
        var partName = base + "_split" + (index + 1);
        if (format === "csv") {
          folder.file(partName + ".csv", Core.toCSV(filtered.headers, rows));
        } else {
          folder.file(partName + ".xlsx", Core.toXLSX(filtered.headers, rows, partName));
        }
      });

      return zip
        .generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } })
        .then(function (blob) {
          downloadBlob(blob, base + "_split.zip");
          var last = batches[batches.length - 1];
          renderStats($("split-stats"), [
            ["Input rows", num(filtered.rows.length)],
            ["Files created", num(batches.length)],
            ["Rows per file", num(Number($("split-size").value))],
            ["Last file rows", num(last ? last.count : 0)],
            ["Columns kept", kept.length + " of " + entry.table.headers.length],
            ["Format", format === "csv" ? "CSV" : "Excel (.xlsx)"],
            ["Download", base + "_split.zip", true],
          ]);
          show(splitResultCard, true);
          splitResultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
          toast("Split complete — " + num(batches.length) + " files downloaded as a ZIP.", "success");
        });
    }).catch(function (err) {
      toast("Split failed:\n" + ((err && err.message) || err), "error", true);
    });
  });

  /* ====================================================================== *
   *  CONSOLIDATE
   * ====================================================================== */

  var consState = { items: [], nextId: 1 };

  var consOptionsCard = $("cons-options-card");
  var consReviewCard = $("cons-review-card");
  var consResultCard = $("cons-result-card");

  var MODE_HINTS = {
    header:
      "Columns are aligned by header text, so column order can differ between files. The output has every column found in any file.",
    stack:
      "Rows are appended in file order with no remapping — every file must already share the first file's column layout. The header row comes from the first file.",
  };

  wireDropzone($("cons-drop"), $("cons-input"), function (files) {
    addConsFiles(files);
  });

  function addConsFiles(files) {
    var usable = [];
    files.forEach(function (file) {
      if (!isAccepted(file.name)) {
        toast('"' + file.name + '" skipped — not a supported type (.xlsx, .xls, .csv).', "warn");
      } else {
        warnIfHuge(file);
        usable.push(file);
      }
    });
    if (!usable.length) return;

    runBusy("Reading " + usable.length + (usable.length === 1 ? " file" : " files") + "\u2026", function () {
      var chain = Promise.resolve();
      usable.forEach(function (file) {
        chain = chain.then(function () {
          busyText.textContent = 'Reading "' + file.name + '"\u2026';
          return parseFileToEntry(file)
            .then(function (entry) {
              entry.id = consState.nextId++;
              consState.items.push(entry);
            })
            .catch(function (err) {
              consState.items.push({
                id: consState.nextId++,
                name: file.name,
                size: file.size,
                error: String((err && err.message) || err),
              });
            });
        });
      });
      return chain;
    }).then(function () {
      show(consResultCard, false);
      renderCons();
    });
  }

  $("cons-clear").addEventListener("click", function () {
    consState.items = [];
    renderCons();
  });

  $("cons-again").addEventListener("click", function () {
    consState.items = [];
    show(consResultCard, false);
    renderCons();
  });

  document.querySelectorAll('input[name="cons-mode"]').forEach(function (radio) {
    radio.addEventListener("change", function () {
      $("cons-mode-hint").textContent = MODE_HINTS[consMode()];
      renderConsReview();
    });
  });

  document.querySelectorAll('input[name="cons-format"]').forEach(function (radio) {
    radio.addEventListener("change", renderConsReview);
  });

  function consMode() {
    return document.querySelector('input[name="cons-mode"]:checked').value;
  }

  function consFormat() {
    return document.querySelector('input[name="cons-format"]:checked').value;
  }

  function validConsTables() {
    return consState.items
      .filter(function (item) {
        return !item.error;
      })
      .map(function (item) {
        return item.table;
      });
  }

  var FILE_ICON =
    '<svg class="fl-icon" viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="M6 2h6l4 4v11a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1z" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><path d="M12 2v4h4" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>';
  var TRASH_ICON =
    '<svg viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M2.5 4.5h11M6.5 2.5h3M5.5 4.5V13a1 1 0 0 0 1 1h3a1 1 0 0 0 1-1V4.5M6.8 7v4.5M9.2 7v4.5" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  function renderCons() {
    var list = $("cons-list");
    list.textContent = "";

    consState.items.forEach(function (item) {
      var li = document.createElement("li");
      if (item.error) li.classList.add("error");
      li.insertAdjacentHTML("afterbegin", FILE_ICON);

      var text = el("div", "fl-text");
      text.appendChild(el("span", "fl-name", item.name));
      var meta = item.error
        ? "Could not read this file: " + item.error
        : num(item.table.rows.length) + " rows \u00D7 " + num(item.table.headers.length) +
          " cols \u00B7 " + formatBytes(item.size);
      text.appendChild(el("span", "fl-meta", meta));
      li.appendChild(text);

      if (!item.error && item.sheetNames.length > 1) {
        var sheetLabel = el("label", "sheet-select", "Sheet ");
        var select = document.createElement("select");
        item.sheetNames.forEach(function (name) {
          var option = el("option", null, name);
          option.value = name;
          option.selected = name === item.sheetName;
          select.appendChild(option);
        });
        select.addEventListener("change", function () {
          switchEntrySheet(item, select.value);
          renderCons();
        });
        sheetLabel.appendChild(select);
        li.appendChild(sheetLabel);
      }

      var remove = el("button", "fl-remove");
      remove.type = "button";
      remove.setAttribute("aria-label", 'Remove "' + item.name + '" from the list');
      remove.insertAdjacentHTML("afterbegin", TRASH_ICON);
      remove.addEventListener("click", function () {
        consState.items = consState.items.filter(function (other) {
          return other.id !== item.id;
        });
        renderCons();
      });
      li.appendChild(remove);

      list.appendChild(li);
    });

    var valid = validConsTables();
    $("cons-count").textContent =
      consState.items.length + (consState.items.length === 1 ? " file" : " files") +
      (consState.items.length !== valid.length
        ? " (" + (consState.items.length - valid.length) + " unreadable)"
        : "");
    show($("cons-list-actions"), consState.items.length > 0);
    show(consOptionsCard, valid.length > 0);
    renderConsReview();
  }

  var MAP_MAX_COLUMNS = 60;

  function renderConsReview() {
    var tables = validConsTables();
    if (!tables.length) {
      show(consReviewCard, false);
      return;
    }

    var mode = consMode();
    var merged = mode === "header" ? Core.mergeByHeader(tables) : Core.appendByPosition(tables);

    /* column map: one row per output column, one column per file */
    var map = $("cons-map");
    map.textContent = "";
    var thead = document.createElement("thead");
    var headRow = document.createElement("tr");
    headRow.appendChild(el("th", "map-col-header", "Output column"));
    tables.forEach(function (table) {
      var th = el("th", null, table.name);
      th.title = table.name;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    map.appendChild(thead);

    var tbody = document.createElement("tbody");
    var mapRows = Math.min(merged.columnMap.length, MAP_MAX_COLUMNS);
    for (var i = 0; i < mapRows; i++) {
      var colInfo = merged.columnMap[i];
      var tr = document.createElement("tr");
      var nameCell = el("th", "map-col-header", colInfo.header);
      nameCell.scope = "row";
      nameCell.title = colInfo.header;
      tr.appendChild(nameCell);
      colInfo.presentIn.forEach(function (present) {
        tr.appendChild(el("td", present ? "map-yes" : "map-no", present ? "\u2713" : "\u2014"));
      });
      tbody.appendChild(tr);
    }
    if (merged.columnMap.length > MAP_MAX_COLUMNS) {
      var moreRow = document.createElement("tr");
      var moreCell = el("td", "blank-cell", "\u2026 +" + (merged.columnMap.length - MAP_MAX_COLUMNS) + " more columns");
      moreCell.colSpan = tables.length + 1;
      moreRow.appendChild(moreCell);
      tbody.appendChild(moreRow);
    }
    map.appendChild(tbody);

    renderPreviewTable($("cons-preview"), merged.headers, merged.rows);

    /* warnings: per-file parse warnings + merge warnings */
    var warnings = [];
    consState.items.forEach(function (item) {
      if (item.error) {
        warnings.push('"' + item.name + '" could not be read and will be left out.');
      } else if (item.table.warnings.length) {
        item.table.warnings.forEach(function (w) {
          warnings.push(item.name + ": " + w);
        });
      }
    });
    warnings = warnings.concat(merged.warnings);
    renderWarnings($("cons-warnings"), warnings);

    $("cons-summary").textContent =
      num(tables.length) + (tables.length === 1 ? " file" : " files") + " \u00B7 " +
      num(merged.rows.length) + " total rows \u00B7 " + num(merged.headers.length) +
      " output columns \u00B7 " + (mode === "header" ? "match by header name" : "simple stack") +
      " \u00B7 " + (consFormat() === "csv" ? "CSV" : "Excel") + " output";

    $("cons-run").disabled = merged.rows.length === 0 && merged.headers.length === 0;
    show(consReviewCard, true);
  }

  $("cons-run").addEventListener("click", function () {
    var tables = validConsTables();
    if (!tables.length) return;
    var mode = consMode();
    var format = consFormat();
    /* ".x" is a throwaway extension so sanitizeFileBase strips it and keeps
       any dots the user typed in the name itself. */
    var outName = Core.sanitizeFileBase(($("cons-name").value || "combined") + ".x");

    runBusy("Combining\u2026", function () {
      return new Promise(function (resolve) {
        setTimeout(resolve, 10);
      }).then(function () {
        var merged = mode === "header" ? Core.mergeByHeader(tables) : Core.appendByPosition(tables);
        var filename;
        var blob;
        if (format === "csv") {
          filename = outName + ".csv";
          blob = new Blob([Core.toCSV(merged.headers, merged.rows)], { type: "text/csv;charset=utf-8" });
        } else {
          filename = outName + ".xlsx";
          blob = new Blob([Core.toXLSX(merged.headers, merged.rows, outName)], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          });
        }
        downloadBlob(blob, filename);

        var statPairs = [
          ["Files merged", num(merged.perFile.length)],
          ["Total rows", num(merged.rows.length)],
          ["Output columns", num(merged.headers.length)],
          ["Mode", mode === "header" ? "Match by header name" : "Simple stack"],
          ["Format", format === "csv" ? "CSV" : "Excel (.xlsx)"],
          ["Download", filename, true],
        ];
        var perFileText = merged.perFile
          .map(function (pf) {
            return pf.name + " (" + num(pf.rows) + ")";
          })
          .join(", ");
        statPairs.push(["Rows per file", perFileText, true]);
        renderStats($("cons-stats"), statPairs);
        show(consResultCard, true);
        consResultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
        toast("Combined " + num(merged.perFile.length) + " files into " + filename + ".", "success");
      });
    }).catch(function (err) {
      toast("Combine failed:\n" + ((err && err.message) || err), "error", true);
    });
  });

  /* ------------------------------ startup ------------------------------ */

  if (typeof XLSX === "undefined" || typeof JSZip === "undefined" || typeof Core === "undefined") {
    toast(
      "Some script files failed to load. Make sure the lib/ and js/ folders sit next to index.html.",
      "error",
      true
    );
  }
})();
