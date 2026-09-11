/*
 * tests.js — test suite for core.js. Environment-agnostic: runs in the
 * browser via tests.html (and in Node if one is ever available).
 * Returns { results: [{name, ok, error}], passed, failed }.
 */
(function (root, factory) {
  if (typeof module === "object" && typeof module.exports === "object") {
    module.exports = factory(require("./core.js"), require("../lib/xlsx.full.min.js"));
  } else {
    root.runCoreTests = factory(root.Core, root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (Core, XLSX) {
  "use strict";

  return function runCoreTests() {
    var results = [];

    function test(name, fn) {
      try {
        fn();
        results.push({ name: name, ok: true });
      } catch (e) {
        results.push({ name: name, ok: false, error: String((e && e.message) || e) });
      }
    }
    function eq(actual, expected, label) {
      var a = JSON.stringify(actual);
      var b = JSON.stringify(expected);
      if (a !== b) throw new Error((label || "eq") + " — expected " + b + " but got " + a);
    }
    function ok(cond, label) {
      if (!cond) throw new Error(label || "expected truthy");
    }
    function throws(fn, label) {
      try {
        fn();
      } catch (e) {
        return;
      }
      throw new Error((label || "throws") + " — expected an error");
    }
    function csvTable(csvText, name) {
      var wb = Core.readWorkbook(csvText, name || "test.csv");
      var t = Core.tableFromSheet(wb, wb.SheetNames[0]);
      t.name = name || "test.csv";
      return t;
    }

    /* ---------------- formatCell ---------------- */

    test("formatCell: null/undefined -> empty string", function () {
      eq(Core.formatCell(null), "");
      eq(Core.formatCell(undefined), "");
    });

    test("formatCell: big integers keep full precision", function () {
      eq(Core.formatCell(123456789012345), "123456789012345");
    });

    test("formatCell: strings pass through untouched (007 stays 007)", function () {
      eq(Core.formatCell("007"), "007");
      eq(Core.formatCell(" spaced "), " spaced ");
    });

    test("formatCell: booleans -> TRUE/FALSE", function () {
      eq(Core.formatCell(true), "TRUE");
      eq(Core.formatCell(false), "FALSE");
    });

    test("formatCell: date at midnight -> YYYY-MM-DD", function () {
      eq(Core.formatCell(new Date(2024, 0, 5)), "2024-01-05");
    });

    test("formatCell: date with time -> YYYY-MM-DD HH:MM:SS", function () {
      eq(Core.formatCell(new Date(2024, 0, 5, 13, 7, 9)), "2024-01-05 13:07:09");
    });

    /* ---------------- headers ---------------- */

    test("cleanHeaders: trims whitespace", function () {
      var r = Core.cleanHeaders(["  Name  ", "Age"], 2);
      eq(r.headers, ["Name", "Age"]);
      eq(r.warnings.length, 0);
    });

    test("cleanHeaders: blank headers get column-letter names", function () {
      var r = Core.cleanHeaders(["A", null, "C"], 3);
      eq(r.headers, ["A", "(Column B)", "C"]);
      eq(r.warnings.length, 1);
    });

    test("cleanHeaders: case-insensitive duplicates renamed", function () {
      var r = Core.cleanHeaders(["UBI", "ubi", "Ubi"], 3);
      eq(r.headers, ["UBI", "ubi (2)", "Ubi (3)"]);
      eq(r.warnings.length, 2);
    });

    test("cleanHeaders: width wider than header row adds unnamed columns", function () {
      var r = Core.cleanHeaders(["A"], 3);
      eq(r.headers, ["A", "(Column B)", "(Column C)"]);
    });

    test("canonKey: collapses case and whitespace", function () {
      ok(Core.canonKey(" First   Name ") === Core.canonKey("first name"), "keys should match");
    });

    /* ---------------- makeBatches ---------------- */

    test("makeBatches: exact multiple", function () {
      var b = Core.makeBatches(1500, 500);
      eq(b.length, 3);
      eq(b[2], { start: 1000, count: 500 });
    });

    test("makeBatches: remainder in last batch", function () {
      var b = Core.makeBatches(1643, 500);
      eq(b.length, 4);
      eq(b[3], { start: 1500, count: 143 });
    });

    test("makeBatches: batch size 1 / single row", function () {
      eq(Core.makeBatches(1, 1), [{ start: 0, count: 1 }]);
    });

    test("makeBatches: zero rows -> no batches", function () {
      eq(Core.makeBatches(0, 500), []);
    });

    test("makeBatches: invalid sizes throw", function () {
      throws(function () { Core.makeBatches(10, 0); }, "0");
      throws(function () { Core.makeBatches(10, -5); }, "-5");
      throws(function () { Core.makeBatches(10, 2.5); }, "2.5");
      throws(function () { Core.makeBatches(10, "abc"); }, "abc");
    });

    /* ---------------- CSV output ---------------- */

    test("csvEscape: quotes fields with commas, quotes, newlines", function () {
      eq(Core.csvEscape("a,b"), '"a,b"');
      eq(Core.csvEscape('say "hi"'), '"say ""hi"""');
      eq(Core.csvEscape("line1\nline2"), '"line1\nline2"');
      eq(Core.csvEscape("plain"), "plain");
    });

    test("toCSV: BOM prefix, CRLF endings, trailing newline", function () {
      var csv = Core.toCSV(["A", "B"], [["1", "2"]]);
      ok(csv.charCodeAt(0) === 0xfeff, "BOM expected");
      eq(csv.slice(1), "A,B\r\n1,2\r\n");
    });

    test("toCSV: null cells render blank, dates/numbers formatted", function () {
      var csv = Core.toCSV(["A", "B", "C"], [[null, 42, new Date(2024, 11, 31)]]);
      eq(csv.slice(1), "A,B,C\r\n,42,2024-12-31\r\n");
    });

    /* ---------------- parsing ---------------- */

    test("parse CSV: leading zeros preserved, all cells text", function () {
      var t = csvTable("id,name\r\n007,Ann\r\n2,Bob");
      eq(t.headers, ["id", "name"]);
      eq(t.rows, [["007", "Ann"], ["2", "Bob"]]);
    });

    test("parse CSV: quoted fields with commas", function () {
      var t = csvTable('id,note\r\n1,"hello, world"');
      eq(t.rows, [["1", "hello, world"]]);
    });

    test("parse CSV: blank rows skipped, header-only file -> 0 rows", function () {
      var t1 = csvTable("a,b\r\n\r\n1,2\r\n\r\n");
      eq(t1.rows.length, 1);
      var t2 = csvTable("a,b\r\n");
      eq(t2.rows.length, 0);
    });

    test("parse XLSX round trip: types survive, strings stay strings", function () {
      var bytes = Core.toXLSX(["id", "qty", "when"], [["007", 5, new Date(2024, 3, 15)]], "Data");
      var wb = Core.readWorkbook(bytes, "round.xlsx");
      var t = Core.tableFromSheet(wb, wb.SheetNames[0]);
      eq(t.headers, ["id", "qty", "when"]);
      eq(t.rows[0][0], "007", "string id preserved");
      eq(t.rows[0][1], 5, "number stays a number");
      ok(t.rows[0][2] instanceof Date, "date stays a Date");
      eq(Core.formatCell(t.rows[0][2]), "2024-04-15");
    });

    test("tableFromSheet: blank + duplicate headers cleaned with warnings", function () {
      var ws = XLSX.utils.aoa_to_sheet([["Name", "", "name"], ["a", "b", "c"]]);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "S1");
      var t = Core.tableFromSheet(wb, "S1");
      eq(t.headers, ["Name", "(Column B)", "name (2)"]);
      eq(t.warnings.length, 2);
      eq(t.rows, [["a", "b", "c"]]);
    });

    test("tableFromSheet: data wider than header row keeps extra cells", function () {
      var t = csvTable("a,b\r\n1,2,3");
      eq(t.headers, ["a", "b", "(Column C)"]);
      eq(t.rows, [["1", "2", "3"]]);
    });

    /* ---------------- selectColumns ---------------- */

    test("selectColumns: keeps subset in given order", function () {
      var t = { headers: ["A", "B", "C"], rows: [["1", "2", "3"], ["4", "5", "6"]] };
      var r = Core.selectColumns(t, [2, 0]);
      eq(r.headers, ["C", "A"]);
      eq(r.rows, [["3", "1"], ["6", "4"]]);
    });

    /* ---------------- mergeByHeader ---------------- */

    test("mergeByHeader: swapped column order aligns by header name", function () {
      var t1 = csvTable("A,B\r\ndata1,data2", "file1.csv");
      var t2 = csvTable("B,A\r\ndata2x,data1x", "file2.csv");
      var m = Core.mergeByHeader([t1, t2]);
      eq(m.headers, ["A", "B"]);
      eq(m.rows, [["data1", "data2"], ["data1x", "data2x"]]);
      eq(m.warnings.length, 0);
    });

    test("mergeByHeader: case/whitespace header variants merge into one column", function () {
      var t1 = csvTable('"UBI ",x\r\n1,2', "f1.csv");
      var t2 = csvTable("ubi,x\r\n3,4", "f2.csv");
      var m = Core.mergeByHeader([t1, t2]);
      eq(m.headers, ["UBI", "x"]);
      eq(m.rows, [["1", "2"], ["3", "4"]]);
      ok(m.warnings.some(function (w) { return w.indexOf("treated as the same column") >= 0; }), "variant warning expected");
    });

    test("mergeByHeader: union of columns, missing cells blank, column map correct", function () {
      var t1 = csvTable("A,B\r\n1,2", "f1.csv");
      var t2 = csvTable("B,C\r\n3,4", "f2.csv");
      var m = Core.mergeByHeader([t1, t2]);
      eq(m.headers, ["A", "B", "C"]);
      eq(m.rows, [["1", "2", null], [null, "3", "4"]]);
      eq(m.columnMap.map(function (c) { return c.presentIn; }), [[true, false], [true, true], [false, true]]);
      ok(m.warnings.some(function (w) { return w.indexOf('Column "A" is missing') >= 0; }), "missing-column warning expected");
    });

    test("mergeByHeader: perFile stats", function () {
      var t1 = csvTable("A\r\n1\r\n2", "f1.csv");
      var t2 = csvTable("A\r\n3", "f2.csv");
      var m = Core.mergeByHeader([t1, t2]);
      eq(m.perFile, [{ name: "f1.csv", rows: 2, columns: 1 }, { name: "f2.csv", rows: 1, columns: 1 }]);
    });

    /* ---------------- appendByPosition ---------------- */

    test("appendByPosition: same layout stacks with no warnings", function () {
      var t1 = csvTable("A,B\r\n1,2", "f1.csv");
      var t2 = csvTable("X,Y\r\n3,4", "f2.csv");
      var m = Core.appendByPosition([t1, t2]);
      eq(m.headers, ["A", "B"]);
      eq(m.rows, [["1", "2"], ["3", "4"]]);
      eq(m.warnings.length, 0);
    });

    test("appendByPosition: narrower file padded with blanks + warning", function () {
      var t1 = csvTable("A,B,C\r\n1,2,3", "f1.csv");
      var t2 = csvTable("A\r\n9", "f2.csv");
      var m = Core.appendByPosition([t1, t2]);
      eq(m.rows, [["1", "2", "3"], ["9", null, null]]);
      eq(m.warnings.length, 1);
      ok(m.warnings[0].indexOf("padded") >= 0, "pad warning expected");
    });

    test("appendByPosition: wider file keeps data under blank headers + warning", function () {
      var t1 = csvTable("A,B\r\n1,2", "f1.csv");
      var t2 = csvTable("A,B,C\r\n3,4,5", "f2.csv");
      var m = Core.appendByPosition([t1, t2]);
      eq(m.headers, ["A", "B", ""]);
      eq(m.rows, [["1", "2", null], ["3", "4", "5"]]);
      eq(m.warnings.length, 1);
      eq(m.columnMap[2].header, "(blank)");
    });

    /* ---------------- names ---------------- */

    test("sanitizeSheetName: strips illegal chars, caps at 31", function () {
      eq(Core.sanitizeSheetName("bad[name]:with*chars?"), "bad name with chars");
      ok(Core.sanitizeSheetName("x".repeat(50)).length <= 31, "length cap");
      eq(Core.sanitizeSheetName(""), "Data");
    });

    test("sanitizeFileBase: strips extension and illegal filename chars", function () {
      eq(Core.sanitizeFileBase("my.data.xlsx"), "my.data");
      eq(Core.sanitizeFileBase('we/ird:name.csv'), "we_ird_name");
      eq(Core.sanitizeFileBase(".csv"), ".csv" === "" ? "file" : Core.sanitizeFileBase(".csv"));
      eq(Core.sanitizeFileBase(""), "file");
    });

    test("extensionOf: detects extensions case-insensitively", function () {
      eq(Core.extensionOf("Report.XLSX"), "xlsx");
      eq(Core.extensionOf("data.csv"), "csv");
      eq(Core.extensionOf("noext"), "");
    });

    /* ---------------- integration: split then consolidate ---------------- */

    test("integration: split into batches then merge restores all rows", function () {
      var lines = ["id,value"];
      for (var i = 1; i <= 7; i++) lines.push("row" + i + "," + i);
      var t = csvTable(lines.join("\r\n"), "big.csv");
      var batches = Core.makeBatches(t.rows.length, 3);
      eq(batches.length, 3);

      var parts = batches.map(function (b, bi) {
        var csv = Core.toCSV(t.headers, t.rows.slice(b.start, b.start + b.count));
        return csvTable(csv, "part" + (bi + 1) + ".csv");
      });
      var merged = Core.mergeByHeader(parts);
      eq(merged.rows.length, 7);
      eq(merged.rows[6], ["row7", "7"]);
    });

    var passed = results.filter(function (r) { return r.ok; }).length;
    return { results: results, passed: passed, failed: results.length - passed };
  };
});
