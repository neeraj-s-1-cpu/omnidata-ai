"use strict";

const assert = require("node:assert/strict");
const http = require("node:http");
const { after, before, test } = require("node:test");
const ExcelJS = require("exceljs");
const { app, normalizeTable } = require("../server");

let server;
let baseUrl;

before(async () => {
  server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  baseUrl = `http://127.0.0.1:${server.address().port}`;
});

after(async () => {
  if (server) await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
});

async function upload(name, content, type, fields = {}) {
  const form = new FormData();
  form.append("file", new Blob([content], { type }), name);
  Object.entries(fields).forEach(([key, value]) => form.append(key, value));
  const response = await fetch(`${baseUrl}/api/analyze`, { method: "POST", body: form });
  return { status: response.status, body: await response.json() };
}

test("normalizes duplicate headers and reports missing, duplicate, and outlier data", () => {
  const result = normalizeTable([
    ["Name", "Name", "Amount"],
    ["A", "A2", "10"],
    ["B", "B2", "12"],
    ["C", "C2", "11"],
    ["D", "D2", "10"],
    ["E", "E2", "1000"],
    ["F", "F2", ""],
    ["F", "F2", ""]
  ]);
  assert.deepEqual(result.headers, ["Name", "Name (2)", "Amount"]);
  assert.equal(result.summary.records, 7);
  assert.equal(result.summary.duplicateRows, 1);
  assert.equal(result.summary.missingCells, 2);
  assert.equal(result.summary.columnStats.Amount.outliers, 1);
  assert.ok(result.insights.some((item) => item.title === "Duplicate rows"));
});

test("analyzes quoted CSV and returns chart-ready statistics", async () => {
  const result = await upload(
    "sales.csv",
    'City,Revenue,Note\n"New York",1500,"line 1, line 2"\nParis,2100,ok\n',
    "text/csv"
  );
  assert.equal(result.status, 200);
  assert.equal(result.body.totalRows, 2);
  assert.equal(result.body.rows[0].City, "New York");
  assert.equal(result.body.rows[0].Note, "line 1, line 2");
  assert.deepEqual(result.body.summary.numericFields, ["Revenue"]);
});

test("analyzes JSON records and XLSX worksheets", async () => {
  const json = await upload("data.json", JSON.stringify([
    { name: "A", value: 3 },
    { name: "B", value: 9 }
  ]), "application/json");
  assert.equal(json.status, 200);
  assert.equal(json.body.summary.columnStats.value.max, 9);

  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet("Data").addRows([["Name", "Amount"], ["One", 4], ["Two", 8]]);
  const excel = await upload(
    "data.xlsx",
    await workbook.xlsx.writeBuffer(),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(excel.status, 200);
  assert.equal(excel.body.summary.columnStats.Amount.max, 8);
});

test("returns actionable errors for unsupported and unconsented audio uploads", async () => {
  const unsupported = await upload("legacy.xls", "data", "application/octet-stream");
  assert.equal(unsupported.status, 415);

  const audio = await upload("recording.wav", new Uint8Array([0, 1, 2]), "audio/wav");
  assert.equal(audio.status, 400);
  assert.match(audio.body.error, /consent/i);
});
