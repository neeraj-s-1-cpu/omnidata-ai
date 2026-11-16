"use strict";

require("dotenv").config();

const path = require("node:path");
const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const mammoth = require("mammoth");
const OpenAI = require("openai");
const { parse: parseCsv } = require("csv-parse/sync");
const { createWorker } = require("tesseract.js");

const app = express();
const port = Number(process.env.PORT || 3000);
const configuredMaxFileSize = Number(process.env.MAX_FILE_SIZE_MB || 20);
const maxFileSizeMb = Number.isFinite(configuredMaxFileSize) ? configuredMaxFileSize : 20;
const maxFileBytes = Math.max(1, Math.min(maxFileSizeMb, 100)) * 1024 * 1024;
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: maxFileBytes, files: 1 }
});
const openai = process.env.OPENAI_API_KEY ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY }) : null;
const supportedExtensions = new Set([
  ".csv", ".tsv", ".json", ".xlsx", ".txt", ".docx", ".pdf",
  ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff",
  ".mp3", ".wav", ".m4a", ".mp4", ".mpeg", ".mpga", ".webm"
]);
const imageExtensions = new Set([".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"]);
const audioExtensions = new Set([".mp3", ".wav", ".m4a", ".mp4", ".mpeg", ".mpga", ".webm"]);
let ocrWorkerPromise;

app.disable("x-powered-by");
app.use(express.json({ limit: "1mb" }));

function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function cellValue(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "number") return Number.isFinite(value) ? value : "";
  if (typeof value === "boolean") return value;
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) return value.richText.map((item) => item.text || "").join("");
    if (value.result !== undefined) return cellValue(value.result);
    if (value.text !== undefined) return String(value.text);
    return JSON.stringify(value);
  }
  const text = String(value).trim();
  if (text && /^-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) {
    const number = Number(text);
    if (Number.isFinite(number)) return number;
  }
  return text;
}

function normalizeTable(matrix) {
  if (!Array.isArray(matrix) || matrix.length < 2) {
    throw httpError(422, "This file needs a header row and at least one data row.");
  }
  const seen = new Map();
  const headers = matrix[0].map((value, index) => {
    const base = String(value == null ? "" : value).trim() || `Column ${index + 1}`;
    const count = (seen.get(base) || 0) + 1;
    seen.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
  const rows = matrix.slice(1).filter((row) =>
    Array.isArray(row) && row.some((value) => String(value == null ? "" : value).trim() !== "")
  ).map((row) => {
    const record = {};
    headers.forEach((header, index) => { record[header] = cellValue(row[index]); });
    return record;
  });
  if (!rows.length) throw httpError(422, "No data rows were found in this file.");
  return analyzeTable(headers, rows);
}

function round(value) {
  return Number.isFinite(value) ? Number(value.toFixed(2)) : null;
}

function quantile(sorted, percentile) {
  const index = (sorted.length - 1) * percentile;
  const lower = Math.floor(index);
  const fraction = index - lower;
  return sorted[lower + 1] === undefined
    ? sorted[lower]
    : sorted[lower] + fraction * (sorted[lower + 1] - sorted[lower]);
}

function analyzeTable(headers, rows, extra = {}) {
  const totalCells = headers.length * rows.length;
  let filledCells = 0;
  const numericFields = [];
  const columnStats = {};
  const rowKeys = new Set();
  let duplicateRows = 0;
  rows.forEach((row) => {
    const key = JSON.stringify(headers.map((header) => row[header]));
    if (rowKeys.has(key)) duplicateRows++;
    else rowKeys.add(key);
  });

  headers.forEach((header) => {
    let filled = 0;
    const numbers = [];
    rows.forEach((row) => {
      const value = row[header];
      if (value !== "" && value !== null && value !== undefined) {
        filled++;
        filledCells++;
      }
      if (typeof value === "number" && Number.isFinite(value)) numbers.push(value);
    });
    const missing = rows.length - filled;
    if (numbers.length >= 2 && numbers.length / rows.length >= 0.5) {
      numericFields.push(header);
      const sorted = numbers.slice().sort((a, b) => a - b);
      const middle = Math.floor(sorted.length / 2);
      const q1 = quantile(sorted, 0.25);
      const q3 = quantile(sorted, 0.75);
      const iqr = q3 - q1;
      const outliers = iqr === 0
        ? 0
        : numbers.filter((value) => value < q1 - 1.5 * iqr || value > q3 + 1.5 * iqr).length;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      columnStats[header] = {
        type: "number",
        count: numbers.length,
        missing,
        min: round(sorted[0]),
        max: round(sorted[sorted.length - 1]),
        mean: round(mean),
        median: round(sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2),
        standardDeviation: round(Math.sqrt(numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / numbers.length)),
        q1: round(q1),
        q3: round(q3),
        outliers
      };
    } else {
      columnStats[header] = { type: "text or mixed", count: filled, missing };
    }
  });

  const completeness = totalCells ? Math.round((filledCells / totalCells) * 100) : 0;
  const summary = {
    records: rows.length,
    fields: headers.length,
    completeness,
    missingCells: totalCells - filledCells,
    duplicateRows,
    numericFields,
    columnStats
  };
  const insights = [];
  const missingPercent = 100 - completeness;
  insights.push(missingPercent
    ? { title: "Missing data", text: `${missingPercent}% of cells are empty. Review gaps before drawing conclusions.` }
    : { title: "Complete dataset", text: `All cells are populated across ${rows.length.toLocaleString()} records.` });
  if (duplicateRows) {
    insights.push({ title: "Duplicate rows", text: `${duplicateRows} repeated row(s) were found. Review them before aggregating.` });
  }
  if (numericFields.length) {
    const field = numericFields[0];
    const stats = columnStats[field];
    const outlierNote = stats.outliers
      ? ` The IQR rule flags ${stats.outliers} potential outlier(s); check them, as outliers are not necessarily errors.`
      : "";
    insights.push({ title: `${field} at a glance`, text: `Average ${stats.mean}; median ${stats.median}; range ${stats.min} to ${stats.max}.${outlierNote}` });
  } else {
    const field = headers.find((header) => rows.some((row) => typeof row[header] === "string" && row[header] !== ""));
    if (field) {
      const counts = new Map();
      rows.forEach((row) => {
        const value = String(row[field] || "");
        if (value) counts.set(value, (counts.get(value) || 0) + 1);
      });
      const mostCommon = Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0];
      if (mostCommon) insights.push({ title: `Most common ${field}`, text: `"${mostCommon[0]}" appears ${mostCommon[1]} times.` });
    }
  }
  return {
    kind: "table",
    headers,
    rows: rows.slice(0, 1000),
    totalRows: rows.length,
    summary,
    insights,
    ...extra
  };
}

function textAnalysis(text, fileName, kind = "document") {
  const cleanText = String(text || "").replace(/\u0000/g, "").trim();
  if (!cleanText) throw httpError(422, "No readable text was found in this file.");
  const words = cleanText.match(/\S+/g) || [];
  const paragraphs = cleanText.split(/\n\s*\n/).filter((item) => item.trim()).length;
  const headers = ["Metric", "Value"];
  const rows = [
    { Metric: "Words", Value: words.length },
    { Metric: "Characters", Value: cleanText.length },
    { Metric: "Paragraphs", Value: paragraphs }
  ];
  return analyzeTable(headers, rows, {
    kind,
    text: cleanText.slice(0, 20000),
    insights: [
      { title: "Text extracted", text: `${words.length.toLocaleString()} words across ${paragraphs.toLocaleString()} paragraphs.` },
      { title: "Source", text: `Readable text was extracted from ${fileName}.` }
    ]
  });
}

function parseJson(buffer) {
  let data;
  try {
    data = JSON.parse(buffer.toString("utf8").replace(/^\uFEFF/, ""));
  } catch {
    throw httpError(422, "This file contains invalid JSON.");
  }
  if (data && !Array.isArray(data) && Array.isArray(data.data)) data = data.data;
  if (Array.isArray(data) && data.length && data.every((item) => Array.isArray(item))) {
    return normalizeTable(data);
  }
  if (Array.isArray(data) && data.length && data.every((item) => item && typeof item === "object")) {
    const headers = Array.from(new Set(data.flatMap((item) => Object.keys(item))));
    const matrix = [headers, ...data.map((item) => headers.map((header) => item[header]))];
    return normalizeTable(matrix);
  }
  throw httpError(422, "JSON must contain an array of objects or an array of rows.");
}

function parseDelimited(buffer, delimiter) {
  let matrix;
  try {
    matrix = parseCsv(buffer, {
      bom: true,
      delimiter,
      skip_empty_lines: true,
      relax_column_count: true,
      trim: true
    });
  } catch {
    throw httpError(422, "The delimited file could not be parsed. Check its quoting and separator.");
  }
  return normalizeTable(matrix);
}

async function parseExcel(buffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
  } catch {
    throw httpError(422, "This Excel workbook could not be opened. Upload a valid .xlsx file.");
  }
  const worksheet = workbook.worksheets.find((sheet) => sheet.actualRowCount > 0);
  if (!worksheet) throw httpError(422, "This workbook does not contain a populated worksheet.");
  const matrix = [];
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    matrix.push(Array.from({ length: row.cellCount }, (_, index) => cellValue(row.getCell(index + 1).value)));
  });
  return normalizeTable(matrix);
}

async function parsePdf(buffer, fileName) {
  const { PDFParse } = await import("pdf-parse");
  const parser = new PDFParse({ data: new Uint8Array(buffer) });
  try {
    const result = await parser.getText();
    return textAnalysis(result.text, fileName, "document");
  } catch (error) {
    if (error.status) throw error;
    throw httpError(422, "The PDF could not be read. Scanned PDFs need OCR and are not supported yet.");
  } finally {
    await parser.destroy();
  }
}

async function parseImage(buffer) {
  try {
    if (!ocrWorkerPromise) {
      ocrWorkerPromise = createWorker("eng").catch((error) => {
        ocrWorkerPromise = null;
        throw error;
      });
    }
    const worker = await ocrWorkerPromise;
    const result = await worker.recognize(buffer);
    if (!result.data.text.trim()) throw httpError(422, "No readable text was detected in this image.");
    return textAnalysis(result.data.text, "image", "image");
  } catch (error) {
    if (error.status) throw error;
    throw httpError(503, "Image OCR could not start. Check the server connection and retry so Tesseract can download its language data.");
  }
}

async function transcribeAudio(buffer, file) {
  if (!openai) {
    throw httpError(503, "Audio transcription needs an OpenAI API key. Add it to .env and restart the app.");
  }
  try {
    const audioFile = new File([buffer], path.basename(file.originalname), { type: file.mimetype || "application/octet-stream" });
    const response = await openai.audio.transcriptions.create({
      file: audioFile,
      model: process.env.OPENAI_TRANSCRIPTION_MODEL || "gpt-4o-mini-transcribe"
    });
    return textAnalysis(response.text, file.originalname, "audio");
  } catch (error) {
    if (error.status) throw error;
    console.error("Audio transcription failed:", error.message);
    throw httpError(502, "Audio transcription failed. Check the configured OpenAI model, API key, and audio format.");
  }
}

async function analyzeFile(file, transcribeAudioConsent) {
  if (!file) throw httpError(400, "Choose a file to analyze.");
  const extension = path.extname(file.originalname).toLowerCase();
  if (!supportedExtensions.has(extension)) {
    throw httpError(415, "This file format is not supported. Try CSV, TSV, JSON, XLSX, TXT, DOCX, PDF, an image, or supported audio.");
  }
  if (extension === ".csv") return parseDelimited(file.buffer, ",");
  if (extension === ".tsv") return parseDelimited(file.buffer, "\t");
  if (extension === ".json") return parseJson(file.buffer);
  if (extension === ".xlsx") return parseExcel(file.buffer);
  if (extension === ".txt") return textAnalysis(file.buffer.toString("utf8"), file.originalname);
  if (extension === ".docx") {
    try {
      const result = await mammoth.extractRawText({ buffer: file.buffer });
      return textAnalysis(result.value, file.originalname);
    } catch {
      throw httpError(422, "This Word document could not be read. Upload a valid .docx file.");
    }
  }
  if (extension === ".pdf") return parsePdf(file.buffer, file.originalname);
  if (imageExtensions.has(extension)) return parseImage(file.buffer);
  if (audioExtensions.has(extension)) {
    if (transcribeAudioConsent !== true) {
      throw httpError(400, "Audio must be sent to OpenAI for transcription in this MVP. Explicitly enable audio transcription consent and retry.");
    }
    return transcribeAudio(file.buffer, file);
  }
  throw httpError(415, "This file format is not supported.");
}

function apiErrorHandler(error, req, res, next) {
  if (res.headersSent) return next(error);
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    return res.status(413).json({ error: `File is too large. The current limit is ${Math.round(maxFileBytes / 1024 / 1024)} MB.` });
  }
  const status = Number.isInteger(error.status) ? error.status : 500;
  if (status >= 500) console.error("Request failed:", error.message);
  return res.status(status).json({
    error: status >= 500 ? "The analysis service encountered an error. Check the server log for details." : error.message
  });
}

app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    aiConfigured: Boolean(openai),
    limits: { maxFileSizeMb: Math.round(maxFileBytes / 1024 / 1024) },
    capabilities: {
      tabular: [".csv", ".tsv", ".json", ".xlsx"],
      documents: [".txt", ".docx", ".pdf"],
      images: [".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"],
      audio: [".mp3", ".wav", ".m4a", ".mp4", ".mpeg", ".mpga", ".webm"]
    }
  });
});

app.post("/api/analyze", upload.single("file"), async (req, res, next) => {
  try {
    const result = await analyzeFile(req.file, req.body && req.body.transcribeAudio === "true");
    result.aiConfigured = Boolean(openai);
    result.fileName = path.basename(req.file.originalname);
    res.json(result);
  } catch (error) {
    next(error);
  }
});

app.post("/api/insights", async (req, res, next) => {
  try {
    if (!openai) throw httpError(503, "AI insights are not configured. Add an OpenAI API key to .env and restart the app.");
    const analysis = req.body;
    if (!analysis || typeof analysis !== "object" || !Array.isArray(analysis.headers) || !Array.isArray(analysis.rows)) {
      throw httpError(400, "A valid dataset analysis is required.");
    }
    const context = {
      kind: String(analysis.kind || "table").slice(0, 30),
      summary: analysis.summary || {},
      headers: analysis.headers.slice(0, 80),
      rows: analysis.rows.slice(0, 20),
      extractedText: typeof analysis.text === "string" ? analysis.text.slice(0, 6000) : ""
    };
    const response = await openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || "gpt-4o-mini",
      temperature: 0.2,
      messages: [
        {
          role: "system",
          content: "You are a careful data analyst. Report only patterns supported by the supplied data. Mention limitations and uncertainty. Do not invent causal claims or recommendations unsupported by the data. Use concise plain text with a short headline and 2-4 bullet points."
        },
        { role: "user", content: `Analyze this user-provided data and summarize useful, cautious insights:\n${JSON.stringify(context)}` }
      ]
    });
    const text = response.choices[0] && response.choices[0].message.content;
    if (!text) throw httpError(502, "The AI provider returned an empty analysis.");
    res.json({ text });
  } catch (error) {
    if (!error.status && error instanceof Error) {
      console.error("AI insights failed:", error.message);
      error.status = 502;
      error.message = "AI analysis failed. Check the API key, configured model, and provider response.";
    }
    next(error);
  }
});

app.get("/", (req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.get("/index.html", (req, res) => res.sendFile(path.join(__dirname, "index.html")));
app.use(apiErrorHandler);

if (require.main === module) {
  app.listen(port, "127.0.0.1", () => {
    console.log(`OmniData AI is running at http://localhost:${port}`);
    console.log(`AI provider ${openai ? "configured" : "not configured (local analysis only)"}.`);
  });
}

module.exports = { app, analyzeTable, normalizeTable, textAnalysis };
