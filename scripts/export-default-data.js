#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const Analytics = require("../assets/analytics.js");

const sourcePath = process.argv[2] || "/Users/user/Downloads/2026-05-29消息记录.xlsx";
const outputPath = process.argv[3] || path.join(__dirname, "../data/default-records.json");

const workbook = XLSX.read(fs.readFileSync(sourcePath), { type: "buffer", cellDates: false, dense: false });
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const rows = Analytics.rowsFromWorksheet(XLSX, worksheet);
const parsed = Analytics.recordsFromRows(rows);

if (!parsed.headerOk) {
  throw new Error(`Header mismatch: ${parsed.headerMissing.join(", ")}`);
}

fs.writeFileSync(outputPath, `${JSON.stringify(parsed.records)}\n`);

const analysis = Analytics.analyzeRecords(parsed.records, {
  fileName: path.basename(sourcePath),
  sheetName,
  rowsRead: rows.length
});

console.log(JSON.stringify({
  outputPath,
  sheetName,
  records: parsed.records.length,
  projects: analysis.totals.projects,
  groups: analysis.totals.groups,
  users: analysis.totals.users,
  topProject: analysis.projectStats[0] && analysis.projectStats[0].name,
  topProjectMessages: analysis.projectStats[0] && analysis.projectStats[0].messages,
  topGroup: analysis.groupStats[0] && analysis.groupStats[0].name,
  topGroupMessages: analysis.groupStats[0] && analysis.groupStats[0].messages,
  dateRange: Analytics.dateRangeText(analysis)
}, null, 2));
