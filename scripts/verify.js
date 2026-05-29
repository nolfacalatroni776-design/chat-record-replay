#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const Analytics = require("../assets/analytics.js");

const demoPath = path.join(__dirname, "../data/demo-records.json");
const records = JSON.parse(fs.readFileSync(demoPath, "utf8"));
const analysis = Analytics.analyzeRecords(records, { fileName: "demo-records.json" });

const expected = {
  messages: 12,
  projects: 2,
  groups: 3
};

const result = {
  messages: analysis.totals.messages,
  projects: analysis.totals.projects,
  groups: analysis.totals.groups,
  topProject: analysis.projectStats[0] && analysis.projectStats[0].name,
  topGroup: analysis.groupStats[0] && analysis.groupStats[0].name,
  dateRange: Analytics.dateRangeText(analysis)
};

console.log(JSON.stringify(result, null, 2));

Object.entries(expected).forEach(([key, value]) => {
  if (result[key] !== value) {
    throw new Error(`${key} expected ${value}, got ${result[key]}`);
  }
});
