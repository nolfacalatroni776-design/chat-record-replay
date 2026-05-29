#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const Analytics = require("../assets/analytics.js");

const defaultPath = path.join(__dirname, "../data/default-records.json");
const demoPath = path.join(__dirname, "../data/demo-records.json");
const records = JSON.parse(fs.readFileSync(defaultPath, "utf8"));
const demoRecords = JSON.parse(fs.readFileSync(demoPath, "utf8"));
const analysis = Analytics.analyzeRecords(records, { fileName: "default-records.json" });
const demoAnalysis = Analytics.analyzeRecords(demoRecords, { fileName: "demo-records.json" });

const expected = {
  messages: 4995,
  projects: 118,
  groups: 191,
  topProject: "5b3359ce-5ad7-44cc-b27c-a2058941f7cd",
  topGroup: "[A12852]TTS/MOS project_Turkish"
};

const result = {
  messages: analysis.totals.messages,
  projects: analysis.totals.projects,
  groups: analysis.totals.groups,
  topProject: analysis.projectStats[0] && analysis.projectStats[0].name,
  topGroup: analysis.groupStats[0] && analysis.groupStats[0].name,
  dateRange: Analytics.dateRangeText(analysis),
  demoMessages: demoAnalysis.totals.messages,
  demoProjects: demoAnalysis.totals.projects
};

console.log(JSON.stringify(result, null, 2));

Object.entries(expected).forEach(([key, value]) => {
  if (result[key] !== value) {
    throw new Error(`${key} expected ${value}, got ${result[key]}`);
  }
});
