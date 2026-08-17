import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const sampleCsvPath = path.join(repoRoot, "examples", "sample_jobs.csv");
const outputPath = process.argv[2] ?? path.join(repoRoot, "output", "Job_Search_Tracker_Sample.xlsx");
const assetDir = path.join(repoRoot, "assets");

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.mkdir(assetDir, { recursive: true });

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"') {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (character !== "\r") {
      field += character;
    }
  }

  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }

  const headers = rows.shift();
  return rows
    .filter((values) => values.some(Boolean))
    .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])));
}

function toDate(value) {
  return value ? new Date(value + "T00:00:00Z") : null;
}

function addTextRule(range, text, fill, color, bold = false) {
  range.conditionalFormats.add("containsText", {
    text,
    format: { fill, font: { color, bold } },
  });
}

const sampleRows = parseCsv(await fs.readFile(sampleCsvPath, "utf8"));
const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Dashboard");
const tracker = workbook.worksheets.add("Tracker");
const lists = workbook.worksheets.add("Lists");

for (const sheet of [dashboard, tracker, lists]) {
  sheet.showGridLines = false;
}

// Auditable dropdown sources.
lists.getRange("A1:C1").values = [["Work mode", "Priority", "Status"]];
lists.getRange("A2:A5").values = [["On-site"], ["Hybrid"], ["Remote"], ["Flexible"]];
lists.getRange("B2:B6").values = [["Apply immediately"], ["Strong option"], ["Possible option"], ["Low priority"], ["Do not apply"]];
lists.getRange("C2:C9").values = [["Open"], ["Saved"], ["Applied"], ["Interview"], ["Offer"], ["Closed"], ["Ineligible"], ["Needs verification"]];
lists.getRange("A1:C1").format = {
  fill: "#1F2937",
  font: { bold: true, color: "#FFFFFF" },
};
lists.getRange("A1:C9").format.borders = {
  insideHorizontal: { style: "thin", color: "#E5E7EB" },
};
lists.getRange("A1:A9").format.columnWidth = 18;
lists.getRange("B1:C9").format.columnWidth = 22;
lists.freezePanes.freezeRows(1);

const columns = [
  "Date discovered",
  "Employer",
  "Job title",
  "Location",
  "Work mode",
  "Application URL",
  "Posting date",
  "Closing date",
  "Days open",
  "Freshness",
  "Match score",
  "Priority",
  "Status",
  "Salary (EUR)",
  "Source URL",
  "Notes",
  "Duplicate check",
];

tracker.getRange("A1:Q1").values = [columns];
tracker.getRange("A1:Q1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  rowHeight: 36,
  borders: { preset: "outside", style: "thin", color: "#0F766E" },
};

const values = sampleRows.map((item) => [
  toDate(item.discovered_date),
  item.employer,
  item.job_title,
  item.location,
  item.work_mode,
  item.application_url,
  toDate(item.posting_date),
  toDate(item.closing_date),
  null,
  null,
  Number(item.match_score),
  item.priority,
  item.status,
  Number(item.salary_eur),
  item.source_url,
  item.notes,
  null,
]);

if (values.length) {
  tracker.getRangeByIndexes(1, 0, values.length, columns.length).values = values;
}

tracker.getRange("I2").formulas = [['=IF(A2="","",MAX(0,TODAY()-A2))']];
tracker.getRange("I2:I101").fillDown();
tracker.getRange("J2").formulas = [['=IF(A2="","",IF(TODAY()-A2<=1,"Today",IF(TODAY()-A2<=7,"Last 7 days","Older")))']];
tracker.getRange("J2:J101").fillDown();
tracker.getRange("Q2").formulas = [['=IF(OR(B2="",C2="",F2=""),"",IF(COUNTIFS($B$2:B2,B2,$C$2:C2,C2,$F$2:F2,F2)>1,"Duplicate",""))']];
tracker.getRange("Q2:Q101").fillDown();

tracker.getRange("A2:A101").format.numberFormat = "yyyy-mm-dd";
tracker.getRange("G2:H101").format.numberFormat = "yyyy-mm-dd";
tracker.getRange("I2:I101").format.numberFormat = "0";
tracker.getRange("K2:K101").format.numberFormat = "0";
tracker.getRange("N2:N101").format.numberFormat = '"â‚¬"#,##0';
tracker.getRange("A2:Q101").format = {
  font: { color: "#111827", size: 10 },
  wrapText: true,
  rowHeight: 30,
};

const widths = [15, 22, 28, 20, 13, 31, 14, 14, 11, 15, 12, 20, 18, 15, 31, 36, 17];
const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];
widths.forEach((width, index) => {
  tracker.getRange(letters[index] + "1:" + letters[index] + "101").format.columnWidth = width;
});

const table = tracker.tables.add("A1:Q101", true, "JobTrackerTable");
table.style = "TableStyleMedium4";
table.showFilterButton = true;
table.showBandedRows = true;

tracker.freezePanes.freezeRows(1);
tracker.freezePanes.freezeColumns(3);

tracker.getRange("E2:E101").dataValidation = {
  rule: { type: "list", formula1: "'Lists'!$A$2:$A$5" },
};
tracker.getRange("K2:K101").dataValidation = {
  rule: { type: "whole", operator: "between", formula1: 0, formula2: 100 },
};
tracker.getRange("L2:L101").dataValidation = {
  rule: { type: "list", formula1: "'Lists'!$B$2:$B$6" },
};
tracker.getRange("M2:M101").dataValidation = {
  rule: { type: "list", formula1: "'Lists'!$C$2:$C$9" },
};

addTextRule(tracker.getRange("J2:J101"), "Today", "#DBEAFE", "#1D4ED8", true);
addTextRule(tracker.getRange("J2:J101"), "Last 7 days", "#FEF3C7", "#92400E");
addTextRule(tracker.getRange("J2:J101"), "Older", "#F1F5F9", "#475569");
tracker.getRange("K2:K101").conditionalFormats.add("colorScale", {
  colors: ["#FEE2E2", "#FEF3C7", "#DCFCE7"],
  thresholds: ["min", "50%", "max"],
});
addTextRule(tracker.getRange("L2:L101"), "Apply immediately", "#FFE4E6", "#BE123C", true);
addTextRule(tracker.getRange("L2:L101"), "Strong option", "#CCFBF1", "#0F766E", true);
addTextRule(tracker.getRange("M2:M101"), "Applied", "#DBEAFE", "#1D4ED8", true);
addTextRule(tracker.getRange("M2:M101"), "Interview", "#EDE9FE", "#6D28D9", true);
addTextRule(tracker.getRange("Q2:Q101"), "Duplicate", "#FEE2E2", "#991B1B", true);
tracker.getRange("H2:H101").conditionalFormats.addCustom('=AND($H2<>"",$H2>=TODAY(),$H2<=TODAY()+7)', {
  fill: "#FED7AA",
  font: { color: "#9A3412", bold: true },
});

// Dashboard.
dashboard.getRange("A1:J1").merge();
dashboard.getRange("A1").values = [["Ireland Data & Analytics Job Tracker"]];
dashboard.getRange("A1:J1").format = {
  fill: "#111827",
  font: { bold: true, color: "#FFFFFF", size: 20 },
  rowHeight: 34,
};
dashboard.getRange("A2:J2").merge();
dashboard.getRange("A2").values = [["Portfolio-safe sample using fictional vacancies. Keep real application data in a private local copy."]];
dashboard.getRange("A2:J2").format = {
  fill: "#E0F2FE",
  font: { color: "#0F172A", italic: true },
  rowHeight: 28,
};

for (const range of ["A4:B4", "C4:D4", "E4:F4", "G4:H4", "I4:J4", "A5:B5", "C5:D5", "E5:F5", "G5:H5", "I5:J5"]) {
  dashboard.getRange(range).merge();
}
dashboard.getRange("A4:J4").values = [[
  "Total tracked", null,
  "Apply immediately", null,
  "Fresh this week", null,
  "Closing in 7 days", null,
  "Applied", null,
]];
dashboard.getRange("A5:J5").formulas = [[
  "=COUNTA('Tracker'!$B$2:$B$101)", null,
  '=COUNTIF(\'Tracker\'!$L$2:$L$101,"Apply immediately")', null,
  '=COUNTIF(\'Tracker\'!$J$2:$J$101,"Last 7 days")+COUNTIF(\'Tracker\'!$J$2:$J$101,"Today")', null,
  '=COUNTIFS(\'Tracker\'!$H$2:$H$101,">="&TODAY(),\'Tracker\'!$H$2:$H$101,"<="&TODAY()+7)', null,
  '=COUNTIF(\'Tracker\'!$M$2:$M$101,"Applied")', null,
]];
dashboard.getRange("A4:J5").format = {
  fill: "#F8FAFC",
  font: { color: "#111827", bold: true },
  borders: { preset: "all", style: "thin", color: "#CBD5E1" },
  rowHeight: 28,
};
dashboard.getRange("A5:J5").format.font = { bold: true, color: "#0F172A", size: 18 };

dashboard.getRange("A7:C7").merge();
dashboard.getRange("A7").values = [["Freshness pipeline"]];
dashboard.getRange("A7:C7").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
};
dashboard.getRange("A8:C11").values = [
  ["Bucket", "Count", "Suggested action"],
  ["Today", null, "Review first"],
  ["Last 7 days", null, "Prioritise strong matches"],
  ["Older", null, "Check only if fit is compelling"],
];
dashboard.getRange("B9:B11").formulas = [
  ['=COUNTIF(\'Tracker\'!$J$2:$J$101,A9)'],
  ['=COUNTIF(\'Tracker\'!$J$2:$J$101,A10)'],
  ['=COUNTIF(\'Tracker\'!$J$2:$J$101,A11)'],
];

dashboard.getRange("E7:G7").merge();
dashboard.getRange("E7").values = [["Priority snapshot"]];
dashboard.getRange("E7:G7").format = {
  fill: "#2563EB",
  font: { bold: true, color: "#FFFFFF" },
};
dashboard.getRange("E8:G12").values = [
  ["Priority", "Count", "Meaning"],
  ["Apply immediately", null, "Best timing and match"],
  ["Strong option", null, "Likely worth applying"],
  ["Possible option", null, "Worth a second look"],
  ["Low priority", null, "Track only"],
];
dashboard.getRange("F9:F12").formulas = [
  ['=COUNTIF(\'Tracker\'!$L$2:$L$101,E9)'],
  ['=COUNTIF(\'Tracker\'!$L$2:$L$101,E10)'],
  ['=COUNTIF(\'Tracker\'!$L$2:$L$101,E11)'],
  ['=COUNTIF(\'Tracker\'!$L$2:$L$101,E12)'],
];

dashboard.getRange("I7:J7").merge();
dashboard.getRange("I7").values = [["Status snapshot"]];
dashboard.getRange("I7:J7").format = {
  fill: "#7C3AED",
  font: { bold: true, color: "#FFFFFF" },
};
dashboard.getRange("I8:J12").values = [
  ["Status", "Count"],
  ["Open", null],
  ["Saved", null],
  ["Applied", null],
  ["Interview", null],
];
dashboard.getRange("J9:J12").formulas = [
  ['=COUNTIF(\'Tracker\'!$M$2:$M$101,I9)'],
  ['=COUNTIF(\'Tracker\'!$M$2:$M$101,I10)'],
  ['=COUNTIF(\'Tracker\'!$M$2:$M$101,I11)'],
  ['=COUNTIF(\'Tracker\'!$M$2:$M$101,I12)'],
];

for (const range of ["A8:C8", "E8:G8", "I8:J8"]) {
  dashboard.getRange(range).format = {
    fill: "#1F2937",
    font: { bold: true, color: "#FFFFFF" },
  };
}
for (const range of ["A9:C11", "E9:G12", "I9:J12"]) {
  dashboard.getRange(range).format = {
    fill: "#FFFFFF",
    borders: { preset: "all", style: "thin", color: "#E5E7EB" },
  };
}

dashboard.getRange("A15:J15").merge();
dashboard.getRange("A15").values = [["Recommended review order"]];
dashboard.getRange("A15:J15").format = {
  fill: "#7C3AED",
  font: { bold: true, color: "#FFFFFF" },
};
dashboard.getRange("A16:J19").values = [
  ["1", "Freshness", "Today, then Last 7 days", null, null, null, null, null, null, null],
  ["2", "Match score", "Highest first", null, null, null, null, null, null, null],
  ["3", "Closing date", "Soonest valid deadline first", null, null, null, null, null, null, null],
  ["4", "Application status", "Avoid duplicate effort and follow up consistently", null, null, null, null, null, null, null],
];
dashboard.getRange("C16:J19").merge(true);
dashboard.getRange("A16:J19").format = {
  fill: "#FAFAFA",
  borders: { preset: "all", style: "thin", color: "#E5E7EB" },
  wrapText: true,
  rowHeight: 30,
};
dashboard.getRange("A16:A19").format = {
  fill: "#F5F3FF",
  font: { bold: true, color: "#5B21B6" },
};

dashboard.getRange("A21:J21").merge();
dashboard.getRange("A21").values = [["Responsible-use notes"]];
dashboard.getRange("A21:J21").format = {
  fill: "#B45309",
  font: { bold: true, color: "#FFFFFF" },
};
dashboard.getRange("A22:J24").values = [
  ["Verify that each vacancy is still open before applying.", null, null, null, null, null, null, null, null, null],
  ["Keep application submission and recruiter outreach manual.", null, null, null, null, null, null, null, null, null],
  ["Store personal CV, contact, and immigration details outside a public repository.", null, null, null, null, null, null, null, null, null],
];
dashboard.getRange("A22:J24").merge(true);
dashboard.getRange("A22:J24").format = {
  fill: "#FFF7ED",
  font: { color: "#7C2D12" },
  borders: { preset: "all", style: "thin", color: "#FDBA74" },
  wrapText: true,
  rowHeight: 24,
};

const dashboardWidths = [14, 16, 20, 4, 20, 12, 24, 4, 20, 12];
dashboardWidths.forEach((width, index) => {
  dashboard.getRange(letters[index] + "1:" + letters[index] + "24").format.columnWidth = width;
});

addTextRule(dashboard.getRange("A9:A11"), "Today", "#DBEAFE", "#1D4ED8", true);
addTextRule(dashboard.getRange("E9:E12"), "Apply immediately", "#FFE4E6", "#BE123C", true);
addTextRule(dashboard.getRange("E9:E12"), "Strong option", "#CCFBF1", "#0F766E", true);

const dashboardCheck = await workbook.inspect({
  kind: "table",
  sheetId: "Dashboard",
  range: "A1:J24",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 10,
  maxChars: 5000,
});
console.log(dashboardCheck.ndjson);

const trackerCheck = await workbook.inspect({
  kind: "table",
  sheetId: "Tracker",
  range: "A1:Q10",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 17,
  maxChars: 5000,
});
console.log(trackerCheck.ndjson);

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 2000,
});
console.log(formulaErrors.ndjson);

for (const [sheetName, range, filename] of [
  ["Dashboard", "A1:J24", "dashboard.png"],
  ["Tracker", "A1:Q10", "tracker.png"],
  ["Lists", "A1:C9", "lists.png"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(path.join(assetDir, filename), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log("Saved " + outputPath);

