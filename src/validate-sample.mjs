import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const samplePath = path.resolve(scriptDir, "..", "examples", "sample_jobs.csv");

const requiredColumns = [
  "discovered_date",
  "employer",
  "job_title",
  "location",
  "work_mode",
  "application_url",
  "posting_date",
  "closing_date",
  "match_score",
  "priority",
  "status",
  "salary_eur",
  "source_url",
  "notes",
];

const allowedWorkModes = new Set(["On-site", "Hybrid", "Remote", "Flexible"]);
const allowedPriorities = new Set([
  "Apply immediately",
  "Strong option",
  "Possible option",
  "Low priority",
  "Do not apply",
]);
const allowedStatuses = new Set([
  "Open",
  "Saved",
  "Applied",
  "Interview",
  "Offer",
  "Closed",
  "Ineligible",
  "Needs verification",
]);

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

  if (quoted) {
    throw new Error("CSV contains an unclosed quoted field.");
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }

  const headers = rows.shift();
  if (!headers) {
    throw new Error("CSV is empty.");
  }
  const duplicateHeaders = headers.filter((header, index) => headers.indexOf(header) !== index);
  if (duplicateHeaders.length) {
    throw new Error(`CSV has duplicate header(s): ${[...new Set(duplicateHeaders)].join(", ")}`);
  }
  if (headers.length !== requiredColumns.length || requiredColumns.some((column, index) => headers[index] !== column)) {
    throw new Error("CSV headers must match the documented sample schema exactly.");
  }

  return rows
    .filter((values) => values.some(Boolean))
    .map((values, index) => {
      if (values.length !== headers.length) {
        throw new Error(`Row ${index + 2} has ${values.length} values; expected ${headers.length}.`);
      }
      return Object.fromEntries(headers.map((header, column) => [header, values[column]]));
    });
}

function isIsoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value) && !Number.isNaN(Date.parse(`${value}T00:00:00Z`));
}

function assertExampleUrl(value, field, rowNumber) {
  let url;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`Row ${rowNumber}: ${field} must be a valid URL.`);
  }
  if (url.protocol !== "https:" || url.hostname !== "example.com") {
    throw new Error(`Row ${rowNumber}: ${field} must use the fictional https://example.com domain.`);
  }
}

const rows = parseCsv(await fs.readFile(samplePath, "utf8"));
if (!rows.length) {
  throw new Error("CSV needs at least one fictional sample row.");
}

const duplicateKeys = new Set();
for (const [index, row] of rows.entries()) {
  const rowNumber = index + 2;
  for (const [field, value] of Object.entries(row)) {
    if (!value.trim()) {
      throw new Error(`Row ${rowNumber}: ${field} must not be blank.`);
    }
  }
  for (const field of ["discovered_date", "posting_date", "closing_date"]) {
    if (!isIsoDate(row[field])) {
      throw new Error(`Row ${rowNumber}: ${field} must be an ISO date (YYYY-MM-DD).`);
    }
  }
  if (Date.parse(`${row.closing_date}T00:00:00Z`) < Date.parse(`${row.posting_date}T00:00:00Z`)) {
    throw new Error(`Row ${rowNumber}: closing_date cannot be before posting_date.`);
  }
  if (!allowedWorkModes.has(row.work_mode)) {
    throw new Error(`Row ${rowNumber}: work_mode is not in the documented list.`);
  }
  if (!allowedPriorities.has(row.priority) || !allowedStatuses.has(row.status)) {
    throw new Error(`Row ${rowNumber}: priority or status is not in the documented list.`);
  }
  const matchScore = Number(row.match_score);
  const salary = Number(row.salary_eur);
  if (!Number.isInteger(matchScore) || matchScore < 0 || matchScore > 100) {
    throw new Error(`Row ${rowNumber}: match_score must be an integer from 0 to 100.`);
  }
  if (!Number.isFinite(salary) || salary < 0) {
    throw new Error(`Row ${rowNumber}: salary_eur must be a non-negative number.`);
  }
  assertExampleUrl(row.application_url, "application_url", rowNumber);
  assertExampleUrl(row.source_url, "source_url", rowNumber);
  if (!/fictional/i.test(row.notes)) {
    throw new Error(`Row ${rowNumber}: notes must state that the row is fictional.`);
  }
  const duplicateKey = [row.employer, row.job_title, row.application_url].map((value) => value.toLowerCase()).join("|");
  if (duplicateKeys.has(duplicateKey)) {
    throw new Error(`Row ${rowNumber}: duplicate employer, title, and application URL.`);
  }
  duplicateKeys.add(duplicateKey);
}

console.log(`Validated ${rows.length} fictional public sample row(s).`);
