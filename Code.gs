// ============================================================
//  Health Office Sankhuwasabha – Training Registration Backend
//  Google Apps Script  |  Paste into script.google.com
// ============================================================

const SHEET_NAME = "Submissions";

// Notification recipients (both will be emailed on every new submission)
const ADMIN_EMAILS = [
  "dpesh.stha2016@gmail.com",
  "dhosankhuwasabha@gmail.com"
];

// Column order — matches form.html field order exactly
const COLUMNS = [
  "submitted_at", "status",
  // Training Information
  "training_name", "role", "training_site", "training_province", "training_district",
  "start_date", "end_date", "fiscal_year",
  // Personal Information
  "name_english", "name_nepali", "sex", "sex_other", "dob_bs",
  // Permanent Address
  "perm_province", "perm_district", "perm_local_level", "perm_ward",
  "contact", "email",
  // Caste
  "caste", "caste_other",
  // Cadre
  "cadre", "cadre_other", "qualification",
  // Sponsored
  "sponsor", "sponsor_details",
  // Working Place
  "work_office", "work_district", "work_province", "work_local_level",
  "work_contact", "designation", "level",
  "pis_no", "citizenship", "council_reg"
];

// ── Entry points ────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const action = (e && e.parameter && e.parameter.action) || "get";
  let result;
  try {
    if (action === "get")  result = getData();
    if (action === "add")  result = addRecord(e);
    if (action === "init") result = initSheet();
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Return ONLY approved rows (public portfolio) ────────────
function getData() {
  const sheet = getOrCreateSheet();
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return { data: [] };

  const headers = rows[0];
  const data = rows.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    })
    .filter(r => String(r.status).toLowerCase() === "approved");

  return { data };
}

// ── Add new record with status = Pending ───────────────────
function addRecord(e) {
  let body = {};
  try {
    if (e.postData && e.postData.contents)
      body = JSON.parse(e.postData.contents);
    else if (e.parameter)
      body = e.parameter;
  } catch(_) { body = e.parameter || {}; }

  // Honeypot — silently drop bot submissions
  if (body.website || body.url_field) return { success: true };

  const sheet = getOrCreateSheet();
  const row   = COLUMNS.map(col => {
    if (col === "submitted_at") return new Date().toLocaleString("en-GB");
    if (col === "status")       return "Pending";
    return body[col] !== undefined ? body[col] : "";
  });

  sheet.appendRow(row);
  notifyAdmins(body);

  return { success: true };
}

// ── Email alert when a new submission arrives ───────────────
function notifyAdmins(data) {
  try {
    const subject = "New Training Registration – " + (data.name_english || "Unknown");
    const body = [
      "A new training registration has been submitted and is awaiting approval.",
      "",
      "Name        : " + (data.name_english || "–"),
      "Designation : " + (data.designation || "–"),
      "Office      : " + (data.work_office || "–"),
      "Local Level : " + (data.work_local_level || "–"),
      "Training    : " + (data.training_name || "–"),
      "Role        : " + (data.role || "–"),
      "Dates       : " + (data.start_date || "–") + " to " + (data.end_date || "–"),
      "Contact     : " + (data.contact || "–"),
      "Email       : " + (data.email || "–"),
      "",
      "To approve:",
      "1. Open the Google Sheet (Submissions tab).",
      "2. Find this row (Status = Pending).",
      "3. Change Status to: Approved",
      "",
      "The record will appear on the public portfolio immediately after approval."
    ].join("\n");

    MailApp.sendEmail(ADMIN_EMAILS.join(","), subject, body);
  } catch(_) {
    // Email is optional — fail silently
  }
}

// ── Initialize sheet with headers + Status dropdown ────────
function initSheet() {
  const sheet = getOrCreateSheet();
  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  styleHeader(sheet);
  return { success: true, message: "Sheet initialized with " + COLUMNS.length + " columns." };
}

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    styleHeader(sheet);
  }
  return sheet;
}

function styleHeader(sheet) {
  const hr = sheet.getRange(1, 1, 1, COLUMNS.length);
  hr.setBackground("#1a3c6e");
  hr.setFontColor("#ffffff");
  hr.setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Highlight Status column (col 2)
  sheet.getRange(1, 2).setBackground("#c0392b");
  sheet.setColumnWidth(2, 100);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pending", "Approved", "Rejected"], true)
    .build();
  sheet.getRange(2, 2, 1000, 1).setDataValidation(rule);
}
