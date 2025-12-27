/**
 * Google Sheets Automation — practical snippets
 * Use case: simple CRM/lead log + status workflow + reporting
 *
 * How to use:
 * 1) Create a Google Sheet
 * 2) Extensions -> Apps Script -> paste this file
 * 3) Reload the spreadsheet to see the custom menu "Automation"
 */

const CFG = {
  SHEET_LEADS: "Leads",
  SHEET_LOGS: "Logs",
  TZ: Session.getScriptTimeZone(),
  STATUSES: ["NEW", "IN_PROGRESS", "DONE", "CLOSED"],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Automation")
    .addItem("Setup sheets", "setupSheets")
    .addSeparator()
    .addItem("Add test lead", "addTestLead")
    .addItem("Generate weekly report", "generateWeeklyReport")
    .addSeparator()
    .addItem("Archive DONE leads", "archiveDoneLeads")
    .addToUi();
}

/**
 * Creates base sheets and headers if not present.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const leads = getOrCreateSheet_(ss, CFG.SHEET_LEADS);
  const logs = getOrCreateSheet_(ss, CFG.SHEET_LOGS);

  ensureHeaders_(leads, [
    "ID",
    "CreatedAt",
    "Name",
    "Phone",
    "Source",
    "Message",
    "Status",
    "Assignee",
    "LastUpdate",
  ]);

  ensureHeaders_(logs, ["Timestamp", "Action", "Details"]);

  log_("SETUP", "Sheets initialized");
}

/**
 * Adds a lead row with unique ID and default status.
 * You can call this from a webhook integration or manual input.
 */
function addLead(name, phone, source, message, assignee) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, CFG.SHEET_LEADS);

  if (sheet.getLastRow() === 0) setupSheets();

  const id = makeId_();
  const now = new Date();

  const row = [
    id,
    now,
    name || "",
    phone || "",
    source || "manual",
    message || "",
    "NEW",
    assignee || "",
    now,
  ];

  sheet.appendRow(row);
  log_("ADD_LEAD", `ID=${id} name=${name || ""} phone=${phone || ""} source=${source || "manual"}`);

  return id;
}

/**
 * Demo: Adds a realistic test lead.
 */
function addTestLead() {
  const id = addLead(
    "Test User",
    "+48 600 000 000",
    "telegram",
    "Hello! I want to rent a car. What are the conditions?",
    "manager_1"
  );

  SpreadsheetApp.getUi().alert(`Test lead added. ID: ${id}`);
}

/**
 * Updates status by Lead ID and touches LastUpdate.
 */
function updateLeadStatus(id, newStatus) {
  if (!id) throw new Error("ID is required");
  if (CFG.STATUSES.indexOf(newStatus) === -1) {
    throw new Error(`Invalid status: ${newStatus}. Allowed: ${CFG.STATUSES.join(", ")}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, CFG.SHEET_LEADS);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("No leads found");

  const header = data[0];
  const idxId = header.indexOf("ID");
  const idxStatus = header.indexOf("Status");
  const idxLastUpdate = header.indexOf("LastUpdate");

  if (idxId < 0 || idxStatus < 0 || idxLastUpdate < 0) {
    throw new Error("Headers not found. Run 'Setup sheets' first.");
  }

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idxId]) === String(id)) {
      sheet.getRange(r + 1, idxStatus + 1).setValue(newStatus);
      sheet.getRange(r + 1, idxLastUpdate + 1).setValue(new Date());
      log_("UPDATE_STATUS", `ID=${id} -> ${newStatus}`);
      return true;
    }
  }

  throw new Error(`Lead with ID=${id} not found`);
}

/**
 * Creates/updates a weekly report sheet:
 * counts leads by status for the last 7 days.
 */
function generateWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leads = getOrCreateSheet_(ss, CFG.SHEET_LEADS);
  if (leads.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No leads to report.");
    return;
  }

  const reportName = `Report_${formatDate_(new Date())}`;
  const report = getOrCreateSheet_(ss, reportName);

  report.clear();
  report.getRange(1, 1, 1, 4).setValues([["Status", "Count (7d)", "From", "To"]]);

  const now = new Date();
  const from = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  const data = leads.getDataRange().getValues();
  const header = data[0];

  const idxCreated = header.indexOf("CreatedAt");
  const idxStatus = header.indexOf("Status");

  if (idxCreated < 0 || idxStatus < 0) {
    throw new Error("Headers not found. Run 'Setup sheets' first.");
  }

  const counts = {};
  CFG.STATUSES.forEach((s) => (counts[s] = 0));

  for (let r = 1; r < data.length; r++) {
    const created = data[r][idxCreated];
    const status = String(data[r][idxStatus] || "NEW");

    if (created instanceof Date && created >= from && created <= now) {
      if (!counts[status]) counts[status] = 0;
      counts[status]++;
    }
  }

  const rows = [];
  Object.keys(counts).forEach((status) => {
    rows.push([status, counts[status], from, now]);
  });

  report.getRange(2, 1, rows.length, 4).setValues(rows);
  report.autoResizeColumns(1, 4);

  log_("REPORT", `Weekly report generated: ${reportName}`);
  SpreadsheetApp.getUi().alert(`Weekly report generated: ${reportName}`);
}

/**
 * Moves DONE leads older than 7 days to "Archive" sheet.
 */
function archiveDoneLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leads = getOrCreateSheet_(ss, CFG.SHEET_LEADS);
  if (leads.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No leads to archive.");
    return;
  }

  const archive = getOrCreateSheet_(ss, "Archive");

  const data = leads.getDataRange().getValues();
  const header = data[0];

  ensureHeaders_(archive, header);

  const idxStatus = header.indexOf("Status");
  const idxCreated = header.indexOf("CreatedAt");

  if (idxStatus < 0 || idxCreated < 0) {
    throw new Error("Headers not found. Run 'Setup sheets' first.");
  }

  const now = new Date();
  const cutoff = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  const toArchive = [];
  const rowsToDelete = [];

  for (let r = 1; r < data.length; r++) {
    const status = String(data[r][idxStatus] || "");
    const created = data[r][idxCreated];

    if (status === "DONE" && created instanceof Date && created < cutoff) {
      toArchive.push(data[r]);
      rowsToDelete.push(r + 1); // sheet rows are 1-based
    }
  }

  if (toArchive.length === 0) {
    SpreadsheetApp.getUi().alert("Nothing to archive.");
    return;
  }

  // Append to archive
  archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);

  // Delete from bottom to top to keep indices valid
  rowsToDelete.sort((a, b) => b - a).forEach((rowNum) => leads.deleteRow(rowNum));

  log_("ARCHIVE", `Archived ${toArchive.length} DONE leads`);
  SpreadsheetApp.getUi().alert(`Archived ${toArchive.length} leads to Archive sheet.`);
}

/* =========================
   Helpers
========================= */

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  const existing = range.getValues()[0];

  const isEmpty = existing.every((v) => !v);
  if (sheet.getLastRow() === 0 || isEmpty) {
    range.setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function log_(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = getOrCreateSheet_(ss, CFG.SHEET_LOGS);
  if (logs.getLastRow() === 0) setupSheets();

  logs.appendRow([new Date(), action, details]);
}

function makeId_() {
  // Example: L-20251227-AB12CD
  const date = formatDate_(new Date());
  const rand = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `L-${date}-${rand}`;
}

function formatDate_(d) {
  return Utilities.formatDate(d, CFG.TZ, "yyyyMMdd");
}
