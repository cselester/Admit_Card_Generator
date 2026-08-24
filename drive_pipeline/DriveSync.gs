/**
 * Admit Card Drive Sync + Export Report
 *
 * Run this AFTER generate_admit_cards.py has written PDFs + batch_manifest.xlsx
 * into a folder that Google Drive for Desktop is syncing to your Drive.
 *
 * Setup (one-time):
 * 1. Install Google Drive for Desktop (free) and let it sync a folder,
 *    e.g. "AdmitCards/output_cards", to a Drive folder.
 * 2. Open the spreadsheet you want this bound to (or any spreadsheet) ->
 *    Extensions -> Apps Script. Paste this file in alongside your existing
 *    portal backend (Code.gs from the earlier phase).
 * 3. Set DRIVE_FOLDER_ID and MANIFEST_FILE_NAME below.
 * 4. Reload the spreadsheet -> you'll see an "Admit Cards" menu.
 *
 * What it does:
 * - buildIndexFromFolder(): shares the synced Drive folder once (so every file
 *   inside inherits "anyone with link can view" — no per-file sharing calls,
 *   which is what makes this fast at scale), scans it for PDFs, matches each
 *   to the batch_manifest.xlsx uploaded alongside them, and writes/updates
 *   the "Index" sheet in one batched write:
 *   Registration Number, Student Name, Batch Code, Registration Status,
 *   Center, CBT Center, CBT Center Address, CBT Exam Timings, Url Link,
 *   Mobile Last4, File Name, Source Sheet, Uploaded At.
 *   This is the same Index sheet the student portal (Code.gs) reads from.
 * - exportReport(): copies the Index sheet into a fresh spreadsheet, exports
 *   it as .xlsx into a "Backups" Drive folder, and returns the file — this is
 *   your permanent, downloadable link report for the batch.
 */

// INDEX_SHEET_NAME is declared in Code.gs (both files share scope in Apps Script)
const DRIVE_FOLDER_ID = "PASTE_YOUR_SYNCED_FOLDER_ID_HERE";
const MANIFEST_FILE_NAME = "batch_manifest.xlsx";
const BACKUPS_FOLDER_NAME = "Admit Card Backups";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Admit Cards")
    .addItem("1. Build/update Index from Drive folder", "buildIndexFromFolder")
    .addItem("2. Export link report (.xlsx)", "exportReport")
    .addSeparator()
    .addItem(
      "Backfill missing details for existing rows",
      "backfillMissingDetails",
    )
    .addItem("Debug: list folder contents", "debugListFolderContents")
    .addToUi();
}

function buildIndexFromFolder() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const manifest = readManifest(folder);

  const ui = SpreadsheetApp.getUi();

  if (Object.keys(manifest).length === 0) {
    const proceed = ui.alert(
      "Manifest not found or empty",
      `Couldn't read student details from "${MANIFEST_FILE_NAME}" in this folder. ` +
        "Registration Number and the Drive link will still be filled in, but Student Name, " +
        'Center, etc. will be blank.\n\nRun "Debug: list folder contents" from the menu to ' +
        "check the manifest file actually synced, then try again.\n\nContinue anyway?",
      ui.ButtonSet.YES_NO,
    );
    if (proceed !== ui.Button.YES) return;
  }

  // Share the FOLDER once — files inside inherit it automatically. This is
  // the single most important speed fix: one API call instead of one per
  // file. Calling file.setSharing() per PDF is what made 400 files take ~10
  // minutes, and would make 30k-100k files take hours and blow past Apps
  // Script's execution time limit entirely.
  try {
    folder.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW,
    );
  } catch (e) {
    ui.alert(
      "Couldn't set folder sharing (org policy may block it): " +
        e.message +
        "\n\nSee the note in the README about domain-only sharing as a fallback.",
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INDEX_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INDEX_SHEET_NAME);
    sheet.appendRow([
      "Registration Number",
      "Student Name",
      "Batch Code",
      "Registration Status",
      "Center",
      "CBT Center",
      "CBT Center Address",
      "CBT Exam Timings",
      "Url Link",
      "Mobile Last4",
      "File Name",
      "Source Sheet",
      "Uploaded At",
    ]);
  }

  const existing = sheet.getDataRange().getValues();
  const existingRegNos = new Set(
    existing.slice(1).map((r) => normalizeRegNo(r[0])),
  );

  const files = folder.getFilesByType(MimeType.PDF);
  const newRows = [];
  const now = new Date();
  while (files.hasNext()) {
    const file = files.next();
    const regNo = normalizeRegNo(file.getName().replace(/\.pdf$/i, ""));
    if (existingRegNos.has(regNo)) continue; // already indexed
    existingRegNos.add(regNo); // guard against duplicate filenames in the same run

    const info = manifest[regNo] || {};
    newRows.push([
      regNo,
      info.studentName || "",
      info.batchCode || "",
      info.registrationStatus || "",
      info.center || "",
      info.cbtCenter || "",
      info.cbtCenterAddress || "",
      info.cbtExamTimings || "",
      file.getUrl(),
      info.mobileLast4 || "",
      file.getName(),
      info.sourceSheet || "",
      now,
    ]);
  }

  // One write for the whole batch instead of one appendRow() per file
  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet
      .getRange(startRow, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  ui.alert(`Index updated: ${newRows.length} new admit card(s) linked.`);
}

/** Reads batch_manifest.xlsx (or its auto-converted Google Sheet form) into
 *  {regNo: {studentName, batchCode, ...}}. Handles both cases because Drive
 *  sometimes auto-converts uploaded Office files to native Google formats on
 *  sync (a common Workspace default) — if that happens, the file is already
 *  a Sheet and doesn't need (and will fail) the xlsx->Sheet conversion step. */
function readManifest(folder) {
  const files = folder.getFilesByName(MANIFEST_FILE_NAME);
  if (!files.hasNext()) {
    Logger.log(
      `No file named "${MANIFEST_FILE_NAME}" found in folder "${folder.getName()}".`,
    );
    return {};
  }

  // If multiple copies exist (re-synced runs, conflicted copies), use the newest
  let file = files.next();
  while (files.hasNext()) {
    const next = files.next();
    if (next.getLastUpdated() > file.getLastUpdated()) file = next;
  }
  Logger.log(
    `Using manifest file: ${file.getName()} (${file.getMimeType()}), last updated ${file.getLastUpdated()}`,
  );

  let dataRows;
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    // Already a native Sheet — open directly, no conversion needed
    const ss = SpreadsheetApp.openById(file.getId());
    dataRows = ss.getSheets()[0].getDataRange().getValues();
  } else {
    // Raw .xlsx — convert via a temporary Google Sheet, then clean up
    const blob = file.getBlob();
    const resource = {
      title: "temp_manifest_import",
      mimeType: MimeType.GOOGLE_SHEETS,
    };
    const tempFile = Drive.Files.insert(resource, blob, { convert: true });
    const tempSs = SpreadsheetApp.openById(tempFile.id);
    dataRows = tempSs.getSheets()[0].getDataRange().getValues();
    DriveApp.getFileById(tempFile.id).setTrashed(true);
  }

  if (dataRows.length < 2) {
    Logger.log("Manifest file was read but contains no data rows.");
    return {};
  }

  const headers = dataRows[0];
  const col = {};
  headers.forEach((h, i) => (col[String(h).trim()] = i));

  const required = ["Registration Number", "Student Name", "Batch Code"];
  const missing = required.filter((h) => !(h in col));
  if (missing.length) {
    Logger.log(
      `Manifest is missing expected column(s): ${missing.join(", ")}. Found: ${headers.join(", ")}`,
    );
  }

  const manifest = {};
  for (let i = 1; i < dataRows.length; i++) {
    const row = dataRows[i];
    const regNo = normalizeRegNo(row[col["Registration Number"]]);
    if (!regNo) continue;
    manifest[regNo] = {
      studentName: row[col["Student Name"]],
      batchCode: row[col["Batch Code"]],
      registrationStatus: row[col["Registration Status"]],
      center: row[col["Center"]],
      cbtCenter: row[col["CBT Center"]],
      cbtCenterAddress: row[col["CBT Center Address"]],
      cbtExamTimings: row[col["CBT Exam Timings"]],
      sourceSheet: row[col["Source Sheet"]],
    };
  }
  Logger.log(`Manifest parsed: ${Object.keys(manifest).length} row(s).`);
  return manifest;
}

/** Normalizes a registration number to a plain digit string so PDF filenames
 *  (always plain text) reliably match manifest values (which may come back
 *  as numbers, e.g. 22290819 vs "22290819.0"). */
function normalizeRegNo(value) {
  if (value === null || value === undefined || value === "") return "";
  return String(value).trim().replace(/\.0$/, "");
}

/** For rows already in the Index sheet with blank Student Name / Center / etc.
 *  (e.g. from a run before the manifest was readable) — re-matches them
 *  against the manifest and fills in the gaps, without touching Registration
 *  Number, Url Link, or Uploaded At, and without adding duplicate rows. */
function backfillMissingDetails() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const manifest = readManifest(folder);
  const ui = SpreadsheetApp.getUi();

  if (Object.keys(manifest).length === 0) {
    ui.alert(
      "Manifest still not found or empty — nothing to backfill with. " +
        'Run "Debug: list folder contents" to check the manifest synced correctly.',
    );
    return;
  }

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET_NAME);
  if (!sheet) {
    ui.alert('No Index sheet found — run "Build/update Index" first.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach((h, i) => (col[h] = i));

  let filled = 0;
  for (let i = 1; i < data.length; i++) {
    const regNo = normalizeRegNo(data[i][col["Registration Number"]]);
    const info = manifest[regNo];
    if (!info) continue;
    if (!data[i][col["Student Name"]]) {
      sheet
        .getRange(i + 1, col["Student Name"] + 1)
        .setValue(info.studentName || "");
      sheet
        .getRange(i + 1, col["Batch Code"] + 1)
        .setValue(info.batchCode || "");
      sheet
        .getRange(i + 1, col["Registration Status"] + 1)
        .setValue(info.registrationStatus || "");
      sheet.getRange(i + 1, col["Center"] + 1).setValue(info.center || "");
      sheet
        .getRange(i + 1, col["CBT Center"] + 1)
        .setValue(info.cbtCenter || "");
      sheet
        .getRange(i + 1, col["CBT Center Address"] + 1)
        .setValue(info.cbtCenterAddress || "");
      sheet
        .getRange(i + 1, col["CBT Exam Timings"] + 1)
        .setValue(info.cbtExamTimings || "");
      sheet
        .getRange(i + 1, col["Source Sheet"] + 1)
        .setValue(info.sourceSheet || "");
      filled++;
    }
  }

  ui.alert(`Backfilled details for ${filled} row(s).`);
}

/** Debug helper: lists every file in the synced folder with its name and
 *  type, so you can confirm the manifest actually made it there and see
 *  exactly what name/type it has. Run from the Apps Script editor (select
 *  this function, then Run) or add it to the menu if you want it handy. */
function debugListFolderContents() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const files = folder.getFiles();
  const lines = [];
  while (files.hasNext()) {
    const f = files.next();
    lines.push(
      `${f.getName()}  |  ${f.getMimeType()}  |  updated ${f.getLastUpdated()}`,
    );
  }
  Logger.log(lines.join("\n"));
  SpreadsheetApp.getUi().alert(
    `${lines.length} file(s) in folder — see Extensions > Apps Script > Executions/Logs for details.`,
  );
}

function exportReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName(INDEX_SHEET_NAME);
  if (!indexSheet) {
    SpreadsheetApp.getUi().alert(
      "No Index sheet found — run 'Build/update Index' first.",
    );
    return;
  }

  const backupsFolder = getOrCreateFolder(BACKUPS_FOLDER_NAME);

  // Copy just the Index sheet into a standalone spreadsheet
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_HHmm",
  );
  const reportName = `Admit_Card_Links_${timestamp}`;
  const tempSs = SpreadsheetApp.create(reportName);
  const importedSheet = indexSheet.copyTo(tempSs);
  tempSs.deleteSheet(tempSs.getSheets()[0]); // remove default blank sheet
  importedSheet.setName("Admit Card Links");

  // Export as .xlsx via the Sheets export endpoint, using the script's own auth
  const ssId = tempSs.getId();
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  const xlsxFile = backupsFolder.createFile(
    response.getBlob().setName(`${reportName}.xlsx`),
  );
  DriveApp.getFileById(ssId).setTrashed(true); // clean up the intermediate Google Sheet

  SpreadsheetApp.getUi().alert(
    `Report exported: ${xlsxFile.getName()}\n\n${xlsxFile.getUrl()}`,
  );
  return xlsxFile.getUrl();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}
