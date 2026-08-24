/**
 * Admit Card Drive Sync + Export Report
 *
 * Run this AFTER generate_admit_cards.py has written PDFs + a manifest_*.xlsx
 * file into a folder that Google Drive for Desktop is syncing to your Drive.
 *
 * Setup (one-time):
 * 1. Install Google Drive for Desktop (free) and let it sync a folder,
 *    e.g. "AdmitCards/output_cards", to a Drive folder.
 * 2. Open the spreadsheet you want this bound to (or any spreadsheet) ->
 *    Extensions -> Apps Script. Paste this file in alongside your existing
 *    portal backend (Code.gs from the earlier phase).
 * 3. Set DRIVE_FOLDER_ID below.
 * 4. Reload the spreadsheet -> you'll see an "Admit Cards" menu.
 *
 * What it does:
 * - buildIndexFromFolder(): shares the synced Drive folder once (so every file
 *   inside inherits "anyone with link can view" — no per-file sharing calls,
 *   which is what makes this fast at scale), scans it for PDFs, matches each
 *   to its batch's manifest_*.xlsx (every manifest file in the folder is read
 *   and merged — each generation run writes a uniquely-named manifest, e.g.
 *   manifest_CBT_20260824_153300.xlsx, so multiple batches synced into the
 *   same folder don't overwrite each other's student details), and writes/
 *   updates the "Index" sheet in one batched write:
 *   Registration Number, Student Name, Batch Code, Registration Status,
 *   Center, CBT Center, CBT Center Address, CBT Exam Timings, Url Link,
 *   Mobile Last4, File Name, Source Sheet, Uploaded At.
 *   This is the same Index sheet the student portal (Code.gs) reads from.
 * - exportReport(): copies the Index sheet into a fresh spreadsheet, exports
 *   it as .xlsx into a "Backups" Drive folder, and returns the file — this is
 *   your permanent, downloadable link report for the batch.
 */

// INDEX_SHEET_NAME is declared in Code.gs (both files share scope in Apps Script)
const DRIVE_FOLDER_ID = "18CJW5pTsq8V2hyH8t5Yja85JaqsI5P2m";
const MANIFEST_NAME_CONTAINS = "manifest"; // matches manifest_CBT_20260824_153300.xlsx etc.
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
      `Couldn't read student details from any manifest file in this folder. ` +
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

/** Finds every file in the folder whose name contains "manifest", reads each
 *  one (raw .xlsx or auto-converted Google Sheet — Drive sometimes converts
 *  uploaded Office files to native Google formats on sync, a common
 *  Workspace default), and merges them into one {regNo: {studentName, ...}}
 *  map. Reading ALL of them (not just the most recent) matters because each
 *  generate_admit_cards.py run writes a uniquely-named manifest — if your
 *  Drive folder holds PDFs from several different batches, this is what
 *  makes every one of them resolve correctly instead of only the latest. */
function readManifest(folder) {
  const search = folder.searchFiles(
    `title contains "${MANIFEST_NAME_CONTAINS}"`,
  );
  const manifestFiles = [];
  while (search.hasNext()) manifestFiles.push(search.next());

  if (manifestFiles.length === 0) {
    Logger.log(
      `No file with "${MANIFEST_NAME_CONTAINS}" in the name found in folder "${folder.getName()}".`,
    );
    return {};
  }
  Logger.log(
    `Found ${manifestFiles.length} manifest file(s): ${manifestFiles.map((f) => f.getName()).join(", ")}`,
  );

  const manifest = {};
  let totalRows = 0;
  for (const file of manifestFiles) {
    const rows = readManifestFile(file);
    totalRows += rows;
  }
  Logger.log(
    `Manifest parsed: ${Object.keys(manifest).length} unique registration number(s) from ${totalRows} total row(s) across ${manifestFiles.length} file(s).`,
  );
  return manifest;

  // Reads one manifest file's rows into the shared `manifest` object above.
  // Defined inline so it can close over `manifest` without a global.
  function readManifestFile(file) {
    let dataRows;
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const ss = SpreadsheetApp.openById(file.getId());
      dataRows = ss.getSheets()[0].getDataRange().getValues();
    } else {
      const blob = file.getBlob();
      const resource = {
        name: "temp_manifest_import",
        mimeType: MimeType.GOOGLE_SHEETS,
      };
      const tempFile = Drive.Files.create(resource, blob, { convert: true });
      const tempSs = SpreadsheetApp.openById(tempFile.id);
      dataRows = tempSs.getSheets()[0].getDataRange().getValues();
      DriveApp.getFileById(tempFile.id).setTrashed(true);
    }

    if (dataRows.length < 2) {
      Logger.log(`${file.getName()}: read but contains no data rows.`);
      return 0;
    }

    const headers = dataRows[0];
    const col = {};
    headers.forEach((h, i) => (col[String(h).trim()] = i));

    const required = ["Registration Number", "Student Name", "Batch Code"];
    const missing = required.filter((h) => !(h in col));
    if (missing.length) {
      Logger.log(
        `${file.getName()}: missing expected column(s): ${missing.join(", ")}. Found: ${headers.join(", ")}`,
      );
      return 0;
    }

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
    return dataRows.length - 1;
  }
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

/** Debug helper: searches specifically for the manifest file (rather than
 *  dumping the whole folder, which can be too large to read if the folder
 *  has thousands of PDFs) and reports exactly what it finds. Safe to run
 *  either from the "Admit Cards" menu or directly in the Apps Script editor
 *  (Run button) — it detects which context it's in. */
function debugListFolderContents() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // Count total files (cheap — just iterating an ID list, no content reads)
  let totalCount = 0;
  const allFiles = folder.getFiles();
  while (allFiles.hasNext()) {
    allFiles.next();
    totalCount++;
  }

  // Search specifically for anything with "manifest" in the name
  const matches = [];
  const search = folder.searchFiles(
    `title contains "${MANIFEST_NAME_CONTAINS}"`,
  );
  while (search.hasNext()) {
    const f = search.next();
    matches.push(
      `${f.getName()}  |  ${f.getMimeType()}  |  updated ${f.getLastUpdated()}  |  id: ${f.getId()}`,
    );
  }

  const summary = matches.length
    ? `Found ${matches.length} file(s) matching "manifest":\n\n${matches.join("\n")}`
    : `No file with "manifest" in the name found in this folder.\n\n` +
      `Folder has ${totalCount} file(s) total. Either the manifest hasn't finished ` +
      `syncing yet, or generate_admit_cards.py wrote it to a different folder than ` +
      `the one DRIVE_FOLDER_ID points to.`;

  Logger.log(summary);
  Logger.log(`Total files in folder: ${totalCount}`);

  // Only show a UI alert if we're actually running inside the spreadsheet
  // (running this directly from the script editor has no UI context)
  try {
    SpreadsheetApp.getUi().alert(summary);
  } catch (e) {
    Logger.log(
      "(No UI context — this is expected when run from the script editor directly. Check the log above instead.)",
    );
  }
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
