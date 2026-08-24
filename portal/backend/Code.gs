/**
 * Admit Card Lookup — Apps Script Web App backend.
 *
 * Deploy this as a Web App (Deploy > New deployment > Web app,
 * execute as "Me", access "Anyone"). The deployed URL goes into
 * BACKEND_URL in portal/index.html.
 *
 * Expects an "Index" sheet in this spreadsheet with columns:
 *   Registration Number | Student Name | Batch Code | Registration Status |
 *   Center | CBT Center | CBT Center Address | CBT Exam Timings | Url Link |
 *   Mobile Last4 | File Name | Source Sheet | Uploaded At
 *
 * This is the same index sheet DriveSync.gs writes to after each batch
 * (registration number -> Drive file link).
 */

const INDEX_SHEET_NAME = "Index";

function doGet(e) {
  const reg = (e.parameter.reg || "").trim();
  const mobile = (e.parameter.mobile || "").trim();

  if (!reg || !mobile) {
    return jsonOut({
      found: false,
      message: "Missing registration number or mobile digits.",
    });
  }

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = {};
  headers.forEach((h, i) => (col[h] = i));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowReg = String(row[col["Registration Number"]]).replace(/\.0$/, "");
    const rowMobile = String(row[col["Mobile Last4"]]);

    if (rowReg === reg && rowMobile === mobile) {
      return jsonOut({
        found: true,
        name: row[col["Student Name"]],
        link: row[col["Url Link"]],
      });
    }
  }

  return jsonOut({
    found: false,
    message: "No admit card found for those details.",
  });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
