// Google Apps Script - paste into script.google.com (Extensions > Apps Script from your Google Sheet)

const CV_FOLDER_NAME = "Student CVs";

const HEADERS = [
  "Timestamp", "Name", "Number", "Date of birth", "Age", "Email ID", "Place",
  "Highest qualification", "10th pass-out year", "10th percentage",
  "12th pass-out year", "12th percentage", "12th stream", "Work experience",
  "Interested country", "Interested course", "Preferred intake", "Funding",
  "Language proficiency", "CV link"
];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    // Save CV into a Drive folder, named with student name + number
    const folders = DriveApp.getFoldersByName(CV_FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CV_FOLDER_NAME);
    const safe = s => String(s || "").replace(/[^\w\-]+/g, "_");
    const blob = Utilities.newBlob(Utilities.base64Decode(d.cvBase64), d.cvType,
      safe(d["Name"]) + "_" + safe(d["Number"]) + "_" + d.cvName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const row = HEADERS.map(h => {
      if (h === "Timestamp") return new Date();
      if (h === "CV link") return file.getUrl();
      return d[h] || "";
    });
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
