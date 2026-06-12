/*
  Google Sheets receiver for the NexCore consultation form.

  Setup:
  1. Create or open the Google Sheet that should store submissions.
  2. Go to Extensions > Apps Script.
  3. Paste this file into the Apps Script editor.
  4. Deploy > New deployment > Web app.
  5. Set "Execute as" to "Me" and "Who has access" to "Anyone".
  6. Copy the Web App URL into index.html as the contact form action.

  If this script is not bound to the target Sheet, paste the spreadsheet ID below.
*/

const SPREADSHEET_ID = "";
const SHEET_NAME = "Consultation Requests";
const HEADERS = [
  "Submitted At",
  "Full Name",
  "Company Name",
  "Business Email",
  "Phone Number",
  "Industry",
  "Company Size",
  "Services Interested In",
  "Current Challenge",
  "Project Budget Range",
  "Source Page"
];

function doPost(e) {
  try {
    const sheet = getSheet_();
    const data = parsePayload_(e);
    const row = HEADERS.map(function(header) {
      if (header === "Submitted At") return new Date();
      return data[header] || "";
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet_() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("No spreadsheet found. Bind the script to a Google Sheet or set SPREADSHEET_ID.");
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  const existingHeaders = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getValues()[0];

  const headersMissing = HEADERS.some(function(header, index) {
    return existingHeaders[index] !== header;
  });

  if (headersMissing) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function parsePayload_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (error) {
      return e.parameter || {};
    }
  }

  return (e && e.parameter) || {};
}
