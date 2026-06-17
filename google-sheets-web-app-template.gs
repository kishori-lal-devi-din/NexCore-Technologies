/*
  Google Sheets receiver for the NexCore consultation form.

  Setup:
  1. Create or open the Google Sheet that should store submissions.
  2. Go to Extensions > Apps Script.
  3. Paste this file into the Apps Script editor.
  4. Deploy > New deployment > Web app.
  5. Set "Execute as" to "Me" and "Who has access" to "Anyone".
  6. Copy the Web App URL into index.html as the contact form action.

  WhatsApp notification setup:
  1. Create an approved WhatsApp template named "new_consultation_alert".
  2. Add these Apps Script Project Settings > Script Properties:
     WA_ACCESS_TOKEN, WA_PHONE_NUMBER_ID, WA_RECIPIENT_PHONE.
     WA_RECIPIENT_PHONE must include the country code, for example 919821912471.
  3. Optional Script Properties:
     WA_API_VERSION, WA_TEMPLATE_NAME, WA_TEMPLATE_LANGUAGE.

  If this script is not bound to the target Sheet, paste the spreadsheet ID below.
*/

const SPREADSHEET_ID = "";
const SHEET_NAME = "Consultation Requests";
const DEFAULT_WHATSAPP_API_VERSION = "v25.0";
const DEFAULT_WHATSAPP_TEMPLATE_NAME = "new_consultation_alert";
const DEFAULT_WHATSAPP_TEMPLATE_LANGUAGE = "en_US";
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

    try {
      sendWhatsAppNotification_(data);
    } catch (notificationError) {
      console.error("WhatsApp notification failed: " + notificationError.message);
    }

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

function sendWhatsAppNotification_(data) {
  const properties = PropertiesService.getScriptProperties();
  const accessToken = getRequiredProperty_(properties, "WA_ACCESS_TOKEN");
  const phoneNumberId = getRequiredProperty_(properties, "WA_PHONE_NUMBER_ID");
  const recipientPhone = getRequiredProperty_(properties, "WA_RECIPIENT_PHONE");
  const apiVersion = properties.getProperty("WA_API_VERSION") || DEFAULT_WHATSAPP_API_VERSION;
  const templateName = properties.getProperty("WA_TEMPLATE_NAME") || DEFAULT_WHATSAPP_TEMPLATE_NAME;
  const templateLanguage =
    properties.getProperty("WA_TEMPLATE_LANGUAGE") || DEFAULT_WHATSAPP_TEMPLATE_LANGUAGE;

  const url =
    "https://graph.facebook.com/" +
    apiVersion +
    "/" +
    phoneNumberId +
    "/messages";

  const payload = {
    messaging_product: "whatsapp",
    recipient_type: "individual",
    to: normalizePhoneNumber_(recipientPhone),
    type: "template",
    template: {
      name: templateName,
      language: {
        code: templateLanguage
      },
      components: [
        {
          type: "body",
          parameters: [
            templateText_(data["Full Name"]),
            templateText_(data["Company Name"]),
            templateText_(data["Phone Number"]),
            templateText_(data["Business Email"]),
            templateText_(data["Services Interested In"]),
            templateText_(data["Industry"]),
            templateText_(data["Company Size"]),
            templateText_(data["Project Budget Range"]),
            templateText_(data["Current Challenge"], 600),
            templateText_(data["Source Page"], 500)
          ]
        }
      ]
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + accessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("WhatsApp API returned " + statusCode + ": " + responseBody);
  }

  console.log("WhatsApp notification sent: " + responseBody);
}

function testWhatsAppNotification() {
  sendWhatsAppNotification_({
    "Full Name": "Test Lead",
    "Company Name": "NexCore Test Company",
    "Phone Number": "+91 98765 43210",
    "Business Email": "test@example.com",
    "Services Interested In": "Hybrid Automation",
    "Industry": "Manufacturing",
    "Company Size": "26 - 100 employees",
    "Project Budget Range": "$5,000 - $15,000",
    "Current Challenge": "Testing WhatsApp notification from Apps Script.",
    "Source Page": "https://nexcoretechnologies.com/"
  });
}

function templateText_(value, maximumLength) {
  const normalized = String(value || "Not provided")
    .replace(/\s+/g, " ")
    .trim();

  return {
    type: "text",
    text: normalized.substring(0, maximumLength || 250)
  };
}

function normalizePhoneNumber_(phoneNumber) {
  const normalized = String(phoneNumber || "").replace(/\D/g, "");

  if (!normalized) {
    throw new Error("WA_RECIPIENT_PHONE must include a country code and phone number.");
  }

  return normalized;
}

function getRequiredProperty_(properties, name) {
  const value = properties.getProperty(name);

  if (!value) {
    throw new Error("Missing Script Property: " + name);
  }

  return value;
}
