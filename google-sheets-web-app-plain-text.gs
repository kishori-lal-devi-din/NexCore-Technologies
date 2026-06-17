/*
  NexCore Technologies — Google Sheets + WhatsApp Notification
  Plain Text version (no template approval needed)

  Setup:
  1. Paste this into Extensions > Apps Script
  2. Go to Project Settings > Script Properties and add:
     WA_ACCESS_TOKEN     → Your Meta access token
     WA_PHONE_NUMBER_ID  → Your Phone Number ID (1209960635525372)
     WA_RECIPIENT_PHONE  → Your WhatsApp number with country code (e.g. 919XXXXXXXXX)
  3. Deploy > New deployment > Web app
     Execute as: Me | Who has access: Anyone
  4. Copy the Web App URL into your contact form action
*/

const SPREADSHEET_ID = "";
const SHEET_NAME = "Consultation Requests";
const DEFAULT_WHATSAPP_API_VERSION = "v25.0";

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

// ============================================================
// WEB APP ENTRY POINT — Receives form submissions
// ============================================================
function doPost(e) {
  try {
    const sheet = getSheet_();
    const data  = parsePayload_(e);

    // Build row matching HEADERS order
    const row = HEADERS.map(function(header) {
      if (header === "Submitted At") return new Date();
      return data[header] || "";
    });

    sheet.appendRow(row);

    // Send WhatsApp — non-blocking (won't fail the form submission)
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

// ============================================================
// SHEET HELPER — Gets or creates the sheet with headers
// ============================================================
function getSheet_() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("No spreadsheet found. Bind the script or set SPREADSHEET_ID.");
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  // Auto-create headers if missing
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

// ============================================================
// PAYLOAD PARSER — Handles both JSON and form-encoded data
// ============================================================
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

// ============================================================
// WHATSAPP SENDER — Plain text, no template needed
// ============================================================
function sendWhatsAppNotification_(data) {
  const properties    = PropertiesService.getScriptProperties();
  const accessToken   = getRequiredProperty_(properties, "WA_ACCESS_TOKEN");
  const phoneNumberId = getRequiredProperty_(properties, "WA_PHONE_NUMBER_ID");
  const recipientPhone = getRequiredProperty_(properties, "WA_RECIPIENT_PHONE");
  const apiVersion    = properties.getProperty("WA_API_VERSION") || DEFAULT_WHATSAPP_API_VERSION;

  const url = "https://graph.facebook.com/" + apiVersion + "/" + phoneNumberId + "/messages";

  // Format submitted date
  const submittedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd MMM yyyy, hh:mm a"
  );

  // Truncate long fields to avoid WhatsApp limits
  const challenge = truncate_(data["Current Challenge"], 500);
  const services  = truncate_(data["Services Interested In"], 300);

  // Rich card-style plain text message
  const message =
    "🔔 *NEW CONSULTATION REQUEST*\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
    "👤 *Full Name:*        " + (data["Full Name"]       || "N/A") + "\n" +
    "🏢 *Company:*          " + (data["Company Name"]    || "N/A") + "\n" +
    "📧 *Email:*            " + (data["Business Email"]  || "N/A") + "\n" +
    "📱 *Phone:*            " + (data["Phone Number"]    || "N/A") + "\n" +
    "🏭 *Industry:*         " + (data["Industry"]        || "N/A") + "\n" +
    "👥 *Company Size:*     " + (data["Company Size"]    || "N/A") + "\n\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "🛠️ *Services Required:*\n" + services + "\n\n" +
    "💡 *Current Challenge:*\n" + challenge + "\n\n" +
    "💰 *Budget Range:*     " + (data["Project Budget Range"] || "N/A") + "\n" +
    "🌐 *Source Page:*      " + (data["Source Page"]     || "N/A") + "\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "🕐 *Submitted At:*     " + submittedAt + "\n\n" +
    "_Automated notification — NexCore Technologies Consultation Form_";

  const payload = {
    messaging_product: "whatsapp",
    recipient_type:    "individual",
    to:                normalizePhoneNumber_(recipientPhone),
    type:              "text",
    text: {
      preview_url: false,
      body:        message
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method:          "post",
    contentType:     "application/json",
    headers: {
      Authorization: "Bearer " + accessToken
    },
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode   = response.getResponseCode();
  const responseBody = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("WhatsApp API returned " + statusCode + ": " + responseBody);
  }

  console.log("✅ WhatsApp notification sent: " + responseBody);
}

// ============================================================
// TEST FUNCTION — Run manually to verify WhatsApp works
// ============================================================
function testWhatsAppNotification() {
  sendWhatsAppNotification_({
    "Full Name":              "Rahul Sharma",
    "Company Name":           "ABC Technologies Pvt Ltd",
    "Business Email":         "rahul@abctech.com",
    "Phone Number":           "+91 98765 43210",
    "Industry":               "Banking and Finance",
    "Company Size":           "50-200 Employees",
    "Services Interested In": "RPA Automation, AI Chatbot Development",
    "Current Challenge":      "Manual data entry consuming too much time and resources.",
    "Project Budget Range":   "$5,000 - $10,000",
    "Source Page":            "https://nexcoretechnologies.com/consultation"
  });
}

// ============================================================
// HELPERS
// ============================================================
function truncate_(value, maxLength) {
  const str = String(value || "N/A").replace(/\s+/g, " ").trim();
  return str.length > maxLength ? str.substring(0, maxLength) + "..." : str;
}

function normalizePhoneNumber_(phoneNumber) {
  const normalized = String(phoneNumber || "").replace(/\D/g, "");
  if (!normalized) {
    throw new Error("WA_RECIPIENT_PHONE must include country code and number.");
  }
  return normalized;
}

function getRequiredProperty_(properties, name) {
  const value = properties.getProperty(name);
  if (!value) {
    throw new Error("Missing Script Property: " + name + ". Add it in Project Settings.");
  }
  return value;
}