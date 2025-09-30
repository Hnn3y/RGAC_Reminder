import { google } from "googleapis";
import { DateTime } from "luxon";
import nodemailer from "nodemailer";
import sgMail from "@sendgrid/mail";

// ENV VARS: GOOGLE_CLIENT_EMAIL, GOOGLE_PRIVATE_KEY, GOOGLE_PROJECT_ID, EMAIL_PROVIDER, EMAIL_USER, EMAIL_PASS, SENDGRID_API_KEY, GOOGLE_SHEET_ID
const {
  GOOGLE_CLIENT_EMAIL,
  GOOGLE_PRIVATE_KEY,
  GOOGLE_PROJECT_ID,
  EMAIL_PROVIDER,
  EMAIL_USER,
  EMAIL_PASS,
  SENDGRID_API_KEY,
  GOOGLE_SHEET_ID,
} = process.env;

if (!GOOGLE_SHEET_ID) throw new Error("Missing GOOGLE_SHEET_ID env variable");

const SHEET_NAMES = {
  MASTER: "ALL AMC CLIENT",
  REMINDERS: "REMINDER SHEET",
  STATUS_LOG: "Status Log",
};

const REQUIRED_COLUMNS = [
  "Name",
  "Plate Number",
  "Email",
  "Phone",
  "Last Service Date",
  "Next Reminder Date",
  "Manual Contact",
  "Status",
];

export async function mainSync() {
  // 1. Authenticate Google Sheets
  const sheets = await getSheetsClient();

  // 2. Fetch and ensure columns in Master
  let { rows: masterRows, header: masterHeader } = await fetchSheetRows(sheets, SHEET_NAMES.MASTER);
  const { header: ensuredHeader, changed: headerChanged } = ensureColumns(masterHeader, REQUIRED_COLUMNS);

  if (headerChanged) {
    await updateSheetHeader(sheets, SHEET_NAMES.MASTER, ensuredHeader);
    masterHeader = ensuredHeader;
  }

  // 3. Process customers
  const processedCustomers = processCustomers(masterRows, masterHeader);

  // 4. Alphabetically sort
  processedCustomers.sort((a, b) => (a["Name"] || "").localeCompare(b["Name"] || ""));

  // 5. Write processed customers to Reminders and update Master (reminder fields only)
  await writeProcessedData(sheets, processedCustomers, ensuredHeader, SHEET_NAMES.REMINDERS);
  await writeProcessedData(sheets, processedCustomers, ensuredHeader, SHEET_NAMES.MASTER);

  // 6. Send due emails and log results
  const today = DateTime.now().toISODate();
  const dueCustomers = processedCustomers.filter(
    c => c["Next Reminder Date"] && DateTime.fromISO(c["Next Reminder Date"]).toISODate() === today
  );
  const emailResults = await sendReminders(dueCustomers);

  // 7. Log to Status Log
  const logRow = [
    DateTime.now().toISO({ suppressMilliseconds: true }),
    processedCustomers.length,
    emailResults.sent,
    emailResults.failed,
    emailResults.failures.join("; ")
  ];
  await appendSheetRow(sheets, SHEET_NAMES.STATUS_LOG, logRow);

  // 8. Return summary
  return {
    processed: processedCustomers.length,
    remindersSent: emailResults.sent,
    remindersFailed: emailResults.failed,
    failures: emailResults.failures,
  };
}

// ==== GOOGLE SHEETS HELPERS ====
async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: GOOGLE_CLIENT_EMAIL,
      private_key: (GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      project_id: GOOGLE_PROJECT_ID,
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const sheets = google.sheets({ version: "v4", auth });
  return sheets;
}

async function fetchSheetRows(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName,
    majorDimension: "ROWS",
  });
  const values = res.data.values || [];
  if (values.length === 0) return { header: [], rows: [] };
  const [header, ...rows] = values;
  return { header, rows };
}

function ensureColumns(header, required) {
  const newHeader = [...header];
  let changed = false;
  required.forEach(col => {
    if (!newHeader.includes(col)) {
      newHeader.push(col);
      changed = true;
    }
  });
  return { header: newHeader, changed };
}

async function updateSheetHeader(sheets, sheetName, header) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${sheetName}!1:1`,
    valueInputOption: "RAW",
    requestBody: { values: [header] },
  });
}

function processCustomers(rows, header) {
  // Convert to objects
  const idx = Object.fromEntries(header.map((h, i) => [h, i]));
  const customers = rows.map(row => {
    const obj = {};
    header.forEach((col, i) => { obj[col] = (row[i] || "").trim(); });
    return obj;
  });

  // Ensure all required fields & logic
  for (const customer of customers) {
    // Recalculate Next Reminder Date if Last Service Date changed or missing
    let lastService = parseDate(customer["Last Service Date"]);
    if (lastService) {
      const nextReminder = lastService.plus({ months: 3 });
      customer["Next Reminder Date"] = nextReminder.toISODate();
    } else {
      customer["Next Reminder Date"] = "";
    }

    // Manual Contact if missing both email and phone
    const hasEmail = Boolean(customer["Email"]);
    const hasPhone = Boolean(customer["Phone"]);
    customer["Manual Contact"] = (!hasEmail && !hasPhone) ? "MISSING CONTACT" : "";

    // Status could be customized here if needed
  }
  return customers;
}

function parseDate(str) {
  if (!str) return null;
  // Accepts ISO, or dd/MM/yyyy, or MM/dd/yyyy
  let dt = DateTime.fromISO(str);
  if (!dt.isValid) dt = DateTime.fromFormat(str, "dd/MM/yyyy");
  if (!dt.isValid) dt = DateTime.fromFormat(str, "MM/dd/yyyy");
  return dt.isValid ? dt : null;
}

async function writeProcessedData(sheets, customers, header, sheetName) {
  // Overwrite all rows (including header)
  const values = [header].concat(
    customers.map(c => header.map(h => c[h] || ""))
  );
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName,
    valueInputOption: "RAW",
    requestBody: { values }
  });
}

async function updateReminderFieldsInMaster(sheets, customers, header, sheetName) {
  // Only update Next Reminder Date & Manual Contact in Master
  const fieldsToUpdate = ["Next Reminder Date", "Manual Contact"];
  // Fetch current data to get correct row positions
  const { rows, header: masterHeader } = await fetchSheetRows(sheets, sheetName);
  const idx = Object.fromEntries(masterHeader.map((h, i) => [h, i]));

  // Map by Name for update
  const byName = Object.fromEntries(customers.map(c => [c["Name"], c]));
  const updatedRows = rows.map(row => {
    const name = row[idx["Name"]];
    const customer = byName[name];
    if (!customer) return row;
    const newRow = [...row];
    for (const field of fieldsToUpdate) {
      if (idx[field] !== undefined) newRow[idx[field]] = customer[field] || "";
    }
    return newRow;
  });
  // Re-write (excluding header)
  await sheets.spreadsheets.values.update({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: `${sheetName}!2:${updatedRows.length + 1}`,
    valueInputOption: "RAW",
    requestBody: { values: updatedRows },
  });
}

async function appendSheetRow(sheets, sheetName, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: GOOGLE_SHEET_ID,
    range: sheetName,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [row] }
  });
}

// ==== EMAIL NOTIFICATION ====
async function sendReminders(customers) {
  let sent = 0, failed = 0, failures = [];
  for (const customer of customers) {
    if (customer["Manual Contact"] === "MISSING CONTACT") continue;
    const to = customer["Email"];
    if (!to) continue; // skip no-email
    try {
      await sendEmail(to, emailTemplate(customer));
      sent++;
    } catch (e) {
      failed++;
      failures.push(`To:${to} ${e.message}`);
    }
  }
  return { sent, failed, failures };
}

function emailTemplate(customer) {
  // Editable template
  return {
    subject: `Service Reminder for ${customer["Name"]}`,
    text: (
      `Dear ${customer["Name"] || "Customer"},\n\n` +
      `This is a friendly reminder that your vehicle (Plate Number: ${customer["Plate Number"]}) is due for service.\n` +
      `Last serviced: ${customer["Last Service Date"]}\n` +
      `Recommended next service: ${customer["Next Reminder Date"]}\n\n` +
      `Please contact us to schedule your appointment.\n\n` +
      `Best regards,\nService Team`
    )
  };
}

async function sendEmail(to, { subject, text }) {
  if (EMAIL_PROVIDER === "gmail") {
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: EMAIL_USER,
        pass: EMAIL_PASS,
      }
    });
    await transporter.sendMail({
      from: EMAIL_USER,
      to,
      subject,
      text
    });
  } else if (EMAIL_PROVIDER === "sendgrid") {
    sgMail.setApiKey(SENDGRID_API_KEY);
    await sgMail.send({
      to,
      from: EMAIL_USER, // must be a verified sender
      subject,
      text
    });
  } else {
    throw new Error("Unknown EMAIL_PROVIDER: " + EMAIL_PROVIDER);
  }
}