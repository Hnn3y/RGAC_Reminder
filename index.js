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
 "Veh. Reg. No.",
 "Email Add.", 
 "Phone Number", 
 "Last Visit",
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
  const processedCustomers = processCustomers(masterRows, ensuredHeader);

  // 4. Alphabetically sort for reminders sheet
  const sortedCustomers = [...processedCustomers].sort((a, b) => 
    (a["Name"] || "").localeCompare(b["Name"] || "")
  );

  // 5. Write sorted data to Reminders sheet
  await writeProcessedData(sheets, sortedCustomers, ensuredHeader, SHEET_NAMES.REMINDERS);

  // 6. Update only reminder fields in Master (preserve original order)
  await updateReminderFieldsInMaster(sheets, processedCustomers, ensuredHeader, SHEET_NAMES.MASTER);

  // 7. Send due/overdue emails and log results
  const emailResults = await sendReminders(processedCustomers);

  // 8. Log to Status Log
  const logRow = [
    DateTime.now().toISO({ suppressMilliseconds: true }),
    processedCustomers.length,
    emailResults.sent,
    emailResults.failed,
    emailResults.failures.join("; ")
  ];
  await appendSheetRow(sheets, SHEET_NAMES.STATUS_LOG, logRow);

  // 9. Return summary
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
  const customers = rows.map(row => {
    const obj = {};
    header.forEach((col, i) => { obj[col] = (row[i] || "").trim(); });

    // Ensure all required fields exist
    REQUIRED_COLUMNS.forEach(col => {
      if (!(col in obj)) obj[col] = "";
    });

    // Calculate Next Reminder Date (3 months after Last Visit)
    let lastVisit = parseDate(obj["Last Visit"]);
    let nextReminder = "";
    if (lastVisit) {
      nextReminder = lastVisit.plus({ months: 3 }).toISODate();
      obj["Next Reminder Date"] = nextReminder;
      console.log(`Customer: ${obj["Name"]}, Last Visit: ${obj["Last Visit"]}, Next Reminder: ${nextReminder}`);
    } else {
      obj["Next Reminder Date"] = "";
      console.log(`Customer: ${obj["Name"]}, No valid Last Visit date found: "${obj["Last Visit"]}"`);
    }

    // Manual Contact if no email/phone
    const hasEmail = Boolean(obj["Email Add."]);
    const hasPhone = Boolean(obj["Phone Number"]);
    obj["Manual Contact"] = (!hasEmail && !hasPhone) ? "MISSING CONTACT" : "";

    // Status can be updated elsewhere if needed
    return obj;
  });
  return customers;
}

function parseDate(str) {
  if (!str) return null;
  
  // Try multiple date formats
  const formats = [
    "yyyy-MM-dd",      // ISO: 2025-12-09
    "dd/MM/yyyy",      // 09/12/2025
    "MM/dd/yyyy",      // 12/09/2025
    "dd-MM-yyyy",      // 09-12-2025 (YOUR FORMAT!)
    "MM-dd-yyyy",      // 12-09-2025
    "d/M/yyyy",        // 9/12/2025 (single digit)
    "d-M-yyyy",        // 9-12-2025 (single digit)
  ];
  
  // First try ISO parsing
  let dt = DateTime.fromISO(str);
  if (dt.isValid) return dt;
  
  // Try each format
  for (const format of formats) {
    dt = DateTime.fromFormat(str, format);
    if (dt.isValid) return dt;
  }
  
  return null;
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

  // Map by Name for update (use combination of Name + Plate for uniqueness)
  const byKey = Object.fromEntries(
    customers.map(c => [`${c["Name"]}|${c["Veh. Reg. No."]}`, c])
  );
  
  const updatedRows = rows.map(row => {
    const name = (row[idx["Name"]] || "").trim();
    const plate = (row[idx["Veh. Reg. No."]] || "").trim();
    const key = `${name}|${plate}`;
    const customer = byKey[key];
    
    if (!customer) return row;
    
    // Ensure row has enough cells
    const newRow = [...row];
    while (newRow.length < masterHeader.length) {
      newRow.push("");
    }
    
    for (const field of fieldsToUpdate) {
      if (idx[field] !== undefined) {
        newRow[idx[field]] = customer[field] || "";
      }
    }
    return newRow;
  });
  
  // Re-write data rows (excluding header)
  if (updatedRows.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: `${sheetName}!A2`,
      valueInputOption: "RAW",
      requestBody: { values: updatedRows },
    });
  }
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
  const today = DateTime.now().toISODate();
  
  console.log(`\n=== EMAIL SENDING PROCESS ===`);
  console.log(`Today's date: ${today}`);
  console.log(`Total customers: ${customers.length}`);
  
  for (const customer of customers) {
    const name = customer["Name"] || "Unknown";
    
    // Skip customers with missing contact info
    if (customer["Manual Contact"] === "MISSING CONTACT") {
      console.log(`⏭️  SKIP: ${name} - Missing contact info`);
      continue;
    }
    
    const to = customer["Email Add."];
    if (!to) {
      console.log(`⏭️  SKIP: ${name} - No email address`);
      continue;
    }

    let nextReminder = customer["Next Reminder Date"];
    let status = (customer["Status"] || "").toUpperCase();
    
    console.log(`\n📋 Checking: ${name}`);
    console.log(`   Email: ${to}`);
    console.log(`   Status: ${status}`);
    console.log(`   Next Reminder: ${nextReminder}`);
    console.log(`   Today: ${today}`);
    
    let overdue = false;

    // Check if overdue
    if (nextReminder) {
      const next = DateTime.fromISO(nextReminder);
      if (next.isValid && status === "NOT SERVICED" && next < DateTime.now().startOf("day")) {
        overdue = true;
        console.log(`  OVERDUE! (${nextReminder} < ${today})`);
      }
    }

    // Determine if we should send email
    let template;
    if (overdue) {
      template = overdueEmailTemplate(customer);
      console.log(`   📧 Sending OVERDUE email...`);
    } else if (nextReminder === today && status === "NOT SERVICED") {
      template = regularEmailTemplate(customer);
      console.log(`   📧 Sending DUE TODAY email...`);
    } else {
      console.log(`   ⏭️  Not due yet or already serviced`);
      continue; // Skip this customer
    }

    // Send email
    try {
      await sendEmail(to, template);
      sent++;
      console.log(`   ✅ Email sent successfully!`);
    } catch (e) {
      failed++;
      const errorMsg = `To:${to} ${e.message}`;
      failures.push(errorMsg);
      console.error(`   ❌ Failed to send: ${e.message}`);
    }
  }
  
  console.log(`\n=== EMAIL SUMMARY ===`);
  console.log(`✅ Sent: ${sent}`);
  console.log(`❌ Failed: ${failed}`);
  if (failures.length > 0) {
    console.log(`Failures: ${failures.join("; ")}`);
  }
  
  return { sent, failed, failures };
}

function regularEmailTemplate(customer) {
  return {
    subject: `Service Reminder for ${customer["Name"]}`,
    text: (
      `Dear ${customer["Name"] || "Customer"},\n\n` +
      `This is a friendly reminder that your vehicle (Plate Number: ${customer["Veh. Reg. No."]}) is due for service at Royal Gem AutoCare.\n` +
      `Your Last serviced date was on: ${customer["Last Visit"]}\n` +
      `Recommended next service date: ${customer["Next Reminder Date"]}\n\n` +
      `Please contact us to schedule your appointment.\n\n` +
      `Best regards,\nRoyal Gem Auto Care Service Team`
    )
  };
}

function overdueEmailTemplate(customer) {
  return {
    subject: `Overdue Service Reminder for ${customer["Name"]}`,
    text: (
      `Dear ${customer["Name"] || "Customer"},\n\n` +
      `Our records show that your vehicle (Plate Number: ${customer["Veh. Reg No."]}) has missed its scheduled service.\n` +
      `Last serviced: ${customer["Last Visit"]}\n` +
      `Recommended service was due: ${customer["Next Reminder Date"]}\n\n` +
      `Please contact us as soon as possible to schedule your overdue service and ensure your vehicle remains in top condition.\n\n` +
      `Best regards,\nService Team`
    )
  };
}

async function sendEmail(to, { subject, text }) {
  if (EMAIL_PROVIDER === "smtp") {
    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,          
      port: Number(process.env.SMTP_PORT),  // 465 or 587
      secure: process.env.SMTP_SECURE === "true", 
      auth: {
        user: EMAIL_USER,  
        pass: EMAIL_PASS,  
      },
    });

    await transporter.sendMail({
      from: `"Royal Gem AutoCare Nigeria Limited" <${EMAIL_USER}>`, 
      to,
      subject,
      text,
    });

  } else if (EMAIL_PROVIDER === "sendgrid") {
    sgMail.setApiKey(SENDGRID_API_KEY);
    await sgMail.send({
      to,
      from: EMAIL_USER, // must be a verified sender in SendGrid
      subject,
      text,
    });

  } else {
    throw new Error("Unknown EMAIL_PROVIDER: " + EMAIL_PROVIDER);
  }
}
