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
 "Last Email Sent",        // Track when we last sent email
 "Email Type",             // Track what type of email was sent
 "Subscription",           // NEW: Track if customer is subscribed to reminders
];

export async function mainSync() {

  console.log('\n🎯 mainSync() STARTED');
  console.log('Current Time:', new Date().toISOString());

  // 1. Authenticate Google Sheets
  console.log('\n📝 Step 1: Authenticating...');
  const sheets = await getSheetsClient();
  console.log('✅ Authenticated');

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

  // 6. Send reminders and get updated customers with email tracking
  const emailResults = await sendReminders(processedCustomers);

  // 7. Update Master with reminder fields AND email tracking
  await updateReminderFieldsInMaster(sheets, emailResults.updatedCustomers, ensuredHeader, SHEET_NAMES.MASTER);

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
      // Format as dd-MM-yyyy
      nextReminder = lastVisit.plus({ months: 3 }).toFormat("dd-MM-yyyy");
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
    "dd-MM-yyyy",      // 09-12-2025
    "MM-dd-yyyy",      // 12-09-2025
    "d/M/yyyy",        // 9/12/2025
    "d-M-yyyy",        // 9-12-2025
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
  // Update reminder fields AND email tracking (including Subscription)
  const fieldsToUpdate = ["Next Reminder Date", "Manual Contact", "Last Email Sent", "Email Type", "Subscription"];
  
  const { rows, header: masterHeader } = await fetchSheetRows(sheets, sheetName);
  const idx = Object.fromEntries(masterHeader.map((h, i) => [h, i]));

  const byKey = Object.fromEntries(
    customers.map(c => [`${c["Name"]}|${c["Veh. Reg. No."]}`, c])
  );
  
  const updatedRows = rows.map(row => {
    const name = (row[idx["Name"]] || "").trim();
    const plate = (row[idx["Veh. Reg. No."]] || "").trim();
    const key = `${name}|${plate}`;
    const customer = byKey[key];
    
    if (!customer) return row;
    
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

// ==== EMAIL NOTIFICATION (IMPROVED VERSION) ====
async function sendReminders(customers) {
  let sent = 0, failed = 0, failures = [];
  const today = DateTime.now().startOf("day");
  const todayStr = today.toISODate();
  
  console.log(`\n=== EMAIL SENDING PROCESS ===`);
  console.log(`Today's date: ${todayStr}`);
  console.log(`Total customers: ${customers.length}`);
  
  for (const customer of customers) {
    const name = customer["Name"] || "Unknown";
    
    // Check subscription status FIRST
    const subscription = (customer["Subscription"] || "").trim().toUpperCase();
    
    console.log(`\n📋 Checking: ${name}`);
    console.log(`   Subscription Status: ${subscription}`);
    
    // Skip if not subscribed
    if (subscription === "NOT SUBSCRIBED" || subscription === "UNSUBSCRIBED") {
      console.log(`   🚫 NOT SUBSCRIBED - Skipping email`);
      continue;
    }
    
    // If subscription field is empty or anything other than "NOT SUBSCRIBED", treat as subscribed
    if (subscription !== "SUBSCRIBED" && subscription !== "") {
      console.log(`   ⚠️  Unknown subscription status: "${subscription}" - Treating as SUBSCRIBED`);
    }
    
    // Skip customers with missing contact info
    if (customer["Manual Contact"] === "MISSING CONTACT") {
      console.log(`   ⏭️  SKIP: Missing contact info`);
      continue;
    }
    
    const to = customer["Email Add."];
    if (!to) {
      console.log(`   ⏭️  SKIP: No email address`);
      continue;
    }

    const nextReminderStr = customer["Next Reminder Date"];
    const lastEmailSent = customer["Last Email Sent"] || "";
    const lastEmailType = customer["Email Type"] || "";
    
    console.log(`   Email: ${to}`);
    console.log(`   Next Reminder: ${nextReminderStr}`);
    console.log(`   Last Email: ${lastEmailSent} (${lastEmailType})`);
    console.log(`   Today: ${todayStr}`);
    
    // Parse reminder date
    if (!nextReminderStr) {
      console.log(`   ⏭️  No reminder date set - skipping`);
      continue;
    }

    // Parse the reminder date (now in dd-MM-yyyy format)
    const nextReminder = parseDate(nextReminderStr);
    if (!nextReminder || !nextReminder.isValid) {
      console.log(`   ⚠️  Invalid reminder date: ${nextReminderStr}`);
      continue;
    }

    const reminderDate = nextReminder.startOf("day");
    const daysUntilDue = reminderDate.diff(today, "days").days;
    
    console.log(`   Days until due: ${Math.round(daysUntilDue)}`);
    
    // Determine email type based on timing
    let emailType = null;
    let template = null;
    
    if (daysUntilDue < 0) {
      // OVERDUE
      emailType = "OVERDUE";
      template = overdueEmailTemplate(customer);
      console.log(`   🔴 OVERDUE by ${Math.abs(Math.round(daysUntilDue))} days`);
    } else if (daysUntilDue === 0) {
      // DUE TODAY
      emailType = "DUE_TODAY";
      template = dueTodayEmailTemplate(customer);
      console.log(`   🟡 DUE TODAY`);
    } else if (daysUntilDue <= 7 && daysUntilDue > 0) {
      // 7-DAY ADVANCE WARNING
      emailType = "ADVANCE_7DAY";
      template = advanceReminderEmailTemplate(customer, Math.round(daysUntilDue));
      console.log(`   🟢 Due in ${Math.round(daysUntilDue)} days - sending advance reminder`);
    } else {
      console.log(`   ⏭️  Too early to send reminder (due in ${Math.round(daysUntilDue)} days)`);
      continue;
    }
    
    // Check if we already sent this type of email today
    if (lastEmailSent === todayStr && lastEmailType === emailType) {
      console.log(`   ⏭️  Already sent ${emailType} email today - skipping`);
      continue;
    }
    
    // Send email
    console.log(`   📧 Sending ${emailType} email...`);
    try {
      await sendEmail(to, template);
      sent++;
      
      // Update tracking fields
      customer["Last Email Sent"] = todayStr;
      customer["Email Type"] = emailType;
      
      console.log(`   ✅ Email sent successfully!`);
    } catch (e) {
      failed++;
      const errorMsg = `${name} (${to}): ${e.message}`;
      failures.push(errorMsg);
      console.error(`   ❌ Failed to send: ${e.message}`);
    }
  }
  
  console.log(`\n=== EMAIL SUMMARY ===`);
  console.log(`✅ Sent: ${sent}`);
  console.log(`❌ Failed: ${failed}`);
  if (failures.length > 0) {
    console.log(`Failures:\n  - ${failures.join("\n  - ")}`);
  }
  
  return { sent, failed, failures, updatedCustomers: customers };
}

function advanceReminderEmailTemplate(customer, daysUntil) {
  return {
    subject: `Upcoming Service Reminder - ${customer["Name"]}`,
    text: (
      `Dear ${customer["Name"] || "Customer"},\n\n` +
      `This is a friendly advance reminder that your vehicle (${customer["Veh. Reg. No."]}) is due for service in ${daysUntil} day(s).\n\n` +
      `Service Details:\n` +
      `- Last Service: ${customer["Last Visit"]}\n` +
      `- Next Service Due: ${customer["Next Reminder Date"]}\n\n` +
      `We recommend booking your appointment early to ensure availability.\n\n` +
      `Contact us at Royal Gem AutoCare to schedule your service.\n\n` +
      `Best regards,\n` +
      `Royal Gem Auto Care Service Team`
    )
  };
}

function dueTodayEmailTemplate(customer) {
  return {
    subject: `Service Due Today - ${customer["Name"]}`,
    text: (
      `Dear ${customer["Name"] || "Customer"},\n\n` +
      `Your vehicle (${customer["Veh. Reg. No."]}) is due for service TODAY.\n\n` +
      `Service Details:\n` +
      `- Last Service: ${customer["Last Visit"]}\n` +
      `- Service Due: ${customer["Next Reminder Date"]}\n\n` +
      `Please contact us to schedule your appointment as soon as possible.\n\n` +
      `Best regards,\n` +
      `Royal Gem Auto Care Service Team`
    )
  };
}

function overdueEmailTemplate(customer) {
  return {
    subject: `⚠️ Overdue Service Notice - ${customer["Name"]}`,
    text: (
      `Dear ${customer["Name"] || "Customer"},\n\n` +
      `URGENT: Our records show your vehicle (${customer["Veh. Reg. No."]}) has missed its scheduled service.\n\n` +
      `Service Details:\n` +
      `- Last Service: ${customer["Last Visit"]}\n` +
      `- Service Was Due: ${customer["Next Reminder Date"]}\n\n` +
      `Regular maintenance is essential for your vehicle's safety and performance. ` +
      `Please contact us IMMEDIATELY to schedule your overdue service.\n\n` +
      `Don't risk your vehicle's condition - book your appointment today!\n\n` +
      `Best regards,\n` +
      `Royal Gem Auto Care Service Team`
    )
  };
}

async function sendEmail(to, { subject, text }) {
  if (EMAIL_PROVIDER === "smtp") {
    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,          
      port: Number(process.env.SMTP_PORT),
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
      from: EMAIL_USER,
      subject,
      text,
    });

  } else {
    throw new Error("Unknown EMAIL_PROVIDER: " + EMAIL_PROVIDER);
  }
}