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
 "Last Email Sent",
 "Email Type",
 "Subscription",
];

export async function mainSync() {
  console.log('\n' + '='.repeat(80));
  console.log('🎯 mainSync() STARTED');
  console.log('Current Time:', new Date().toISOString());
  console.log('Current Date (ISO):', DateTime.now().toISODate());
  console.log('Current Date (dd-MM-yyyy):', DateTime.now().toFormat('dd-MM-yyyy'));
  console.log('Timezone:', Intl.DateTimeFormat().resolvedOptions().timeZone);
  console.log('='.repeat(80) + '\n');
  
  try {
    // 1. Authenticate Google Sheets
    console.log('📝 STEP 1: Authenticating Google Sheets...');
    const sheets = await getSheetsClient();
    console.log('✅ Authentication successful\n');

    // 2. Fetch and ensure columns in Master
    console.log('📝 STEP 2: Fetching Master Sheet Data...');
    let { rows: masterRows, header: masterHeader } = await fetchSheetRows(sheets, SHEET_NAMES.MASTER);
    console.log(`✅ Fetched ${masterRows.length} rows from Master Sheet`);
    console.log('Current headers:', masterHeader.join(', '));
    
    const { header: ensuredHeader, changed: headerChanged } = ensureColumns(masterHeader, REQUIRED_COLUMNS);

    if (headerChanged) {
      console.log('⚠️  Adding missing columns to Master Sheet...');
      console.log('Missing columns:', REQUIRED_COLUMNS.filter(col => !masterHeader.includes(col)).join(', '));
      await updateSheetHeader(sheets, SHEET_NAMES.MASTER, ensuredHeader);
      masterHeader = ensuredHeader;
      console.log('✅ Headers updated\n');
    } else {
      console.log('✅ All required columns present\n');
    }

    // 3. Process customers
    console.log('📝 STEP 3: Processing customers...');
    const processedCustomers = processCustomers(masterRows, ensuredHeader);
    console.log(`✅ Processed ${processedCustomers.length} customers\n`);

    // 4. Alphabetically sort for reminders sheet
    console.log('📝 STEP 4: Sorting customers alphabetically...');
    const sortedCustomers = [...processedCustomers].sort((a, b) => 
      (a["Name"] || "").localeCompare(b["Name"] || "")
    );
    console.log('✅ Customers sorted\n');

    // 5. Write sorted data to Reminders sheet
    console.log('📝 STEP 5: Writing to Reminder Sheet...');
    await writeProcessedData(sheets, sortedCustomers, ensuredHeader, SHEET_NAMES.REMINDERS);
    console.log('✅ Reminder Sheet updated\n');

    // 6. Send reminders and get updated customers with email tracking
    console.log('📝 STEP 6: Sending email reminders...');
    const emailResults = await sendReminders(processedCustomers);
    console.log(`✅ Email process complete: ${emailResults.sent} sent, ${emailResults.failed} failed\n`);

    // 7. Update Master with reminder fields AND email tracking
    console.log('📝 STEP 7: Updating Master Sheet with email tracking...');
    await updateReminderFieldsInMaster(sheets, emailResults.updatedCustomers, ensuredHeader, SHEET_NAMES.MASTER);
    console.log('✅ Master Sheet updated\n');

    // 8. Log to Status Log
    console.log('📝 STEP 8: Writing to Status Log...');
    const logRow = [
      DateTime.now().toISO({ suppressMilliseconds: true }),
      processedCustomers.length,
      emailResults.sent,
      emailResults.failed,
      emailResults.failures.join("; ")
    ];
    await appendSheetRow(sheets, SHEET_NAMES.STATUS_LOG, logRow);
    console.log('✅ Status Log updated\n');

    // 9. Return summary
    console.log('='.repeat(80));
    console.log('✅ mainSync() COMPLETED SUCCESSFULLY');
    const summary = {
      processed: processedCustomers.length,
      remindersSent: emailResults.sent,
      remindersFailed: emailResults.failed,
      failures: emailResults.failures,
    };
    console.log('FINAL SUMMARY:', JSON.stringify(summary, null, 2));
    console.log('='.repeat(80) + '\n');
    
    return summary;
    
  } catch (error) {
    console.error('\n' + '='.repeat(80));
    console.error('❌ ERROR IN mainSync():');
    console.error('Message:', error.message);
    console.error('Stack trace:', error.stack);
    console.error('='.repeat(80) + '\n');
    throw error;
  }
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
  console.log('\n--- PROCESSING CUSTOMERS ---');
  const customers = rows.map((row, index) => {
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
      nextReminder = lastVisit.plus({ months: 3 }).toFormat("dd-MM-yyyy");
      obj["Next Reminder Date"] = nextReminder;
      console.log(`✅ Customer ${index + 1}: ${obj["Name"]}, Last Visit: ${obj["Last Visit"]}, Next Reminder: ${nextReminder}`);
    } else {
      obj["Next Reminder Date"] = "";
      console.log(`⚠️  Customer ${index + 1}: ${obj["Name"]}, No valid Last Visit date found: "${obj["Last Visit"]}"`);
    }

    // Manual Contact if no email/phone
    const hasEmail = Boolean(obj["Email Add."]);
    const hasPhone = Boolean(obj["Phone Number"]);
    obj["Manual Contact"] = (!hasEmail && !hasPhone) ? "MISSING CONTACT" : "";

    return obj;
  });
  console.log('--- END PROCESSING CUSTOMERS ---\n');
  return customers;
}

function parseDate(str) {
  if (!str) return null;
  
  const formats = [
    "yyyy-MM-dd",
    "dd/MM/yyyy",
    "MM/dd/yyyy",
    "dd-MM-yyyy",
    "MM-dd-yyyy",
    "d/M/yyyy",
    "d-M-yyyy",
  ];
  
  let dt = DateTime.fromISO(str);
  if (dt.isValid) return dt;
  
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
  const todayFormatted = today.toFormat("dd-MM-yyyy");
  
  console.log('\n' + '='.repeat(80));
  console.log('📧 EMAIL SENDING PROCESS STARTED');
  console.log(`Today's date (ISO): ${todayStr}`);
  console.log(`Today's date (dd-MM-yyyy): ${todayFormatted}`);
  console.log(`Total customers to check: ${customers.length}`);
  console.log('='.repeat(80));
  
  let checkedCount = 0;
  let skippedNotSubscribed = 0;
  let skippedNoContact = 0;
  let skippedNoEmail = 0;
  let skippedNoReminderDate = 0;
  let skippedNotDue = 0;
  let skippedAlreadySent = 0;
  
  for (const customer of customers) {
    checkedCount++;
    const name = customer["Name"] || "Unknown";
    
    console.log(`\n${'─'.repeat(60)}`);
    console.log(`📋 Customer ${checkedCount}/${customers.length}: ${name}`);
    
    // Check subscription status FIRST
    const subscription = (customer["Subscription"] || "").trim().toUpperCase();
    console.log(`   Subscription: "${subscription}"`);
    
    if (subscription === "NOT SUBSCRIBED" || subscription === "UNSUBSCRIBED") {
      console.log(`   🚫 NOT SUBSCRIBED - Skipping`);
      skippedNotSubscribed++;
      continue;
    }
    
    if (subscription !== "SUBSCRIBED" && subscription !== "") {
      console.log(`   ⚠️  Unknown subscription: "${subscription}" - Treating as SUBSCRIBED`);
    }
    
    // Skip customers with missing contact info
    if (customer["Manual Contact"] === "MISSING CONTACT") {
      console.log(`   ⏭️  MISSING CONTACT - Skipping`);
      skippedNoContact++;
      continue;
    }
    
    const to = customer["Email Add."];
    if (!to) {
      console.log(`   ⏭️  NO EMAIL ADDRESS - Skipping`);
      skippedNoEmail++;
      continue;
    }

    const nextReminderStr = customer["Next Reminder Date"];
    const lastEmailSent = customer["Last Email Sent"] || "";
    const lastEmailType = customer["Email Type"] || "";
    
    console.log(`   Email: ${to}`);
    console.log(`   Vehicle: ${customer["Veh. Reg. No."]}`);
    console.log(`   Last Visit: ${customer["Last Visit"]}`);
    console.log(`   Next Reminder: ${nextReminderStr}`);
    console.log(`   Last Email Sent: ${lastEmailSent} (${lastEmailType})`);
    
    // Parse reminder date
    if (!nextReminderStr) {
      console.log(`   ⏭️  NO REMINDER DATE SET - Skipping`);
      skippedNoReminderDate++;
      continue;
    }

    const nextReminder = parseDate(nextReminderStr);
    if (!nextReminder || !nextReminder.isValid) {
      console.log(`   ⚠️  INVALID REMINDER DATE: ${nextReminderStr} - Skipping`);
      skippedNoReminderDate++;
      continue;
    }

    const reminderDate = nextReminder.startOf("day");
    const daysUntilDue = reminderDate.diff(today, "days").days;
    
    console.log(`   Days until due: ${Math.round(daysUntilDue)}`);
    console.log(`   Reminder date: ${reminderDate.toISODate()} vs Today: ${todayStr}`);
    
    // Determine email type based on timing
    let emailType = null;
    let template = null;
    
    if (daysUntilDue < 0) {
      emailType = "OVERDUE";
      template = overdueEmailTemplate(customer);
      console.log(`   🔴 OVERDUE by ${Math.abs(Math.round(daysUntilDue))} days`);
    } else if (daysUntilDue === 0) {
      emailType = "DUE_TODAY";
      template = dueTodayEmailTemplate(customer);
      console.log(`   🟡 DUE TODAY`);
    } else if (daysUntilDue <= 7 && daysUntilDue > 0) {
      emailType = "ADVANCE_7DAY";
      template = advanceReminderEmailTemplate(customer, Math.round(daysUntilDue));
      console.log(`   🟢 DUE IN ${Math.round(daysUntilDue)} DAYS - Advance reminder`);
    } else {
      console.log(`   ⏭️  TOO EARLY (${Math.round(daysUntilDue)} days away) - Skipping`);
      skippedNotDue++;
      continue;
    }
    
    // Check if we already sent this type of email today
    if (lastEmailSent === todayStr && lastEmailType === emailType) {
      console.log(`   ⏭️  ALREADY SENT ${emailType} TODAY - Skipping`);
      skippedAlreadySent++;
      continue;
    }
    
    // Send email
    console.log(`   📧 SENDING ${emailType} EMAIL...`);
    try {
      await sendEmail(to, template);
      sent++;
      
      // Update tracking fields
      customer["Last Email Sent"] = todayStr;
      customer["Email Type"] = emailType;
      
      console.log(`   ✅ EMAIL SENT SUCCESSFULLY!`);
    } catch (e) {
      failed++;
      const errorMsg = `${name} (${to}): ${e.message}`;
      failures.push(errorMsg);
      console.error(`   ❌ FAILED: ${e.message}`);
    }
  }
  
  console.log('\n' + '='.repeat(80));
  console.log('📊 EMAIL SENDING SUMMARY');
  console.log(`Total customers checked: ${checkedCount}`);
  console.log(`✅ Emails sent: ${sent}`);
  console.log(`❌ Emails failed: ${failed}`);
  console.log(`\nSkip Reasons:`);
  console.log(`   🚫 Not subscribed: ${skippedNotSubscribed}`);
  console.log(`   📭 No contact info: ${skippedNoContact}`);
  console.log(`   📧 No email address: ${skippedNoEmail}`);
  console.log(`   📅 No reminder date: ${skippedNoReminderDate}`);
  console.log(`   ⏰ Not due yet: ${skippedNotDue}`);
  console.log(`   🔁 Already sent today: ${skippedAlreadySent}`);
  
  if (failures.length > 0) {
    console.log(`\nFailures:\n   - ${failures.join("\n   - ")}`);
  }
  console.log('='.repeat(80) + '\n');
  
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
  console.log(`      → Email Provider: ${EMAIL_PROVIDER}`);
  console.log(`      → Sending to: ${to}`);
  console.log(`      → Subject: ${subject}`);
  
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