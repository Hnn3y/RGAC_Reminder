import { google } from 'googleapis';
import { DateTime } from 'luxon';
import nodemailer from 'nodemailer';
import sgMail from '@sendgrid/mail';

// Use only env, do not read local files
function getConfig(env = process.env) {
  // Google credentials via env, for Vercel
  const {
    SPREADSHEET_ID,
    MASTER_SHEET = 'Master',
    REMINDERS_SHEET = 'Reminders',
    STATUS_LOG_SHEET = 'Status Log',
    EMAIL_PROVIDER = 'gmail',
    EMAIL_USER,
    EMAIL_PASS,
    SENDGRID_API_KEY,
    EMAIL_FROM,
    GOOGLE_CLIENT_EMAIL,
    GOOGLE_PRIVATE_KEY,
    GOOGLE_PROJECT_ID,
  } = env;
  if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID missing');
  if (!GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY || !GOOGLE_PROJECT_ID)
    throw new Error('Google Service Account env vars missing');
  return {
    SPREADSHEET_ID,
    MASTER_SHEET,
    REMINDERS_SHEET,
    STATUS_LOG_SHEET,
    EMAIL_PROVIDER,
    EMAIL_USER,
    EMAIL_PASS,
    SENDGRID_API_KEY,
    EMAIL_FROM: EMAIL_FROM || EMAIL_USER,
    GOOGLE_CLIENT_EMAIL,
    GOOGLE_PRIVATE_KEY: GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    GOOGLE_PROJECT_ID,
  };
}

function getSheetsClient(env) {
  const auth = new google.auth.JWT(
    env.GOOGLE_CLIENT_EMAIL,
    null,
    env.GOOGLE_PRIVATE_KEY,
    ['https://www.googleapis.com/auth/spreadsheets']
  );
  return google.sheets({ version: 'v4', auth });
}

function parseDate(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    return new Date(Date.UTC(1899, 11, 30) + val * 86400000);
  }
  if (typeof val === 'string') {
    let n = Number(val);
    if (!isNaN(n) && n > 40000 && n < 60000) return new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    let d = new Date(val);
    if (!isNaN(d.getTime())) return d;
    let m = val.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (m) {
      let dd = parseInt(m[1], 10), mm = parseInt(m[2], 10), yyyy = parseInt(m[3], 10);
      if (dd > 12) return new Date(yyyy, mm - 1, dd);
      if (mm > 12) return new Date(yyyy, dd - 1, mm);
      return new Date(yyyy, mm - 1, dd);
    }
  }
  return null;
}

function formatDateISO(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  return date.toISOString().slice(0, 10);
}

function add3Months(date) {
  const d = new Date(date.getTime());
  d.setMonth(d.getMonth() + 3);
  if (d.getDate() !== date.getDate()) d.setDate(0);
  return d;
}

// Main logic: sync sheets, update, return result
export async function syncAndNotify(envOverride) {
  const env = getConfig(envOverride);
  const sheets = getSheetsClient(env);
  // 1. Read master
  const masterRes = await sheets.spreadsheets.values.get({
    spreadsheetId: env.SPREADSHEET_ID,
    range: env.MASTER_SHEET,
  });
  let [headersRaw, ...rowsRaw] = masterRes.data.values || [[]];
  const headers = headersRaw.map(h => h.trim());
  // Header mapping
  const colMap = {};
  for (const [std, variants] of Object.entries({
    name: ['name', 'customer name', 'full name'],
    plate: ['plate number', 'plate', 'plate_no'],
    email: ['email', 'email address'],
    phone: ['phone', 'phone number', 'mobile'],
    lastService: ['last service date', 'lastservicedate', 'last visit', 'last visit date'],
    nextReminder: ['next reminder date', 'nextreminderdate', 'next visit', 'next visit date'],
    manualContact: ['manual contact', 'manual_contact', 'manualcontact'],
  })) {
    colMap[std] = headers.findIndex(h => variants.includes(h.trim().toLowerCase()));
  }
  // Add columns if missing
  let updateHeaders = false;
  if (colMap.nextReminder === -1) {
    headers.push('Next Reminder Date');
    colMap.nextReminder = headers.length - 1;
    updateHeaders = true;
  }
  if (colMap.manualContact === -1) {
    headers.push('Manual Contact');
    colMap.manualContact = headers.length - 1;
    updateHeaders = true;
  }
  // Update master headers if needed
  if (updateHeaders) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: env.SPREADSHEET_ID,
      range: `${env.MASTER_SHEET}!1:1`,
      valueInputOption: 'RAW',
      requestBody: { values: [headers] },
    });
  }
  // Process rows
  const processed = [];
  for (let row of rowsRaw) {
    while (row.length < headers.length) row.push('');
    const name = row[colMap.name] || '';
    const plate = row[colMap.plate] || '';
    const email = (row[colMap.email] || '').trim();
    const phone = (row[colMap.phone] || '').trim();
    const lastVal = row[colMap.lastService];
    const lastDate = parseDate(lastVal);
    const nextDate = lastDate ? add3Months(lastDate) : null;
    const nextIso = nextDate ? formatDateISO(nextDate) : '';
    const missingContact = !email && !phone ? 'MISSING CONTACT' : '';
    row[colMap.nextReminder] = nextIso;
    row[colMap.manualContact] = missingContact;
    processed.push({
      Name: name,
      'Plate Number': plate,
      Email: email,
      Phone: phone,
      'Last Service Date': lastVal,
      'Next Reminder Date': nextIso,
      'Manual Contact': missingContact,
      Status: '',
    });
  }
  // Write updated master
  await sheets.spreadsheets.values.update({
    spreadsheetId: env.SPREADSHEET_ID,
    range: `${env.MASTER_SHEET}`,
    valueInputOption: 'RAW',
    requestBody: { values: [headers, ...rowsRaw] },
  });
  // Write reminders (sorted)
  const remindersHeaders = [
    'Name', 'Plate Number', 'Email', 'Phone', 'Last Service Date', 'Next Reminder Date', 'Manual Contact', 'Status',
  ];
  processed.sort((a, b) => (a.Name || '').localeCompare(b.Name || ''));
  await sheets.spreadsheets.values.update({
    spreadsheetId: env.SPREADSHEET_ID,
    range: env.REMINDERS_SHEET,
    valueInputOption: 'RAW',
    requestBody: { values: [remindersHeaders, ...processed.map(row => remindersHeaders.map(h => row[h] || ''))] },
  });
  // Status log
  await sheets.spreadsheets.values.append({
    spreadsheetId: env.SPREADSHEET_ID,
    range: env.STATUS_LOG_SHEET,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [[new Date().toISOString(), 'success', `Processed ${processed.length} customers`]] },
  });
  // Return summary
  return { ok: true, processed: processed.length };
}

// Email sending (optional, ready-to-use)
export async function sendEmail({ to, subject, text, html }, envOverride) {
  const env = getConfig(envOverride);
  if (env.EMAIL_PROVIDER === 'gmail') {
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: env.EMAIL_USER, pass: env.EMAIL_PASS },
    });
    await transporter.sendMail({ from: env.EMAIL_FROM, to, subject, text, html });
  } else if (env.EMAIL_PROVIDER === 'sendgrid') {
    sgMail.setApiKey(env.SENDGRID_API_KEY);
    await sgMail.send({ from: env.EMAIL_FROM, to, subject, text, html });
  } else {
    throw new Error('Unsupported EMAIL_PROVIDER');
  }
}

export default syncAndNotify;

console.log("GOOGLE_CLIENT_EMAIL:", process.env.GOOGLE_CLIENT_EMAIL);
console.log("GOOGLE_PRIVATE_KEY starts with:", process.env.GOOGLE_PRIVATE_KEY && process.env.GOOGLE_PRIVATE_KEY.slice(0, 30));
console.log("GOOGLE_PROJECT_ID:", process.env.GOOGLE_PROJECT_ID);