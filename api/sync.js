import 'dotenv/config';
import { google } from 'googleapis';
import { DateTime } from 'luxon';
import nodemailer from 'nodemailer';
import sgMail from '@sendgrid/mail';

// ---------- CONFIG ----------
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const MASTER_SHEET = process.env.MASTER_SHEET || 'Master';
const REMINDERS_SHEET = process.env.REMINDERS_SHEET || 'Reminders';
const MASTER_SORTED_SHEET = process.env.MASTER_SORTED_SHEET || 'Master_Sorted';
const STATUS_LOG_SHEET = process.env.STATUS_LOG_SHEET || 'Status Log';

const EMAIL_PROVIDER = (process.env.EMAIL_PROVIDER || 'gmail').toLowerCase();
const EMAIL_USER = process.env.EMAIL_USER || '';
const EMAIL_PASS = process.env.EMAIL_PASS || '';
const SENDGRID_API_KEY = process.env.SENDGRID_API_KEY || '';
const EMAIL_FROM = process.env.EMAIL_FROM || EMAIL_USER;

const OVERWRITE_MASTER = (process.env.OVERWRITE_MASTER || 'false').toLowerCase() === 'true';
const SEND_OFFSET_DAYS = Number(process.env.SEND_OFFSET_DAYS || 0);

// ---------- GOOGLE SHEETS AUTH ----------
if (!SPREADSHEET_ID) throw new Error('SPREADSHEET_ID must be set in .env');

const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  },
  projectId: process.env.GOOGLE_PROJECT_ID,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

// ---------- EMAIL SENDER ----------
let emailSender = null;

if (EMAIL_PROVIDER === 'gmail') {
  if (EMAIL_USER && EMAIL_PASS) {
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: EMAIL_USER, pass: EMAIL_PASS },
    });
    emailSender = async (to, subject, text) => transporter.sendMail({ from: EMAIL_USER, to, subject, text });
  }
} else if (EMAIL_PROVIDER === 'sendgrid') {
  if (SENDGRID_API_KEY) {
    sgMail.setApiKey(SENDGRID_API_KEY);
    emailSender = async (to, subject, text) => sgMail.send({ to, from: EMAIL_FROM, subject, text });
  }
}

// ---------- UTILITY FUNCTIONS ----------
function serialToJSDate(serial) {
  const n = Number(serial);
  if (!Number.isFinite(n)) return null;
  const ms = Math.round((n - 25569) * 86400 * 1000);
  return new Date(ms);
}

function isoDateFromAny(value) {
  if (value === null || value === undefined || value === '') return null;

  if (typeof value === 'number' || (!isNaN(value) && typeof value !== 'object')) {
    const maybeDate = serialToJSDate(value);
    if (maybeDate instanceof Date && !Number.isNaN(maybeDate.getTime())) {
      return DateTime.fromJSDate(maybeDate, { zone: 'utc' }).toISODate();
    }
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) return DateTime.fromJSDate(value, { zone: 'utc' }).toISODate();

  const s = String(value).trim();
  if (!s) return null;

  let dt = DateTime.fromISO(s, { zone: 'utc' });
  if (dt.isValid) return dt.toISODate();

  dt = DateTime.fromFormat(s, 'dd-MM-yyyy', { zone: 'utc' });
  if (dt.isValid) return dt.toISODate();

  dt = DateTime.fromFormat(s, 'dd/MM/yyyy', { zone: 'utc' });
  if (dt.isValid) return dt.toISODate();

  dt = DateTime.fromFormat(s, 'MM/dd/yyyy', { zone: 'utc' });
  if (dt.isValid) return dt.toISODate();

  dt = DateTime.fromRFC2822(s, { zone: 'utc' });
  if (dt.isValid) return dt.toISODate();

  const js = new Date(s);
  if (!Number.isNaN(js.getTime())) return DateTime.fromJSDate(js, { zone: 'utc' }).toISODate();

  return null;
}

function isoToDisplayDDMMYYYY(iso) {
  if (!iso) return '';
  const dt = DateTime.fromISO(iso, { zone: 'utc' });
  if (!dt.isValid) return '';
  return dt.toFormat('dd-MM-yyyy');
}

async function ensureSheetExists(sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const found = (meta.data.sheets || []).find(s => s.properties.title === sheetName);
  if (!found) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
    });
  }
}

async function readMaster() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: MASTER_SHEET,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const values = res.data.values || [];
  if (values.length === 0) return { headers: [], rows: [] };
  const headers = values[0].map(h => (h === undefined ? '' : String(h).trim()));
  const rows = values.slice(1).map(r => r.map(cell => (cell === undefined ? '' : cell)));
  return { headers, rows };
}

async function writeSheet(sheetName, values2D) {
  await ensureSheetExists(sheetName);
  await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: sheetName });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetName,
    valueInputOption: 'RAW',
    requestBody: { values: values2D }
  });
}

async function appendStatusLog(row) {
  await ensureSheetExists(STATUS_LOG_SHEET);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: STATUS_LOG_SHEET,
    valueInputOption: 'RAW',
    requestBody: { values: [row] }
  });
}

async function updateMasterColumn(headerIndex, rowsData) {
  const toCol = (n) => {
    let s = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };
  const colLetter = toCol(headerIndex + 1);
  const startRow = 2;
  const endRow = startRow + rowsData.length - 1;
  const range = `${MASTER_SHEET}!${colLetter}${startRow}:${colLetter}${endRow}`;
  const values = rowsData.map(v => [v || '']);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'RAW',
    requestBody: { values }
  });
}

// ---------- MAIN SYNC FUNCTION ----------
async function syncAndNotify() {
  const { headers, rows } = await readMaster();
  if (!headers.length) throw new Error('Master sheet empty or missing headers.');

  const lowered = headers.map(h => (h || '').toString().toLowerCase());
  const pickIndex = (variants) => variants.map(v => lowered.indexOf(v.toLowerCase())).find(i => i >= 0) ?? -1;

  const idxName = pickIndex(['name', 'customer name', 'full name']);
  const idxPlate = pickIndex(['plate number', 'plate', 'plate_no']);
  const idxEmail = pickIndex(['email', 'email address']);
  const idxPhone = pickIndex(['phone', 'phone number', 'mobile']);
  const idxLast = pickIndex(['last service date', 'lastservicedate', 'last visit', 'last visit date']);
  let idxNext = pickIndex(['next reminder date', 'nextreminderdate', 'next visit', 'next visit date']);
  let idxManual = pickIndex(['manual contact', 'manual_contact', 'manualcontact']);

  if (idxNext === -1) { idxNext = headers.length; headers.push('Next Reminder Date'); }
  if (idxManual === -1) { idxManual = headers.length; headers.push('Manual Contact'); }

  const customers = rows.map((row, i) => ({
    originalRow: i + 2,
    name: row[idxName] || '',
    plate: row[idxPlate] || '',
    email: row[idxEmail] || '',
    phone: row[idxPhone] || '',
    lastRaw: row[idxLast] || ''
  }));

  const nextDatesDisplay = [];
  const manualFlags = [];

  for (const c of customers) {
    let isoLast = !isNaN(c.lastRaw) ? DateTime.fromJSDate(serialToJSDate(c.lastRaw), { zone: 'utc' }).toISODate() : isoDateFromAny(c.lastRaw);
    let nextDisplay = '';
    if (isoLast) {
      nextDisplay = isoToDisplayDDMMYYYY(DateTime.fromISO(isoLast, { zone: 'utc' }).plus({ months: 3 }).toISODate());
    }
    nextDatesDisplay.push(nextDisplay);
    manualFlags.push(!c.email && !c.phone ? 'MISSING CONTACT' : '');
  }

  await updateMasterColumn(idxNext, nextDatesDisplay);
  await updateMasterColumn(idxManual, manualFlags);

  const remindersHeader = ['Name','Plate Number','Email','Phone','Last Service Date','Next Reminder Date','Manual Contact','Status'];
  const remindersRows = customers.map((c, i) => [c.name, c.plate, c.email, c.phone, c.lastRaw, nextDatesDisplay[i], manualFlags[i], '']);
  remindersRows.sort((a,b) => String(a[0]||'').toLowerCase().localeCompare(String(b[0]||'').toLowerCase()));
  await writeSheet(REMINDERS_SHEET, [remindersHeader, ...remindersRows]);

  const masterRowsWithHeader = [headers, ...customers.map((c, i) => {
    const originalRow = rows[c.originalRow - 2] || [];
    const rowCopy = originalRow.slice();
    while (rowCopy.length < headers.length) rowCopy.push('');
    rowCopy[idxNext] = nextDatesDisplay[i];
    rowCopy[idxManual] = manualFlags[i];
    return rowCopy;
  })];

  const masterSortedValues = [headers, ...masterRowsWithHeader.slice(1).slice().sort((r1, r2) => String(r1[idxName]||'').toLowerCase().localeCompare(String(r2[idxName]||'').toLowerCase()))];
  await writeSheet(MASTER_SORTED_SHEET, masterSortedValues);
  if (OVERWRITE_MASTER) await writeSheet(MASTER_SHEET, masterSortedValues);

  const todayISO = DateTime.utc().toISODate();
  for (let i = 0; i < customers.length; i++) {
    const c = customers[i];
    const nextDisplay = nextDatesDisplay[i];
    if (!nextDisplay) continue;
    const parsed = DateTime.fromFormat(nextDisplay, 'dd-MM-yyyy', { zone: 'utc' });
    if (!parsed.isValid) continue;

    const dueISO = parsed.minus({ days: SEND_OFFSET_DAYS }).toISODate();
    if (dueISO <= todayISO) {
      let sendResult = 'skipped', note = '';
      if (emailSender && c.email) {
        try {
          const subject = `Service Reminder for ${c.name || c.plate || ''}`;
          const body = `Dear ${c.name || ''},\n\nThis is a reminder that your vehicle (${c.plate || ''}) is due for service on ${nextDisplay}.\n\nKindly book an appointment.\n\nBest regards,\nAuto Shop`;
          await emailSender(c.email, subject, body);
          sendResult = 'sent';
          note = 'Email sent';
        } catch (err) {
          sendResult = 'failed';
          note = String(err.message || err);
        }
      } else {
        sendResult = 'no-email';
        note = 'No email or email sender not configured';
      }
      await appendStatusLog([DateTime.utc().toISO(), c.name, c.plate, 'email', c.email || '', sendResult, note]);
    }
  }

  return { message: 'Sync complete. Reminders sheet & Master_Sorted updated. Status Log appended.' };
}

// ---------- VERCEL HANDLER ----------
export default async function handler(req, res) {
  try {
    const result = await syncAndNotify();
    res.status(200).json(result);
  } catch (err) {
    console.error('Error in sync API:', err);
    res.status(500).json({ error: err.message });
  }
}

// ---------- OPTIONAL CRON CONFIG ----------
export const config = {
  runtime: 'edge',
  schedule: '0 9 * * *', // Daily at 9 AM UTC
};
