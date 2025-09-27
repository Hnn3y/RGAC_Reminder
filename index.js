import 'dotenv/config';
import fs from 'fs';
import { google } from 'googleapis';
import { DateTime } from 'luxon';
import nodemailer from 'nodemailer';
import sgMail from '@sendgrid/mail';


function convertSerialToDate(serial) {
  const baseDate = new Date(1899, 11, 30);
  const days = Math.floor(serial);
  const milliseconds = days * 24 * 60 * 60 * 1000;
  return new Date(baseDate.getTime() + milliseconds);
}


const KEYFILE = process.env.GOOGLE_CREDENTIALS_PATH || './config/google.json';
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

if (!SPREADSHEET_ID) {
  console.error('SPREADSHEET_ID must be set in .env');
  process.exit(1);
}
if (!fs.existsSync(KEYFILE)) {
  console.error('Google credentials file not found at', KEYFILE);
  process.exit(1);
}


const auth = new google.auth.GoogleAuth({
  keyFile: KEYFILE,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });


let emailSender = null;
if (EMAIL_PROVIDER === 'gmail') {
  if (!EMAIL_USER || !EMAIL_PASS) {
    console.warn('Gmail configured but EMAIL_USER/EMAIL_PASS missing.');
  } else {
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: { user: EMAIL_USER, pass: EMAIL_PASS },
    });
    emailSender = async (to, subject, text) => {
      return transporter.sendMail({ from: EMAIL_USER, to, subject, text });
    };
  }
} else if (EMAIL_PROVIDER === 'sendgrid') {
  if (!SENDGRID_API_KEY) {
    console.warn('SendGrid configured but SENDGRID_API_KEY missing.');
  } else {
    sgMail.setApiKey(SENDGRID_API_KEY);
    emailSender = async (to, subject, text) => {
      return sgMail.send({ to, from: EMAIL_FROM, subject, text });
    };
  }
} else {
  console.warn('Unknown EMAIL_PROVIDER, no email will be sent.');
}


function serialToJSDate(serial) {
  const n = Number(serial);
  if (Number.isFinite(n)) {
    const ms = Math.round((n - 25569) * 86400 * 1000);
    return new Date(ms);
  }
  return null;
}


function isoDateFromAny(value) {
  if (value === null || value === undefined || value === '') return null;

  if (typeof value === 'number' || (!isNaN(value) && typeof value !== 'object')) {
    const maybeDate = serialToJSDate(value);
    if (maybeDate instanceof Date && !Number.isNaN(maybeDate.getTime())) {
      return DateTime.fromJSDate(maybeDate, { zone: 'utc' }).toISODate(); // YYYY-MM-DD
    }
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return DateTime.fromJSDate(value, { zone: 'utc' }).toISODate();
  }

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


async function syncAndNotify() {
  const { headers, rows } = await readMaster();
  if (!headers.length) {
    console.error('Master sheet empty or missing headers.');
    return;
  }

  const lowered = headers.map(h => (h || '').toString().toLowerCase());
  const pickIndex = (variants) => {
    for (const v of variants) {
      const i = lowered.indexOf(v.toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };

  const idxName = pickIndex(['name', 'customer name', 'full name']);
  const idxPlate = pickIndex(['plate number', 'plate', 'plate_no']);
  const idxEmail = pickIndex(['email', 'email address']);
  const idxPhone = pickIndex(['phone', 'phone number', 'mobile']);
  const idxLast = pickIndex(['last service date', 'lastservicedate', 'last visit', 'last visit date']);
  let idxNext = pickIndex(['next reminder date', 'nextreminderdate', 'next visit', 'next visit date']);
  let idxManual = pickIndex(['manual contact', 'manual_contact', 'manualcontact']);
  const willAppendNext = idxNext === -1;
  if (willAppendNext) { idxNext = headers.length; headers.push('Next Reminder Date'); }
  const willAppendManual = idxManual === -1;
  if (willAppendManual) { idxManual = headers.length; headers.push('Manual Contact'); }

  const customers = rows.map((row, i) => {
    const name = row[idxName] || '';
    const plate = row[idxPlate] || '';
    const email = row[idxEmail] || '';
    const phone = row[idxPhone] || '';
    const lastRaw = row[idxLast] || '';
    return {
      originalRow: i + 2, 
      name, plate, email, phone, lastRaw,
      existingNext: row[idxNext] || '',
      existingManual: row[idxManual] || ''
    };
  });

  const nextDatesDisplay = []; 
  const manualFlags = [];      

  for (const c of customers) {
    let isoLast = null;
if (!isNaN(c.lastRaw)) {
  const d = serialToJSDate(c.lastRaw);
  if (d) isoLast = DateTime.fromJSDate(d, { zone: 'utc' }).toISODate();
} else {
  isoLast = isoDateFromAny(c.lastRaw);
}
    let nextDisplay = '';
    if (isoLast) {
      const nextIso = DateTime.fromISO(isoLast, { zone: 'utc' }).plus({ months: 3 }).toISODate();
      nextDisplay = isoToDisplayDDMMYYYY(nextIso); 
    }
    nextDatesDisplay.push(nextDisplay);
    const manual = (!c.email && !c.phone) ? 'MISSING CONTACT' : '';
    manualFlags.push(manual);
  }

  await updateMasterColumn(idxNext, nextDatesDisplay);
  await updateMasterColumn(idxManual, manualFlags);

  const remindersHeader = ['Name','Plate Number','Email','Phone','Last Service Date','Next Reminder Date','Manual Contact','Status'];
  const remindersRows = customers.map((c, i) => [
    c.name || '', c.plate || '', c.email || '', c.phone || '', c.lastRaw || '', nextDatesDisplay[i] || '', manualFlags[i] || '', ''
  ]);
  remindersRows.sort((a,b) => String(a[0]||'').toLowerCase().localeCompare(String(b[0]||'').toLowerCase()));

  await writeSheet(REMINDERS_SHEET, [remindersHeader, ...remindersRows]);

  const masterRowsWithHeader = [headers, ...customers.map((c, i) => {
    const originalRow = rows[c.originalRow - 2] || [];
    const rowCopy = originalRow.slice();
    while (rowCopy.length < headers.length) rowCopy.push('');
    rowCopy[idxNext] = nextDatesDisplay[c.originalRow - 2] || '';
    rowCopy[idxManual] = manualFlags[c.originalRow - 2] || '';
    return rowCopy;
  })];

  const masterSortedContentRows = masterRowsWithHeader.slice(1).slice().sort((r1, r2) => {
    const n1 = String(r1[idxName] || '').toLowerCase();
    const n2 = String(r2[idxName] || '').toLowerCase();
    return n1.localeCompare(n2);
  });
  const masterSortedValues = [headers, ...masterSortedContentRows];
  await writeSheet(MASTER_SORTED_SHEET, masterSortedValues);

  if (OVERWRITE_MASTER) {
    await writeSheet(MASTER_SHEET, masterSortedValues);
  }

  const todayISO = DateTime.utc().toISODate();
  for (let i=0; i<customers.length; i++) {
    const c = customers[i];
    const nextDisplay = nextDatesDisplay[i];
    if (!nextDisplay) continue; 

    const parsed = DateTime.fromFormat(nextDisplay, 'dd-MM-yyyy', { zone: 'utc' });
    if (!parsed.isValid) continue;
    const dueISO = parsed.minus({ days: SEND_OFFSET_DAYS }).toISODate();
    if (dueISO <= todayISO) {
      let sendResult = 'skipped';
      let note = '';
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

  console.log('Sync complete. Reminders sheet & Master_Sorted updated. Status Log appended for sends.');
}


syncAndNotify().catch(err => {
  console.error('Fatal error in syncAndNotify:', err);
  process.exit(1);
});
