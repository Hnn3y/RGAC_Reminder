import { google } from 'googleapis';
import { DateTime } from 'luxon';
import nodemailer from 'nodemailer';
import sgMail from '@sendgrid/mail';

export async function syncAndNotify(env) {
  const {
    SPREADSHEET_ID,
    MASTER_SHEET = 'Master',
    REMINDERS_SHEET = 'Reminders',
    STATUS_LOG_SHEET = 'Status Log',
    EMAIL_PROVIDER = 'gmail',
    EMAIL_USER = '',
    EMAIL_PASS = '',
    SENDGRID_API_KEY = '',
    EMAIL_FROM = EMAIL_USER,
    GOOGLE_CLIENT_EMAIL,
    GOOGLE_PRIVATE_KEY,
    GOOGLE_PROJECT_ID,
  } = env;

  if (!SPREADSHEET_ID || !GOOGLE_CLIENT_EMAIL || !GOOGLE_PRIVATE_KEY || !GOOGLE_PROJECT_ID) {
    throw new Error('Missing required environment variables.');
  }

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: GOOGLE_CLIENT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    projectId: GOOGLE_PROJECT_ID,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  // Email sender setup (optional)
  let emailSender = null;
  const provider = EMAIL_PROVIDER.toLowerCase();
  if (provider === 'gmail' && EMAIL_USER && EMAIL_PASS) {
    const transporter = nodemailer.createTransport({ service: 'gmail', auth: { user: EMAIL_USER, pass: EMAIL_PASS } });
    emailSender = async (to, subject, text) => transporter.sendMail({ from: EMAIL_USER, to, subject, text });
  } else if (provider === 'sendgrid' && SENDGRID_API_KEY) {
    sgMail.setApiKey(SENDGRID_API_KEY);
    emailSender = async (to, subject, text) => sgMail.send({ to, from: EMAIL_FROM, subject, text });
  }

  // --- Utility functions ---
  const serialToJSDate = serial => (!isNaN(Number(serial)) ? new Date(Math.round((serial - 25569) * 86400 * 1000)) : null);

  const isoDateFromAny = value => {
    if (!value && value !== 0) return null;
    if (typeof value === 'number') return DateTime.fromJSDate(serialToJSDate(value), { zone: 'utc' }).toISODate();
    const s = String(value).trim();
    if (!s) return null;
    const formats = [
      () => DateTime.fromISO(s, { zone: 'utc' }),
      () => DateTime.fromFormat(s, 'dd-MM-yyyy', { zone: 'utc' }),
      () => DateTime.fromFormat(s, 'dd/MM/yyyy', { zone: 'utc' }),
      () => DateTime.fromFormat(s, 'MM/dd/yyyy', { zone: 'utc' }),
      () => DateTime.fromRFC2822(s, { zone: 'utc' }),
      () => DateTime.fromJSDate(new Date(s), { zone: 'utc' }),
    ];
    for (const f of formats) {
      const dt = f();
      if (dt.isValid) return dt.toISODate();
    }
    return null;
  };

  const isoToDisplayDDMMYYYY = iso => {
    if (!iso) return '';
    const dt = DateTime.fromISO(iso, { zone: 'utc' });
    return dt.isValid ? dt.toFormat('dd-MM-yyyy') : '';
  };

  // --- Sheet helpers ---
  const readMaster = async () => {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: MASTER_SHEET, valueRenderOption: 'UNFORMATTED_VALUE' });
    const values = res.data.values || [];
    if (!values.length) return { headers: [], rows: [] };
    const headers = values[0].map(h => String(h || '').trim());
    const rows = values.slice(1).map(r => r.map(c => c || ''));
    return { headers, rows };
  };

  const ensureSheetExists = async sheetName => {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    if (!(meta.data.sheets || []).some(s => s.properties.title === sheetName)) {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] } });
    }
  };

  const writeSheet = async (sheetName, values2D) => {
    await ensureSheetExists(sheetName);
    await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: sheetName });
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: sheetName, valueInputOption: 'RAW', requestBody: { values: values2D } });
  };

  const updateMasterColumn = async (headerIndex, rowsData) => {
    const colLetter = n => {
      let s = '';
      while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); }
      return s;
    };
    const letter = colLetter(headerIndex + 1);
    const startRow = 2;
    const endRow = startRow + rowsData.length - 1;
    const range = `${MASTER_SHEET}!${letter}${startRow}:${letter}${endRow}`;
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range, valueInputOption: 'RAW', requestBody: { values: rowsData.map(v => [v || '']) } });
  };

  // --- Main logic ---
  const { headers, rows } = await readMaster();
  if (!headers.length) throw new Error('Master sheet empty or missing headers.');

  const lowered = headers.map(h => h.toLowerCase());
  const pickIndex = variants => variants.map(v => lowered.indexOf(v.toLowerCase())).find(i => i >= 0) ?? -1;

  const idxName = pickIndex(['name','customer name','full name']);
  const idxPlate = pickIndex(['plate number','plate','plate_no']);
  const idxEmail = pickIndex(['email','email address']);
  const idxPhone = pickIndex(['phone','phone number','mobile']);
  const idxLast = pickIndex(['last service date','lastservicedate','last visit','last visit date']);
  let idxNext = pickIndex(['next reminder date','nextreminderdate','next visit','next visit date']);
  let idxManual = pickIndex(['manual contact','manual_contact','manualcontact']);

  if (idxNext === -1) { idxNext = headers.length; headers.push('Next Reminder Date'); }
  if (idxManual === -1) { idxManual = headers.length; headers.push('Manual Contact'); }

  const customers = rows.map((row, i) => ({
    originalRow: i + 2,
    name: row[idxName] || '',
    plate: row[idxPlate] || '',
    email: row[idxEmail] || '',
    phone: row[idxPhone] || '',
    lastRaw: row[idxLast] || '',
    existingNext: row[idxNext] || '',
    existingManual: row[idxManual] || ''
  }));

  const nextDatesDisplay = [];
  const manualFlags = [];

  for (const c of customers) {
    const isoLast = isoDateFromAny(c.lastRaw);
    const nextISO = isoLast ? DateTime.fromISO(isoLast, { zone: 'utc' }).plus({ months: 3 }).toISODate() : null;
    nextDatesDisplay.push(isoToDisplayDDMMYYYY(nextISO));
    manualFlags.push(!c.email && !c.phone ? 'MISSING CONTACT' : '');
  }

  await updateMasterColumn(idxNext, nextDatesDisplay);
  await updateMasterColumn(idxManual, manualFlags);

  const remindersHeader = ['Name','Plate Number','Email','Phone','Last Service Date','Next Reminder Date','Manual Contact','Status'];
  const remindersRows = customers.map((c, i) => [
    c.name, c.plate, c.email, c.phone, c.lastRaw, nextDatesDisplay[i], manualFlags[i], ''
  ]).sort((a, b) => String(a[0]).toLowerCase().localeCompare(String(b[0]).toLowerCase()));

  await writeSheet(REMINDERS_SHEET, [remindersHeader, ...remindersRows]);
}
