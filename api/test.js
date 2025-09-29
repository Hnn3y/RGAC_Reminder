import { google } from 'googleapis';

export default async function handler(req, res) {
  try {
    console.log('API called!');

    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_CLIENT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    const sheet = await sheets.spreadsheets.get({
      spreadsheetId: process.env.SPREADSHEET_ID,
    });

    console.log('Sheets data:', sheet.data.sheets.map(s => s.properties.title));

    res.status(200).json({ ok: true, sheets: sheet.data.sheets.map(s => s.properties.title) });
  } catch (err) {
    console.error('Crash point:', err);
    res.status(500).json({ error: err.message });
  }
}
