import { syncAndNotify } from '../../sync.js'; // your existing logic

export default async function handler(req, res) {
  try {
    await syncAndNotify();
    res.status(200).json({ message: 'Sync complete. Emails sent & sheets updated.' });
  } catch (err) {
    console.error('Error in /api/sync:', err);
    res.status(500).json({ error: err.message });
  }
}
