import { syncAndNotify } from '../lib/syncLogic';

export default async function handler(req, res) {
  try {
    await syncAndNotify(process.env);
    res.status(200).json({ status: 'success', message: 'Sync completed' });
  } catch (err) {
    console.error('Error in /sync:', err);
    res.status(500).json({ status: 'error', message: err.message });
  }
}
