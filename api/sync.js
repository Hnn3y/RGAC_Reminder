import { syncAndNotify } from '../index';

export default async function handler(req, res) {
  try {
    await syncAndNotify();
    res.status(200).json({ status: 'success', message: 'Sync completed' });
  } catch (err) {
    console.error('Error in /api/sync:', err);
    res.status(500).json({ error: err.message });
  }
}
