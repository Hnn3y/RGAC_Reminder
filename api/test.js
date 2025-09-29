import { syncAndNotify } from '../index.js';

export default async function handler(req, res) {
  try {
    await syncAndNotify();
    res.status(200).json({ message: "Sync complete" });
  } catch (err) {
    console.error('Error in sync API:', err);
    res.status(500).json({ error: err.message });
  }
}
