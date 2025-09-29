import { syncAndNotify } from '../index.js';

export default async function handler(req, res) {
  try {
    const result = await syncAndNotify(process.env);
    res.status(200).json(result);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
} 