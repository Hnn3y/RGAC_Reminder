import { syncAndNotify } from '../src/index.js';

export const config = {
  runtime: 'edge',
  schedule: '0 9 * * *', // run daily at 9 AM UTC
};

export default async function handler() {
  try {
    await syncAndNotify();
    return new Response(JSON.stringify({ message: 'Sync complete via cron' }), { status: 200 });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), { status: 500 });
  }
}
