// POST /api/send-push
// Body: { subscription: { endpoint, keys: { p256dh, auth } }, payload: { title, body, url, tag } }
//
// Called by the NEXUS Apps Script backend (never by the browser directly).
// Holds the VAPID private key as a Vercel env var — this is the one piece
// Apps Script can't do itself (Web Push requires ECDH + AES-GCM payload
// encryption, which Apps Script's crypto tools don't cover).
//
// Setup (one-time):
//   1. vercel env add VAPID_PUBLIC_KEY       (paste the public key)
//   2. vercel env add VAPID_PRIVATE_KEY      (paste the private key — keep this secret)
//   3. vercel env add VAPID_SUBJECT          (e.g. mailto:you@abion.org)
//   4. Add "web-push" to package.json dependencies, then redeploy.

const webpush = require('web-push');

module.exports = async function handler(req, res) {
  // Basic CORS: Apps Script's UrlFetchApp doesn't send an Origin header that
  // needs CORS, but this keeps the endpoint usable from a browser too if needed.
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, VAPID_SUBJECT } = process.env;
  if (!VAPID_PUBLIC_KEY || !VAPID_PRIVATE_KEY) {
    return res.status(500).json({ error: 'VAPID keys not configured on the server.' });
  }
  webpush.setVapidDetails(VAPID_SUBJECT || 'mailto:admin@abion.org', VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);

  let body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (e) { return res.status(400).json({ error: 'Invalid JSON body' }); }
  }
  const { subscription, payload } = body || {};
  if (!subscription || !subscription.endpoint) {
    return res.status(400).json({ error: 'Missing subscription' });
  }

  try {
    await webpush.sendNotification(subscription, JSON.stringify(payload || {}));
    return res.status(200).json({ success: true });
  } catch (err) {
    // 404/410 means the subscription is stale (user revoked, browser data cleared, etc.)
    // — that's expected over time and not something to alert on.
    const stale = err.statusCode === 404 || err.statusCode === 410;
    return res.status(stale ? 200 : 500).json({ success: stale, error: stale ? undefined : err.message });
  }
};
