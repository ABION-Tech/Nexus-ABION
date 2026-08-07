// api/gas-proxy.js
//
// Same-origin proxy for the NEXUS Apps Script web app.
// The browser calls THIS endpoint (same domain as the frontend, e.g.
// https://nexus-abion.vercel.app/api/gas-proxy), so the browser never makes a
// cross-origin request and CORS never applies. This function then does a
// plain server-to-server fetch to Apps Script — server-to-server calls are
// not subject to CORS at all — and pipes the raw response straight back.
//
// This preserves full JSON responses and real error messages from doGet/
// doPost, unlike a `mode: 'no-cors'` workaround.

export const config = {
  api: {
    bodyParser: false, // we forward the raw request body untouched
  },
};

const GAS_URL =
  'https://script.google.com/macros/s/AKfycbzrR4cvDD0pwNkJQgOlVLufIzjaJADcdz4g2FAOy--3I3jdeB4yBvpe5CahfhiLWqti/exec';

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

export default async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      const qs = req.url.includes('?') ? req.url.slice(req.url.indexOf('?')) : '';
      const gasRes = await fetch(GAS_URL + qs, {
        method: 'GET',
        redirect: 'follow',
      });
      const text = await gasRes.text();
      res.status(gasRes.status).send(text);
      return;
    }

    if (req.method === 'POST') {
      const rawBody = await readRawBody(req);
      const gasRes = await fetch(GAS_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: rawBody,
      });
      const text = await gasRes.text();
      res.status(gasRes.status).send(text);
      return;
    }
    

    res.status(405).send('Method not allowed');
  } catch (e) {
    res.status(502).json({ error: 'gas-proxy failed: ' + e.message });
  }
}
