// server_email_via_graph.js
// Express endpoint to send email from a 365 shared mailbox via Microsoft Graph
// - Supports multiple attachments (small files inline)
// - Subject/body come from client; also persists note via existing notes route (optional)
//
// Env vars required:
//   PORT=5000
//   ALLOWED_ORIGINS=http://localhost:5173,http://localhost:3000
//   GRAPH_TENANT_ID=<tenantId>
//   GRAPH_CLIENT_ID=<appId>
//   GRAPH_CLIENT_SECRET=<secret>
//   SHARED_MAILBOXES=ap@gear4music.com,another@gear4music.com
//
// Permissions (App registration in Entra ID / Azure AD):
//   - Application permission: Mail.Send  (admin consent granted)
//   (This allows POST /users/{sharedMailbox}/sendMail)

import express from 'express';
import cors from 'cors';
import multer from 'multer';
import fetch from 'node-fetch';
import { ClientSecretCredential } from '@azure/identity';

const app = express();
const upload = multer({ limits: { fileSize: 8 * 1024 * 1024 } }); // 8MB per file (Graph inline limit ~3-4MB; we keep 8MB but recommend smaller)

app.use(express.json());
app.use(cors({
  origin: (origin, cb) => {
    if (!origin) return cb(null, true);
    const allowed = (process.env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
    if (!allowed.length || allowed.includes(origin)) return cb(null, true);
    cb(new Error('CORS not allowed'));
  },
  credentials: true,
}));

// --- Graph auth helper ---
const credential = new ClientSecretCredential(
  process.env.GRAPH_TENANT_ID,
  process.env.GRAPH_CLIENT_ID,
  process.env.GRAPH_CLIENT_SECRET
);
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

async function getGraphToken() {
  const token = await credential.getToken(GRAPH_SCOPE);
  return token?.token;
}

function parseEmails(str = '') {
  return str
    .split(/[;,]/)
    .map(s => s.trim())
    .filter(Boolean)
    .map(addr => ({ emailAddress: { address: addr } }));
}

function b64(arrayBuffer) {
  return Buffer.from(arrayBuffer).toString('base64');
}

function allowedFromAddress(addr) {
  const allowed = (process.env.SHARED_MAILBOXES || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  if (!allowed.length) return true; // fallback: allow anything if not configured
  return allowed.includes(String(addr || '').toLowerCase());
}

// --- API: send email ---
app.post('/api/email/send', upload.array('attachments'), async (req, res) => {
  try {
    const {
      from = '',
      to = '',
      cc = '',
      bcc = '',
      subject = '',
      body = '',
      invoiceId = '',
    } = req.body;

    if (!from || !allowedFromAddress(from)) {
      return res.status(400).send('Invalid or unauthorized from address.');
    }
    if (!to) return res.status(400).send('Missing to recipients.');

    // Build message object per Graph sendMail
    const message = {
      subject: subject || '',
      body: { contentType: 'Text', content: body || '' }, // You can switch to 'HTML'
      toRecipients: parseEmails(to),
      ccRecipients: parseEmails(cc),
      bccRecipients: parseEmails(bcc),
      attachments: [],
    };

    // Small attachments as FileAttachment base64 (each < 3 MB recommended)
    for (const file of req.files || []) {
      message.attachments.push({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: file.originalname,
        contentType: file.mimetype || 'application/octet-stream',
        contentBytes: b64(file.buffer),
      });
    }

    const accessToken = await getGraphToken();
    if (!accessToken) throw new Error('Failed to acquire Graph token');

    // Use the shared mailbox address in the URL to send as that mailbox
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(from)}/sendMail`;

    const graphResp = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ message, saveToSentItems: true }),
    });

    if (!graphResp.ok) {
      const errText = await graphResp.text();
      return res.status(graphResp.status).send(errText);
    }

    // Optionally: persist to your own store here if required
    // e.g., create a Note record; the frontend already posts to /notes after success.

    res.json({ ok: true, message: 'sent' });
  } catch (e) {
    console.error('send error', e);
    res.status(500).send(e?.message || 'Internal error');
  }
});

// --- OPTIONAL: helper to proxy an invoice file as a temporary URL ---
// If you already expose GET /api/invoices/:id/blob-url, keep your implementation.
// Here is a placeholder showing expected shape.
app.get('/api/invoices/:id/blob-url', async (req, res) => {
  // TODO: implement with your storage (SharePoint/S3/Blob). Return { url }.
  res.json({ url: '' });
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`API listening on ${PORT}`));

/*
Notes:
- If you need HTML body support, set body: { contentType: 'HTML', content: '<p>…</p>' }
- For large attachments (> 3-4MB), use Graph upload sessions:
  1) POST /users/{from}/messages to create draft
  2) POST /users/{from}/messages/{id}/attachments/createUploadSession
  3) PUT chunks
  4) POST /users/{from}/messages/{id}/send
- App permissions path is simplest for shared mailboxes. Alternatively, use delegated auth + Exchange "Send As" permission for the user.
*/
