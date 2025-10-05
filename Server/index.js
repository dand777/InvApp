// Server/index.js
// Load .env only when NOT running on Azure App Service
if (!process.env.WEBSITE_INSTANCE_ID) {
  try {
    await import('dotenv/config');
  } catch {
    // dotenv is optional in prod; ignore if not installed
  }
}

import express from 'express'
import cors from 'cors'
import multer from 'multer'
import fetch from 'node-fetch'
import path from 'path'
import fs from 'fs'
import { fileURLToPath } from 'url'
import { pool } from './db.js'
import { ClientSecretCredential } from '@azure/identity'
import {
  StorageSharedKeyCredential,
  BlobSASPermissions,
  generateBlobSASQueryParameters
} from '@azure/storage-blob'

// ESM __dirname
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const app = express()
const upload = multer({ limits: { fileSize: 8 * 1024 * 1024 } }) // 8 MB per file

app.use(express.json())

// ---- CORS (allow localhost + this Azure site automatically) ----
const azureOrigin = process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : null
const allowedOrigins = ['http://localhost:5173', 'http://localhost:3000', ...(azureOrigin ? [azureOrigin] : [])]
app.use(cors({ origin: allowedOrigins, credentials: false }))

// ---- Optional health check ----
app.get('/healthz', (_req, res) => res.type('text/plain').send('ok'))

// ------------------------- React Frontend (static) -------------------------
// (We’ll mount the catch-all **after** API routes so it won’t swallow /api/*)
const spaCandidates = [
  path.join(__dirname, '../client/dist'),
  path.join(__dirname, '../dist'),
  path.join(__dirname, 'dist'),
]
const clientDir = spaCandidates.find(p => fs.existsSync(path.join(p, 'index.html')))
if (clientDir) {
  app.use(express.static(clientDir))
  console.log('Serving SPA from:', clientDir)
} else {
  console.log('No SPA build found. Running API-only.')
}

// ---------------------------------------------------------------------------

// Tag outbound subjects with a stable token we can find in replies
function withRefTag(subject, invoiceId) {
  if (!invoiceId) return subject || ''
  const tag = `[#INV:${invoiceId}]`
  return (subject || '').includes(tag) ? subject : `${subject || ''} ${tag}`.trim()
}

// 🔹 Azure Blob SAS helpers
const AZ_ACCOUNT   = process.env.AZURE_STORAGE_ACCOUNT
const AZ_KEY       = process.env.AZURE_STORAGE_ACCOUNT_KEY
const AZ_CONTAINER = process.env.AZURE_STORAGE_CONTAINER
const AZ_BASE_DIR  = (process.env.AZURE_BLOB_BASE_DIR || '').replace(/^\/+|\/+$/g, '')

const hasAzureCreds = Boolean(AZ_ACCOUNT && AZ_KEY && AZ_CONTAINER)
const sharedKey = hasAzureCreds ? new StorageSharedKeyCredential(AZ_ACCOUNT, AZ_KEY) : null

// Encode each path segment safely
const encodePath = (p) =>
  (p || '')
    .split('/')
    .filter(Boolean)
    .map(s => encodeURIComponent(s))
    .join('/')

// Turn stored relative path into "<base_dir>/<relative>" if base dir is set
function normalizeBlobPath(storedPath) {
  const rel = String(storedPath || '').replace(/^\/+/, '')
  if (!AZ_BASE_DIR) return rel
  if (rel.toLowerCase().startsWith(AZ_BASE_DIR.toLowerCase() + '/')) return rel
  return `${AZ_BASE_DIR}/${rel}`
}

// ------------------------- Microsoft Graph -------------------------
const credential = new ClientSecretCredential(
  process.env.GRAPH_TENANT_ID,
  process.env.GRAPH_CLIENT_ID,
  process.env.GRAPH_CLIENT_SECRET
)
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default'

async function getGraphToken() {
  const token = await credential.getToken(GRAPH_SCOPE)
  return token?.token
}

function parseEmails(str = '') {
  return str
    .split(/[;,]/)
    .map(s => s.trim())
    .filter(Boolean)
    .map(address => ({ emailAddress: { address } }))
}

function allowedFromAddress(addr) {
  const allowed = (process.env.SHARED_MAILBOXES || '')
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean)
  if (!allowed.length) return true
  return allowed.includes(String(addr || '').toLowerCase())
}

// Helper: shape invoice row
const mapInvoiceRow = r => ({
  id: r.id,
  supplier: r.supplier,
  hub: r.hub,
  type: r.type,
  invoiceno: r.invoiceno,
  invoice_date: r.invoice_date,
  po: r.po,
  folder: r.folder,
  assigned: r.assigned,
  ref: r.ref,
  last_modified: r.last_modified,
  created_on: r.created_on,
  status: r.status,
  notes: r.notes || []
})

// ------------------------- API Routes -------------------------

// GET /api/invoices
app.get('/api/invoices', async (_req, res) => {
  try {
    const { rows } = await pool.query(`
      SELECT i.*,
             COALESCE(
               json_agg(
                 json_build_object('id', n.id, 'text', n.text, 'date', to_char(n.date, 'YYYY-MM-DD'))
               ) FILTER (WHERE n.id IS NOT NULL), '[]'
             ) AS notes
      FROM invoice i
      LEFT JOIN note n ON n.invoice_id = i.id
      GROUP BY i.id
      ORDER BY i.created_on DESC;
    `)
    res.json(rows.map(mapInvoiceRow))
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Failed to fetch invoices' })
  }
})

// POST /api/invoices/:id/notes
app.post('/api/invoices/:id/notes', async (req, res) => {
  const { id } = req.params
  const { text, date } = req.body
  if (!text) return res.status(400).json({ error: 'text is required' })
  try {
    const { rows } = await pool.query(
      `INSERT INTO note (invoice_id, text, date)
       VALUES ($1, $2, COALESCE($3, now()::date))
       RETURNING id, text, to_char(date, 'YYYY-MM-DD') AS date`,
      [id, text, date || null]
    )
    res.status(201).json(rows[0])
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Failed to add note' })
  }
})

// PUT /api/invoices/:id/notes/:noteId
app.put('/api/invoices/:id/notes/:noteId', async (req, res) => {
  const { noteId } = req.params
  const { text } = req.body
  if (!text) return res.status(400).json({ error: 'text is required' })
  try {
    const { rows } = await pool.query(
      `UPDATE note SET text = $1
       WHERE id = $2
       RETURNING id, text, to_char(date, 'YYYY-MM-DD') AS date`,
      [text, noteId]
    )
    if (!rows[0]) return res.status(404).json({ error: 'Note not found' })
    res.json(rows[0])
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Failed to update note' })
  }
})

// DELETE /api/invoices/:id/notes/:noteId
app.delete('/api/invoices/:id/notes/:noteId', async (req, res) => {
  const { noteId } = req.params
  try {
    const { rowCount } = await pool.query(`DELETE FROM note WHERE id = $1`, [noteId])
    if (!rowCount) return res.status(404).json({ error: 'Note not found' })
    res.status(204).send()
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Failed to delete note' })
  }
})

// DELETE /api/invoices/:id
app.delete('/api/invoices/:id', async (req, res) => {
  const { id } = req.params
  try {
    const result = await pool.query('DELETE FROM invoice WHERE id = $1 RETURNING id', [id])
    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Invoice not found' })
    }
    return res.status(204).send()
  } catch (e) {
    console.error(e)
    return res.status(500).json({ error: 'Failed to delete invoice' })
  }
})

// GET /api/invoices/:id/blob-url
app.get('/api/invoices/:id/blob-url', async (req, res) => {
  const { id } = req.params
  try {
    const { rows } = await pool.query(`SELECT bloburl FROM invoice WHERE id = $1`, [id])
    const row = rows[0]
    if (!row) return res.status(404).json({ error: 'Invoice not found' })
    if (!row.bloburl) return res.status(400).json({ error: 'bloburl is empty for this invoice' })

    const blobPath = normalizeBlobPath(row.bloburl)

    if (!hasAzureCreds) {
      const bestEffortUrl =
        `https://${AZ_ACCOUNT}.blob.core.windows.net/` +
        `${encodeURIComponent(AZ_CONTAINER)}/` +
        `${encodePath(blobPath)}`
      res.set('Cache-Control', 'no-store')
      return res.json({ url: bestEffortUrl })
    }

    const expiresOn = new Date(Date.now() + 15 * 60 * 1000)
    const startsOn  = new Date(Date.now() - 60 * 1000)

    const sas = generateBlobSASQueryParameters(
      {
        containerName: AZ_CONTAINER,
        blobName: blobPath,
        permissions: BlobSASPermissions.parse('r'),
        startsOn,
        expiresOn
      },
      sharedKey
    ).toString()

    const url =
      `https://${AZ_ACCOUNT}.blob.core.windows.net/` +
      `${encodeURIComponent(AZ_CONTAINER)}/` +
      `${encodePath(blobPath)}?${sas}`

    res.set('Cache-Control', 'no-store')
    return res.json({ url })
  } catch (e) {
    console.error('blob-url error:', e)
    return res.status(500).json({ error: 'Failed to build access URL' })
  }
})

// POST /api/email/send
app.post('/api/email/send', upload.array('attachments'), async (req, res) => {
  try {
    const { from = '', to = '', cc = '', bcc = '', subject = '', body = '', invoiceId = '' } = req.body

    if (!from || !allowedFromAddress(from)) return res.status(400).send('Invalid from address.')
    if (!to) return res.status(400).send('Missing to recipients.')

    const taggedSubject = withRefTag(subject, invoiceId)

    const message = {
      subject: taggedSubject,
      body: { contentType: 'Text', content: body || '' },
      toRecipients: parseEmails(to),
      ccRecipients: parseEmails(cc),
      bccRecipients: parseEmails(bcc),
      attachments: [],
    }

    for (const file of req.files || []) {
      message.attachments.push({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: file.originalname,
        contentType: file.mimetype || 'application/octet-stream',
        contentBytes: Buffer.from(file.buffer).toString('base64'),
      })
    }

    const token = await getGraphToken()
    if (!token) throw new Error('Failed to acquire Graph token')

    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(from)}/sendMail`
    const resp = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, saveToSentItems: true }),
    })

    if (!resp.ok) {
      const text = await resp.text()
      return res.status(resp.status).send(text)
    }

    res.json({ ok: true, sent: true, invoiceId })
  } catch (e) {
    console.error('send error', e)
    res.status(500).send(e?.message || 'Internal error')
  }
})

// PATCH /api/invoices/:id
app.patch('/api/invoices/:id', async (req, res) => {
  const { id } = req.params
  const allowed = ['status', 'folder', 'assigned', 'ref']
  const patch = {}
  for (const k of allowed) if (k in req.body) patch[k] = req.body[k]

  if (Object.keys(patch).length === 0) return res.status(400).json({ error: 'No updatable fields supplied' })
  if (patch.status && !['New', 'Matched', 'Posting', 'Completed'].includes(patch.status)) {
    return res.status(400).json({ error: 'Invalid status' })
  }

  const sets = []
  const vals = []
  let idx = 1
  for (const [k, v] of Object.entries(patch)) {
    sets.push(`${k} = $${idx++}`)
    vals.push(v)
  }
  sets.push(`last_modified = now()`)

  try {
    const { rows } = await pool.query(
      `UPDATE invoice SET ${sets.join(', ')} WHERE id = $${idx} RETURNING *`,
      [...vals, id]
    )
    if (!rows[0]) return res.status(404).json({ error: 'Invoice not found' })
    res.json(mapInvoiceRow({ ...rows[0], notes: [] }))
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Failed to update invoice' })
  }
})

// ------------------------- Reply ingestion poller -------------------------
const REPLY_MAILBOX =
  (process.env.REPLY_MAILBOX ||
    (process.env.SHARED_MAILBOXES || '').split(',')[0] ||
    ''
  ).trim()

function extractInvoiceIdFromSubject(subject = '') {
  const m = subject.match(/\[#INV:(\d+)\]/i)
  return m ? m[1] : null
}
function stripHtml(html = '') {
  return html
    .replace(/<style[^>]*>.*?<\/style>/gis, '')
    .replace(/<script[^>]*>.*?<\/script>/gis, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/\u00A0/g, ' ')
    .trim()
}

let REPLY_SCHEMA_READY = false
async function ensureReplySchema() {
  if (REPLY_SCHEMA_READY) return
  await pool.query(`CREATE TABLE IF NOT EXISTS mailbox_cursor (mailbox TEXT PRIMARY KEY, delta_link TEXT)`)
  await pool.query(`ALTER TABLE note ADD COLUMN IF NOT EXISTS message_id TEXT`)
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS note_message_id_uidx ON note (message_id)`)
  REPLY_SCHEMA_READY = true
}

async function loadDeltaLink(mailbox) {
  const { rows } = await pool.query(`SELECT delta_link FROM mailbox_cursor WHERE mailbox=$1`, [mailbox])
  return rows[0]?.delta_link || null
}
async function saveDeltaLink(mailbox, link) {
  await pool.query(
    `INSERT INTO mailbox_cursor (mailbox, delta_link)
     VALUES ($1,$2)
     ON CONFLICT (mailbox) DO UPDATE SET delta_link=EXCLUDED.delta_link`,
    [mailbox, link]
  )
}

async function pollRepliesOnce() {
  if (!REPLY_MAILBOX) return
  try {
    await ensureReplySchema()

    const token = await getGraphToken()
    if (!token) {
      console.error('pollRepliesOnce error: no Graph token')
      return
    }
    const headers = { Authorization: `Bearer ${token}` }

    // Build safe default URL
    const defaultUrl = new URL(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(REPLY_MAILBOX)}/mailFolders('Inbox')/messages/delta`
    )
    defaultUrl.searchParams.set('$select', 'subject,from,receivedDateTime,body,internetMessageId')

    // Load cursor and coerce to URL
    let cursor = await loadDeltaLink(REPLY_MAILBOX)
    if (cursor && typeof cursor !== 'string') cursor = String(cursor)

    let next
    try {
      next = new URL(cursor || defaultUrl.toString())
    } catch {
      next = new URL(defaultUrl.toString())
    }

    while (next) {
      const resp = await fetch(next.toString(), { headers })
      if (!resp.ok) {
        const text = await resp.text()
        throw new Error(`Graph delta failed ${resp.status}: ${text}`)
      }
      const json = await resp.json()

      for (const m of (json.value || [])) {
        const invoiceId = extractInvoiceIdFromSubject(m.subject || '')
        if (!invoiceId) continue

        const raw = m?.body?.content || ''
        const contentType = (m?.body?.contentType || '').toLowerCase()
        const bodyText = contentType === 'html' ? stripHtml(raw) : raw

        try {
          await pool.query(
            `INSERT INTO note (invoice_id, text, date, message_id)
             VALUES ($1, $2, now()::date, $3)
             ON CONFLICT (message_id) DO NOTHING`,
            [invoiceId, bodyText.slice(0, 8000), m.internetMessageId || null]
          )
        } catch (e) {
          console.error('note insert failed', e)
        }
      }

      if (json['@odata.nextLink']) {
        try {
          next = new URL(String(json['@odata.nextLink']))
        } catch {
          next = null
        }
      } else {
        if (json['@odata.deltaLink']) {
          await saveDeltaLink(REPLY_MAILBOX, String(json['@odata.deltaLink']))
        }
        next = null
      }
    }
  } catch (e) {
    console.error('pollRepliesOnce error:', e && e.message ? e.message : e)
  }
}

// Enable poller only if config is complete (define ONCE)
const ENABLE_REPLY_POLLER =
  !!REPLY_MAILBOX &&
  !!process.env.GRAPH_TENANT_ID &&
  !!process.env.GRAPH_CLIENT_ID &&
  !!process.env.GRAPH_CLIENT_SECRET

if (ENABLE_REPLY_POLLER) {
  setInterval(pollRepliesOnce, 60_000)
  pollRepliesOnce()
} else {
  console.log('Reply poller disabled: missing Graph config or REPLY_MAILBOX')
}

// ------------------------- Catch-all for React (Express 5-safe) -------------------------
if (clientDir) {
  // after API routes so it doesn't intercept /api/*
  app.get(/^(?!\/api\/).*/, (_req, res) => {
    res.sendFile(path.join(clientDir, 'index.html'))
  })
}

// ------------------------- Start Server -------------------------
const PORT = process.env.PORT || 5000
app.listen(PORT, '0.0.0.0', () => {
  console.log(`API listening on http://localhost:${PORT}`)
})
