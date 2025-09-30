import 'dotenv/config'
import express from 'express'
import cors from 'cors'
import { pool } from './db.js'

const app = express()
app.use(express.json())

// CORS: allow local React dev + (later) your deployed frontend
app.use(cors({
  origin: [
    'http://localhost:5173', // Vite default
    'http://localhost:3000'  // CRA default (just in case)
  ],
  credentials: false
}))

// Helper: shape each invoice row and attach notes array
const mapInvoiceRow = r => ({
  id: r.id,
  supplier: r.supplier,
  hub: r.hub,
  type: r.type,
  invoice_date: r.invoice_date,     // your UI formats to UK date
  po: r.po,
  folder: r.folder,
  assigned: r.assigned,
  ref: r.ref,
  last_modified: r.last_modified,
  created_on: r.created_on,
  status: r.status,
  notes: r.notes || []              // array of {id, text, date}
})

// GET /api/invoices  -> returns invoices with notes[] (to match your UI)
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
  const { text, date } = req.body   // date optional; default today in DB
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
    const { rowCount } = await pool.query(
      `DELETE FROM note WHERE id = $1`,
      [noteId]
    )
    if (!rowCount) return res.status(404).json({ error: 'Note not found' })
    res.status(204).send()
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'Failed to delete note' })
  }
})

const PORT = process.env.PORT || 5000
app.listen(PORT, '0.0.0.0', () => {
  console.log(`API listening on http://localhost:${PORT}`)
})
