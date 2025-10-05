<!-- .github/copilot-instructions.md
     Purpose: give an AI-coding agent the exact, discoverable knowledge
     it needs to make useful changes in this repo (InvApp).
-->

# Quick orientation for AI agents

This repo is a small two-part web app: a Vite + React client under `Client/` and
an Express API under `Server/`. The server persists to PostgreSQL and integrates
with Microsoft Graph + Azure Blob Storage for email and attachments.

Key things an agent should know (concise):

- Architecture

  - Client: `Client/` — Vite + React (MUI) app. Entry: `Client/src/main.jsx`, API
    helper `Client/src/api.js`. Dev server runs on port 5173 by default.
  - Server: `Server/` — Express API implemented in `Server/index.js`. Database
    connection in `Server/db.js` (uses `DATABASE_URL` or PG\_\* env vars). Default
    API port is 5000.
  - Production hosting: build the client (`Client: npm run build`) and the
    server will serve `../client/dist` (or `../dist` / `./dist`) if `index.html`
    is present.

- How the client talks to the API

  - `Client/src/api.js` chooses the base URL:
    - In PROD: use same-origin (''), so requests are like `/api/invoices`.
    - In DEV: uses `VITE_API_URL` or falls back to `http://localhost:5000`.
  - So for local dev: run the server, then run the client with `VITE_API_URL`
    pointing at the server.

- Important server integrations and behaviors

  - Postgres pool in `Server/db.js` — the server will fail-fast if no DB
    configuration is available. Use `DATABASE_URL` (recommended) or set
    `PGHOST/PGUSER/PGPASSWORD/PGDATABASE`.
  - Azure Blob Storage: SAS URLs are generated in `Server/index.js` when
    `AZURE_STORAGE_ACCOUNT` + `AZURE_STORAGE_ACCOUNT_KEY` +
    `AZURE_STORAGE_CONTAINER` are present.
  - Microsoft Graph email sending and reply ingestion require
    `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET` and
    optionally `REPLY_MAILBOX` / `SHARED_MAILBOXES`.
  - The reply poller (reads mailbox delta links) is enabled only when Graph
    config + `REPLY_MAILBOX` are set.

- Run & build commands (local)

  - Client: (from `Client/`)
    ```powershell
    npm install
    npm run dev       # start Vite dev server (port 5173)
    npm run build     # produce production files in Client/dist
    npm run preview   # preview the built site
    ```
  - Server: (from `Server/`)
    ```powershell
    npm install
    npm start         # runs node index.js (API defaults to port 5000)
    ```
  - Typical local dev flow: start the Server, then in Client start Vite.
    Optionally set `VITE_API_URL=http://localhost:5000` in `Client/.env`.

- Key API surface (implemented in `Server/index.js`)

  - GET /api/invoices — list invoices (includes notes)
  - PATCH /api/invoices/:id — update {status,folder,assigned,ref}
  - POST /api/invoices/:id/notes — add a note
  - PUT /api/invoices/:id/notes/:noteId — edit a note
  - DELETE /api/invoices/:id — delete invoice
  - GET /api/invoices/:id/blob-url — get a time-limited blob URL (SAS)
  - POST /api/email/send — send email via Microsoft Graph

- Frontend patterns to follow (concrete examples)

  - Polling & optimistic updates: `Client/src/components/InvoiceDashboard.jsx`
    polls `/api/invoices` every 5s and uses an optimistic `applyPatch()` which
    updates UI immediately then PATCHes the server.
  - API helper: prefer `Client/src/api.js` wrappers (they set JSON headers and
    centralize the base URL); mimic its error handling when adding new client
    calls.
  - Shared mailbox list: `Client/src/components/InvoiceDashboard.jsx` reads
    `import.meta.env.VITE_SHARED_MAILBOXES` and splits on commas.

- Environment variables to check or set (local `.env` in `Server/`/`Client/`)

  - Server: DATABASE_URL (preferred) OR PGHOST/PGUSER/PGPASSWORD/PGDATABASE,
    PORT (optional), AZURE_STORAGE_ACCOUNT, AZURE_STORAGE_ACCOUNT_KEY,
    AZURE_STORAGE_CONTAINER, AZURE_BLOB_BASE_DIR, GRAPH_TENANT_ID,
    GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, REPLY_MAILBOX, SHARED_MAILBOXES
  - Client: VITE_API_URL, VITE_SHARED_MAILBOXES
  - Note: `Server/index.js` only loads `dotenv` if not running inside Azure
    App Service (it checks `WEBSITE_INSTANCE_ID`).

- Debugging tips

  - Server logs connection status at boot (see `Server/db.js`); check server
    console for "API listening on http://localhost:5000" or database errors.
  - Client logs `API_BASE` in dev (see `Client/src/components/InvoiceDashboard.jsx`)
    — useful to confirm the Vite -> API wiring.
  - If blob SAS URLs are empty, verify `AZURE_*` env vars and `AZ_CONTAINER`.

- Conventions & gotchas
  - The client expects certain invoice fields (see `mapInvoiceRow` in
    `Server/index.js`) — when changing the API shape, keep keys: `id,
supplier,hub,type,invoiceno,invoice_date,po,folder,assigned,ref,last_modified,
created_on,status,notes`.
  - `PATCH /api/invoices/:id` only accepts `['status','folder','assigned','ref']`.
  - The server will create small schema helpers at runtime (e.g. `mailbox_cursor`
    table) — but the main tables (`invoice`, `note`) are assumed to exist.

If anything here is unclear or you want the agent to include extra examples
(e.g. sample `.env` files or a minimal end-to-end local run script), tell me
which part to expand and I will iterate.
