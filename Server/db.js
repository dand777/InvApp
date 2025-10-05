// Server/db.js
import pg from 'pg'
const { Pool } = pg

// Prefer a single DATABASE_URL. Fall back to PG* parts if needed.
const CONNECTION_URL =
  (process.env.DATABASE_URL && process.env.DATABASE_URL.trim()) ||
  (process.env.DATABASE_URL_POOL && process.env.DATABASE_URL_POOL.trim()) ||
  null

if (
  !CONNECTION_URL &&
  !(process.env.PGHOST && process.env.PGUSER && process.env.PGPASSWORD && process.env.PGDATABASE)
) {
  // Fail fast with a clear message in App Service logs
  throw new Error(
    '[db] No DB settings found. Set DATABASE_URL (recommended) or PGHOST/PGUSER/PGPASSWORD/PGDATABASE in Azure → Your Web App → Configuration → Application settings.'
  )
}

export const pool = new Pool(
  CONNECTION_URL
    ? {
        connectionString: CONNECTION_URL,
        // Azure PG requires TLS; using a permissive CA check in App Service.
        ssl: { rejectUnauthorized: false },
      }
    : {
        host: process.env.PGHOST,
        port: +(process.env.PGPORT || 5432),
        user:
          process.env.PGUSER, // Single Server: user@servername, Flexible: user
        password: process.env.PGPASSWORD,
        database: process.env.PGDATABASE,
        ssl: { rejectUnauthorized: false },
      }
)

// Optional: log a one-time connection status at boot
pool
  .query('select current_database() db, current_user as user, now() as now')
  .then((r) => console.log('[db] Connected:', r.rows[0]))
  .catch((e) => console.error('[db] Connection FAILED:', e.message))
