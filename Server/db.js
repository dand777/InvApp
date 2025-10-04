// Server/db.js
import 'dotenv/config';
import pg from 'pg';
const { Pool } = pg;

const connectionString =
  process.env.DATABASE_URL_POOL || process.env.DATABASE_URL;

export const pool = new Pool({
  connectionString,
  // Azure PG requires TLS. For quick dev you can skip cert verification:
  ssl: { rejectUnauthorized: false },
  // In production, prefer a real CA bundle instead of disabling verification (see note below).
});
