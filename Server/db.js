import pg from 'pg'
const { Pool } = pg

const connectionString =
  process.env.DATABASE_URL_POOL || process.env.DATABASE_URL

export const pool = new Pool({
  connectionString,
  ssl: { rejectUnauthorized: false } // Neon uses TLS
})
