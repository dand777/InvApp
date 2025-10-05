// src/api.js
const API_BASE = (import.meta.env.VITE_API_URL || '').trim() || ''; // same-origin in prod

export async function api(path, init) {
  const isGet = !init?.method || init.method.toUpperCase() === 'GET';
  const headers = new Headers(init?.headers || {});
  if (!isGet && !headers.has('Content-Type')) headers.set('Content-Type', 'application/json');

  const res = await fetch(`${API_BASE}${path}`, { ...init, headers, credentials: 'omit' });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}

export const getInvoices = () => api('/api/invoices');
