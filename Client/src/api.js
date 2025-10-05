// src/api.js
// In Azure (production) call the same origin: e.g. https://<yourapp>.azurewebsites.net
// In local dev, let VITE_API_URL point to your API (e.g. http://localhost:5000)
const API_BASE = import.meta.env.PROD
  ? ''                                  // makes requests like fetch('/api/invoices')
  : (import.meta.env.VITE_API_URL || 'http://localhost:5000');


export async function api(path, init) {
  const isGet = !init?.method || init.method.toUpperCase() === 'GET';
  const headers = new Headers(init?.headers || {});
  if (!isGet && !headers.has('Content-Type')) headers.set('Content-Type', 'application/json');

  const res = await fetch(`${API_BASE}${path}`, { ...init, headers, credentials: 'omit' });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}

export const getInvoices = () => api('/api/invoices');
