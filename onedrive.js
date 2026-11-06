// OneDrive-Anbindung über Microsoft Graph.
// Anmeldung: OAuth 2.0 Authorization Code Flow mit PKCE (ohne Server, ohne Bibliothek).
// Die App-Registrierung muss die Umleitungs-URI als Plattform "Single-Page-Anwendung" enthalten.

import { isDocx, extractDocxText } from './docx.js';

const AUTH_BASE ='https://login.microsoftonline.com/common/oauth2/v2.0';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = 'openid offline_access Files.Read';
const KEY_TOKENS = 'mp.od.tokens';
const KEY_PKCE = 'mp.od.pkce';
const SCRIPT_EXT = /\.(txt|md|text|docx)$/i;

let clientId = '';

function err(code, message) {
  const e = new Error(message || code);
  e.code = code;
  return e;
}

export function configure({ clientId: id }) {
  clientId = (id || '').trim();
}

export function isConfigured() {
  return !!clientId;
}

export function redirectUri() {
  const u = new URL(location.href);
  u.search = '';
  u.hash = '';
  u.pathname = u.pathname.replace(/index\.html$/, '');
  return u.toString();
}

function loadTokens() {
  try { return JSON.parse(localStorage.getItem(KEY_TOKENS)); } catch { return null; }
}
function saveTokens(t) {
  localStorage.setItem(KEY_TOKENS, JSON.stringify(t));
}

export function isConnected() {
  const t = loadTokens();
  return !!(t && (t.refresh_token || t.access_token));
}

export function logout() {
  localStorage.removeItem(KEY_TOKENS);
}

function b64url(bytes) {
  let s = '';
  bytes.forEach((b) => { s += String.fromCharCode(b); });
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function randomString(n = 48) {
  const a = new Uint8Array(n);
  crypto.getRandomValues(a);
  return b64url(a);
}

/** Leitet zur Microsoft-Anmeldung weiter. */
export async function login() {
  if (!clientId) throw err('config', 'Client-ID fehlt');
  const verifier = randomString(48);
  const state = randomString(16);
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  const challenge = b64url(new Uint8Array(digest));
  localStorage.setItem(KEY_PKCE, JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri(),
    response_mode: 'query',
    scope: SCOPES,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    state,
  });
  location.assign(`${AUTH_BASE}/authorize?${p}`);
}

/** Wertet die Rückkehr von der Microsoft-Anmeldung aus (falls vorhanden). */
export async function handleRedirect() {
  const url = new URL(location.href);
  const code = url.searchParams.get('code');
  const error = url.searchParams.get('error');
  if (!code && !error) return null;
  const state = url.searchParams.get('state');
  history.replaceState(null, '', redirectUri());

  if (error) {
    return { error: url.searchParams.get('error_description') || error };
  }
  let pk = null;
  try { pk = JSON.parse(localStorage.getItem(KEY_PKCE)); } catch { /* leer */ }
  localStorage.removeItem(KEY_PKCE);
  if (!pk || pk.state !== state) {
    return { error: 'Die Anmeldung konnte nicht zugeordnet werden. Bitte noch einmal verbinden.' };
  }
  await tokenRequest({
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri(),
    code_verifier: pk.verifier,
  });
  return { ok: true };
}

async function tokenRequest(params) {
  const body = new URLSearchParams({ client_id: clientId, scope: SCOPES, ...params });
  let res;
  try {
    res = await fetch(`${AUTH_BASE}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body,
    });
  } catch {
    throw err('network', 'Keine Verbindung');
  }
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw err('login', data.error_description || data.error || 'Anmeldung fehlgeschlagen');
  const old = loadTokens() || {};
  const tokens = {
    access_token: data.access_token,
    refresh_token: data.refresh_token || old.refresh_token,
    expires_at: Date.now() + (data.expires_in || 3600) * 1000,
  };
  saveTokens(tokens);
  return tokens;
}

async function getAccessToken(forceRefresh = false) {
  const t = loadTokens();
  if (!t) throw err('login');
  if (!forceRefresh && t.access_token && t.expires_at - 60000 > Date.now()) return t.access_token;
  if (!t.refresh_token) { logout(); throw err('login'); }
  try {
    return (await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh_token })).access_token;
  } catch (e) {
    if (e.code === 'login') logout();
    throw e;
  }
}

async function graph(url, { raw = false } = {}) {
  for (let attempt = 0; attempt < 2; attempt++) {
    const token = await getAccessToken(attempt > 0);
    let res;
    try {
      res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    } catch {
      throw err('network', 'Keine Verbindung');
    }
    if (res.status === 401 && attempt === 0) continue;
    if (res.status === 404) throw err('notfound');
    if (!res.ok) throw err('http', `Fehler ${res.status}`);
    return raw ? res : res.json();
  }
  throw err('login');
}

/** Liefert alle Textdateien im angegebenen OneDrive-Ordner. */
export async function listScripts(folder) {
  const clean = (folder || '')
    .split('/')
    .map((s) => s.trim())
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
  let url = clean
    ? `${GRAPH}/me/drive/root:/${clean}:/children?$top=200`
    : `${GRAPH}/me/drive/root/children?$top=200`;
  const items = [];
  while (url) {
    const data = await graph(url);
    items.push(...(data.value || []));
    url = data['@odata.nextLink'];
  }
  return items
    .filter((i) => i.file && SCRIPT_EXT.test(i.name))
    .map((i) => ({
      id: i.id,
      name: i.name,
      title: i.name.replace(SCRIPT_EXT, ''),
      modified: i.lastModifiedDateTime,
      size: i.size,
      downloadUrl: i['@microsoft.graph.downloadUrl'] || '',
    }));
}

/** Text erkennen: UTF-8, UTF-16 (Windows-Editor "Unicode") oder Windows-1252. */
export function decodeText(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/** Lädt den Inhalt einer Skriptdatei. */
export async function fetchContent(item) {
  const toText = (buf) => (isDocx(item.name) ? extractDocxText(buf) : decodeText(buf));
  if (item.downloadUrl) {
    let buf = null;
    try {
      const res = await fetch(item.downloadUrl);
      if (res.ok) buf = await res.arrayBuffer();
    } catch { /* Link abgelaufen oder offline – zweiter Weg unten */ }
    if (buf) return toText(buf);
  }
  const res = await graph(`${GRAPH}/me/drive/items/${encodeURIComponent(item.id)}/content`, { raw: true });
  return toText(await res.arrayBuffer());
}
