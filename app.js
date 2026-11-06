import * as OD from './onedrive.js';
import { SpeechFollower, speechSupported, normalizeWord, createRecognizer } from './speech.js';

const APP_VERSION = '1.0.0';
const cfg = window.MARY_CONFIG || {};
const $ = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

const FONT_SIZES = [28, 34, 40, 48, 56, 66, 78, 90];
const LINE_HEIGHT = 1.3;
const LIMITS = { fontStep: [0, FONT_SIZES.length - 1], lines: [2, 8], speed: [1, 10] };

// ---------------------------------------------------------------------------
// Speicher
// ---------------------------------------------------------------------------
const store = {
  get(key, fallback) {
    try {
      const v = localStorage.getItem('mp.' + key);
      return v === null ? fallback : JSON.parse(v);
    } catch { return fallback; }
  },
  set(key, val) {
    try { localStorage.setItem('mp.' + key, JSON.stringify(val)); } catch { /* voll */ }
  },
};

const DEFAULTS = {
  fontStep: 4,
  lines: 4,
  speed: 4,
  countdown: 10,
  mode: 'auto',
  promptTop: null,
  mirror: true,
  folder: cfg.folder || 'Teleprompter',
  clientId: cfg.clientId || '',
};
const settings = Object.assign({}, DEFAULTS, store.get('settings', {}));
if (!settings.clientId && cfg.clientId) settings.clientId = cfg.clientId;
const saveSettings = () => store.set('settings', settings);

// ---------------------------------------------------------------------------
// Hilfen
// ---------------------------------------------------------------------------
let toastTimer;
function toast(msg, ms = 3500) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), ms);
}
function hideToast() {
  clearTimeout(toastTimer);
  $('#toast').classList.remove('show');
}

const fmtDate = (iso) => {
  try {
    return new Date(iso).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  } catch { return ''; }
};

/** Markdown-Zeichen entfernen, Zeilenumbrüche vereinheitlichen. */
function cleanScript(text) {
  return (text || '')
    .replace(/\r\n?/g, '\n')
    .replace(/^\s{0,3}#{1,6}\s+/gm, '')
    .replace(/^\s*[-*•]\s+/gm, '')
    .replace(/(\*\*|__|\*|`)/g, '')
    .trim();
}

// ---------------------------------------------------------------------------
// Bildschirme
// ---------------------------------------------------------------------------
let currentScreen = 'library';
let previousScreen = 'library';
function showScreen(name) {
  if (name !== currentScreen) previousScreen = currentScreen;
  $$('.screen').forEach((s) => s.classList.toggle('active', s.id === 'screen-' + name));
  currentScreen = name;
  window.scrollTo(0, 0);
}
$$('[data-back]').forEach((b) => b.addEventListener('click', () => {
  showScreen(previousScreen === currentScreen ? 'library' : previousScreen);
  if (currentScreen === 'library') renderLibrary();
}));

// ---------------------------------------------------------------------------
// Skriptliste
// ---------------------------------------------------------------------------
let localScripts = store.get('localScripts', []);
let cloudScripts = store.get('cloudScripts', []);
let cloudTexts = store.get('cloudTexts', {});

function pruneCloudTexts() {
  const ids = new Set(cloudScripts.map((s) => s.id));
  for (const id of Object.keys(cloudTexts)) if (!ids.has(id)) delete cloudTexts[id];
  store.set('cloudTexts', cloudTexts);
}

function renderLibrary() {
  const list = $('#script-list');
  list.textContent = '';
  const items = [
    ...cloudScripts.map((s) => ({ ...s, source: 'cloud' })),
    ...localScripts.map((s) => ({ ...s, source: 'local' })),
  ].sort((a, b) => (b.modified || '').localeCompare(a.modified || ''));
  $('#empty-hint').hidden = items.length > 0;

  for (const it of items) {
    const li = document.createElement('li');
    const btn = document.createElement('button');
    btn.className = 'script-item';
    const title = document.createElement('span');
    title.className = 'script-title';
    title.textContent = it.title || 'Ohne Titel';
    const meta = document.createElement('span');
    meta.className = 'script-meta';
    meta.textContent = `${it.source === 'cloud' ? 'OneDrive' : 'Eingefügt'} · ${fmtDate(it.modified)}`;
    btn.append(title, meta);
    btn.addEventListener('click', () => openScript(it));
    li.append(btn);

    if (it.source === 'local') {
      const del = document.createElement('button');
      del.className = 'btn-delete';
      del.textContent = 'Löschen';
      del.setAttribute('aria-label', `„${it.title}“ löschen`);
      del.addEventListener('click', () => {
        if (!confirm(`„${it.title}“ wirklich löschen?`)) return;
        localScripts = localScripts.filter((s) => s.id !== it.id);
        store.set('localScripts', localScripts);
        renderLibrary();
      });
      li.append(del);
    }
    list.append(li);
  }
}

async function openScript(it) {
  let text = it.text;
  if (it.source === 'cloud') {
    const cached = cloudTexts[it.id];
    if (cached && cached.modified === it.modified) {
      text = cached.text;
    } else {
      toast('Skript wird geladen …', 15000);
      try {
        text = await OD.fetchContent(it);
        cloudTexts[it.id] = { modified: it.modified, text };
        store.set('cloudTexts', cloudTexts);
        hideToast();
      } catch {
        if (cached) {
          text = cached.text;
          toast('Keine Verbindung. Die zuletzt geladene Fassung wird verwendet.');
        } else {
          toast('Das Skript konnte nicht geladen werden. Bitte die Internetverbindung prüfen.', 5000);
          return;
        }
      }
    }
  }
  if (!cleanScript(text)) {
    toast('Dieses Skript ist leer.');
    return;
  }
  enterRecord(text);
}

// ---------------------------------------------------------------------------
// OneDrive
// ---------------------------------------------------------------------------
function renderOD(msg) {
  const connected = OD.isConnected();
  $('#btn-od-connect').hidden = connected;
  $('#btn-od-refresh').hidden = !connected;
  $('#od-status').textContent = msg || (connected
    ? 'OneDrive ist verbunden.'
    : settings.clientId
      ? 'OneDrive ist nicht verbunden.'
      : 'OneDrive ist noch nicht eingerichtet. Die Client-ID fehlt (siehe Einstellungen).');
}

let syncing = false;
async function syncOneDrive() {
  if (!OD.isConnected()) { renderOD(); return; }
  if (syncing) return;
  syncing = true;
  renderOD('OneDrive wird abgerufen …');
  try {
    const files = await OD.listScripts(settings.folder);
    cloudScripts = files;
    store.set('cloudScripts', files);
    pruneCloudTexts();
    renderLibrary();
    const time = new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
    renderOD(`OneDrive verbunden · Ordner „${settings.folder}“ · ${files.length} ${files.length === 1 ? 'Skript' : 'Skripte'} · Stand ${time} Uhr`);
  } catch (e) {
    renderOD(
      e.code === 'login'
        ? 'Die OneDrive-Anmeldung ist abgelaufen. Bitte „Mit OneDrive verbinden“ tippen.'
        : e.code === 'notfound'
          ? `Der Ordner „${settings.folder}“ wurde in OneDrive nicht gefunden. Name in den Einstellungen prüfen.`
          : 'OneDrive ist gerade nicht erreichbar. Die zuletzt geladenen Skripte werden angezeigt.',
    );
  } finally {
    syncing = false;
  }
}

$('#btn-od-connect').addEventListener('click', async () => {
  if (!settings.clientId) {
    toast('Bitte zuerst die Client-ID in den Einstellungen eintragen.', 5000);
    openSettings();
    return;
  }
  if (!window.isSecureContext) {
    toast('Die Anmeldung funktioniert nur, wenn die App über https geöffnet wird.', 6000);
    return;
  }
  OD.configure({ clientId: settings.clientId });
  renderOD('Weiterleitung zu Microsoft …');
  try { await OD.login(); } catch { renderOD('Die Anmeldung konnte nicht gestartet werden.'); }
});
$('#btn-od-refresh').addEventListener('click', syncOneDrive);

// ---------------------------------------------------------------------------
// Text einfügen
// ---------------------------------------------------------------------------
$('#btn-open-paste').addEventListener('click', () => {
  $('#paste-title').value = '';
  $('#paste-text').value = '';
  showScreen('paste');
});
$('#btn-clipboard').addEventListener('click', async () => {
  try {
    const t = await navigator.clipboard.readText();
    if (t) $('#paste-text').value = t;
    else toast('Die Zwischenablage ist leer.');
  } catch {
    toast('Kein Zugriff auf die Zwischenablage. Bitte ins Textfeld tippen und „Einsetzen“ wählen.', 5000);
  }
});
function pastedScript() {
  const text = $('#paste-text').value;
  if (!cleanScript(text)) {
    toast('Bitte zuerst einen Text einfügen.');
    return null;
  }
  const title = $('#paste-title').value.trim()
    || cleanScript(text).split(/\s+/).slice(0, 5).join(' ').slice(0, 40);
  return { title, text };
}
$('#btn-paste-save').addEventListener('click', () => {
  const s = pastedScript();
  if (!s) return;
  localScripts.push({ id: 'local-' + Date.now(), title: s.title, text: s.text, modified: new Date().toISOString() });
  store.set('localScripts', localScripts);
  renderLibrary();
  showScreen('library');
  toast('Skript gespeichert.');
});
$('#btn-paste-use').addEventListener('click', () => {
  const s = pastedScript();
  if (s) enterRecord(s.text);
});

// ---------------------------------------------------------------------------
// Einstellungen
// ---------------------------------------------------------------------------
function openSettings() {
  renderSettings();
  showScreen('settings');
}
$('#btn-open-settings').addEventListener('click', openSettings);

function renderSettings() {
  $('#val-fontStep').textContent = `Stufe ${settings.fontStep + 1} von ${FONT_SIZES.length}`;
  $('#font-sample').style.fontSize = Math.min(FONT_SIZES[settings.fontStep], 60) + 'px';
  $('#val-lines').textContent = `${settings.lines} Zeilen`;
  $('#val-speed').textContent = `Stufe ${settings.speed} von ${LIMITS.speed[1]}`;
  $$('#countdown-choices .choice').forEach((b) => b.setAttribute('aria-pressed', String(+b.dataset.val === settings.countdown)));
  $$('input[name="mode"]').forEach((r) => { r.checked = r.value === settings.mode; });
  $('#set-mirror').checked = settings.mirror;
  $('#set-folder').value = settings.folder;
  $('#set-client-id').value = settings.clientId;
  $('#redirect-uri').textContent = OD.redirectUri();
  $('#btn-od-logout').hidden = !OD.isConnected();
}

for (let s = 5; s <= 10; s++) {
  const b = document.createElement('button');
  b.className = 'choice';
  b.dataset.val = s;
  b.textContent = s;
  b.setAttribute('aria-label', `${s} Sekunden`);
  b.addEventListener('click', () => { settings.countdown = s; saveSettings(); renderSettings(); });
  $('#countdown-choices').append(b);
}

function changeSetting(key, dir) {
  const [min, max] = LIMITS[key];
  const val = Math.max(min, Math.min(max, settings[key] + dir));
  if (val === settings[key]) return false;
  settings[key] = val;
  saveSettings();
  return true;
}
$$('.step-btn').forEach((b) => b.addEventListener('click', () => {
  changeSetting(b.dataset.step, +b.dataset.dir);
  renderSettings();
}));
$$('input[name="mode"]').forEach((r) => r.addEventListener('change', () => {
  settings.mode = r.value;
  saveSettings();
  if (r.value === 'speech' && !speechSupported()) {
    toast('Dieser Browser bietet keine Spracherkennung. Der Text läuft automatisch.', 5000);
  }
}));
$('#set-mirror').addEventListener('change', (e) => { settings.mirror = e.target.checked; saveSettings(); });
$('#btn-reset-pos').addEventListener('click', () => {
  settings.promptTop = null;
  saveSettings();
  toast('Der Text steht wieder ganz oben.');
});
$('#set-folder').addEventListener('change', (e) => {
  settings.folder = e.target.value.trim() || DEFAULTS.folder;
  saveSettings();
  cloudScripts = [];
  store.set('cloudScripts', cloudScripts);
  syncOneDrive();
});
$('#set-client-id').addEventListener('change', (e) => {
  const v = e.target.value.trim();
  if (v !== settings.clientId) {
    settings.clientId = v;
    saveSettings();
    OD.logout();
    OD.configure({ clientId: v });
    renderOD();
    renderSettings();
  }
});
$('#btn-od-logout').addEventListener('click', () => {
  if (!confirm('Von OneDrive abmelden? Die Skripte aus OneDrive verschwinden dann aus der Liste.')) return;
  OD.logout();
  cloudScripts = [];
  cloudTexts = {};
  store.set('cloudScripts', cloudScripts);
  store.set('cloudTexts', cloudTexts);
  renderOD();
  renderLibrary();
  renderSettings();
});

// Spracherkennung testen
let testRec = null;
$('#btn-speech-test').addEventListener('click', () => {
  const out = $('#speech-test-out');
  const btn = $('#btn-speech-test');
  if (testRec) {
    testRec.onend = null;
    try { testRec.stop(); } catch { /* egal */ }
    testRec = null;
    btn.textContent = 'Spracherkennung testen';
    return;
  }
  out.hidden = false;
  if (!speechSupported()) {
    out.textContent = 'Dieser Browser bietet keine Spracherkennung.';
    return;
  }
  out.textContent = 'Bitte jetzt sprechen …';
  try {
    testRec = createRecognizer();
    testRec.onresult = (e) => {
      let t = '';
      for (let i = 0; i < e.results.length; i++) t += e.results[i][0].transcript + ' ';
      out.textContent = t.trim();
    };
    testRec.onerror = (e) => { out.textContent = `Fehler: ${e.error}. Auf dem iPhone muss „Diktieren“ eingeschaltet sein.`; };
    testRec.onend = () => { testRec = null; btn.textContent = 'Spracherkennung testen'; };
    testRec.start();
    btn.textContent = 'Test beenden';
  } catch (e) {
    out.textContent = 'Die Spracherkennung konnte nicht gestartet werden.';
    testRec = null;
  }
});

// ---------------------------------------------------------------------------
// Teleprompter
// ---------------------------------------------------------------------------
const P = {
  words: [],
  offset: 0,
  target: 0,
  running: false,
  paused: false,
  lastTs: 0,
  raf: 0,
  mode: 'auto',
  follower: null,
  doneUpTo: 0,
};

const fontPx = () => FONT_SIZES[settings.fontStep];
const lineHeightPx = () => fontPx() * LINE_HEIGHT;
const speedPx = () => 0.1 * settings.speed * lineHeightPx(); // Zeilen pro Sekunde × Zeilenhöhe

function buildPrompter(text) {
  const container = $('#prompter-text');
  container.textContent = '';
  P.words = [];
  const paragraphs = text.split(/\n+/).map((p) => p.trim()).filter(Boolean);
  for (const para of paragraphs) {
    const p = document.createElement('p');
    const sub = document.createElement('span');
    sub.className = 'sub';
    const parts = para.split(/\s+/);
    parts.forEach((w, i) => {
      const ws = document.createElement('span');
      ws.className = 'w';
      ws.textContent = w;
      sub.append(ws);
      if (i < parts.length - 1) sub.append(' ');
      P.words.push({ el: ws, norm: normalizeWord(w) });
    });
    p.append(sub);
    container.append(p);
  }
  const end = document.createElement('div');
  end.className = 'prompter-end';
  container.append(end);
}

function maxOffset() {
  const last = P.words[P.words.length - 1];
  return last ? last.el.offsetTop : 0;
}

function renderOffset() {
  $('#prompter-text').style.transform = `translate3d(0, ${-P.offset}px, 0)`;
}

function clampTop(top) {
  const pr = $('#prompter');
  const minTop = 0;
  const maxTop = Math.max(minTop, window.innerHeight - pr.offsetHeight - 150);
  return Math.max(minTop, Math.min(maxTop, top));
}

function applyPrompterLayout() {
  const root = document.documentElement.style;
  root.setProperty('--fs', fontPx() + 'px');
  root.setProperty('--lh', String(LINE_HEIGHT));
  root.setProperty('--lines', String(settings.lines));
  const pr = $('#prompter');
  if (settings.promptTop == null) pr.style.top = '';
  else pr.style.top = clampTop(settings.promptTop) + 'px';
  $('#cam').classList.toggle('mirror', settings.mirror);
  $('#speed-label').textContent = `Tempo ${settings.speed}`;
}

/** Schriftgröße ändern, ohne die Leseposition zu verlieren. */
function relayoutKeepingPosition() {
  const anchor = P.words.find((w) => w.el.offsetTop + 2 >= P.offset);
  applyPrompterLayout();
  if (anchor) {
    P.offset = Math.min(anchor.el.offsetTop, maxOffset());
    P.target = P.offset;
  }
  renderOffset();
}

function tick(ts) {
  const dt = P.lastTs ? Math.min(0.1, (ts - P.lastTs) / 1000) : 0;
  P.lastTs = ts;
  if (P.running && !P.paused) {
    if (P.mode === 'auto') {
      P.offset += speedPx() * dt;
    } else {
      P.offset += (P.target - P.offset) * Math.min(1, dt * 5);
    }
    P.offset = Math.max(0, Math.min(P.offset, maxOffset()));
    renderOffset();
  }
  P.raf = requestAnimationFrame(tick);
}
function startLoop() {
  cancelAnimationFrame(P.raf);
  P.lastTs = 0;
  P.raf = requestAnimationFrame(tick);
}

function resetPrompter() {
  P.offset = 0;
  P.target = 0;
  P.doneUpTo = 0;
  P.words.forEach((w) => w.el.classList.remove('done'));
  renderOffset();
}

function markDone(upTo) {
  for (let i = P.doneUpTo; i < upTo && i < P.words.length; i++) P.words[i].el.classList.add('done');
  P.doneUpTo = Math.max(P.doneUpTo, upTo);
}

function updateBadges() {
  $('#paused-badge').hidden = !(P.running && P.paused);
  const mb = $('#mode-badge');
  if (P.running && P.mode === 'speech' && !P.paused) {
    mb.textContent = 'Sprachsteuerung aktiv';
    mb.hidden = false;
  } else {
    mb.hidden = true;
  }
}

function startSpeech() {
  if (!speechSupported()) {
    P.mode = 'auto';
    toast('Keine Spracherkennung verfügbar. Der Text läuft automatisch.', 5000);
    return;
  }
  P.follower = new SpeechFollower(P.words.map((w) => w.norm), {
    onProgress: (idx) => {
      markDone(idx + 1);
      const el = P.words[idx].el;
      P.target = Math.min(maxOffset(), Math.max(0, el.offsetTop - fontPx() * 0.1));
    },
    onFail: () => {
      P.follower = null;
      P.mode = 'auto';
      updateBadges();
      toast('Die Sprachsteuerung funktioniert hier nicht zuverlässig. Der Text läuft jetzt automatisch.', 6000);
    },
  });
  P.follower.start();
}

function startRun() {
  P.running = true;
  P.paused = false;
  P.mode = settings.mode;
  P.lastTs = 0;
  if (P.mode === 'speech') startSpeech();
  updateBadges();
}

function stopRun() {
  P.running = false;
  P.paused = false;
  if (P.follower) { P.follower.stop(); P.follower = null; }
  updateBadges();
}

function togglePause() {
  if (!P.running) {
    if (recState === 'idle') toast('Zum Starten den orangen Aufnahme-Knopf tippen.');
    return;
  }
  P.paused = !P.paused;
  if (P.follower) {
    if (P.paused) P.follower.stop();
    else P.follower.start();
  }
  updateBadges();
}

// Tippen = Pause/Weiter, Ziehen am Text = Position verschieben
(() => {
  const screen = $('#screen-record');
  const pr = $('#prompter');
  let g = null;
  screen.addEventListener('pointerdown', (e) => {
    if (e.target.closest('button, .rec-controls, .cam-error')) return;
    g = { id: e.pointerId, x: e.clientX, y: e.clientY, startTop: pr.offsetTop, onPrompter: !!e.target.closest('#prompter'), dragging: false };
  });
  screen.addEventListener('pointermove', (e) => {
    if (!g || e.pointerId !== g.id) return;
    const dy = e.clientY - g.y;
    if (!g.dragging && g.onPrompter && Math.abs(dy) > 10) {
      g.dragging = true;
      pr.classList.add('dragging');
      try { screen.setPointerCapture(e.pointerId); } catch { /* egal */ }
    }
    if (g.dragging) {
      pr.style.top = clampTop(g.startTop + dy) + 'px';
      e.preventDefault();
    }
  });
  const end = (e, cancelled) => {
    if (!g || e.pointerId !== g.id) return;
    const gg = g;
    g = null;
    if (gg.dragging) {
      pr.classList.remove('dragging');
      settings.promptTop = pr.offsetTop;
      saveSettings();
      return;
    }
    if (!cancelled && Math.hypot(e.clientX - gg.x, e.clientY - gg.y) < 12) togglePause();
  };
  screen.addEventListener('pointerup', (e) => end(e, false));
  screen.addEventListener('pointercancel', (e) => end(e, true));
})();

// Schnellregler auf dem Aufnahme-Bildschirm
$$('[data-quick]').forEach((b) => b.addEventListener('click', () => {
  const key = b.dataset.quick;
  if (!changeSetting(key, +b.dataset.dir)) {
    toast(key === 'fontStep' ? 'Grenze der Schriftgröße erreicht.' : 'Grenze des Tempos erreicht.', 1500);
    return;
  }
  if (key === 'fontStep') {
    relayoutKeepingPosition();
    toast(`Schrift: Stufe ${settings.fontStep + 1} von ${FONT_SIZES.length}`, 1500);
  } else {
    $('#speed-label').textContent = `Tempo ${settings.speed}`;
    toast(`Tempo: Stufe ${settings.speed} von ${LIMITS.speed[1]}`, 1500);
  }
}));

// ---------------------------------------------------------------------------
// Kamera und Aufnahme
// ---------------------------------------------------------------------------
let stream = null;
let recorder = null;
let chunks = [];
let recBlob = null;
let previewUrl = null;
let recState = 'idle'; // idle | countdown | recording
let cdTimer = null;
let recTimer = null;
let recStartedAt = 0;
let canvasRaf = 0;

function camErrorText(e) {
  if (!window.isSecureContext) return 'Die Kamera funktioniert nur, wenn die App über https geöffnet wird.';
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return 'Dieser Browser erlaubt keinen Kamerazugriff.';
  switch (e && e.name) {
    case 'NotAllowedError':
      return 'Kamera oder Mikrofon wurden nicht erlaubt. Auf dem iPhone: Einstellungen › Apps › Safari › Kamera und Mikrofon auf „Fragen“ oder „Erlauben“ stellen. Danach „Erneut versuchen“ tippen.';
    case 'NotFoundError':
      return 'Es wurde keine Kamera gefunden.';
    case 'NotReadableError':
      return 'Die Kamera wird gerade von einer anderen App benutzt. Bitte die andere App schließen.';
    default:
      return 'Die Kamera konnte nicht gestartet werden.';
  }
}

async function startCamera() {
  if (stream) return true;
  $('#cam-error').hidden = true;
  try {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('unsupported');
    stream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: 'user',
        width: { ideal: 1920 },
        height: { ideal: 1080 },
        frameRate: { ideal: 30 },
      },
      audio: { echoCancellation: true, noiseSuppression: true },
    });
  } catch (e) {
    stream = null;
    $('#cam-error-text').textContent = camErrorText(e);
    $('#cam-error').hidden = false;
    return false;
  }
  const cam = $('#cam');
  cam.srcObject = stream;
  await cam.play().catch(() => {});
  return true;
}
$('#btn-cam-retry').addEventListener('click', startCamera);

function stopCamera() {
  if (stream) stream.getTracks().forEach((t) => t.stop());
  stream = null;
  $('#cam').srcObject = null;
}

function pickMime() {
  if (typeof MediaRecorder === 'undefined' || !MediaRecorder.isTypeSupported) return '';
  const types = [
    'video/mp4;codecs=avc1.42E01E,mp4a.40.2',
    'video/mp4;codecs=avc1',
    'video/mp4',
    'video/webm;codecs=vp9,opus',
    'video/webm;codecs=vp8,opus',
    'video/webm',
  ];
  return types.find((t) => MediaRecorder.isTypeSupported(t)) || '';
}

/**
 * Liefert einen Stream im Hochformat 9:16. Liefert die Kamera bereits 9:16,
 * wird sie direkt aufgenommen, sonst wird die Mitte über ein Canvas ausgeschnitten.
 * Der Teleprompter-Text ist nie Teil der Aufnahme.
 */
function recordingStream() {
  const cam = $('#cam');
  const w = cam.videoWidth;
  const h = cam.videoHeight;
  if (!w || !h) return stream;
  const ratio = w / h;
  if (ratio > 0.55 && ratio < 0.575) return stream;

  let sx, sy, sw, sh;
  if (ratio > 9 / 16) { sh = h; sw = Math.round(h * 9 / 16); sx = Math.round((w - sw) / 2); sy = 0; }
  else { sw = w; sh = Math.round(w * 16 / 9); sx = 0; sy = Math.round((h - sh) / 2); }

  const canvas = $('#rec-canvas');
  const outW = Math.min(1080, sw - (sw % 2));
  const outH = Math.round(outW * 16 / 9);
  canvas.width = outW;
  canvas.height = outH - (outH % 2);
  const ctx = canvas.getContext('2d');
  const draw = () => {
    ctx.drawImage(cam, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height);
    canvasRaf = requestAnimationFrame(draw);
  };
  draw();
  const cs = canvas.captureStream(30);
  stream.getAudioTracks().forEach((t) => cs.addTrack(t));
  return cs;
}
function stopCanvas() {
  cancelAnimationFrame(canvasRaf);
  canvasRaf = 0;
}

function setRecUi() {
  const btn = $('#btn-record');
  btn.classList.toggle('counting', recState === 'countdown');
  btn.classList.toggle('recording', recState === 'recording');
  btn.setAttribute('aria-label',
    recState === 'idle' ? 'Aufnahme starten' : recState === 'countdown' ? 'Countdown abbrechen' : 'Aufnahme beenden');
  $('#btn-rec-back').disabled = recState !== 'idle';
  $('#rec-time').hidden = recState !== 'recording';
}

function updateRecTime() {
  const s = Math.floor((Date.now() - recStartedAt) / 1000);
  $('#rec-time').textContent = `● ${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`;
}

$('#btn-record').addEventListener('click', () => {
  if (recState === 'idle') startCountdown();
  else if (recState === 'countdown') cancelCountdown();
  else stopRecording();
});

async function startCountdown() {
  if (!stream && !(await startCamera())) return;
  resetPrompter();
  recState = 'countdown';
  setRecUi();
  let n = settings.countdown;
  $('#countdown-num').textContent = n;
  $('#countdown').hidden = false;
  cdTimer = setInterval(() => {
    n -= 1;
    if (n <= 0) {
      clearInterval(cdTimer);
      cdTimer = null;
      $('#countdown').hidden = true;
      beginRecording();
    } else {
      $('#countdown-num').textContent = n;
    }
  }, 1000);
}

function cancelCountdown() {
  if (cdTimer) { clearInterval(cdTimer); cdTimer = null; }
  $('#countdown').hidden = true;
  if (recState === 'countdown') { recState = 'idle'; setRecUi(); }
}

function beginRecording() {
  try {
    if (typeof MediaRecorder === 'undefined') throw new Error('Dieser Browser kann keine Videos aufnehmen.');
    const rs = recordingStream();
    const mime = pickMime();
    const opts = { videoBitsPerSecond: 8000000, audioBitsPerSecond: 128000 };
    if (mime) opts.mimeType = mime;
    chunks = [];
    recorder = new MediaRecorder(rs, opts);
    recorder.ondataavailable = (e) => { if (e.data && e.data.size) chunks.push(e.data); };
    recorder.onstop = () => {
      stopCanvas();
      clearInterval(recTimer);
      const type = recorder.mimeType || mime || 'video/mp4';
      recBlob = new Blob(chunks, { type });
      chunks = [];
      recorder = null;
      if (recState === 'recording') { recState = 'idle'; stopRun(); setRecUi(); }
      showPreview();
    };
    recorder.start();
  } catch (e) {
    stopCanvas();
    recState = 'idle';
    setRecUi();
    toast('Die Aufnahme konnte nicht gestartet werden. ' + (e.message || ''), 6000);
    return;
  }
  recStartedAt = Date.now();
  updateRecTime();
  recTimer = setInterval(updateRecTime, 500);
  recState = 'recording';
  setRecUi();
  startRun();
}

function stopRecording() {
  recState = 'idle';
  stopRun();
  setRecUi();
  if (recorder && recorder.state !== 'inactive') recorder.stop();
}

// ---------------------------------------------------------------------------
// Bildschirm wach halten
// ---------------------------------------------------------------------------
let wakeLock = null;
async function requestWakeLock() {
  try {
    if ('wakeLock' in navigator && !wakeLock) {
      wakeLock = await navigator.wakeLock.request('screen');
      wakeLock.addEventListener('release', () => { wakeLock = null; });
    }
  } catch { /* nicht unterstützt */ }
}
function releaseWakeLock() {
  if (wakeLock) wakeLock.release().catch(() => {});
  wakeLock = null;
}

// ---------------------------------------------------------------------------
// Aufnahme-Bildschirm betreten / verlassen
// ---------------------------------------------------------------------------
async function enterRecord(text) {
  buildPrompter(cleanScript(text));
  recState = 'idle';
  stopRun();
  setRecUi();
  showScreen('record');
  applyPrompterLayout();
  resetPrompter();
  startLoop();
  requestWakeLock();
  await startCamera();
}

function leaveRecord() {
  cancelCountdown();
  stopRun();
  cancelAnimationFrame(P.raf);
  stopCamera();
  releaseWakeLock();
  showScreen('library');
  renderLibrary();
  syncOneDrive();
}
$('#btn-rec-back').addEventListener('click', () => { if (recState === 'idle') leaveRecord(); });

// ---------------------------------------------------------------------------
// Vorschau, Speichern, Verwerfen
// ---------------------------------------------------------------------------
function showPreview() {
  const v = $('#preview-video');
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = URL.createObjectURL(recBlob);
  v.src = previewUrl;
  $('#preview-unsaved').hidden = false;
  $('#preview-saved').hidden = true;
  showScreen('preview');
  v.play().catch(() => {});
}

function clearPreview() {
  const v = $('#preview-video');
  v.pause();
  v.removeAttribute('src');
  v.load();
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = null;
  recBlob = null;
}

function backToRecord() {
  clearPreview();
  showScreen('record');
  applyPrompterLayout();
  resetPrompter();
  requestWakeLock();
  if (!stream) startCamera();
}

function fileName(ext) {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `Mary-Prompt_${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}.${ext}`;
}

async function saveVideo() {
  if (!recBlob) return;
  const type = (recBlob.type || 'video/mp4').split(';')[0];
  const ext = type.includes('mp4') ? 'mp4' : type.includes('quicktime') ? 'mov' : 'webm';
  const name = fileName(ext);
  const file = new File([recBlob], name, { type });

  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({ files: [file] });
      $('#preview-unsaved').hidden = true;
      $('#preview-saved').hidden = false;
      return;
    } catch (e) {
      if (e.name === 'AbortError') {
        toast('Nicht gespeichert. Du kannst es noch einmal versuchen.');
        return;
      }
      // sonst: Herunterladen versuchen
    }
  }
  const a = document.createElement('a');
  a.href = previewUrl;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  $('#preview-unsaved').hidden = true;
  $('#preview-saved').hidden = false;
  toast('Das Video wurde heruntergeladen.');
}

$('#btn-save').addEventListener('click', saveVideo);
$('#btn-save-again').addEventListener('click', saveVideo);
$('#btn-discard').addEventListener('click', () => {
  if (confirm('Aufnahme wirklich verwerfen? Sie wird nicht gespeichert.')) backToRecord();
});
$('#btn-new-take').addEventListener('click', backToRecord);
$('#btn-to-library').addEventListener('click', () => {
  clearPreview();
  leaveRecord();
});

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------
async function init() {
  $('#version-info').textContent = `Mary Prompt · Version ${APP_VERSION}`;
  OD.configure({ clientId: settings.clientId });
  applyPrompterLayout();
  renderLibrary();
  renderOD();

  let redirectFailed = false;
  try {
    const r = await OD.handleRedirect();
    if (r && r.error) { renderOD('Die Anmeldung hat nicht geklappt: ' + r.error); redirectFailed = true; }
    else if (r && r.ok) toast('OneDrive ist verbunden.');
  } catch (e) {
    renderOD('Die Anmeldung hat nicht geklappt: ' + (e.message || 'unbekannter Fehler'));
    redirectFailed = true;
  }
  if (!redirectFailed) syncOneDrive();

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'visible') return;
    if (currentScreen === 'library') syncOneDrive();
    if (currentScreen === 'record' || currentScreen === 'preview') requestWakeLock();
  });
  window.addEventListener('resize', () => {
    if (currentScreen === 'record') applyPrompterLayout();
  });

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }
}

init();
