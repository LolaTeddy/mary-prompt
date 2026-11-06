// Sprachgesteuerter Textlauf: vergleicht das Gesprochene mit dem Skript
// und meldet, bis zu welchem Wort gelesen wurde.

const FATAL = new Set(['not-allowed', 'service-not-allowed', 'audio-capture', 'language-not-supported']);

export function speechSupported() {
  return !!(window.SpeechRecognition || window.webkitSpeechRecognition);
}

export function normalizeWord(w) {
  return (w || '')
    .toLowerCase()
    .replace(/ß/g, 'ss')
    .normalize('NFC')
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

function lev(a, b) {
  const m = a.length, n = b.length;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[n];
}

function similar(a, b) {
  if (a === b) return true;
  const min = Math.min(a.length, b.length);
  if (min >= 4 && (a.startsWith(b) || b.startsWith(a))) return true;
  if (min >= 5 && Math.abs(a.length - b.length) <= 1 && lev(a, b) <= 1) return true;
  if (min >= 8 && lev(a, b) <= 2) return true;
  return false;
}

export function createRecognizer() {
  const R = window.SpeechRecognition || window.webkitSpeechRecognition;
  const rec = new R();
  rec.lang = 'de-DE';
  rec.continuous = true;
  rec.interimResults = true;
  rec.maxAlternatives = 1;
  return rec;
}

export class SpeechFollower {
  /**
   * @param {string[]} norms  normalisierte Wörter des Skripts (leere Einträge = Satzzeichen)
   * @param {{onProgress:(wordIndex:number)=>void, onFail:(reason:string)=>void}} cb
   */
  constructor(norms, { onProgress, onFail }) {
    this.map = [];
    this.seq = [];
    norms.forEach((n, i) => { if (n) { this.map.push(i); this.seq.push(n); } });
    this.pos = 0;
    this.active = false;
    this.rec = null;
    this.onProgress = onProgress;
    this.onFail = onFail;
  }

  start() {
    if (this.active) return;
    this.active = true;
    this.restarts = [];
    try { this._spawn(); } catch { this._fail('start'); }
  }

  stop() {
    this.active = false;
    const r = this.rec;
    this.rec = null;
    if (r) {
      r.onend = r.onresult = r.onerror = null;
      try { r.abort(); } catch { /* egal */ }
    }
  }

  _spawn() {
    const rec = createRecognizer();
    rec.onresult = (e) => {
      let t = '';
      for (let i = 0; i < e.results.length; i++) t += ' ' + e.results[i][0].transcript;
      this._consume(t);
    };
    rec.onerror = (e) => { if (FATAL.has(e.error)) this._fail(e.error); };
    rec.onend = () => {
      // Safari beendet die Erkennung regelmäßig – dann neu starten.
      if (!this.active || this.rec !== rec) return;
      const now = Date.now();
      this.restarts = this.restarts.filter((t) => now - t < 20000);
      this.restarts.push(now);
      if (this.restarts.length > 8) { this._fail('restarts'); return; }
      setTimeout(() => {
        if (this.active && this.rec === rec) {
          try { this._spawn(); } catch { this._fail('start'); }
        }
      }, 200);
    };
    this.rec = rec;
    rec.start();
  }

  _fail(reason) {
    if (!this.active) return;
    this.stop();
    if (this.onFail) this.onFail(reason);
  }

  _consume(transcript) {
    const tail = transcript.split(/\s+/).map(normalizeWord).filter(Boolean).slice(-6);
    if (!tail.length) return;
    const seq = this.seq;
    const last = tail.length - 1;
    const from = Math.max(0, this.pos - 3);
    const to = Math.min(seq.length - 1, this.pos + 40);
    let best = -1, bestScore = 0, bestDist = Infinity;

    for (let e = from; e <= to; e++) {
      if (!similar(seq[e], tail[last])) continue;
      let score = 1;
      for (let k = 1; k <= last && e - k >= 0; k++) {
        if (similar(seq[e - k], tail[last - k])) score++;
      }
      const dist = Math.abs(e - this.pos);
      const needed = e > this.pos + 12 ? 3 : 2;
      const ok = score >= needed || (score === 1 && e >= this.pos && e <= this.pos + 2 && seq[e].length >= 3);
      if (!ok) continue;
      if (score > bestScore || (score === bestScore && dist < bestDist)) {
        best = e; bestScore = score; bestDist = dist;
      }
    }
    if (best >= 0 && best + 1 > this.pos) {
      this.pos = best + 1;
      if (this.onProgress) this.onProgress(this.map[best]);
    }
  }
}
