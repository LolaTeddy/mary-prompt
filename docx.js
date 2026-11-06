// Liest den reinen Text aus einer Word-Datei (.docx), ohne Bibliothek.
// Eine .docx ist ein ZIP-Archiv; der Text steht in word/document.xml.
// Entpacken über DecompressionStream (iPhone ab iOS 16.4).

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function isDocx(name) {
  return /\.docx$/i.test(name || '');
}

function err(code) {
  const e = new Error(code);
  e.code = code;
  return e;
}

async function inflateRaw(bytes) {
  if (typeof DecompressionStream === 'undefined') throw err('unsupported');
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Sucht eine Datei im ZIP-Archiv und gibt ihren entpackten Inhalt zurück. */
async function readZipEntry(buffer, wanted) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  // Ende des zentralen Verzeichnisses suchen
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw err('notzip');
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (view.getUint32(p, true) !== 0x02014b50) throw err('notzip');
    const method = view.getUint16(p + 10, true);
    const compSize = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localOff = view.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    if (name === wanted) {
      const lNameLen = view.getUint16(localOff + 26, true);
      const lExtraLen = view.getUint16(localOff + 28, true);
      const start = localOff + 30 + lNameLen + lExtraLen;
      const data = bytes.subarray(start, start + compSize);
      if (method === 0) return data;
      if (method === 8) return inflateRaw(data);
      throw err('method');
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  throw err('nodocument');
}

function paragraphText(p) {
  let s = '';
  const walk = (node) => {
    for (const c of node.children) {
      if (c.namespaceURI === W) {
        if (c.localName === 'p') continue; // verschachtelte Absätze kommen separat
        if (c.localName === 't') { s += c.textContent; continue; }
        if (c.localName === 'tab') { s += ' '; continue; }
        if (c.localName === 'br' || c.localName === 'cr') { s += '\n'; continue; }
        if (c.localName === 'delText' || c.localName === 'instrText') continue;
      }
      walk(c);
    }
  };
  walk(p);
  return s;
}

/** Liefert den Text einer .docx-Datei, Absätze durch Zeilenumbrüche getrennt. */
export async function extractDocxText(buffer) {
  const xmlBytes = await readZipEntry(buffer, 'word/document.xml');
  const xml = new TextDecoder('utf-8').decode(xmlBytes);
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw err('xml');
  const paras = Array.from(doc.getElementsByTagNameNS(W, 'p')).filter((p) => {
    // Textfelder stehen doppelt im Dokument (Choice + Fallback) – Fallback auslassen
    for (let a = p.parentElement; a; a = a.parentElement) if (a.localName === 'Fallback') return false;
    return true;
  });
  return paras.map(paragraphText).join('\n');
}
