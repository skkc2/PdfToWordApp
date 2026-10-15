import * as pdfjsLib from './vendor/pdf.min.mjs';

const MAX_PDF_BYTES = 30 * 1024 * 1024;
const MAX_PAGES = 300;
const MAX_TOTAL_CHARS = 3000000;
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

pdfjsLib.GlobalWorkerOptions.workerSrc =
  new URL('./vendor/pdf.worker.min.mjs', import.meta.url).href;

const fileInput = document.getElementById('file');
const convertBtn = document.getElementById('convert');
const progressEl = document.getElementById('progress');
const statusEl = document.getElementById('status');
let busy = false;

class UserError extends Error {}

function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.classList.toggle('error', isError);
}

fileInput.addEventListener('change', () => {
  const f = fileInput.files && fileInput.files[0];
  convertBtn.disabled = !f || busy;
  setStatus(f ? 'PDF selected. Tap step 2 to convert.' : 'Select a PDF to begin.');
});

convertBtn.addEventListener('click', async () => {
  if (busy) return;
  const file = fileInput.files && fileInput.files[0];
  if (!file) return;

  busy = true;
  convertBtn.disabled = true;
  fileInput.disabled = true;
  progressEl.hidden = false;
  progressEl.value = 0;
  setStatus('Starting local conversion...');

  try {
    const pages = await extractPages(file);
    const blob = buildDocx(pages);
    saveBlob(blob, outputName(file.name));
    setStatus('Done! Your Word (.docx) file was downloaded.');
  } catch (err) {
    setStatus(friendlyMessage(err), true);
  } finally {
    busy = false;
    fileInput.disabled = false;
    convertBtn.disabled = !(fileInput.files && fileInput.files[0]);
    progressEl.hidden = true;
  }
});

function friendlyMessage(err) {
  if (err instanceof UserError) return err.message;
  if (err && err.name === 'PasswordException') {
    return "This PDF is password-protected and can't be converted.";
  }
  if (err && (err.name === 'InvalidPDFException' || err.name === 'FormatError')) {
    return 'Could not read the PDF. It may be damaged or not a valid PDF.';
  }
  if (err instanceof RangeError) {
    return 'Not enough memory to convert this PDF. Try a smaller file.';
  }
  return 'Conversion failed. The file may be damaged or unsupported.';
}

function hasPdfHeader(bytes) {
  const limit = Math.min(bytes.length - 4, 1024);
  for (let i = 0; i <= limit; i++) {
    if (bytes[i] === 0x25 && bytes[i + 1] === 0x50 && bytes[i + 2] === 0x44 &&
        bytes[i + 3] === 0x46 && bytes[i + 4] === 0x2d) return true;
  }
  return false;
}

async function extractPages(file) {
  if (file.size > MAX_PDF_BYTES) {
    throw new UserError('This PDF is larger than the ' + (MAX_PDF_BYTES / 1048576) + ' MB limit.');
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (!hasPdfHeader(bytes)) {
    throw new UserError('This file does not look like a PDF.');
  }

  const task = pdfjsLib.getDocument({
    data: bytes,
    cMapUrl: new URL('./vendor/cmaps/', import.meta.url).href,
    cMapPacked: true,
    isEvalSupported: false,
    enableXfa: false,
    disableFontFace: true,
    useSystemFonts: false
  });

  const pages = [];
  let totalChars = 0;
  try {
    const pdf = await task.promise;
    const total = pdf.numPages;
    if (total > MAX_PAGES) {
      throw new UserError('This PDF has ' + total + ' pages. The limit is ' + MAX_PAGES + '.');
    }
    for (let p = 1; p <= total; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      let annots = [];
      try { annots = await page.getAnnotations(); } catch (e) { annots = []; }
      const lines = itemsToLines(content.items, annots);
      page.cleanup();
      for (const line of lines) totalChars += line.length;
      if (totalChars > MAX_TOTAL_CHARS) {
        throw new UserError('This PDF contains too much text to convert safely.');
      }
      pages.push(lines);
      progressEl.value = Math.round((p * 100) / total);
      setStatus('Converting page ' + p + ' of ' + total + '...');
    }
  } finally {
    await task.destroy();
  }
  return pages;
}

function itemsToLines(items, annotations) {
  const pieces = [];
  for (const it of items) {
    if (typeof it.str !== 'string' || it.str === '' || !it.transform) continue;
    pieces.push({ x: it.transform[4], y: it.transform[5], s: it.str, w: it.width || 0 });
  }
  for (const a of annotations || []) {
    if (a.subtype !== 'Widget' || !a.rect || a.fieldType === 'Btn') continue;
    let v = a.fieldValue;
    if (Array.isArray(v)) v = v.join(', ');
    if (typeof v !== 'string') continue;
    const rows = v.split(/\r\n|\r|\n/).map((s) => s.trim()).filter((s) => s.length > 0);
    const top = Math.max(a.rect[1], a.rect[3]);
    const bottom = Math.min(a.rect[1], a.rect[3]);
    const left = Math.min(a.rect[0], a.rect[2]);
    rows.forEach((row, i) => {
      const y = rows.length === 1 ? (top + bottom) / 2 - 4 : top - 10 - i * 12;
      pieces.push({ x: left, y: y, s: row, w: row.length * 5 });
    });
  }
  pieces.sort((a, b) => (b.y - a.y) || (a.x - b.x));
  const lines = [];
  let cur = null;
  for (const p of pieces) {
    if (cur && Math.abs(p.y - cur.y) <= 3) {
      cur.parts.push(p);
    } else {
      cur = { y: p.y, parts: [p] };
      lines.push(cur);
    }
  }
  return lines.map((l) => {
    l.parts.sort((a, b) => a.x - b.x);
    let out = '';
    let end = null;
    for (const p of l.parts) {
      if (end !== null && p.x - end > 1.5 && !/\s$/.test(out) && !/^\s/.test(p.s)) out += ' ';
      out += p.s;
      end = p.x + p.w;
    }
    return out.replace(/\s+/g, ' ').trim();
  }).filter((s) => s.length > 0);
}
/* ---------- .docx writing (plain OpenXML in an uncompressed ZIP) ---------- */

function sanitizeXmlChars(input) {
  let out = '';
  for (const ch of input) {
    const cp = ch.codePointAt(0);
    if (cp === 0x9 || cp === 0xA || cp === 0xD ||
        (cp >= 0x20 && cp <= 0xD7FF) ||
        (cp >= 0xE000 && cp <= 0xFFFD) ||
        (cp >= 0x10000 && cp <= 0x10FFFF)) {
      out += ch;
    }
  }
  return out;
}

function escapeXml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function buildDocumentXml(pages) {
  const parts = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  parts.push('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>');
  pages.forEach((paragraphs, i) => {
    for (const text of paragraphs) {
      parts.push('<w:p><w:r><w:t xml:space="preserve">' +
                 escapeXml(sanitizeXmlChars(text)) + '</w:t></w:r></w:p>');
    }
    if (i < pages.length - 1) {
      parts.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    }
  });
  parts.push('</w:body></w:document>');
  return parts.join('');
}

const CONTENT_TYPES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const encoder = new TextEncoder();

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function zipStore(entries) {
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const e of entries) {
    const name = encoder.encode(e.name);
    const crc = crc32(e.data);
    const size = e.data.length;

    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true);
    lh.setUint16(4, 20, true);
    lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, 0, true);
    lh.setUint16(10, 0, true);
    lh.setUint16(12, 0x21, true);
    lh.setUint32(14, crc, true);
    lh.setUint32(18, size, true);
    lh.setUint32(22, size, true);
    lh.setUint16(26, name.length, true);
    lh.setUint16(28, 0, true);
    chunks.push(new Uint8Array(lh.buffer), name, e.data);

    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true);
    ch.setUint16(4, 20, true);
    ch.setUint16(6, 20, true);
    ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, 0, true);
    ch.setUint16(12, 0, true);
    ch.setUint16(14, 0x21, true);
    ch.setUint32(16, crc, true);
    ch.setUint32(20, size, true);
    ch.setUint32(24, size, true);
    ch.setUint16(28, name.length, true);
    ch.setUint16(30, 0, true);
    ch.setUint16(32, 0, true);
    ch.setUint16(34, 0, true);
    ch.setUint16(36, 0, true);
    ch.setUint32(38, 0, true);
    ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);

    offset += 30 + name.length + size;
  }
  let cdSize = 0;
  for (const c of central) cdSize += c.length;

  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(4, 0, true);
  end.setUint16(6, 0, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  end.setUint16(20, 0, true);

  return new Blob([...chunks, ...central, new Uint8Array(end.buffer)], { type: DOCX_MIME });
}

function buildDocx(pages) {
  return zipStore([
    { name: '[Content_Types].xml', data: encoder.encode(CONTENT_TYPES_XML) },
    { name: '_rels/.rels', data: encoder.encode(RELS_XML) },
    { name: 'word/document.xml', data: encoder.encode(buildDocumentXml(pages)) }
  ]);
}

function outputName(original) {
  const base = String(original || '')
    .replace(/\.pdf$/i, '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .trim()
    .slice(0, 80);
  return (base || 'Converted_Document') + '.docx';
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}