import * as pdfjsLib from './vendor/pdf.min.mjs';

const MAX_PDF_BYTES = 30 * 1024 * 1024;
const MAX_PAGES = 300;
const MAX_TOTAL_CHARS = 3000000;
const MAX_RENDER_PIXELS = 18000000;
const MAX_OUTPUT_BYTES = 180 * 1024 * 1024;
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const encoder = new TextEncoder();
let sourcePdfBytes = null;
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdf.worker.min.mjs', import.meta.url).href;

const fileInput = document.getElementById('file');
const convertBtn = document.getElementById('convert');
const progressEl = document.getElementById('progress');
const statusEl = document.getElementById('status');
const visualAppendixInput = document.getElementById('visual-appendix');
let busy = false;\nlet currentStage = 'startup';
class UserError extends Error {}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle('error', isError);
}
fileInput.addEventListener('change', () => {
  const file = fileInput.files && fileInput.files[0];
  convertBtn.disabled = !file || busy;
  setStatus(file ? 'PDF selected. Ready for local conversion.' : 'Select a PDF to begin.');
});
convertBtn.addEventListener('click', async () => {
  if (busy) return;
  const file = fileInput.files && fileInput.files[0];
  if (!file) return;
  busy = true;
  convertBtn.disabled = true;
  fileInput.disabled = true;
  visualAppendixInput.disabled = true;
  progressEl.hidden = false;
  progressEl.value = 0;
  setStatus('Starting local conversion…');
  try {
    currentStage = 'PDF text analysis';
    const model = await analyzePdf(file);
    currentStage = visualAppendixInput.checked ? 'visual appendix rendering or DOCX packaging' : 'DOCX packaging';
    const blob = await buildDocx(model, visualAppendixInput.checked, (percent, message) => {
      progressEl.value = percent;
      if (message.toLowerCase().includes('rendering visual reference')) currentStage = 'visual appendix rendering';
      else if (message.toLowerCase().includes('packaging')) currentStage = 'DOCX packaging';
      setStatus(message);
    });
    if (blob.size > MAX_OUTPUT_BYTES) throw new UserError('The Word file exceeds the 180 MB safety limit. Disable the visual appendix and try again.');
    saveBlob(blob, outputName(file.name));
    setStatus('Done. Your editable Word document was downloaded locally.');
  } catch (error) {
    setStatus(friendlyMessage(error), true);
  } finally {
    sourcePdfBytes = null;
    busy = false;
    fileInput.disabled = false;
    visualAppendixInput.disabled = false;
    convertBtn.disabled = !(fileInput.files && fileInput.files[0]);
    progressEl.hidden = true;
  }
});
function friendlyMessage(error) {
  if (error instanceof UserError) return error.message;
  if (error?.name === 'PasswordException') return "This PDF is password-protected and can't be converted.";
  if (error?.name === 'InvalidPDFException' || error?.name === 'FormatError') return 'Could not read the PDF. It may be damaged or not a valid PDF.';
  if (error instanceof RangeError) return 'Not enough memory. Try a smaller PDF or disable the visual appendix.';
  const detail = String(error?.message || error || 'Unknown error').replace(/\s+/g, ' ').slice(0, 220);
  return 'Conversion failed during ' + currentStage + '. ' + (error?.name ? error.name + ': ' : '') + detail;
}
function hasPdfHeader(bytes) {
  const limit = Math.min(bytes.length - 5, 1024);
  for (let i = 0; i <= limit; i++) {
    if (bytes[i] === 0x25 && bytes[i + 1] === 0x50 && bytes[i + 2] === 0x44 &&
        bytes[i + 3] === 0x46 && bytes[i + 4] === 0x2d) return true;
  }
  return false;
}

async function analyzePdf(file) {
  if (file.size > MAX_PDF_BYTES) throw new UserError('This PDF is larger than the 30 MB limit.');
  sourcePdfBytes = new Uint8Array(await file.arrayBuffer());
  if (!hasPdfHeader(sourcePdfBytes)) throw new UserError('This file does not look like a PDF.');
  const task = pdfjsLib.getDocument({
    data: sourcePdfBytes.slice(),
    cMapUrl: new URL('./vendor/cmaps/', import.meta.url).href,
    cMapPacked: true,
    isEvalSupported: false,
    enableXfa: false,
    useSystemFonts: true,
  });
  const pages = [];
  let totalChars = 0;
  try {
    const pdf = await task.promise;
    if (pdf.numPages > MAX_PAGES) throw new UserError('This PDF has ' + pdf.numPages + ' pages. The limit is ' + MAX_PAGES + '.');
    for (let n = 1; n <= pdf.numPages; n++) {
      const page = await pdf.getPage(n);
      try {
        const [content, annotations] = await Promise.all([
          page.getTextContent({ includeMarkedContent: true }),
          page.getAnnotations({ intent: 'display' }).catch(() => []),
        ]);
        const model = buildPageModel(page, content, annotations);
        totalChars += model.textLength;
        if (totalChars > MAX_TOTAL_CHARS) throw new UserError('This PDF contains too much text to convert safely.');
        pages.push(model);
      } finally {
        page.cleanup();
      }
      progressEl.value = Math.round(n * 45 / pdf.numPages);
      setStatus('Analyzing page ' + n + ' of ' + pdf.numPages + ' locally…');
    }
  } finally {
    await task.destroy();
  }
  if (!pages.some(p => p.paragraphs.some(line => line.text.trim())) && !visualAppendixInput.checked) {
    throw new UserError('No selectable text was found. This appears to be a scanned or image-only PDF. Enable the visual reference appendix to preserve its appearance; OCR is not included yet.');
  }
  return { pages };
}

function buildPageModel(page, content, annotations) {
  const viewport = page.getViewport({ scale: 1 });
  const items = [];
  for (const item of content.items) {
    if (!item || typeof item.str !== 'string' || !item.str || !item.transform) continue;
    const [x, y] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
    const fontSize = Math.max(1, Math.hypot(item.transform[2], item.transform[3]));
    const style = content.styles?.[item.fontName] || {};
    items.push({
      text: item.str, x, y, width: Math.max(0, item.width || 0),
      height: Math.max(fontSize, item.height || 0), fontSize,
      fontName: style.fontFamily || item.fontName || 'Arial',
      bold: /bold|black|heavy|demi|semibold/i.test(item.fontName || ''),
      italic: /italic|oblique/i.test(item.fontName || ''),
      hasEOL: Boolean(item.hasEOL),
    });
  }
  // Add visible form text when PDF.js provides a text widget value.
  for (const a of annotations || []) {
    if (a.subtype !== 'Widget' || !a.rect || a.fieldType === 'Btn') continue;
    let value = a.fieldValue;
    if (Array.isArray(value)) value = value.join(', ');
    if (typeof value !== 'string' || !value.trim()) continue;
    const [x, y] = viewport.convertToViewportPoint(Math.min(a.rect[0], a.rect[2]), Math.max(a.rect[1], a.rect[3]));
    value.split(/\r\n|\r|\n/).filter(Boolean).forEach((text, i) => items.push({
      text, x, y: y + 10 + i * 12, width: text.length * 5, height: 10,
      fontSize: 10, fontName: 'Arial', bold: false, italic: false, hasEOL: true,
    }));
  }
  items.sort((a, b) => a.y - b.y || a.x - b.x);
  const medianHeight = median(items.map(i => i.height).filter(Number.isFinite)) || 10;
  const tolerance = Math.max(1.5, Math.min(5, medianHeight * 0.35));
  const lines = [];
  for (const item of items) {
    let line = lines[lines.length - 1];
    if (!line || Math.abs(item.y - line.y) > tolerance) {
      line = { y: item.y, items: [] };
      lines.push(line);
    }
    line.items.push(item);
    line.y = line.items.reduce((sum, it) => sum + it.y, 0) / line.items.length;
  }
  const paragraphs = [];
  let previousY = null;
  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
    const runs = [];
    let previousEnd = null;
    for (const item of line.items) {
      const last = runs[runs.length - 1];
      if (previousEnd !== null && item.x - previousEnd > Math.max(1.2, item.fontSize * 0.18) &&
          last && !/\s$/.test(last.text) && !/^\s/.test(item.text)) {
        runs.push({ text: ' ', fontSize: item.fontSize, fontName: item.fontName });
      }
      runs.push({ text: item.text, fontSize: item.fontSize, fontName: item.fontName, bold: item.bold, italic: item.italic });
      previousEnd = item.x + item.width;
    }
    const text = runs.map(r => r.text).join('').replace(/[ \t]+/g, ' ').trim();
    if (!text) continue;
    const fontSize = median(line.items.map(i => i.fontSize)) || 11;
    const x = Math.min(...line.items.map(i => i.x));
    paragraphs.push({
      runs: trimRuns(runs), text, x, y: line.y, fontSize,
      heading: fontSize >= 16 || (fontSize >= 13 && /^[A-Z0-9 \p{P}]{3,}$/u.test(text) && text.length < 100),
      gapBefore: previousY === null ? 0 : Math.max(0, line.y - previousY),
    });
    previousY = line.y;
  }
  return {
    number: page.pageNumber,
    widthPt: viewport.width,
    heightPt: viewport.height,
    rotation: page.rotate || 0,
    paragraphs,
    textLength: paragraphs.reduce((sum, p) => sum + p.text.length, 0),
    confidence: paragraphs.length > 2 ? 0.7 : paragraphs.length ? 0.35 : 0,
    hasAnnotations: (annotations || []).length > 0,
  };
}
function median(values) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}
function trimRuns(runs) {
  const result = runs.map(r => ({ ...r }));
  while (result.length && !result[0].text.trim()) result.shift();
  while (result.length && !result[result.length - 1].text.trim()) result.pop();
  if (result.length) result[0].text = result[0].text.replace(/^\s+/, '');
  if (result.length) result[result.length - 1].text = result[result.length - 1].text.replace(/\s+$/, '');
  return result.filter(r => r.text);
}

/* DOCX package builder: editable reconstruction plus an optional visual appendix. */
function xml(value) {
  let safe = '';
  for (const ch of String(value)) {
    const cp = ch.codePointAt(0);
    if (cp === 9 || cp === 10 || cp === 13 || (cp >= 32 && cp <= 0xD7FF) ||
        (cp >= 0xE000 && cp <= 0xFFFD) || (cp >= 0x10000 && cp <= 0x10FFFF)) safe += ch;
  }
  return safe.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function twips(points) { return Math.max(1, Math.round(points * 20)); }
function runXml(run) {
  const props = [];
  if (run.bold) props.push('<w:b/>');
  if (run.italic) props.push('<w:i/>');
  if (Number.isFinite(run.fontSize)) {
    const size = Math.max(2, Math.min(144, Math.round(run.fontSize * 2)));
    props.push('<w:sz w:val="' + size + '"/><w:szCs w:val="' + size + '"/>');
  }
  if (run.fontName) {
    const font = xml(String(run.fontName).slice(0, 80));
    props.push('<w:rFonts w:ascii="' + font + '" w:hAnsi="' + font + '" w:eastAsia="' + font + '"/>');
  }
  return '<w:r>' + (props.length ? '<w:rPr>' + props.join('') + '</w:rPr>' : '') +
    '<w:t xml:space="preserve">' + xml(run.text) + '</w:t></w:r>';
}
function paragraphXml(p) {
  const props = [];
  if (p.heading) props.push('<w:pStyle w:val="Heading1"/>');
  if (p.x > 3) props.push('<w:ind w:left="' + twips(Math.min(p.x, 240)) + '"/>');
  if (p.gapBefore > 3) props.push('<w:spacing w:before="' + twips(Math.min(p.gapBefore * 0.4, 36)) + '"/>');
  return '<w:p>' + (props.length ? '<w:pPr>' + props.join('') + '</w:pPr>' : '') +
    p.runs.map(runXml).join('') + '</w:p>';
}
function imageParagraph(image) {
  const maxW = 6.5 * 914400, maxH = 9 * 914400, ratio = image.widthPx / image.heightPx;
  let width = maxW, height = width / ratio;
  if (height > maxH) { height = maxH; width = height * ratio; }
  const cx = Math.round(width), cy = Math.round(height), id = image.relationshipId;
  return '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + image.pageNumber + '" name="PDF page ' + image.pageNumber + ' reference"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="' + image.pageNumber + '" name="Page ' + image.pageNumber + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="' + id + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
}
function documentXml(model, images) {
  const body = [];
  model.pages.forEach((page, index) => {
    if (!page.paragraphs.length) {
      body.push('<w:p><w:r><w:rPr><w:i/><w:color w:val="777777"/></w:rPr><w:t>[No extractable text on source page ' + page.number + ']</w:t></w:r></w:p>');
    } else {
      for (const p of page.paragraphs) body.push(paragraphXml(p));
    }
    if (index < model.pages.length - 1) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  });
  if (images.length) {
    body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    body.push('<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Visual reference appendix</w:t></w:r></w:p>');
    body.push('<w:p><w:r><w:t>These rendered page images preserve source appearance. They are reference images, not editable text.</w:t></w:r></w:p>');
    images.forEach((image, index) => {
      body.push('<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Original PDF page ' + image.pageNumber + '</w:t></w:r></w:p>');
      body.push(imageParagraph(image));
      if (index < images.length - 1) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    });
  }
  const firstPage = model.pages[0] || { widthPt: 612, heightPt: 792 };
  const pageW = twips(firstPage.widthPt);
  const pageH = twips(firstPage.heightPt);
  const orientation = pageW > pageH ? ' w:orient="landscape"' : '';
  body.push('<w:sectPr><w:pgSz w:w="' + pageW + '" w:h="' + pageH + '"' + orientation + '/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/></w:sectPr>');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
    body.join('') + '</w:body></w:document>';
}
const STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="30"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style></w:styles>';
const SETTINGS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/><w:compat/></w:settings>';
const ROOT_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
function contentTypes() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>';
}
async function buildDocx(model, includeAppendix, progress) {
  const images = [];
  if (includeAppendix) {
    const task = pdfjsLib.getDocument({
      data: sourcePdfBytes.slice(),
      cMapUrl: new URL('./vendor/cmaps/', import.meta.url).href,
      cMapPacked: true, isEvalSupported: false, enableXfa: false, useSystemFonts: true,
    });
    try {
      const pdf = await task.promise;
      let imageBytes = 0;
      for (let i = 0; i < model.pages.length; i++) {
        progress(45 + Math.round(i * 45 / model.pages.length), 'Rendering visual reference page ' + (i + 1) + ' of ' + model.pages.length + ' locally…');
        const page = await pdf.getPage(i + 1);
        try {
          const rendered = await renderPageToPng(page, 110);
          imageBytes += rendered.blob.size;
          if (imageBytes > MAX_OUTPUT_BYTES) throw new UserError('Visual reference images exceed the output budget. Disable the appendix and try again.');
          images.push({ pageNumber: i + 1, ...rendered, relationshipId: 'rIdImage' + (i + 1) });
        } finally {
          page.cleanup();
        }
      }
    } finally {
      await task.destroy();
    }
  }
  const entries = [
    { name: '[Content_Types].xml', data: encoder.encode(contentTypes()) },
    { name: '_rels/.rels', data: encoder.encode(ROOT_RELS) },
    { name: 'word/document.xml', data: encoder.encode(documentXml(model, images)) },
    { name: 'word/styles.xml', data: encoder.encode(STYLES_XML) },
    { name: 'word/settings.xml', data: encoder.encode(SETTINGS_XML) },
  ];
  const rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'];
  for (const image of images) {
    rels.push('<Relationship Id="' + image.relationshipId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/page-' + image.pageNumber + '.png"/>');
    entries.push({ name: 'word/media/page-' + image.pageNumber + '.png', data: new Uint8Array(await image.blob.arrayBuffer()) });
  }
  rels.push('</Relationships>');
  entries.push({ name: 'word/_rels/document.xml.rels', data: encoder.encode(rels.join('')) });
  progress(93, 'Packaging the editable document and visual reference appendix…');
  return zipStore(entries);
}
async function renderPageToPng(page, dpi = 110) {
  let viewport = page.getViewport({ scale: dpi / 72 });
  const pixels = viewport.width * viewport.height;
  if (pixels > MAX_RENDER_PIXELS) viewport = page.getViewport({ scale: (dpi / 72) * Math.sqrt(MAX_RENDER_PIXELS / pixels) });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  const canvasContext = canvas.getContext('2d', { alpha: false });
  if (!canvasContext) throw new UserError('Your browser could not create a canvas for the visual appendix. Try another browser or disable the appendix.');
  try {
    await page.render({ canvasContext, viewport }).promise;
    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob(
        value => value ? resolve(value) : reject(new UserError('The browser could not encode a rendered PDF page as PNG. Try a smaller PDF or disable the visual appendix.')),
        'image/png'
      );
    });
    return { blob, widthPx: canvas.width, heightPx: canvas.height };
  } catch (error) {
    if (error instanceof UserError) throw error;
    if (error?.name === 'RenderingCancelledException') throw error;
    if (error?.name === 'PasswordException') throw new UserError('The PDF became unavailable while rendering the visual appendix.');
    throw new UserError('Could not render PDF page locally for the visual appendix. Try a smaller PDF, or disable the appendix. Details: ' + String(error?.message || error).slice(0, 180));
  } finally {
    canvas.width = 0;
    canvas.height = 0;
  }
}
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c >>> 0;
  }
  return table;
})();
function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function zipStore(entries) {
  const chunks = [], central = [];
  let offset = 0;
  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const data = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
    const crc = crc32(data), size = data.length;
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
    local.setUint16(8, 0, true); local.setUint16(10, 0, true); local.setUint16(12, 0x21, true);
    local.setUint32(14, crc, true); local.setUint32(18, size, true); local.setUint32(22, size, true);
    local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
    chunks.push(new Uint8Array(local.buffer), name, data);
    const header = new DataView(new ArrayBuffer(46));
    header.setUint32(0, 0x02014b50, true); header.setUint16(4, 20, true); header.setUint16(6, 20, true);
    header.setUint16(8, 0x0800, true); header.setUint16(10, 0, true); header.setUint16(12, 0, true);
    header.setUint16(14, 0x21, true); header.setUint32(16, crc, true); header.setUint32(20, size, true);
    header.setUint32(24, size, true); header.setUint16(28, name.length, true); header.setUint16(30, 0, true);
    header.setUint16(32, 0, true); header.setUint16(34, 0, true); header.setUint16(36, 0, true);
    header.setUint32(38, 0, true); header.setUint32(42, offset, true);
    central.push(new Uint8Array(header.buffer), name);
    offset += 30 + name.length + size;
    if (offset > 0xFFFFFFFF) throw new UserError('The document exceeds the ZIP format size limit.');
  }
  const centralSize = central.reduce((sum, c) => sum + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(4, 0, true); end.setUint16(6, 0, true);
  end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, centralSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
  return new Blob([...chunks, ...central, new Uint8Array(end.buffer)], { type: DOCX_MIME });
}
function outputName(original) {
  const base = String(original || '').replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 80);
  return (base || 'Converted_Document') + '.docx';
}
function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url; anchor.download = filename;
  document.body.appendChild(anchor); anchor.click(); anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

const offlineStatusEl = document.getElementById('offline-status');
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register(new URL('./service-worker.js', import.meta.url), { scope: './' })
    .then(async registration => {
      await navigator.serviceWorker.ready;
      if (offlineStatusEl) offlineStatusEl.textContent = 'Offline app shell cached. PDF conversion runs locally; assets first used online are cached for later offline use.';
      registration.update().catch(() => {});
    })
    .catch(() => {
      if (offlineStatusEl) offlineStatusEl.textContent = 'Offline caching is unavailable in this browser/session. Conversion still runs locally while the app is loaded.';
    });
} else if (offlineStatusEl) {
  offlineStatusEl.textContent = 'Service workers are unavailable here. Conversion still runs locally while the app is loaded.';
}
