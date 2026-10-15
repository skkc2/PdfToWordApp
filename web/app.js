import * as pdfjsLib from './vendor/pdf.min.mjs';
import { escapeXml, pageParagraphXml, sectionPropertiesXml } from './docx-layout.mjs';

const MAX_PDF_BYTES = 30 * 1024 * 1024;
const MAX_PAGES = 300;
const MAX_TOTAL_CHARS = 3000000;
const MAX_TEXT_ITEMS = 90000;
const MAX_ITEMS_PER_PAGE = 6000;
const MAX_RENDER_PIXELS = 12000000;
const MAX_OUTPUT_BYTES = 180 * 1024 * 1024;
const MAX_RENDERED_IMAGE_BYTES = 150 * 1024 * 1024;
const RENDER_DPI = 144;
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const encoder = new TextEncoder();
const OPS = pdfjsLib.OPS || {};
let sourcePdfBytes = null;
let busy = false;
let currentStage = 'startup';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdf.worker.min.mjs', import.meta.url).href;

const fileInput = document.getElementById('file');
const convertBtn = document.getElementById('convert');
const progressEl = document.getElementById('progress');
const statusEl = document.getElementById('status');
const offlineStatusEl = document.getElementById('offline-status');

class UserError extends Error {}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle('error', isError);
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files && fileInput.files[0];
  convertBtn.disabled = !file || busy;
  setStatus(file ? 'PDF selected. Ready for on-device conversion.' : 'Select a PDF to begin.');
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
  setStatus('Starting local conversion…');

  try {
    currentStage = 'PDF analysis';
    const model = await analyzePdf(file);
    currentStage = 'graphics-layer rendering and editable text placement';
    const blob = await buildDocx(model, (percent, message) => {
      progressEl.value = percent;
      if (message.toLowerCase().includes('rendering source page')) currentStage = 'page rendering';
      else if (message.toLowerCase().includes('building editable text')) currentStage = 'DOCX layout generation';
      else if (message.toLowerCase().includes('packaging')) currentStage = 'DOCX packaging';
      setStatus(message);
    });

    if (blob.size > MAX_OUTPUT_BYTES) {
      throw new UserError('The generated Word file exceeds the 180 MB safety limit. Try a smaller PDF.');
    }
    saveBlob(blob, outputName(file.name));
    const fallbackCount = model.pages.filter(page => page.fullPageFallback).length;
    setStatus(fallbackCount
      ? 'Done. Text is editable on most pages; ' + fallbackCount + ' complex page(s) were preserved as page images to avoid damaging their layout.'
      : 'Done. Your Word document has an editable text layer positioned over the preserved page artwork.');
  } catch (error) {
    setStatus(friendlyMessage(error), true);
  } finally {
    sourcePdfBytes = null;
    busy = false;
    fileInput.disabled = false;
    convertBtn.disabled = !(fileInput.files && fileInput.files[0]);
    progressEl.hidden = true;
  }
});

function friendlyMessage(error) {
  if (error instanceof UserError) return error.message;
  if (error?.name === 'PasswordException') return "This PDF is password-protected and can't be converted.";
  if (error?.name === 'InvalidPDFException' || error?.name === 'FormatError') return 'Could not read the PDF. It may be damaged or not a valid PDF.';
  if (error instanceof RangeError) return 'Not enough memory. Try a smaller PDF.';
  const detail = String(error?.message || error || 'Unknown error').replace(/\s+/g, ' ').slice(0, 240);
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

function pdfLoadOptions(bytes) {
  return {
    data: bytes.slice(),
    cMapUrl: new URL('./vendor/cmaps/', import.meta.url).href,
    cMapPacked: true,
    standardFontDataUrl: new URL('./vendor/standard_fonts/', import.meta.url).href,
    wasmUrl: new URL('./vendor/wasm/', import.meta.url).href,
    iccUrl: new URL('./vendor/iccs/', import.meta.url).href,
    useWasm: true,
    isImageDecoderSupported: false,
    isEvalSupported: false,
    enableXfa: false,
    useSystemFonts: true,
  };
}

async function analyzePdf(file) {
  if (file.size > MAX_PDF_BYTES) throw new UserError('This PDF is larger than the 30 MB limit.');
  sourcePdfBytes = new Uint8Array(await file.arrayBuffer());
  if (!hasPdfHeader(sourcePdfBytes)) throw new UserError('This file does not look like a PDF.');

  const task = pdfjsLib.getDocument(pdfLoadOptions(sourcePdfBytes));
  const pages = [];
  let totalChars = 0;
  let totalItems = 0;
  try {
    const pdf = await task.promise;
    if (pdf.numPages > MAX_PAGES) {
      throw new UserError('This PDF has ' + pdf.numPages + ' pages. The limit is ' + MAX_PAGES + '.');
    }

    for (let number = 1; number <= pdf.numPages; number++) {
      const page = await pdf.getPage(number);
      try {
        const [content, annotations, operatorList] = await Promise.all([
          page.getTextContent({ includeMarkedContent: true }),
          page.getAnnotations({ intent: 'display' }).catch(() => []),
          page.getOperatorList(),
        ]);
        const operatorInfo = inspectOperatorList(operatorList);
        const pageModel = buildPageModel(page, content, annotations, operatorInfo);
        totalChars += pageModel.textLength;
        totalItems += pageModel.textItems.length;

        if (totalChars > MAX_TOTAL_CHARS) {
          throw new UserError('This PDF contains too much text to convert safely.');
        }
        if (totalItems > MAX_TEXT_ITEMS) {
          throw new UserError('This PDF contains too many separate text objects for safe layout reconstruction. Try a simpler or smaller PDF.');
        }
        if (pageModel.textItems.length > MAX_ITEMS_PER_PAGE) {
          pageModel.fullPageFallback = true;
          pageModel.fallbackReason = 'too many separate text objects';
        }

        pages.push(pageModel);
      } finally {
        page.cleanup();
      }

      progressEl.value = Math.round(number * 35 / pdf.numPages);
      setStatus('Analyzing page ' + number + ' of ' + pdf.numPages + ' locally…');
    }
  } finally {
    await task.destroy();
  }
  return { pages };
}

function textPaintOperatorIds() {
  return new Set([
    OPS.showText,
    OPS.showSpacedText,
    OPS.nextLineShowText,
    OPS.nextLineSetSpacingShowText,
  ].filter(Number.isInteger));
}

function imageOperatorIds() {
  return new Set([
    'paintImageXObject',
    'paintInlineImageXObject',
    'paintImageMaskXObject',
    'paintImageMaskXObjectGroup',
    'paintImageXObjectRepeat',
    'paintSolidColorImageMask',
    'paintJpegXObject',
  ].map(name => OPS[name]).filter(Number.isInteger));
}

function vectorOperatorIds() {
  return new Set([
    'constructPath', 'stroke', 'closeStroke', 'fill', 'eoFill',
    'fillStroke', 'eoFillStroke', 'shadingFill', 'paintFormXObjectBegin',
  ].map(name => OPS[name]).filter(Number.isInteger));
}

function normalizeColor(args, fallback) {
  const first = args && args.length ? args[0] : null;
  if (typeof first === 'string') {
    const match = first.match(/#?([0-9a-f]{6})\b/i);
    if (match) return match[1].toUpperCase();
    const rgb = first.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
    if (rgb) return [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])].map(n => Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  let values = Array.isArray(first) || ArrayBuffer.isView(first) ? Array.from(first) : args;
  if (values && values.length >= 3 && values.slice(0, 3).every(Number.isFinite)) {
    const rgb = values.slice(0, 3).map(value => {
      const normalized = value <= 1 ? value * 255 : value;
      return Math.max(0, Math.min(255, Math.round(normalized))).toString(16).padStart(2, '0');
    });
    return rgb.join('').toUpperCase();
  }
  return fallback;
}

function textFromOperatorArgs(args) {
  if (!Array.isArray(args)) return '';
  for (let i = args.length - 1; i >= 0; i--) {
    const candidate = args[i];
    if (!Array.isArray(candidate)) continue;
    if (!candidate.some(glyph => glyph && typeof glyph === 'object' && typeof glyph.unicode === 'string')) continue;
    return candidate.map(glyph => glyph && typeof glyph === 'object' && typeof glyph.unicode === 'string' ? glyph.unicode : '').join('');
  }
  return '';
}

/**
 * Read rendering state from PDF.js's display operator list. The resulting text
 * stream is only used to align color/rendering-mode hints with getTextContent;
 * page geometry and the actual background render still come from PDF.js.
 */
function inspectOperatorList(operatorList) {
  const fnArray = operatorList?.fnArray || [];
  const argsArray = operatorList?.argsArray || [];
  const textIds = textPaintOperatorIds();
  const imageIds = imageOperatorIds();
  const vectorIds = vectorOperatorIds();
  const fillIds = new Set(['setFillRGBColor', 'setFillColor'].map(name => OPS[name]).filter(Number.isInteger));
  const strokeIds = new Set(['setStrokeRGBColor', 'setStrokeColor'].map(name => OPS[name]).filter(Number.isInteger));
  const saveId = OPS.save;
  const restoreId = OPS.restore;
  const setTextModeId = OPS.setTextRenderingMode;

  let fillColor = '000000';
  let strokeColor = '000000';
  let textMode = 0;
  const stateStack = [];
  const textPaintRecords = [];
  let rasterImageOps = 0;
  let vectorPaintOps = 0;
  let invisibleTextOpCount = 0;
  let hasTextClipping = false;

  for (let i = 0; i < fnArray.length; i++) {
    const op = fnArray[i];
    const args = argsArray[i] || [];

    if (op === saveId) {
      stateStack.push({ fillColor, strokeColor, textMode });
    } else if (op === restoreId) {
      const state = stateStack.pop();
      if (state) {
        fillColor = state.fillColor;
        strokeColor = state.strokeColor;
        textMode = state.textMode;
      }
    } else if (fillIds.has(op)) {
      fillColor = normalizeColor(args, fillColor);
    } else if (strokeIds.has(op)) {
      strokeColor = normalizeColor(args, strokeColor);
    } else if (op === setTextModeId) {
      textMode = Number(args[0]) || 0;
    }

    if (imageIds.has(op)) rasterImageOps++;
    if (vectorIds.has(op)) vectorPaintOps++;

    if (textIds.has(op)) {
      if (textMode === 3) invisibleTextOpCount++;
      if (textMode >= 4 && textMode <= 7) hasTextClipping = true;
      const text = textFromOperatorArgs(args);
      const color = textMode === 1 || textMode === 5 ? strokeColor : fillColor;
      textPaintRecords.push({ text, color, mode: textMode });
    }
  }

  return {
    textPaintRecords,
    rasterImageOps,
    vectorPaintOps,
    invisibleTextOpCount,
    hasTextClipping,
    operatorListUnavailable: !fnArray.length && !argsArray.length,
  };
}

function normalizedCharacters(value) {
  return Array.from(String(value || '').normalize('NFKC').toLocaleLowerCase()).filter(ch => !/\s/u.test(ch));
}

function findSequence(haystack, needle, start) {
  if (!needle.length || needle.length > haystack.length) return -1;
  const last = haystack.length - needle.length;
  for (let i = Math.max(0, start); i <= last; i++) {
    let matched = true;
    for (let j = 0; j < needle.length; j++) {
      if (haystack[i + j] !== needle[j]) { matched = false; break; }
    }
    if (matched) return i;
  }
  return -1;
}

function stylesForTextItems(items, textPaintRecords) {
  const chars = [];
  for (const record of textPaintRecords) {
    for (const ch of normalizedCharacters(record.text)) chars.push({ ch, color: record.color, mode: record.mode });
  }
  const textStream = chars.map(entry => entry.ch);
  let cursor = 0;
  let previous = { color: '000000', mode: 0 };
  return items.map(item => {
    const needle = normalizedCharacters(item.str);
    let at = findSequence(textStream, needle, cursor);
    if (at < 0) at = findSequence(textStream, needle, Math.max(0, cursor - 128));
    if (at < 0) {
      return { color: previous.color, mode: 0 };
    }
    const style = { color: chars[at]?.color || '000000', mode: chars[at]?.mode ?? 0 };
    cursor = at + needle.length;
    previous = style;
    return style;
  });
}

function buildPageModel(page, content, annotations, operatorInfo) {
  const viewport = page.getViewport({ scale: 1 });
  const sourceItems = (content.items || []).filter(item =>
    item && typeof item.str === 'string' && item.str.length > 0 && Array.isArray(item.transform)
  );
  const styleHints = stylesForTextItems(sourceItems, operatorInfo.textPaintRecords);
  const textItems = [];
  let hasUnsupportedTextTransform = false;

  sourceItems.forEach((item, index) => {
    const transform = item.transform;
    const [x, baselineY] = viewport.convertToViewportPoint(transform[4], transform[5]);
    const fontSize = Math.max(1, Math.hypot(transform[0], transform[1]));
    const style = content.styles?.[item.fontName] || {};
    const ascent = Number.isFinite(style.ascent) ? Math.min(1.25, Math.max(0.45, style.ascent)) : 0.8;
    const rotation = Math.atan2(transform[1], transform[0]) * 180 / Math.PI;
    const skewed = Math.abs(transform[2]) > Math.max(0.1, fontSize * 0.02);
    const rtl = item.dir && item.dir !== 'ltr';

    if (Math.abs(rotation) > 1.2 || skewed || rtl || transform[0] < 0) hasUnsupportedTextTransform = true;

    const fontName = style.fontFamily || 'Arial';
    const combinedFontLabel = String(fontName) + ' ' + String(item.fontName || '');
    const styleHint = styleHints[index] || { color: '000000', mode: 0 };
    textItems.push({
      index,
      text: item.str,
      x,
      top: baselineY - fontSize * ascent,
      baselineY,
      width: Math.max(1.5, Math.abs(Number(item.width) || 0)),
      height: Math.max(fontSize * 1.2, Math.abs(Number(item.height) || 0)),
      fontSize,
      fontName,
      color: styleHint.color,
      hidden: styleHint.mode === 3,
      bold: /bold|black|heavy|demi|semibold/i.test(combinedFontLabel),
      italic: /italic|oblique/i.test(combinedFontLabel),
      rotation,
      dir: item.dir || 'ltr',
    });
  });

  // Widget annotations can contain real editable field values. They are mapped
  // separately because PDF.js does not necessarily emit form values as text items.
  for (const annotation of annotations || []) {
    if (annotation.subtype !== 'Widget' || !annotation.rect || annotation.fieldType === 'Btn') continue;
    let value = annotation.fieldValue;
    if (Array.isArray(value)) value = value.join(', ');
    if (typeof value !== 'string' || !value.trim()) continue;
    const rect = annotation.rect;
    const [x, top] = viewport.convertToViewportPoint(Math.min(rect[0], rect[2]), Math.max(rect[1], rect[3]));
    value.split(/\r\n|\r|\n/).filter(Boolean).forEach((text, lineIndex) => {
      textItems.push({
        index: textItems.length,
        text, x, top: top + lineIndex * 12, baselineY: top + lineIndex * 12 + 10,
        width: Math.max(12, text.length * 5), height: 12, fontSize: 10,
        fontName: 'Arial', color: '000000', hidden: false, bold: false, italic: false,
        rotation: 0, dir: 'ltr',
      });
    });
  }

  const textLength = textItems.reduce((sum, item) => sum + item.text.length, 0);
  let fallbackReason = '';
  if (page.rotate && page.rotate % 360 !== 0) fallbackReason = 'rotated PDF page';
  else if (operatorInfo.hasTextClipping) fallbackReason = 'text used as a clipping path';
  else if (operatorInfo.rasterImageOps > 0 && operatorInfo.invisibleTextOpCount > 0) fallbackReason = 'a scanned page has an invisible OCR layer';
  else if (hasUnsupportedTextTransform) fallbackReason = 'rotated, skewed, or right-to-left text';
  else if (operatorInfo.textPaintRecords.length > 0 && sourceItems.length === 0) fallbackReason = 'PDF.js could not map drawn text to editable text items';
  else if (operatorInfo.operatorListUnavailable && sourceItems.length > 0) fallbackReason = 'PDF drawing instructions could not be analyzed';
  else if (textItems.length > MAX_ITEMS_PER_PAGE) fallbackReason = 'too many separate text objects';

  return {
    number: page.pageNumber,
    widthPt: viewport.width,
    heightPt: viewport.height,
    rotation: page.rotate || 0,
    textItems,
    textLength,
    operatorInfo: {
      rasterImageOps: operatorInfo.rasterImageOps,
      vectorPaintOps: operatorInfo.vectorPaintOps,
      invisibleTextOpCount: operatorInfo.invisibleTextOpCount,
      hasTextClipping: operatorInfo.hasTextClipping,
    },
    fullPageFallback: Boolean(fallbackReason),
    fallbackReason,
  };
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
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 1);
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
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
    local.setUint16(8, 0, true); local.setUint16(10, 0, true); local.setUint16(12, 0x21, true);
    local.setUint32(14, crc, true); local.setUint32(18, size, true); local.setUint32(22, size, true);
    local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
    chunks.push(new Uint8Array(local.buffer), name, data);

    const header = new DataView(new ArrayBuffer(46));
    header.setUint32(0, 0x02014b50, true);
    header.setUint16(4, 20, true); header.setUint16(6, 20, true);
    header.setUint16(8, 0x0800, true); header.setUint16(10, 0, true); header.setUint16(12, 0, true);
    header.setUint16(14, 0x21, true); header.setUint32(16, crc, true);
    header.setUint32(20, size, true); header.setUint32(24, size, true);
    header.setUint16(28, name.length, true); header.setUint16(30, 0, true);
    header.setUint16(32, 0, true); header.setUint16(34, 0, true); header.setUint16(36, 0, true);
    header.setUint32(38, 0, true); header.setUint32(42, offset, true);
    central.push(new Uint8Array(header.buffer), name);
    offset += 30 + name.length + size;
    if (offset > 0xFFFFFFFF) throw new UserError('The generated document exceeds the ZIP file-size limit.');
  }

  const centralSize = central.reduce((sum, chunk) => sum + chunk.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(4, 0, true); end.setUint16(6, 0, true);
  end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, centralSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
  return new Blob([...chunks, ...central, new Uint8Array(end.buffer)], { type: DOCX_MIME });
}

const STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '</w:styles>';
const SETTINGS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/><w:compat/></w:settings>';
const ROOT_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

function contentTypes() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '</Types>';
}

function documentXml(model, images) {
  const body = [];
  const imageByPage = new Map(images.map(image => [image.pageNumber, image]));
  let docPrId = 1;

  model.pages.forEach((page, index) => {
    const image = imageByPage.get(page.number);
    const pageXml = pageParagraphXml(page, image, {
      startDocPrId: docPrId,
      sectionProperties: index < model.pages.length - 1 ? sectionPropertiesXml(page, true) : '',
    });
    docPrId = pageXml.nextDocPrId;
    body.push(pageXml.xml);
  });

  const lastPage = model.pages[model.pages.length - 1] || { widthPt: 612, heightPt: 792 };
  body.push(sectionPropertiesXml(lastPage, false));

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"' +
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<w:body>' + body.join('') + '</w:body></w:document>';
}

async function renderPageToPng(page, modelPage, requestedGraphicsOnly = true) {
  let viewport = page.getViewport({ scale: RENDER_DPI / 72 });
  const initialPixels = viewport.width * viewport.height;
  if (initialPixels > MAX_RENDER_PIXELS) {
    viewport = page.getViewport({ scale: (RENDER_DPI / 72) * Math.sqrt(MAX_RENDER_PIXELS / initialPixels) });
  }

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(viewport.width));
  canvas.height = Math.max(1, Math.ceil(viewport.height));
  const context = canvas.getContext('2d', { alpha: false });
  if (!context) throw new UserError('Your browser could not create the canvas needed to preserve the page layout.');

  try {
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);

    let fullPage = !requestedGraphicsOnly;
    let filterCalls = 0;
    let filterReceivedOperatorList = false;
    let mappedRenderTextStyles = false;
    const renderOptions = { canvasContext: context, viewport };

    if (requestedGraphicsOnly) {
      renderOptions.operationsFilter = (index, renderOperatorList) => {
        filterCalls++;
        if (!renderOperatorList || !Array.isArray(renderOperatorList.fnArray)) {
          return true;
        }
        filterReceivedOperatorList = true;
        if (!mappedRenderTextStyles) {
          mappedRenderTextStyles = true;
          // Use the exact list being painted. A separately requested getOperatorList()
          // can differ from the renderer's list because PDF.js uses different intents.
          const renderInfo = inspectOperatorList(renderOperatorList);
          const hints = stylesForTextItems(modelPage.textItems, renderInfo.textPaintRecords);
          modelPage.textItems.forEach((item, itemIndex) => {
            item.color = hints[itemIndex]?.color || item.color || '000000';
            item.hidden = hints[itemIndex]?.mode === 3 || item.hidden;
          });
        }
        return !textPaintOperatorIds().has(renderOperatorList.fnArray[index]);
      };
    }

    await page.render(renderOptions).promise;

    // If a PDF.js build/browser doesn't provide the render operator list to the
    // filter callback, never overlay editable text on top of already-rendered text.
    if (requestedGraphicsOnly && filterCalls > 0 && !filterReceivedOperatorList) {
      fullPage = true;
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: context, viewport }).promise;
    } else if (requestedGraphicsOnly && !filterCalls && modelPage.textItems.length) {
      fullPage = true;
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: context, viewport }).promise;
    }

    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob(value => value ? resolve(value) : reject(new UserError('The browser could not encode a rendered page image.')), 'image/png');
    });
    return { blob, widthPx: canvas.width, heightPx: canvas.height, fullPage };
  } catch (error) {
    if (error instanceof UserError) throw error;
    throw new UserError('Could not render PDF page ' + modelPage.number + ' locally. Details: ' + String(error?.message || error).slice(0, 180));
  } finally {
    canvas.width = 0;
    canvas.height = 0;
  }
}

async function buildDocx(model, progress) {
  const task = pdfjsLib.getDocument(pdfLoadOptions(sourcePdfBytes));
  const images = [];
  let imageBytes = 0;
  try {
    const pdf = await task.promise;
    for (let i = 0; i < model.pages.length; i++) {
      const modelPage = model.pages[i];
      progress(38 + Math.round(i * 47 / model.pages.length),
        modelPage.fullPageFallback
          ? 'Rendering source page ' + modelPage.number + ' as a visual-safe fallback (' + modelPage.fallbackReason + ')…'
          : 'Rendering source page ' + modelPage.number + ' artwork and editable text positions…');

      const page = await pdf.getPage(modelPage.number);
      try {
        const rendered = await renderPageToPng(page, modelPage, !modelPage.fullPageFallback);
        imageBytes += rendered.blob.size;
        if (imageBytes > MAX_RENDERED_IMAGE_BYTES) {
          throw new UserError('Rendered page artwork exceeds the 150 MB safety limit. Try a smaller PDF.');
        }
        images.push({
          pageNumber: modelPage.number,
          blob: rendered.blob,
          widthPx: rendered.widthPx,
          heightPx: rendered.heightPx,
          fullPage: rendered.fullPage || modelPage.fullPageFallback,
          relationshipId: 'rIdPage' + modelPage.number,
        });
      } finally {
        page.cleanup();
      }
    }
  } finally {
    await task.destroy();
  }

  progress(88, 'Building editable text boxes at their original page coordinates…');
  const entries = [
    { name: '[Content_Types].xml', data: encoder.encode(contentTypes()) },
    { name: '_rels/.rels', data: encoder.encode(ROOT_RELS) },
    { name: 'word/document.xml', data: encoder.encode(documentXml(model, images)) },
    { name: 'word/styles.xml', data: encoder.encode(STYLES_XML) },
    { name: 'word/settings.xml', data: encoder.encode(SETTINGS_XML) },
  ];

  const relationships = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
  ];
  for (const image of images) {
    relationships.push('<Relationship Id="' + image.relationshipId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/page-' + image.pageNumber + '.png"/>');
    entries.push({
      name: 'word/media/page-' + image.pageNumber + '.png',
      data: new Uint8Array(await image.blob.arrayBuffer()),
    });
  }
  relationships.push('</Relationships>');
  entries.push({ name: 'word/_rels/document.xml.rels', data: encoder.encode(relationships.join('')) });

  progress(94, 'Packaging one layout-preserving Word page for each PDF page…');
  const blob = zipStore(entries);
  if (blob.size > MAX_OUTPUT_BYTES) throw new UserError('The generated Word file exceeds the 180 MB safety limit. Try a smaller PDF.');
  return blob;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ c) : (c >>> 1);
    table[n] = c >>> 0;
  }
  return table;
})();

function outputName(original) {
  const base = String(original || '').replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 80);
  return (base || 'Converted_Document') + '.docx';
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function zipStore(entries) {
  const chunks = [], central = [];
  let offset = 0;
  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const data = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
    const crc = crc32(data), size = data.length;

    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
    local.setUint16(8, 0, true); local.setUint16(10, 0, true); local.setUint16(12, 0x21, true);
    local.setUint32(14, crc, true); local.setUint32(18, size, true); local.setUint32(22, size, true);
    local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
    chunks.push(new Uint8Array(local.buffer), name, data);

    const header = new DataView(new ArrayBuffer(46));
    header.setUint32(0, 0x02014b50, true);
    header.setUint16(4, 20, true); header.setUint16(6, 20, true);
    header.setUint16(8, 0x0800, true); header.setUint16(10, 0, true); header.setUint16(12, 0, true);
    header.setUint16(14, 0x21, true); header.setUint32(16, crc, true);
    header.setUint32(20, size, true); header.setUint32(24, size, true);
    header.setUint16(28, name.length, true); header.setUint16(30, 0, true);
    header.setUint16(32, 0, true); header.setUint16(34, 0, true); header.setUint16(36, 0, true);
    header.setUint32(38, 0, true); header.setUint32(42, offset, true);
    central.push(new Uint8Array(header.buffer), name);
    offset += 30 + name.length + size;
    if (offset > 0xFFFFFFFF) throw new UserError('The generated document exceeds the ZIP file-size limit.');
  }

  const centralSize = central.reduce((sum, chunk) => sum + chunk.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(4, 0, true); end.setUint16(6, 0, true);
  end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, centralSize, true); end.setUint32(16, offset, true); end.setUint16(20, 0, true);
  return new Blob([...chunks, ...central, new Uint8Array(end.buffer)], { type: DOCX_MIME });
}

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

const offlineStatusElLocal = document.getElementById('offline-status');
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register(new URL('./service-worker.js', import.meta.url), { scope: './' })
    .then(async registration => {
      await navigator.serviceWorker.ready;
      if (offlineStatusElLocal) offlineStatusElLocal.textContent = 'Offline app shell cached. Conversion is local; required PDF.js assets are bundled with the app.';
      registration.update().catch(() => {});
    })
    .catch(() => {
      if (offlineStatusElLocal) offlineStatusElLocal.textContent = 'Offline caching is unavailable in this browser/session. Conversion still runs locally while the app is loaded.';
    });
} else if (offlineStatusElLocal) {
  offlineStatusElLocal.textContent = 'Service workers are unavailable here. Conversion still runs locally while the app is loaded.';
}
