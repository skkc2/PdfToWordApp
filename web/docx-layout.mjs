const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PICTURE_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const SHAPE_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_POINT = 12700;

export function escapeXml(value) {
  let safe = "";
  for (const ch of String(value ?? "")) {
    const cp = ch.codePointAt(0);
    if (cp === 9 || cp === 10 || cp === 13 ||
        (cp >= 32 && cp <= 0xD7FF) ||
        (cp >= 0xE000 && cp <= 0xFFFD) ||
        (cp >= 0x10000 && cp <= 0x10FFFF)) safe += ch;
  }
  return safe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

export function pointsToTwips(points) {
  return Math.max(1, Math.round(Number(points) * 20));
}

export function pointsToEmu(points) {
  return Math.max(1, Math.round(Number(points) * EMU_PER_POINT));
}

function safeFont(value) {
  return escapeXml(String(value || "Arial").replace(/[\u0000-\u001f]/g, "").slice(0, 80));
}

function safeHexColor(value) {
  const color = String(value || "000000").replace(/^#/, "");
  return /^[0-9a-fA-F]{6}$/.test(color) ? color.toUpperCase() : "000000";
}

function runXml(item) {
  const props = [];
  const font = safeFont(item.fontName);
  props.push('<w:rFonts w:ascii="' + font + '" w:hAnsi="' + font + '" w:eastAsia="' + font + '" w:cs="' + font + '"/>');
  const halfPoints = Math.max(2, Math.min(144, Math.round((Number(item.fontSize) || 11) * 2)));
  props.push('<w:sz w:val="' + halfPoints + '"/><w:szCs w:val="' + halfPoints + '"/>');
  props.push('<w:color w:val="' + safeHexColor(item.color) + '"/>');
  if (item.bold) props.push("<w:b/><w:bCs/>");
  if (item.italic) props.push("<w:i/><w:iCs/>");
  if (item.hidden) props.push("<w:vanish/><w:vanishCs/>");
  const text = escapeXml(item.text).replace(/\r\n|\r|\n/g, "</w:t><w:br/><w:t xml:space=\"preserve\">");
  return "<w:r><w:rPr>" + props.join("") + '</w:rPr><w:t xml:space="preserve">' + text + "</w:t></w:r>";
}

/**
 * An anchored DrawingML textbox: its geometry is fixed to the original PDF
 * page, but its run remains normal editable Word text.
 */
export function textBoxAnchorXml(item, docPrId, pageNumber) {
  const x = Number.isFinite(item.x) ? item.x : 0;
  const y = Number.isFinite(item.top) ? item.top : 0;
  const fontSize = Math.max(1, Number(item.fontSize) || 11);
  const widthPt = Math.max(1.5, Number(item.width) || fontSize * 0.5);
  const heightPt = Math.max(fontSize * 1.25, Number(item.height) || fontSize * 1.25);
  const cx = pointsToEmu(widthPt);
  const cy = pointsToEmu(heightPt);
  const px = Math.round(x * EMU_PER_POINT);
  const py = Math.round(y * EMU_PER_POINT);
  const fontLineTwips = Math.max(2, pointsToTwips(fontSize * 1.12));
  const rotation = Number.isFinite(item.rotation) ? Math.round(item.rotation * 60000) : 0;
  const itemName = escapeXml(("Text p" + pageNumber + " i" + item.index).slice(0, 240));
  const paragraphProps =
    '<w:pPr><w:spacing w:before="0" w:after="0" w:line="' + fontLineTwips + '" w:lineRule="exact"/>' +
    '<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>';
  const bodyPr = '<wps:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:normAutofit fontScale="100000" lnSpcReduction="0"/></wps:bodyPr>';
  return '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' +
    String(1000 + docPrId) + '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>' + px + '</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="page"><wp:posOffset>' + py + '</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
    '<wp:docPr id="' + docPrId + '" name="' + itemName + '"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="' + SHAPE_NS + '"><wps:wsp>' +
    '<wps:cNvSpPr txBox="1"/><wps:spPr bwMode="auto"><a:xfrm' + (rotation ? ' rot="' + rotation + '"' : '') + '><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>' +
    '<wps:txbx><w:txbxContent><w:p>' + paragraphProps + runXml(item) + '</w:p></w:txbxContent></wps:txbx>' +
    bodyPr + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

/**
 * A full-page raster layer. For ordinary digital pages the renderer filters out
 * all text-paint operators, retaining artwork, raster images, rules, borders,
 * lines and other non-text objects. A special-case page can instead use an
 * unfiltered, complete-page image.
 */
export function pageImageAnchorXml(image, page, docPrId) {
  const cx = pointsToEmu(page.widthPt);
  const cy = pointsToEmu(page.heightPt);
  const x = 0;
  const y = 0;
  const behindDoc = image.fullPage ? "0" : "1";
  const name = escapeXml((image.fullPage ? "Rendered PDF page " : "PDF page artwork ") + page.number);
  return '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' +
    String(behindDoc === "1" ? 0 : 500) + '" behindDoc="' + behindDoc + '" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>' + x + '</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="page"><wp:posOffset>' + y + '</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
    '<wp:docPr id="' + docPrId + '" name="' + name + '"/>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="' + docPrId + '" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="' + escapeXml(image.relationshipId) + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
}

export function sectionPropertiesXml(page, nextPage = false) {
  const width = pointsToTwips(page.widthPt);
  const height = pointsToTwips(page.heightPt);
  const orientation = width > height ? ' w:orient="landscape"' : "";
  return '<w:sectPr>' + (nextPage ? '<w:type w:val="nextPage"/>' : "") +
    '<w:pgSz w:w="' + width + '" w:h="' + height + '"' + orientation + '/>' +
    '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/>' +
    '</w:sectPr>';
}

export function pageParagraphXml(page, image, options = {}) {
  let docPrId = options.startDocPrId || 1;
  const parts = [pageImageAnchorXml(image, page, docPrId++)];
  if (!image.fullPage) {
    for (const item of page.textItems || []) {
      parts.push(textBoxAnchorXml(item, docPrId++, page.number));
    }
  }
  const sectPr = options.sectionProperties || "";
  const pPr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>' +
    '<w:ind w:left="0" w:right="0" w:firstLine="0"/>' + sectPr + '</w:pPr>';
  return { xml: '<w:p>' + pPr + parts.join("") + '</w:p>', nextDocPrId: docPrId };
}
