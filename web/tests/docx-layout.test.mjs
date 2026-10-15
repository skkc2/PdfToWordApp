import test from "node:test";
import assert from "node:assert/strict";
import {
  escapeXml,
  pointsToTwips,
  pointsToEmu,
  textBoxAnchorXml,
  pageImageAnchorXml,
  pageParagraphXml,
  sectionPropertiesXml,
} from "../docx-layout.mjs";

const page = { number: 2, widthPt: 612, heightPt: 792, textItems: [] };
const textItem = {
  index: 4, text: 'A & <B> "quoted"',
  x: 72, top: 96, width: 80, height: 14,
  fontSize: 11, fontName: 'Arial', color: '336699',
  bold: true, italic: false, hidden: false, rotation: 0,
};

test("XML escaping prevents text content from breaking DOCX markup", () => {
  assert.equal(escapeXml('&<>"\''), '&amp;&lt;&gt;&quot;&apos;');
  assert.equal(escapeXml('ok\u0000text'), 'oktext');
});

test("point conversions use Word twips and EMUs", () => {
  assert.equal(pointsToTwips(72), 1440);
  assert.equal(pointsToEmu(72), 914400);
});

test("editable text boxes use original PDF coordinates and text content", () => {
  const xml = textBoxAnchorXml(textItem, 9, 2);
  assert.match(xml, /<wp:posOffset>914400<\/wp:posOffset>/);
  assert.match(xml, /<wp:posOffset>1219200<\/wp:posOffset>/);
  assert.match(xml, /A &amp; &lt;B&gt; &quot;quoted&quot;/);
  assert.match(xml, /<w:b\/><w:bCs\/>/);
  assert.match(xml, /txBox="1"/);
  assert.match(xml, /<a:noFill\/>/);
});

test("page artwork is positioned as a page-sized image behind editable text", () => {
  const xml = pageImageAnchorXml({
    fullPage: false, relationshipId: "rIdPage2",
  }, page, 1);
  assert.match(xml, /behindDoc="1"/);
  assert.match(xml, /<wp:posOffset>0<\/wp:posOffset>/);
  assert.match(xml, /<wp:extent cx="7772400" cy="10058400"\/>/);
  assert.match(xml, /r:embed="rIdPage2"/);
});

test("page paragraph includes editable boxes and a section break without an appendix", () => {
  const result = pageParagraphXml({
    ...page, textItems: [textItem],
  }, { fullPage: false, relationshipId: "rIdPage2" }, {
    startDocPrId: 1, sectionProperties: sectionPropertiesXml(page, true),
  });
  assert.equal(result.nextDocPrId, 3);
  assert.match(result.xml, /<w:sectPr><w:type w:val="nextPage"\/>/);
  assert.match(result.xml, /A &amp; &lt;B&gt; &quot;quoted&quot;/);
  assert.doesNotMatch(result.xml, /Visual reference appendix|Original PDF page/);
});

test("full-page fallback does not add editable text over the page image", () => {
  const result = pageParagraphXml({
    ...page, textItems: [textItem],
  }, { fullPage: true, relationshipId: "rIdPage2" });
  assert.match(result.xml, /Rendered PDF page 2/);
  assert.doesNotMatch(result.xml, /A &amp; &lt;B&gt;/);
});
