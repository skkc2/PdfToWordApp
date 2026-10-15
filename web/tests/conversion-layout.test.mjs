import test from "node:test";
import assert from "node:assert/strict";
import { choosePageRepresentation } from "../conversion-layout.mjs";

const simple = {
  textLength: 520,
  textItemCount: 36,
  rasterImageOps: 0,
  vectorPaintOps: 3,
  wideGapLineCount: 0,
  operatorListUnavailable: false,
};

test("hybrid keeps a structurally simple text page editable", () => {
  assert.equal(choosePageRepresentation("hybrid", simple).representation, "text");
});

test("visual mode uses an image for every page", () => {
  assert.equal(choosePageRepresentation("visual", simple).representation, "image");
});

test("editable mode never swaps page text for an image", () => {
  assert.equal(choosePageRepresentation("editable", {
    ...simple, rasterImageOps: 2, vectorPaintOps: 100, wideGapLineCount: 5,
  }).representation, "text");
});

test("hybrid uses an image for image-only/scanned pages", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, textLength: 0, rasterImageOps: 1,
  }).representation, "image");
});

test("hybrid uses an image when raster graphics are present", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, rasterImageOps: 1,
  }).reason, "page contains raster graphics");
});

test("hybrid uses an image when repeated wide gaps suggest columns or a table", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, wideGapLineCount: 3,
  }).representation, "image");
});

test("hybrid uses an image when vector drawing density is high", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, vectorPaintOps: 24,
  }).representation, "image");
});

test("hybrid fails safe to an image if structure analysis fails", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, operatorListUnavailable: true,
  }).representation, "image");
});

test("unknown modes are rejected", () => {
  assert.throws(() => choosePageRepresentation("appendix", simple), TypeError);
});

test("hybrid falls back when text spans almost the full page height", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, pageHeightPt: 792, textVerticalSpan: 770,
  }).representation, "image");
});

test("hybrid falls back when text reaches the page edge", () => {
  assert.equal(choosePageRepresentation("hybrid", {
    ...simple, pageWidthPt: 612, maxTextRightPt: 610,
  }).representation, "image");
});
