/**
 * Decide which representation to use for each original PDF page.
 *
 * "hybrid" preserves source page count by using editable Word text only where
 * the page appears structurally simple; pages with graphics/layout ambiguity
 * become one full-page image in place. "visual" is image-only. "editable"
 * always chooses reconstructed text, including on pages where fidelity may suffer.
 */
export function choosePageRepresentation(mode, metrics) {
  if (!["hybrid", "editable", "visual"].includes(mode)) {
    throw new TypeError("Unknown conversion mode: " + mode);
  }

  if (mode === "visual") {
    return { representation: "image", reason: "visual-fidelity mode" };
  }
  if (mode === "editable") {
    return { representation: "text", reason: "editable-text mode" };
  }

  if (!metrics || metrics.operatorListUnavailable) {
    return { representation: "image", reason: "page layout could not be analyzed safely" };
  }
  if (!metrics.textLength) {
    return { representation: "image", reason: "no extractable text" };
  }
  if (metrics.textVerticalSpan > Math.max(1, (metrics.pageHeightPt || 792) - 36)) {
    return { representation: "image", reason: "text content spans almost the full page height" };
  }
  if (metrics.maxTextRightPt > Math.max(1, (metrics.pageWidthPt || 612) - 12)) {
    return { representation: "image", reason: "text reaches the page edge and may reflow" };
  }
  if (metrics.rasterImageOps > 0) {
    return { representation: "image", reason: "page contains raster graphics" };
  }
  if (metrics.wideGapLineCount >= 3) {
    return { representation: "image", reason: "multi-column or tabular alignment detected" };
  }
  const graphicThreshold = Math.max(24, Math.ceil((metrics.textItemCount || 0) * 0.65));
  if (metrics.vectorPaintOps >= graphicThreshold) {
    return { representation: "image", reason: "dense vector drawing detected" };
  }
  return { representation: "text", reason: "page appears suitable for editable reconstruction" };
}
