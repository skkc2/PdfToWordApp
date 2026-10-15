const assert = require("node:assert/strict");
const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const { execFileSync } = require("node:child_process");
const { chromium } = require("playwright");
const { PNG } = require("pngjs");

const root = path.resolve(process.env.SITE_DIR || path.join(__dirname, "../../site"));
const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".wasm": "application/wasm",
  ".png": "image/png",
  ".bcmap": "application/octet-stream",
  ".icc": "application/octet-stream",
  ".pfb": "application/octet-stream",
  ".ttf": "application/octet-stream",
};

function makeFixturePdf() {
  const stream = "BT /F1 18 Tf 72 700 Td (HELLO WORLD) Tj ET\n1 0 0 rg 72 650 120 30 re f\n";
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    "<< /Length " + Buffer.byteLength(stream, "ascii") + " >>\nstream\n" + stream + "endstream",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  for (let i = 0; i < objects.length; i++) {
    offsets.push(Buffer.byteLength(pdf, "ascii"));
    pdf += (i + 1) + " 0 obj\n" + objects[i] + "\nendobj\n";
  }
  const xrefOffset = Buffer.byteLength(pdf, "ascii");
  pdf += "xref\n0 " + (objects.length + 1) + "\n";
  pdf += "0000000000 65535 f \n";
  for (let i = 1; i < offsets.length; i++) {
    pdf += String(offsets[i]).padStart(10, "0") + " 00000 n \n";
  }
  pdf += "trailer\n<< /Size " + (objects.length + 1) + " /Root 1 0 R >>\n";
  pdf += "startxref\n" + xrefOffset + "\n%%EOF\n";
  return Buffer.from(pdf, "ascii");
}

function readStoredZip(filePath) {
  const data = fs.readFileSync(filePath);
  const files = new Map();
  let offset = 0;
  while (offset + 4 <= data.length && data.readUInt32LE(offset) === 0x04034b50) {
    const method = data.readUInt16LE(offset + 8);
    const compressedSize = data.readUInt32LE(offset + 18);
    const nameLength = data.readUInt16LE(offset + 26);
    const extraLength = data.readUInt16LE(offset + 28);
    const nameStart = offset + 30;
    const name = data.toString("utf8", nameStart, nameStart + nameLength);
    const bodyStart = nameStart + nameLength + extraLength;
    const bodyEnd = bodyStart + compressedSize;
    assert.equal(method, 0, "expected store-only DOCX ZIP entries");
    assert.ok(bodyEnd <= data.length, "ZIP entry must be inside the file");
    files.set(name, data.subarray(bodyStart, bodyEnd));
    offset = bodyEnd;
  }
  assert.ok(files.has("word/document.xml"), "DOCX must contain the Word document part");
  return files;
}

async function main() {
  assert.ok(fs.existsSync(path.join(root, "index.html")), "built site directory is missing");
  const server = http.createServer((request, response) => {
    const pathname = new URL(request.url, "http://127.0.0.1").pathname;
    const resolved = path.resolve(root, "." + decodeURIComponent(pathname));
    if (resolved !== root && !resolved.startsWith(root + path.sep)) {
      response.writeHead(403).end("Forbidden");
      return;
    }
    const file = path.join(resolved, pathname.endsWith("/") ? "index.html" : "");
    const target = pathname.endsWith("/") ? file : resolved;
    fs.readFile(target, (error, bytes) => {
      if (error) {
        response.writeHead(404).end("Not found");
        return;
      }
      response.writeHead(200, {
        "Content-Type": contentTypes[path.extname(target)] || "application/octet-stream",
        "Cache-Control": "no-store",
      });
      response.end(bytes);
    });
  });

  await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
  const baseUrl = "http://127.0.0.1:" + server.address().port + "/";
  const browser = await chromium.launch({ headless: true, args: ["--no-sandbox"] });
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), "pdftoword-test-"));
  try {
    const page = await browser.newPage({ acceptDownloads: true });
    page.on("pageerror", error => { throw error; });
    await page.goto(baseUrl, { waitUntil: "networkidle" });
    await page.waitForFunction(
      () => document.querySelector("#offline-status")?.textContent.includes("Offline app shell cached"),
      null,
      { timeout: 90000 }
    );
    await page.locator("#file").setInputFiles({
      name: "layout-fixture.pdf",
      mimeType: "application/pdf",
      buffer: makeFixturePdf(),
    });
    const downloadPromise = page.waitForEvent("download", { timeout: 90000 });
    await page.locator("#convert").click();
    const download = await downloadPromise;
    const outputPath = path.join(temp, "converted.docx");
    await download.saveAs(outputPath);
    await page.waitForFunction(
      () => document.querySelector("#status")?.textContent.startsWith("Done."),
      null,
      { timeout: 90000 }
    );

    execFileSync("unzip", ["-t", outputPath], { stdio: "pipe" });
    const parts = readStoredZip(outputPath);
    const xml = parts.get("word/document.xml").toString("utf8");
    assert.match(xml, /HELLO WORLD/, "editable text must be present in Word XML");
    assert.match(xml, /<wps:wsp>/, "text must be represented by an editable Word text box");
    assert.match(xml, /behindDoc="1"/, "the artwork layer must sit behind editable text");
    assert.doesNotMatch(xml, /Visual reference appendix|Original PDF page/);
    const pngEntry = [...parts.entries()].find(([name]) => /^word\/media\/page-1\.png$/.test(name));
    assert.ok(pngEntry, "DOCX must contain the rendered artwork layer");
    const png = PNG.sync.read(pngEntry[1]);
    const pixel = (x, y) => {
      const index = (y * png.width + x) * 4;
      return [...png.data.subarray(index, index + 3)];
    };
    const red = pixel(160, 250);
    assert.ok(red[0] > 220 && red[1] < 60 && red[2] < 60, "vector artwork must survive in the background layer");
    let darkPixelsInTextArea = 0;
    for (let y = 170; y <= 220; y++) {
      for (let x = 140; x <= 500; x++) {
        const [r, g, b] = pixel(x, y);
        if (r < 80 && g < 80 && b < 80) darkPixelsInTextArea++;
      }
    }
    assert.equal(darkPixelsInTextArea, 0, "the background layer must not contain duplicate rasterized text");

    // Reload once so this page is controlled by the worker, then shut down the
    // only server. The second conversion must work from the installed cache.
    await page.reload({ waitUntil: "networkidle" });
    await page.waitForFunction(() => Boolean(navigator.serviceWorker.controller), null, { timeout: 30000 });
    await new Promise(resolve => server.close(resolve));
    await page.reload({ waitUntil: "domcontentloaded", timeout: 30000 });
    await page.locator("#file").setInputFiles({
      name: "layout-fixture.pdf",
      mimeType: "application/pdf",
      buffer: makeFixturePdf(),
    });
    const offlineDownloadPromise = page.waitForEvent("download", { timeout: 90000 });
    await page.locator("#convert").click();
    const offlineDownload = await offlineDownloadPromise;
    const offlineOutputPath = path.join(temp, "converted-offline.docx");
    await offlineDownload.saveAs(offlineOutputPath);
    await page.waitForFunction(
      () => document.querySelector("#status")?.textContent.startsWith("Done."),
      { timeout: 90000 }
    );
    execFileSync("unzip", ["-t", offlineOutputPath], { stdio: "pipe" });
    const offlineParts = readStoredZip(offlineOutputPath);
    assert.match(offlineParts.get("word/document.xml").toString("utf8"), /HELLO WORLD/);
    console.log("Browser conversion integration passed: editable text, graphics-only background, valid DOCX and conversion after disconnecting the server.");
  } finally {
    await browser.close();
    if (server.listening) server.close();
    fs.rmSync(temp, { recursive: true, force: true });
  }
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
