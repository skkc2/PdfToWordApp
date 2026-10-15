# PdfToWordApp

Convert a PDF to a Word (`.docx`) file **entirely on your own device**. Nothing is uploaded, and there are no accounts, ads or tracking.

There are two versions, built from this repo:

| Version | Where it runs | Install needed |
|---|---|---|
| **Web app** | Your browser (GitHub Pages) | No |
| **Android app** | Your phone (sideloaded APK) | Yes |

Web app: https://skkc2.github.io/PdfToWordApp/

## Web converter: hybrid reconstruction

- Extracts positioned PDF.js text and reconstructs editable Word paragraphs and text runs, with approximate font family, size, bold/italic, indentation, and vertical spacing.
- Preserves page order and page breaks. It keeps the source PDF in browser memory for optional page rendering.
- Optional **visual reference appendix** renders each source page to an image and appends it after the editable reconstruction. These appendix images preserve appearance but are not editable text.
- Includes a same-origin service worker to cache the application shell and locally requested assets for offline use after the site has loaded.

This is a reconstruction, not a pixel-perfect PDF-to-Word engine. Tables, multi-column reading order, vector shapes, original images, headers/footers, exact colors and complex typography are not yet reconstructed as native Word objects. The visual appendix is the fidelity fallback. Scanned PDFs without a text layer cannot produce editable text because OCR is not included. Password-protected PDFs are rejected.

## Limits (both versions)

- PDF size: 30 MB
- Pages: 300
- Extracted text: about 3 million characters

These limits reduce the risk of crashes from huge or malicious files.

## Privacy and security design

**Android app**
- Requests **no permissions**. `INTERNET`, network state, storage and install-packages permissions are explicitly removed in the manifest.
- Reads only the one PDF you choose and writes only the one file you create, through Android's system file picker.
- Backups are disabled and the APK is not debuggable.
- The build fails if the APK contains any unexpected permission or is debuggable.

**Web app**
- Conversion runs in your browser; the PDF never leaves your device.
- A Content-Security-Policy restricts the page to its own origin. All code is in separate files, with no inline scripts.
- The PDF engine ([pdf.js](https://github.com/mozilla/pdf.js), version 4.10.38, legacy build) is fetched during the build and served from this same site, not from a third-party CDN.
- The build fails if `app.js` contains network calls (`fetch`, `XMLHttpRequest`, `WebSocket`, `sendBeacon`) or if `index.html` references an external resource.
- The service worker caches same-origin application assets. PDF bytes are read from the selected local file and are not sent to any endpoint. Assets not requested before going offline may not yet be cached; reload once while online for the initial shell cache to be installed.

**Build pipeline**
- GitHub Actions run with read-only tokens, no stored credentials, and no secrets.
- Actions are pinned to commit SHAs (the Pages actions are still on version tags).
- Each build prints SHA-256 fingerprints (APK and web files) in the run summary.

## Repository layout

```
build_pdf_to_word_app.py        Generates the whole Android project
web/index.html                  Web app page
web/app.js                      Web app logic (PDF reading + .docx writing)
web/styles.css                  Web app styling
web/service-worker.js            Browser-only offline app-shell cache
.github/workflows/build-apk.yml Builds, tests and checks the APK
.github/workflows/pages.yml     Builds and deploys the web app
```

The Android project is not committed. The Python script writes out the Gradle project and Kotlin source, and the workflow runs it on every build.

## Install the Android app

1. Open the latest successful **Build Android APK** run under the Actions tab.
2. Note the APK's SHA-256 shown in the run summary.
3. Download the **Offline-PdfToWord-APK** artifact (a zip) and unzip it.
4. Hash `app-debug.apk` with any hash-checker app and compare. If it differs, do not install.
5. Uninstall any older version first (each build uses a new debug signing key).
6. Allow "install unknown apps" for your file manager or browser, install, then turn that setting off again.
7. Check Settings → Apps → Offline PDF to Word → Permissions. It should show none requested.

The APK is signed with a debug key, so Android treats it as unverified. That does not change what the app can do.

## Verify the web app is offline

1. Open the page once and convert a harmless PDF.
2. Turn on airplane mode and convert again. It should still work.
3. Compare the file fingerprints in the latest **Deploy Web App to GitHub Pages** run summary.

## Build and deploy

Both workflows run on push to `main` (the web one only when `web/**` or its workflow changes).

- **Android:** generate the project, run unit tests, build the debug APK, verify permissions, upload the APK (kept 14 days).
- **Web:** fetch pinned pdf.js with install scripts disabled, assemble the site, run safety checks, deploy to GitHub Pages.

One-time setup for the web app: Settings → Pages → Source: **GitHub Actions**.

## Tech

- Android: Kotlin, AndroidX, PDFBox-Android 2.0.27.0, Gradle 8.10.2, JDK 17, minSdk 24
- Web: plain JavaScript, pdf.js 4.10.38, a hand-written ZIP/OpenXML writer with editable text runs, styles, image relationships and an optional visual reference appendix

## Known limitations

- Layout reconstruction is heuristic; multi-column pages can have incorrect reading order, and tables are not yet emitted as native Word tables.
- The optional visual appendix is appended after the editable reconstruction rather than interleaved page-by-page.
- Scanned PDFs need a future browser-bundled OCR stage for editable text.
- PDFBox-Android 2.0.27.0 is older. Check for newer releases from time to time.
- Pages actions are pinned to version tags, not commit SHAs.

## Security notes

No app can be guaranteed 100% secure. Third-party libraries (PDFBox, AndroidX, pdf.js) are not independently audited here. Don't use this for highly sensitive documents unless you've reviewed the code yourself.

## License

No license has been chosen yet, so all rights are reserved by default. Add a `LICENSE` file (for example MIT) if you want others to reuse the code.