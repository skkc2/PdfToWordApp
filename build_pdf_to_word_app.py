from pathlib import Path

PROJECT_DIR = Path("PdfToWordApp")

FILES = {
    "settings.gradle.kts": r'''pluginManagement {
    repositories {
        google()
        mavenCentral()
        gradlePluginPortal()
    }
}
dependencyResolutionManagement {
    repositoriesMode.set(RepositoriesMode.FAIL_ON_PROJECT_REPOS)
    repositories {
        google()
        mavenCentral()
    }
}
rootProject.name = "PdfToWord"
include(":app")
''',

    "build.gradle.kts": r'''plugins {
    id("com.android.application") version "8.7.3" apply false
    id("org.jetbrains.kotlin.android") version "2.0.21" apply false
}
''',

    "gradle.properties": r'''org.gradle.jvmargs=-Xmx2048m -Dfile.encoding=UTF-8
android.useAndroidX=true
kotlin.code.style=official
android.nonTransitiveRClass=true
''',

    "app/build.gradle.kts": r'''plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
}

android {
    namespace = "com.example.pdftoword"
    compileSdk = 35

    defaultConfig {
        applicationId = "com.example.pdftoword"
        minSdk = 24
        targetSdk = 35
        versionCode = 2
        versionName = "1.1"
    }

    buildTypes {
        debug {
            // Not debuggable: blocks debugger/adb run-as access to the app.
            isDebuggable = false
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = "17"
    }

    testOptions {
        unitTests.isReturnDefaultValues = true
    }

    packaging {
        resources {
            excludes += "/META-INF/{AL2.0,LGPL2.1}"
        }
    }
}

dependencies {
    implementation("com.tom-roush:pdfbox-android:2.0.27.0")
    implementation("androidx.core:core-ktx:1.15.0")
    implementation("androidx.appcompat:appcompat:1.7.0")
    implementation("androidx.activity:activity-ktx:1.10.0")
    implementation("androidx.lifecycle:lifecycle-runtime-ktx:2.8.7")
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.9.0")

    testImplementation("junit:junit:4.13.2")
}
''',

    "app/src/main/AndroidManifest.xml": r'''<?xml version="1.0" encoding="utf-8"?>
<manifest xmlns:android="http://schemas.android.com/apk/res/android"
    xmlns:tools="http://schemas.android.com/tools">

    <!-- This app requests NO permissions. Strip these even if a library adds them. -->
    <uses-permission android:name="android.permission.INTERNET" tools:node="remove" />
    <uses-permission android:name="android.permission.ACCESS_NETWORK_STATE" tools:node="remove" />
    <uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE" tools:node="remove" />
    <uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" tools:node="remove" />
    <uses-permission android:name="android.permission.MANAGE_EXTERNAL_STORAGE" tools:node="remove" />
    <uses-permission android:name="android.permission.REQUEST_INSTALL_PACKAGES" tools:node="remove" />

    <application
        android:allowBackup="false"
        android:usesCleartextTraffic="false"
        android:label="Offline PDF to Word"
        android:supportsRtl="true"
        android:theme="@style/Theme.AppCompat.DayNight.NoActionBar">

        <activity
            android:name=".MainActivity"
            android:exported="true"
            android:configChanges="orientation|screenSize|keyboardHidden|screenLayout|smallestScreenSize|uiMode">
            <intent-filter>
                <action android:name="android.intent.action.MAIN" />
                <category android:name="android.intent.category.LAUNCHER" />
            </intent-filter>
        </activity>
    </application>
</manifest>
''',

    "app/src/main/java/com/example/pdftoword/PdfToDocxConverter.kt": r'''package com.example.pdftoword

import android.content.Context
import android.net.Uri
import android.provider.DocumentsContract
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.text.PDFTextStripper
import kotlinx.coroutines.CancellationException
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.ensureActive
import kotlinx.coroutines.withContext
import java.io.FilterInputStream
import java.io.IOException
import java.io.InputStream
import java.io.OutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream
import kotlin.coroutines.coroutineContext

class ConversionException(message: String) : Exception(message)

/** Reads at most [limit] bytes; throws if the source is larger. */
private class BoundedInputStream(
    input: InputStream,
    private val limit: Long
) : FilterInputStream(input) {
    private var count = 0L

    override fun read(): Int {
        val b = super.read()
        if (b >= 0) add(1)
        return b
    }

    override fun read(buffer: ByteArray, off: Int, len: Int): Int {
        val n = super.read(buffer, off, len)
        if (n > 0) add(n.toLong())
        return n
    }

    private fun add(n: Long) {
        count += n
        if (count > limit) {
            throw ConversionException(
                "This PDF is larger than the ${limit / (1024 * 1024)} MB limit."
            )
        }
    }
}

class PdfToDocxConverter(private val context: Context? = null) {

    init {
        context?.let {
            if (!PDFBoxResourceLoader.isReady()) {
                PDFBoxResourceLoader.init(it.applicationContext)
            }
        }
    }

    suspend fun convertPdfToDocx(
        pdfUri: Uri,
        docxUri: Uri,
        onProgress: (currentPage: Int, totalPages: Int) -> Unit
    ): Result<Unit> = withContext(Dispatchers.IO) {
        val appContext = context?.applicationContext
            ?: return@withContext Result.failure(ConversionException("Internal error: no context."))
        try {
            val pages = readPdfText(appContext, pdfUri, onProgress)
            writeDocx(appContext, docxUri, pages)
            Result.success(Unit)
        } catch (e: CancellationException) {
            deleteQuietly(appContext, docxUri)
            throw e
        } catch (e: Throwable) {
            deleteQuietly(appContext, docxUri)
            Result.failure(ConversionException(friendlyMessage(e)))
        }
    }

    private suspend fun readPdfText(
        ctx: Context,
        pdfUri: Uri,
        onProgress: (Int, Int) -> Unit
    ): List<List<String>> {
        val raw = ctx.contentResolver.openInputStream(pdfUri)
            ?: throw ConversionException("Unable to open the selected PDF.")
        val pagesText = mutableListOf<List<String>>()
        var totalChars = 0L

        raw.use { stream ->
            PDDocument.load(BoundedInputStream(stream, MAX_PDF_BYTES)).use { document ->
                val totalPages = document.numberOfPages
                if (totalPages > MAX_PAGES) {
                    throw ConversionException(
                        "This PDF has $totalPages pages. The limit is $MAX_PAGES."
                    )
                }
                val stripper = PDFTextStripper().apply { sortByPosition = true }

                for (page in 1..totalPages) {
                    coroutineContext.ensureActive()
                    stripper.startPage = page
                    stripper.endPage = page
                    val rawPageText = stripper.getText(document)
                    totalChars += rawPageText.length
                    if (totalChars > MAX_TOTAL_CHARS) {
                        throw ConversionException("This PDF contains too much text to convert safely.")
                    }
                    val paragraphs = rawPageText
                        .split(Regex("\\r?\\n"))
                        .map { it.trim() }
                        .filter { it.isNotEmpty() }
                    pagesText.add(paragraphs)
                    onProgress(page, totalPages)
                }
            }
        }
        return pagesText
    }

    private fun writeDocx(ctx: Context, docxUri: Uri, pages: List<List<String>>) {
        val out = ctx.contentResolver.openOutputStream(docxUri, "wt")
            ?: throw ConversionException("Unable to create the output Word file.")
        out.use { writeDocxArchive(it, pages) }
    }

    private fun deleteQuietly(ctx: Context, uri: Uri) {
        try {
            DocumentsContract.deleteDocument(ctx.contentResolver, uri)
        } catch (ignored: Exception) {
            // Best effort only.
        }
    }

    private fun friendlyMessage(t: Throwable): String = when {
        t is ConversionException -> t.message ?: "Conversion failed."
        t is OutOfMemoryError -> "Not enough memory to convert this PDF. Try a smaller file."
        t.javaClass.simpleName == "InvalidPasswordException" ->
            "This PDF is password-protected and can't be converted."
        t is IOException -> "Could not read the PDF. It may be damaged or not a valid PDF."
        else -> "Conversion failed (${t.javaClass.simpleName})."
    }

    internal fun writeDocxArchive(outputStream: OutputStream, pages: List<List<String>>) {
        ZipOutputStream(outputStream).use { zip ->
            zip.putNextEntry(ZipEntry("[Content_Types].xml"))
            zip.write(CONTENT_TYPES_XML.toByteArray(Charsets.UTF_8))
            zip.closeEntry()

            zip.putNextEntry(ZipEntry("_rels/.rels"))
            zip.write(RELS_XML.toByteArray(Charsets.UTF_8))
            zip.closeEntry()

            zip.putNextEntry(ZipEntry("word/document.xml"))
            zip.write(buildDocumentXml(pages).toByteArray(Charsets.UTF_8))
            zip.closeEntry()
        }
    }

    internal fun buildDocumentXml(pages: List<List<String>>): String {
        val sb = StringBuilder()
        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
        sb.append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">")
        sb.append("<w:body>")
        pages.forEachIndexed { pageIndex, paragraphs ->
            paragraphs.forEach { paragraphText ->
                val escaped = escapeXml(sanitizeXmlChars(paragraphText))
                sb.append("<w:p><w:r><w:t xml:space=\"preserve\">")
                sb.append(escaped)
                sb.append("</w:t></w:r></w:p>")
            }
            if (pageIndex < pages.lastIndex) {
                sb.append("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>")
            }
        }
        sb.append("</w:body></w:document>")
        return sb.toString()
    }

    /** Keeps only characters that are valid in XML 1.0 (including emoji pairs). */
    internal fun sanitizeXmlChars(input: String): String {
        val sb = StringBuilder(input.length)
        var i = 0
        while (i < input.length) {
            val cp = input.codePointAt(i)
            val len = Character.charCount(cp)
            val ok = cp == 0x9 || cp == 0xA || cp == 0xD ||
                (cp in 0x20..0xD7FF) ||
                (cp in 0xE000..0xFFFD) ||
                (cp in 0x10000..0x10FFFF)
            if (ok) sb.append(input, i, i + len)
            i += len
        }
        return sb.toString()
    }

    private fun escapeXml(input: String): String {
        return buildString(input.length) {
            for (ch in input) {
                when (ch) {
                    '&' -> append("&amp;")
                    '<' -> append("&lt;")
                    '>' -> append("&gt;")
                    '\u0022' -> append("&quot;")
                    '\u0027' -> append("&apos;")
                    else -> append(ch)
                }
            }
        }
    }

    companion object {
        internal const val MAX_PDF_BYTES = 30L * 1024 * 1024
        internal const val MAX_PAGES = 300
        internal const val MAX_TOTAL_CHARS = 3_000_000L

        private const val CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

        private const val RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    }
}
''',

    "app/src/main/java/com/example/pdftoword/MainActivity.kt": r'''package com.example.pdftoword

import android.net.Uri
import android.os.Bundle
import android.view.Gravity
import android.view.View
import android.widget.Button
import android.widget.LinearLayout
import android.widget.ProgressBar
import android.widget.TextView
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.launch

class MainActivity : AppCompatActivity() {

    private var selectedPdfUri: Uri? = null
    private var isConverting = false

    private lateinit var statusText: TextView
    private lateinit var progressBar: ProgressBar
    private lateinit var selectPdfButton: Button
    private lateinit var convertButton: Button
    private lateinit var converter: PdfToDocxConverter

    // Temporary read access to ONLY the file the user picks.
    private val pickPdfLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        if (uri != null) {
            selectedPdfUri = uri
            statusText.text = "PDF selected. Tap step 2 to convert."
            convertButton.isEnabled = true
        }
    }

    // Write access to ONLY the location the user chooses.
    private val createDocxLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    ) { outputUri: Uri? ->
        val inputUri = selectedPdfUri
        if (outputUri != null && inputUri != null) {
            runConversion(inputUri, outputUri)
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        converter = PdfToDocxConverter(this)

        val density = resources.displayMetrics.density
        val pad = (24 * density).toInt()

        val rootLayout = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            gravity = Gravity.CENTER
            setPadding(pad, pad, pad, pad)
        }

        val titleView = TextView(this).apply {
            text = "Offline PDF to Word (.docx)"
            textSize = 22f
            gravity = Gravity.CENTER
        }

        statusText = TextView(this).apply {
            text = "Select a PDF from your device. Nothing leaves your phone."
            textSize = 16f
            gravity = Gravity.CENTER
            setPadding(0, pad / 2, 0, pad)
        }

        progressBar = ProgressBar(this, null, android.R.attr.progressBarStyleHorizontal).apply {
            max = 100
            visibility = View.GONE
        }

        selectPdfButton = Button(this).apply {
            text = "1. Select PDF File"
            setOnClickListener {
                if (!isConverting) pickPdfLauncher.launch(arrayOf("application/pdf"))
            }
        }

        convertButton = Button(this).apply {
            text = "2. Convert & Save Word (.docx)"
            isEnabled = false
            setOnClickListener {
                if (!isConverting) createDocxLauncher.launch("Converted_Document.docx")
            }
        }

        rootLayout.addView(titleView)
        rootLayout.addView(statusText)
        rootLayout.addView(progressBar)
        rootLayout.addView(selectPdfButton)
        rootLayout.addView(convertButton)
        setContentView(rootLayout)
    }

    private fun runConversion(inputUri: Uri, outputUri: Uri) {
        if (isConverting) return
        isConverting = true
        selectPdfButton.isEnabled = false
        convertButton.isEnabled = false
        progressBar.progress = 0
        progressBar.visibility = View.VISIBLE
        statusText.text = "Starting local conversion..."

        lifecycleScope.launch {
            val result = converter.convertPdfToDocx(
                pdfUri = inputUri,
                docxUri = outputUri,
                onProgress = { current, total ->
                    runOnUiThread {
                        progressBar.progress = (current * 100) / total.coerceAtLeast(1)
                        statusText.text = "Converting page $current of $total..."
                    }
                }
            )

            isConverting = false
            progressBar.visibility = View.GONE
            selectPdfButton.isEnabled = true
            convertButton.isEnabled = true

            result.onSuccess {
                statusText.text = "Success! Word (.docx) saved to your device."
            }.onFailure { error ->
                statusText.text = "Error: ${error.message ?: "Conversion failed."}"
            }
        }
    }
}
''',

    "app/src/test/java/com/example/pdftoword/PdfToDocxConverterTest.kt": r'''package com.example.pdftoword

import org.junit.Assert.assertEquals
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.util.zip.ZipInputStream

class PdfToDocxConverterTest {

    @Test
    fun `writeDocxArchive produces valid OpenXML zip with paragraphs and page breaks`() {
        val converter = PdfToDocxConverter()
        val samplePages = listOf(
            listOf("First Page Heading", "Paragraph with special chars: A & B < C > D"),
            listOf("Second Page Paragraph")
        )

        val outputBuffer = ByteArrayOutputStream()
        converter.writeDocxArchive(outputBuffer, samplePages)

        val zipEntries = mutableMapOf<String, String>()
        ZipInputStream(ByteArrayInputStream(outputBuffer.toByteArray())).use { zis ->
            var entry = zis.nextEntry
            while (entry != null) {
                zipEntries[entry.name] = zis.readBytes().toString(Charsets.UTF_8)
                entry = zis.nextEntry
            }
        }

        assertTrue("Missing [Content_Types].xml", zipEntries.containsKey("[Content_Types].xml"))
        assertTrue("Missing _rels/.rels", zipEntries.containsKey("_rels/.rels"))
        assertTrue("Missing word/document.xml", zipEntries.containsKey("word/document.xml"))

        val docXml = zipEntries.getValue("word/document.xml")
        assertTrue("Missing first page text", docXml.contains("First Page Heading"))
        assertTrue("Special characters not escaped", docXml.contains("A &amp; B &lt; C &gt; D"))
        assertTrue("Missing page break between pages", docXml.contains("<w:br w:type=\"page\"/>"))
        assertTrue("Missing second page text", docXml.contains("Second Page Paragraph"))
    }

    @Test
    fun `sanitizeXmlChars removes control characters but keeps text and emoji`() {
        val converter = PdfToDocxConverter()
        val cleaned = converter.sanitizeXmlChars("A\u0000B\u0008C \uD83D\uDE00 end")
        assertEquals("ABC \uD83D\uDE00 end", cleaned)
    }
}
''',
}


def main():
    for rel_path, content in FILES.items():
        full_path = PROJECT_DIR / rel_path
        full_path.parent.mkdir(parents=True, exist_ok=True)
        full_path.write_text(content, encoding="utf-8")
    print(f"Project generated successfully in ./{PROJECT_DIR}/")


if __name__ == "__main__":
    main()