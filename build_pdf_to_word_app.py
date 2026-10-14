import os
from pathlib import Path

PROJECT_DIR = Path("PdfToWordApp")

FILES = {
    "settings.gradle.kts": """pluginManagement {
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
""",

    "build.gradle.kts": """plugins {
    id("com.android.application") version "8.7.3" apply false
    id("org.jetbrains.kotlin.android") version "2.0.21" apply false
}
""",

    "gradle.properties": """org.gradle.jvmargs=-Xmx2048m -Dfile.encoding=UTF-8
android.useAndroidX=true
kotlin.code.style=official
android.nonTransitiveRClass=true
""",

    "app/build.gradle.kts": """plugins {
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
        versionCode = 1
        versionName = "1.0"
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
    implementation("com.google.android.material:material:1.12.0")
    implementation("androidx.activity:activity-ktx:1.10.0")
    implementation("androidx.lifecycle:lifecycle-runtime-ktx:2.8.7")
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.9.0")

    testImplementation("junit:junit:4.13.2")
}
""",

    "app/src/main/AndroidManifest.xml": """<?xml version="1.0" encoding="utf-8"?>
<manifest xmlns:android="http://schemas.android.com/apk/res/android">

    <application
        android:allowBackup="true"
        android:label="Offline PDF to Word"
        android:supportsRtl="true"
        android:theme="@style/Theme.Material3.DayNight.NoActionBar">
        <activity
            android:name=".MainActivity"
            android:exported="true">
            <intent-filter>
                <action android:name="android.intent.action.MAIN" />
                <category android:name="android.intent.category.LAUNCHER" />
            </intent-filter>
        </activity>
    </application>

</manifest>
""",

    "app/src/main/java/com/example/pdftoword/PdfToDocxConverter.kt": """package com.example.pdftoword

import android.content.Context
import android.net.Uri
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.text.PDFTextStripper
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.OutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

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
        runCatching {
            val appContext = requireNotNull(context) { "Android Context required for URI streams" }
            val pagesText = mutableListOf<List<String>>()

            appContext.contentResolver.openInputStream(pdfUri)?.use { inputStream ->
                PDDocument.load(inputStream).use { document ->
                    val totalPages = document.numberOfPages
                    val stripper = PDFTextStripper().apply {
                        sortByPosition = true
                    }

                    for (page in 1..totalPages) {
                        stripper.startPage = page
                        stripper.endPage = page
                        val rawPageText = stripper.getText(document)
                        val paragraphs = rawPageText
                            .split(Regex("\\\\r?\\\\n"))
                            .map { it.trim() }
                            .filter { it.isNotEmpty() }

                        pagesText.add(paragraphs)
                        onProgress(page, totalPages)
                    }
                }
            } ?: throw IllegalStateException("Unable to open selected PDF file.")

            appContext.contentResolver.openOutputStream(docxUri)?.use { outputStream ->
                writeDocxArchive(outputStream, pagesText)
            } ?: throw IllegalStateException("Unable to create output Word (.docx) file.")
        }
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
            val documentXml = buildDocumentXml(pages)
            zip.write(documentXml.toByteArray(Charsets.UTF_8))
            zip.closeEntry()
        }
    }

    internal fun buildDocumentXml(pages: List<List<String>>): String {
        val sb = StringBuilder()
        sb.append("<?xml version=\\"1.0\\" encoding=\\"UTF-8\\" standalone=\\"yes\\"?>")
        sb.append("<w:document xmlns:w=\\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\\">")
        sb.append("<w:body>")

        pages.forEachIndexed { pageIndex, paragraphs ->
            paragraphs.forEach { paragraphText ->
                val escaped = escapeXml(sanitizeXmlChars(paragraphText))
                sb.append("<w:p><w:r><w:t xml:space=\\"preserve\\">")
                sb.append(escaped)
                sb.append("</w:t></w:r></w:p>")
            }
            if (pageIndex < pages.lastIndex) {
                sb.append("<w:p><w:r><w:br w:type=\\"page\\"/></w:r></w:p>")
            }
        }

        sb.append("</w:body></w:document>")
        return sb.toString()
    }

    internal fun sanitizeXmlChars(input: String): String {
        return input.filter { ch ->
            ch == '\\t' || ch == '\\n' || ch == '\\r' ||
                ch in '\\u0020'..'\\uD7FF' ||
                ch in '\\uE000'..'\\uFFFD'
        }
    }

    private fun escapeXml(input: String): String {
        return buildString(input.length) {
            for (ch in input) {
                when (ch) {
                    '&' -> append("&amp;")
                    '<' -> append("&lt;")
                    '>' -> append("&gt;")
                    '\\"' -> append("&quot;")
                    '\\'' -> append("&apos;")
                    else -> append(ch)
                }
            }
        }
    }

    companion object {
        private const val CONTENT_TYPES_XML = \"\"\"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>\"\"\"

        private const val RELS_XML = \"\"\"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>\"\"\"
    }
}
""",

    "app/src/main/java/com/example/pdftoword/MainActivity.kt": """package com.example.pdftoword

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
    private lateinit var statusText: TextView
    private lateinit var progressBar: ProgressBar
    private lateinit var selectPdfButton: Button
    private lateinit var convertButton: Button
    private lateinit var converter: PdfToDocxConverter

    private val pickPdfLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        if (uri != null) {
            selectedPdfUri = uri
            statusText.text = "PDF loaded ready for offline conversion."
            convertButton.isEnabled = true
        }
    }

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
            text = "Select a local PDF file to convert offline."
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
                pickPdfLauncher.launch(arrayOf("application/pdf"))
            }
        }

        convertButton = Button(this).apply {
            text = "2. Convert & Save Word (.docx)"
            isEnabled = false
            setOnClickListener {
                createDocxLauncher.launch("Converted_Document.docx")
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
                        statusText.text = "Converting page $current of$total..."
                    }
                }
            )
            progressBar.visibility = View.GONE
            selectPdfButton.isEnabled = true
            convertButton.isEnabled = true

            result.onSuccess {
                statusText.text = "Success! Word (.docx) saved to your device."
            }.onFailure { error ->
                statusText.text = "Error: ${error.localizedMessage ?: "Conversion failed"}"
            }
        }
    }
}
""",

    "app/src/test/java/com/example/pdftoword/PdfToDocxConverterTest.kt": """package com.example.pdftoword

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
        assertTrue("Missing page break between pages", docXml.contains("<w:br w:type=\\"page\\"/>"))
        assertTrue("Missing second page text", docXml.contains("Second Page Paragraph"))
    }
}
""",

    ".github/workflows/build-apk.yml": """name: Build and Test Android APK

on:
  push:
    branches: [ "main", "master" ]
  workflow_dispatch:

jobs:
  build-apk:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout code
        uses: actions/checkout@v4

      - name: Set up JDK 17
        uses: actions/setup-java@v4
        with:
          distribution: 'temurin'
          java-version: '17'

      - name: Setup Gradle
        uses: gradle/actions/setup-gradle@v4
        with:
          gradle-version: '8.10.2'

      - name: Run Unit Tests
        run: gradle testDebugUnitTest

      - name: Build Debug APK
        run: gradle assembleDebug

      - name: Upload Compiled APK Artifact
        uses: actions/upload-artifact@v4
        with:
          name: Offline-PdfToWord-APK
          path: app/build/outputs/apk/debug/app-debug.apk
"""
}

def main():
    for rel_path, content in FILES.items():
        full_path = PROJECT_DIR / rel_path
        full_path.parent.mkdir(parents=True, exist_ok=True)
        full_path.write_text(content, encoding="utf-8")
    print(f"Project generated successfully in ./{PROJECT_DIR}/")

if __name__ == "__main__":
    main()
