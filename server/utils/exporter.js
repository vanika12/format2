

import puppeteer from "puppeteer"
import fs from "fs";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  convertInchesToTwip,
  Header,
  Footer,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  PageNumber,
  ImageRun,
} from "docx";
// import { Document, Packer, Paragraph, TextRun, AlignmentType, convertInchesToTwip } from "docx"
import { applyLatexFormatting, formatCitations } from "./formatter.js"
import { generateCompleteLatexDocument } from "./latex-converter.js"
import path from "path";

import { fileURLToPath } from "url";

// Fix __dirname for ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ✅ Go two levels up (from server/utils → project root)
const imagePath = path.join(__dirname, "..", "..", "public", "rsis.jpg");
const imageBuffer = fs.readFileSync(imagePath);
const imageBase64 = imageBuffer.toString("base64");
const imageSrc = `data:image/jpeg;base64,${imageBase64}`;

// Fallback styles used when a section/subsection doesn't include formatting
const DEFAULT_HEADING_STYLE = "font-size: 14pt; font-weight: bold; text-transform: uppercase; font-family: 'Times New Roman', serif; text-align: left; margin-top: 20pt; margin-bottom: 10pt;"
const DEFAULT_SUB_HEADING_STYLE = "font-size: 12pt; font-weight: bold; font-family: 'Times New Roman', serif; text-align: left; margin-top: 12pt; margin-bottom: 6pt;"
const DEFAULT_CONTENT_STYLE = "font-size: 12pt; font-family: 'Times New Roman', serif; line-height: 1.0; text-align: justify; margin-bottom: 12pt;"

const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")

export const exportToHTML = async (formattedDocument) => {
   const doc = formattedDocument || {}
  const {
    metadata = { formatSpecs: { margins: { top: "0.76in", bottom: "0.42in", left: "0.42in", right: "0.42in" } } },
    header = {},
    title = { text: "" },
    authors = { text: "", style: "" },
    affiliations = [],
    abstract = { heading: "ABSTRACT", content: "", style: { heading: "", content: "" } },
    keywords = { heading: "Keywords", content: "", style: { heading: "", content: "" } },
    sections = [],
    references = { heading: "REFERENCES", content: [], style: { heading: "", content: "" } },
    footer = {},
    publicationInfo = null,
    imageSrc = ""
  } = formattedDocument || {}

  const headerPartsTmp = []
  if (header?.issn) headerPartsTmp.push(header.issn)
  if (header?.doi) headerPartsTmp.push(header.doi)
  if (header?.volume) headerPartsTmp.push(header.volume)
  // const headerSecondLineHTML = esc(headerPartsTmp.join(" | "))
    const headerSecondLineHTML = (typeof esc === "function") ? esc(headerPartsTmp.join(" | ")) : headerPartsTmp.join(" | ")


  const html = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title.text}</title>
    <style>
        @page {
            size: A4;
            margin: ${metadata.formatSpecs.margins.top} ${metadata.formatSpecs.margins.right} ${metadata.formatSpecs.margins.bottom} ${metadata.formatSpecs.margins.left};
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Times New Roman', serif;
            font-size: 12pt;
            line-height: 1.0;
            color: #000;
            background: #fff;
            text-align: justify;
        }
        
        .document-container {
            max-width: 8.5in;
            margin: 0 auto;
            padding: 0;
        }
        
        .header {
            text-align: center;
            font-size: 10pt;
            font-weight: bold;
            border-bottom: 1px solid #000;
            padding-bottom: 10pt;
            margin-bottom: 20pt;
            border-top:1px solid #000;
        }
        
        .header-line {
            margin-bottom: 2pt;
        }
        
        .title {
            ${title?.style|| ""}
            margin: 12pt 0;
        }
        
        .authors {
            ${authors?.style || ""}
            margin: 10pt 0;
        }

        
        .affiliations {
            text-align: center;
            font-size: 12pt;
            margin: 10pt 0;
        }
        
        .affiliation {
            margin-bottom: 6pt;
        }
        
        .publication-info {
            ${publicationInfo?.style || "text-align: center; font-size: 12pt; margin: 15pt 0;"}
        }
        
        .abstract {
            margin: 20pt 0;
        }
        
        .abstract-heading {
            ${abstract?.style?.heading || ""}
            margin-bottom: 10pt;
        }
        
        .abstract-content {
            ${abstract?.style?.content|| ""}
            margin-bottom: 15pt;
        }

        .abstract-content p, .section-content p {
            margin-bottom: 12pt; /* Adds space between paragraphs */
        }

        .abstract-content p:last-child, .section-content p:last-child {
            margin-bottom: 0; /* Removes space after the last paragraph in a section */
        }
        
        .keywords {
            margin: 15pt 0;
        }
        
        .keywords-heading {
            ${keywords?.style?.heading|| ""}
        }
        
        .keywords-content {
            ${keywords?.style?.content|| ""}
        }
        
        .section {
            margin: 20pt 0;
        }
        
        .section-heading {
            margin-bottom: 10pt;
        }
        
        .section-content {
            text-align: justify;
            margin-bottom: 15pt;
        }
        
        .subsection {
            margin: 15pt 0;
        }
        
        .subsection-heading {
            font-size: 12pt;
            font-weight: bold;
            margin-bottom: 6pt;
        }
        
        .references {
            margin: 20pt 0;
        }
        
        .references-heading {
            ${references?.style?.heading|| ""}
            margin-bottom: 10pt;
        }
        
        /* References aligned number + text (like screenshot) */
        .references-list { margin-top: 2pt;
        padding-left: 20pt; /* shift everything to the right */ }
        .reference-row {
            display: grid;
            grid-template-columns: 24pt 1fr; /* number column + text */
            column-gap: 4pt;
            align-items: start;
        }
        .reference-row .ref-num { text-align: left; }
        .reference-row .ref-text { text-align: justify; }
        
        /* Citation and formatting styles */
        sup {
            font-size: 10pt;
            vertical-align: super;
        }
        
        .equation {
            text-align: center;
            margin: 12pt 0;
            font-style: italic;
        }
        
        .inline-equation {
            font-style: italic;
        }
        
        .table-caption, .figure-caption {
            font-weight: bold;
            text-align: center;
            margin: 12pt 0 6pt 0;
        }

        /* Body header visible in HTML, hidden in print/PDF */
        .body-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-bottom: 1px solid #000;
            padding: 0 0.42in 10pt;
            margin-bottom: 20pt;
            font-family: 'Times New Roman', serif;
        }
        .body-header .logo { flex: 0 0 auto; }
        .body-header .logo img { height: 35px; }
        .body-header .header-text { flex: 1; text-align: right; font-weight: bold; font-size: 10pt; }
        .body-header .header-text span { font-size: 9pt; font-weight: normal; }
        
        /* Print styles */
        @media print {
            .document-container { max-width: none; }
            .section-heading { page-break-after: avoid; }
            .body-header { display: none; }
        }
    </style>
</head>
<body>
    <div class="document-container">
        <!-- HTML-visible header (hidden in print/PDF) -->
        <div class="body-header">
          <div class="logo"><img src="${imageSrc}" alt="Logo"/></div>
          <div class="header-text">
            ${esc(header?.content || "")}<br>
            <span>${headerSecondLineHTML}</span>
          </div>
        </div>
        
        <div class="title">${title.text|| ""}</div>
        
        <div class="authors">${authors.text|| ""}</div>
        
        ${
          affiliations && affiliations.length > 0
            ? `
        <div class="affiliations">
            ${affiliations
              .map(
                (aff) => `
                <div class="affiliation" style="${aff.style || ''}">
                    <sup>${aff.number|| ""}</sup> ${aff.text|| ""}
                </div>
            `,
              )
              .join("")}
        </div>
        `
            : ""
        }
        
        ${
  publicationInfo
    ? `
<div class="publication-info">
    ${
      publicationInfo.doi
        ? `<p style="margin-bottom:8pt;"><strong>DOI:</strong> 
           <a href="${publicationInfo.doi}" style="color: blue; text-decoration: underline;">
             ${publicationInfo.doi}
           </a></p>`
        : ""
    }
    ${
      publicationInfo.received || publicationInfo.accepted || publicationInfo.published
        ? `<p style="margin-top:6pt;">
             ${publicationInfo.received ? `<strong>Received:</strong> ${publicationInfo.received}` : ""}
             ${publicationInfo.accepted ? `; <strong>Accepted:</strong> ${publicationInfo.accepted}` : ""}
             ${publicationInfo.published ? `; <strong>Published:</strong> ${publicationInfo.published}` : ""}
           </p>`
        : ""
    }
</div>
`
    : ""
}

        
        ${
  (abstract?.heading || abstract?.content)
    ? `
    <div class="abstract">
        <div class="abstract-heading">${abstract.heading || "ABSTRACT"}</div>
        <div class="abstract-content">${
          abstract?.content
            ? applyLatexFormatting(formatCitations(abstract.content))
                .split(/\n\s*\n/)
                .map((p) => p.trim())
                .filter((p) => p)
                .map((p) => `<p>${p}</p>`)
                .join("")
            : ""
        }</div>
    </div>
    `
    : ""
}

        
        ${
          keywords?.content
            ? `
        <div class="keywords">
            <span class="keywords-heading">${keywords.heading|| ""}:</span>
            <span class="keywords-content">${keywords.content|| ""}</span>
        </div>
        `
            : ""
        }
        
        ${sections
          .map(
            (section) => `
            <div class="section">
<div class=\"section-heading\" style=\"${(section && section.formatting && section.formatting.headingStyle) || DEFAULT_HEADING_STYLE}\">
${esc(section.heading)}
                </div>
<div class=\"section-content\" style=\"${(section && section.formatting && section.formatting.contentStyle) || DEFAULT_CONTENT_STYLE}\">
${applyLatexFormatting(formatCitations(section.content|| ""))
                      .split(/\n\s*\n/)
                      .map((p) => p.trim())
                      .filter((p) => p)
                      .map((p) => {
                        const plain = p.replace(/<[^>]+>/g, '')
return isKeywordHeadingOnly(plain)
? `<div class=\"section-heading\" style=\"${(section && section.formatting && section.formatting.headingStyle) || DEFAULT_HEADING_STYLE}\">${plain}</div>`
                          : `<p>${p}</p>`
                      })
                      .join("")}
                </div>
                
                ${
                  section.subsections && section.subsections.length > 0
                    ? section.subsections
                        .map(
                          (subsection) => `
                        <div class="subsection">
<div class=\"subsection-heading\" style=\"${(subsection && subsection.formatting && subsection.formatting.headingStyle) || DEFAULT_SUB_HEADING_STYLE}\">
${esc(subsection?.heading|| "")}
                            </div>
<div class=\"section-content\" style=\"${(subsection && subsection.formatting && subsection.formatting.contentStyle) || DEFAULT_CONTENT_STYLE}\">
${applyLatexFormatting(formatCitations(subsection?.content|| ""))
                                  .split(/\n\s*\n/)
                                  .map((p) => p.trim())
                                  .filter((p) => p)
                                  .map((p) => {
                                    const plain = p.replace(/<[^>]+>/g, '')
return isKeywordHeadingOnly(plain)
? `<div class=\"subsection-heading\" style=\"${(subsection && subsection.formatting && subsection.formatting.headingStyle) || DEFAULT_SUB_HEADING_STYLE}\">${plain}</div>`
                                      : `<p>${p}</p>`
                                  })
                                  .join("")}
                            </div>
                        </div>
                    `,
                        )
                        .join("")
                    : ""
                }
            </div>
        `,
          )
          .join("")}
        
        ${
          references.content && references.content.length > 0
            ? `
        <div class="references">
            <div class="references-heading">${references.heading}</div>
            <div class="references-list">
              ${references.content
                .map((ref) => `
                  <div class="reference-row">
                    <div class="ref-num">${ref.number}.</div>
                    <div class="ref-text">${ref.text}</div>
                  </div>
                `)
                .join("")}
            </div>
        </div>
        `
            : ""
        }
    </div>
</body>
</html>
`

  return html
}

export const exportToPDF = async (formattedDocument) => {
  console.log("[v0] Starting PDF generation...")

  const htmlContent = await exportToHTML(formattedDocument)
  console.log("[v0] HTML content generated, length:", htmlContent.length)

  let browser
  try {
    console.log("[v0] Launching Puppeteer browser...")
    browser = await puppeteer.launch({
      headless: "new",
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-accelerated-2d-canvas",
        "--no-first-run",
        "--no-zygote",
        "--disable-gpu",
        "--disable-web-security",
        "--disable-features=VizDisplayCompositor",
      ],
      timeout: 60000,
    })

    console.log("[v0] Browser launched successfully")
    const page = await browser.newPage()

    // Set viewport for consistent rendering
    await page.setViewport({ width: 794, height: 1123 }) // A4 dimensions in pixels
    console.log("[v0] Viewport set")

    console.log("[v0] Setting page content...")
    await page.emulateMediaType("print");
    await page.setContent(htmlContent, {
      waitUntil: "networkidle0",
      timeout: 30000,
    })

    console.log("[v0] Content loaded, generating PDF...")

    const headerParts = []
    const headerData = formattedDocument.header || {}
    if (headerData.issn) {
      headerParts.push(headerData.issn)
    }
    if (headerData.doi) {
      headerParts.push(headerData.doi)
    }
    if (headerData.volume) {
      headerParts.push(headerData.volume)
    }
    const headerSecondLine = headerParts.join(" | ")

    // const startPageNumber = formattedDocument.startPageNumber || 1;
//     const actualPageNumber = startPage + currentPageIndex - 1;
// footerText = footerText.replace("{pageNumber}", actualPageNumber);

    await page.addStyleTag({
  content: `
    @page {
      margin-bottom: 50px !important; /* Ensure enough space for footer */
    }
  `
});
// Adjust page numbers based on startPageNumber
// const startPage = formattedDocument.startPageNumber || 1;

// // Override Puppeteer's pageNumber spans after content loads
// await page.evaluate((startPage) => {
//   const spans = document.querySelectorAll(".pageNumber");
//   spans.forEach((el, i) => {
//     el.textContent = startPage + i; // Shift page numbering
//   });
// }, startPage);



    const pdfBuffer = await page.pdf({
      format: "A4",
      margin: {
        top: "0.76in",
        bottom: "0.42in",
        left: "0.42in",
        right: "0.42in",
        // top: "0.1in",
        // bottom: "0.8in",
        // left: "0.63in",
        // right: "0.63in",
      },
      printBackground: true,
      displayHeaderFooter: true,
       preferCSSPageSize: true,
       

      headerTemplate: `
        <div style="font-size: 10pt; width: 100%; text-align: center; font-family: 'Times New Roman', serif; font-weight: bold; padding: 0 0.42in;">
        <div style="border-bottom: 1px solid #000; padding-bottom: 10pt; display: flex; align-items: center; justify-content: space-between;">
          
           <!-- Logo at top-left -->
      <div style="flex: 0 0 auto; text-align: left;">
        <img src="${imageSrc}" alt="Logo" style="height:35px;"/>
      </div>
             <!-- Journal header content centered -->
      <div style=\"flex: 1; text-align: right;\">
        ${headerData.content || ""}<br>
        <span style="font-size: 9pt; font-weight: normal;">
          ${headerSecondLine}
        </span>
      </div>

      <!-- Empty space on right to balance layout -->
      <div style="flex: 0 0 auto;"></div>
    </div>
        </div>
      `,
     



   footerTemplate: `
  <div style="
    font-size: 10pt; 
    width: 100%; 
    font-family: 'Times New Roman', serif; 
    padding: 0 0.42in; /* Match left/right page margins */
    box-sizing: border-box;
  ">
    <!-- Very visible black line -->
    <div style="
      border-top: 1px solid #000000;
      margin: 0 auto; /* Center the line */
      width: 100%; /* Control the width of the line */
      height: 1px;
    "></div>
    
    <div style="padding-top:2pt; display: flex; justify-content: space-between; align-items: center;">
       <div style="flex:1; text-align:left;">Page <span class="pageNumber"></span></div>
      <div style="flex: 1; text-align: center;">www.rsisinternational.org</div>
      <div style="flex: 1;"></div>
    </div>
  </div>
`,

   




      preferCSSPageSize: true,
      timeout: 60000,
    })

    console.log("[v0] PDF generated successfully, size:", pdfBuffer.length, "bytes")
    
    fs.writeFileSync("debug.html", htmlContent, "utf8");

    return pdfBuffer
  } catch (error) {
    console.error("[v0] PDF generation error:", error.message)
    console.error("[v0] Error stack:", error.stack)
    throw new Error(`PDF generation failed: ${error.message}`)
  } finally {
    if (browser) {
      console.log("[v0] Closing browser...")
      await browser.close()
    }
  }
}

export const exportToLatex = async (formattedDocument) => {
  return generateCompleteLatexDocument(formattedDocument)
}






// Convert a simple HTML string with <sup> tags into an array of docx TextRun objects
// Only handles plain text and <sup>...</sup> markers, which is sufficient for author lines
const supHtmlToRuns = (html, { size = 24, bold = true } = {}) => {
  if (!html || typeof html !== "string") return []
  const runs = []
  // Split on <sup>...</sup> while keeping the tags
  const parts = html.split(/(<sup>.*?<\/sup>)/gi)
  parts.forEach((part) => {
    if (!part) return
    const supMatch = part.match(/^<sup>(.*?)<\/sup>$/i)
    if (supMatch) {
      const text = supMatch[1]
      if (text) {
        runs.push(new TextRun({ text, superScript: true, size: Math.max(16, Math.round(size * 0.7)), bold }))
      }
    } else {
      // Remove any other HTML tags defensively
      const plain = part.replace(/<[^>]+>/g, "")
      if (plain) {
        runs.push(new TextRun({ text: plain, size, bold }))
      }
    }
  })
  return runs
}

// Normalization and keyword-based heading detection
const normalizeHeadingText = (s) =>
  String(s)
    .trim()
    .replace(/^\d+(?:\.\d+)*\s*[-.)]?\s*/, "") // strip leading numbering like 2., 2.1, 3.2.4)
    .replace(/\s+/g, " ") // collapse spaces
    .replace(/[:.;,\-\s]+$/, "") // strip trailing punctuation/colon
    .toLowerCase()

const HEADING_KEYWORDS = new Set([
  "conclusion and recommendation",
  "conclusion and recommendations",
  "conclusion",
  "background study",
  "literature review",
  "material",
  "materials",
  "methods",
  "mehtods",
  "result",
  "results",
  "discussion",
  "findings",
  "recommendation",
  "recommendations",
])

// DOCX: Only promote paragraphs that match explicit keywords
const isKeywordHeadingOnly = (text) => {
  if (!text) return false
  const raw = String(text).trim()
  if (raw.length === 0) return false
  const t = normalizeHeadingText(raw)
  return HEADING_KEYWORDS.has(t)
}

// Heuristic to detect heading-like lines within plain text
const isLikelyHeading = (text) => {
  if (!text) return false
  const raw = String(text).trim()
  if (raw.length === 0) return false
  const t = normalizeHeadingText(raw)

  // Explicit keyword match when the line is just that heading (optionally with numbering)
  if (HEADING_KEYWORDS.has(t)) return true

  // Numbered headings with short remainder
  if (/^\d+(?:\.\d+)*\s+/.test(raw) && t.length <= 80) return true

  // Short line without sentence-ending punctuation is likely a heading
  if (raw.length <= 80 && /[A-Za-z]/.test(raw) && !/[.!?]$/.test(raw)) return true

  return false
}

export const exportToDocx = async (formattedDocument) => {
  const {
    header = {},
    title = { text: "" },
    authors = { text: "" },
    affiliations = [],
    abstract = { heading: "ABSTRACT", content: "" },
    keywords = { heading: "Keywords", content: "" },
    sections = [],
    references = { heading: "REFERENCES", content: [] },
  } = formattedDocument || {}

  const docChildren = [];
  // ---------- Add spacing between header and title ----------
docChildren.push(
  new Paragraph({
    text: "", // blank line
    spacing: { after: 200 }, // ~10pt of space
  }),
);

  // ---------- Title ----------
  docChildren.push(
    new Paragraph({
      children: [new TextRun({ text: title.text, bold: true, size: 36 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
  );

  // ---------- Authors ----------
  docChildren.push(
    new Paragraph({
      children: supHtmlToRuns(authors.text, { size: 24, bold: true }),
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
  );

  // ---------- Affiliations ----------
  if (affiliations && affiliations.length > 0) {
    affiliations.forEach((aff) => {
      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({ text: String(aff.number), superScript: true, size: 18 }),
            new TextRun({ text: " " + aff.text, size: 24, bold: true }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
        }),
      );
    });
    docChildren.push(new Paragraph({ text: "" }));
  }



  // ---------- Abstract ----------
  if (abstract && abstract.content) {
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: abstract.heading,
            bold: true,
            allCaps: true,
            size: 28,
          }),
        ],
        spacing: { before: 400, after: 200 },
      }),
    );

    const abstractParagraphs = (abstract.content || "")
      .replace(/<\/p>/gi, "\n")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .split(/\n+/)
      .map((p) => p.trim())
      .filter((p) => p);

    abstractParagraphs.forEach((p, index) => {
      docChildren.push(
        new Paragraph({
          children: [new TextRun({ text: p, size: 24 })],
          alignment: AlignmentType.JUSTIFIED,
          spacing: {
            after: index === abstractParagraphs.length - 1 ? 300 : 240,
          },
        }),
      );
    });
  }

  // ---------- Keywords ----------
  if (keywords && keywords.content) {
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({ text: `${keywords.heading}: `, bold: true, size: 24 }),
          new TextRun({ text: keywords.content, italic: true, size: 24 }),
        ],
        spacing: { after: 300 },
      }),
    );
  }

  // ---------- Sections ----------
  sections.forEach((section) => {
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: section.heading,
            bold: true,
            allCaps: true,
            size: 28,
          }),
        ],
        spacing: { before: 400, after: 200 },
      }),
    );

    // const cleanText = section.content
    //   .replace(/<[^>]+>/g, " ")
    //   .split(/\n+/)
    //   .map((p) => p.trim())
    //   .filter((p) => p);

    // cleanText.forEach((p) => {
    //   docChildren.push(
    //     new Paragraph({
    //       text: p,
    //       alignment: AlignmentType.JUSTIFIED,
    //       spacing: { after: 240 },
    //     }),
    //   );
    // });

    // Split on <p>, <br>, or multiple newlines
const cleanParagraphs = section.content
  .replace(/<\/p>/gi, "\n") // treat </p> as new line
  .replace(/<br\s*\/?>/gi, "\n") // treat <br> as new line
  .replace(/<[^>]+>/g, "") // remove other HTML tags
  .split(/\n+/) // split on newlines
  .map((p) => p.trim())
  .filter((p) => p.length > 0);

cleanParagraphs.forEach((p) => {
  if (isKeywordHeadingOnly(p)) {
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: p, bold: true, allCaps: true, size: 28 })],
        spacing: { before: 400, after: 200 },
      }),
    );
  } else {
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: p, size: 24 })],
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 240 },
      }),
    );
  }
});
    if (section.subsections && section.subsections.length > 0) {
      section.subsections.forEach((subsection) => {
        docChildren.push(
          new Paragraph({
            children: [
              new TextRun({
                text: subsection.heading,
                bold: true,
                size: 24,
              }),
            ],
            spacing: { before: 240, after: 120 },
          }),
        );

        const subCleanParagraphs = (subsection.content || "")
          .replace(/<\/p>/gi, "\n")
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<[^>]+>/g, "")
          .split(/\n+/)
          .map((p) => p.trim())
          .filter((p) => p.length > 0);

        subCleanParagraphs.forEach((p) => {
          if (isKeywordHeadingOnly(p)) {
            docChildren.push(
              new Paragraph({
                children: [new TextRun({ text: p, bold: true, allCaps: true, size: 28 })],
                spacing: { before: 240, after: 120 },
              }),
            );
          } else {
            docChildren.push(
              new Paragraph({
                children: [new TextRun({ text: p, size: 24 })],
                alignment: AlignmentType.JUSTIFIED,
                spacing: { after: 240 },
              }),
            );
          }
        });
      });
    }
  });

  // ---------- References ----------
  if (references.content && references.content.length > 0) {
    docChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: references.heading,
            bold: true,
            allCaps: true,
            size: 28,
          }),
        ],
        spacing: { before: 400, after: 200 },
      }),
    );
    references.content.forEach((ref) => {
      docChildren.push(
        new Paragraph({
          children: [new TextRun({ text: `${ref.number}. ${ref.text}`, size: 24 })],
          alignment: AlignmentType.JUSTIFIED,
          indent: {
            left: convertInchesToTwip(0.25),
            hanging: convertInchesToTwip(0.25),
          },
          spacing: { after:0 },
        }),
      );
    });
  }

  // Build a minimal header without logo: right-aligned lines only
  const headerChildren = [];

  headerChildren.push(
    new Paragraph({
      children: [new TextRun({ text: header?.content || "", bold: true, size: 22 })],
      alignment: AlignmentType.RIGHT,
      spacing: { after: 50 },
    }),
  );

  const hp = [];
  if (header?.issn) hp.push(header.issn);
  if (header?.doi) hp.push(header.doi);
  if (header?.volume) hp.push(header.volume);
  if (hp.length > 0) {
    headerChildren.push(
      new Paragraph({
        children: [new TextRun({ text: hp.join(" | "), size: 20 })],
        alignment: AlignmentType.RIGHT,
      }),
    );
  }

const docxHeader = new Header({
  children: [
    ...headerChildren,
    new Paragraph({
      border: {
        bottom: {
          color: "000000",
          style: BorderStyle.SINGLE,
          size: 6,
        },
      },
    }),
  ],
});

// chnages started 

const docxFooter = new Footer({
  children: [
    new Paragraph({
      border: {
        top: {
          color: "000000",
          style: BorderStyle.SINGLE,
          size: 6,
        },
        spacing: { before: 0, after: 0 },
      },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Page ", color: "000000" }), new TextRun({ children: [PageNumber.CURRENT], color: "000000" })],
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0 }, // 👈 removes
    }),
    new Paragraph({
      children: [new TextRun({ text: "www.rsisinternational.org", size: 20, color: "000000" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 }, 
    }),
  ],
});
  // ---------- Document ----------
  const doc = new Document({
    sections: [
      {
        headers: { default: docxHeader },
        footers: { default: docxFooter },
        properties: {
          page: {
            pageNumbers: { start: formattedDocument.startPageNumber || 1 },
            margin: {
              top: convertInchesToTwip(0.76),
              right: convertInchesToTwip(0.42),
              bottom: convertInchesToTwip(0.42),
              left: convertInchesToTwip(0.42),
              footer: convertInchesToTwip(0.2), //
            },
          },
        },
        children: docChildren,
      },
    ],
    compatibility: { useOldTables: true },
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer;
};

export const validateExport = (format, content) => {
  const validFormats = ["html", "xml", "pdf", "latex", "docx"]

  if (!validFormats.includes(format.toLowerCase())) {
    throw new Error(`Unsupported export format: ${format}`)
  }

  if (!content || (typeof content === "string" && content.trim().length === 0)) {
    throw new Error("No content to export")
  }

  return true
}
