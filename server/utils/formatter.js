const FONT_FAMILY_TIMES = "font-family: 'Times New Roman';"
const FONT_SIZE_10 = "font-size: 10pt;"
const FONT_SIZE_12 = "font-size: 12pt;"
const FONT_SIZE_14 = "font-size: 14pt;"
const FONT_WEIGHT_BOLD = "font-weight: bold;"
const TEXT_ALIGN_CENTER = "text-align: center;"
const TEXT_ALIGN_JUSTIFY = "text-align: justify;"
const LINE_HEIGHT_1 = "line-height: 1.0;"


// Escape HTML safely inside cells
const escapeHtml = (s = "") =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export const formatToAcademicStandard = async (processedData) => {
  try {
    console.log("[v0] Starting academic formatting...")
    console.log("[v0] Input data structure:", Object.keys(processedData))

    // Default to blank-safe values so missing fields never break downstream exporters
    const title = processedData.title?.text ?? processedData.title ?? ""
    const authors = processedData.authors?.text ?? processedData.authors ?? ""
    const abstract = processedData.abstract ?? { heading: "ABSTRACT", content: "" }
    const keywords = processedData.keywords ?? { heading: "Keywords", content: "" }
    const sections = processedData.sections ?? []
    const references = processedData.references ?? { heading: "REFERENCES", content: [] }
    const metadata = processedData.metadata ?? {}
    const originalFilename = processedData.originalFilename ?? "document"
    const publicationInfo = processedData.publicationInfo ?? null
    const startPage = processedData.startPageNumber ?? 1;

    console.log("[v0] Extracted title:", title)
    console.log("[v0] Extracted authors:", authors)
    console.log("[v0] Sections count:", sections.length)
     

        console.log("[v0] Sections count:", sections.length)

    const parseAuthorsAndAffiliations = (authorText) => {
      if (!authorText || typeof authorText !== "string") return { authors: "", affiliations: [] }

      const lines = authorText
        .split("\n")
        .map((l) => l.trim())
        .filter((l) => l)

      const authorLines = []
      const affiliationLines = []

      let contentStarted = false
      for (const line of lines) {
        // Heuristic: A long line with spaces is likely the start of the abstract.
        if (line.length > 150 && line.includes(" ")) {
          contentStarted = true
        }

        // Heuristic: Abstract often starts with "abstract" or "keywords".
        if (line.toLowerCase().startsWith("abstract") || line.toLowerCase().startsWith("keywords")) {
          contentStarted = true
        }

        if (contentStarted) {
          break // Stop processing at the beginning of the abstract
        }

        // Heuristic: Affiliations contain keywords.
        if (line.match(/department|school|university|college|institute|@/i)) {
          affiliationLines.push(line)
        } else {
          // Otherwise, assume it's part of the author list.
          if (line.length < 150) {
            authorLines.push(line)
          }
        }
      }

      // Join authors with a comma, and remove any trailing comma.
      const authors = authorLines.join(", ").replace(/,$/, "").trim()

      // Process unique affiliations
      const uniqueAffiliations = [...new Set(affiliationLines)]
      const affiliations = uniqueAffiliations.map((aff, index) => ({
        number: index + 1,
        text: aff,
        style: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${TEXT_ALIGN_CENTER} margin-bottom: 6pt;`,
      }))
      return { authors, affiliations }
    }



    // const parseAuthorsAndAffiliations = (authorText) => {
    //   if (!authorText) return { authors: "", affiliations: [] }

    //   // Split by common patterns that indicate affiliation
    //   const lines = authorText.split(/\n|Department|School|University|College|Institute/).filter((line) => line.trim())

    //   if (lines.length === 1) {
    //     // Single line - likely just authors
    //     return { authors: lines[0].trim(), affiliations: [] }
    //   }

    //   // First line is usually authors, rest are affiliations
    //   const authors = lines[0].trim()
    //   const affiliations = lines
    //     .slice(1)
    //     .map((line) => line.trim())
    //     .filter((line) => line.length > 0)
    //     .map((aff, index) => ({
    //       number: index + 1,
    //       text: aff.startsWith("Department") || aff.startsWith("School") ? aff : `Department ${aff}`,
    //       style: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${TEXT_ALIGN_CENTER} margin-bottom: 6pt;`,
    //     }))

    //   return { authors, affiliations }
    // }

    const { authors: parsedAuthors, affiliations: parsedAffiliations } = parseAuthorsAndAffiliations(authors)

    // Prefer parsed authors when available; otherwise fall back to original input
    const authorsForFormatting = (typeof parsedAuthors === "string" && parsedAuthors.trim().length > 0) ? parsedAuthors : authors

    // Build affiliations from parsed data if present; otherwise derive from original input (when it's an array of author objects)
    const affiliationsFinal = (Array.isArray(parsedAffiliations) && parsedAffiliations.length > 0)
      ? parsedAffiliations
      : formatAffiliations(Array.isArray(authors) ? authors : [])

    // Academic formatting specifications from the requirements
    const formatSpecs = {
      pageSize: "A4",
      margins: {
        top: "0.76in",
        bottom: "0.42in",
        left: "0.42in",
        right: "0.42in",
      },
      font: {
        family: "Times New Roman",
        size: "12pt",
        lineHeight: "1.0",
      },
      spacing: {
        paragraph: "12pt",
      },
      header: {
        distance: "0.24in",
        content: "INTERNATIONAL JOURNAL OF RESEARCH AND INNOVATION IN SOCIAL SCIENCE (IJRISS)",
        issn: "ISSN No. 2454-6186",
        doi: "DOI: 10.47772/IJRISS",
        volume: "Volume IX Issue VIII August 2025",
      },
      footer: {
        distance: "0.28in",
      },
    }

    const formatAuthors = (authorsList) => {
      if (!authorsList) return ""

      // Handle string format
      if (typeof authorsList === "string") {
        // Normalize common inline superscript patterns in author lines
        // 1) Caret syntax: Name^1^ or Name^*^
        // 2) Bracket syntax: Name[1] or Name[*]
        // 3) Trailing digit/asterisk stuck to the name: "Mungai1" or "Doe*"
        return authorsList
          .replace(/\^(\*|\d+)\^/g, "<sup>$1</sup>")
          .replace(/\[(\*|\d+)\]/g, "<sup>$1</sup>")
          .replace(/([A-Za-z])((?:\d{1,2}|\*))(?![\w<])/g, "$1<sup>$2</sup>")
      }

      // Handle array format: [{ name, affiliation, corresponding }]
      if (Array.isArray(authorsList)) {
        // Build a stable affiliation index based on first appearance order
        const affIndex = new Map()
        let nextIdx = 1

        const getAffNumber = (aff) => {
          if (!aff) return null
          if (!affIndex.has(aff)) {
            affIndex.set(aff, nextIdx++)
          }
          return affIndex.get(aff)
        }

        return authorsList
          .map((author, index) => {
            if (typeof author === "string") {
              // Normalize any inline markers inside raw strings
              return author
                .replace(/\^(\*|\d+)\^/g, "<sup>$1</sup>")
                .replace(/\[(\*|\d+)\]/g, "<sup>$1</sup>")
            }

            let authorText = author.name || `Author ${index + 1}`

            // Add superscript number mapped from affiliation
            const n = getAffNumber(author.affiliation)
            if (n) {
               authorText += `<sup>${n}</sup>`
            }

            // Mark corresponding author with asterisk
            if (author.corresponding) {
              authorText += "<sup>*</sup>"
            }

            return authorText
          })
          .join(", ")
      }

      return String(authorsList)
    }

    const formatSections = (sectionsList) => {
      if (!sectionsList || sectionsList.length === 0) return []

      return sectionsList.map((section) => {
        const sectionType = section.type?.toLowerCase() || "other"

        return {
          heading: section.heading || "Untitled Section",
          content: preserveOriginalFormatting(section.content || ""),
          type: sectionType,
          formatting: {
            headingStyle: getHeadingStyle(sectionType),
            contentStyle: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1} ${TEXT_ALIGN_JUSTIFY} margin-bottom: 12pt;`,
          },
          subsections:
            section.subsections?.map((sub) => ({
              heading: sub.heading,
              content: preserveOriginalFormatting(sub.content || ""),
              formatting: {
                headingStyle: `${FONT_SIZE_12} ${FONT_WEIGHT_BOLD} ${FONT_FAMILY_TIMES} margin-top: 12pt; margin-bottom: 6pt;`,
                contentStyle: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1} ${TEXT_ALIGN_JUSTIFY} margin-bottom: 12pt;`,
              },
            })) || [],
        }
      })
    }

    const preserveOriginalFormatting = (content) => {
      if (!content) return "";
      // The previous implementation was flawed and broke paragraph structure.
      // The AI model is already instructed to preserve original formatting, including
      // paragraph breaks (which are represented as double newlines).
      // By simply passing the content through, we trust the AI's output and
      // allow the exporter to handle the rendering correctly and robustly.
      return content
    .split(/\n\s*\n/) // split on double newlines (paragraphs)
    .map((para) => `<p style="${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1} ${TEXT_ALIGN_JUSTIFY} margin-bottom: 12pt;">${para.trim()}</p>`)
    .join("");
    }

    const getHeadingStyle = (sectionType) => {
      // Treat all headings uniformly: 14pt, bold, uppercase
      return `${FONT_SIZE_14} ${FONT_WEIGHT_BOLD} text-transform: uppercase; ${FONT_FAMILY_TIMES} text-align: left; margin-top: 20pt; margin-bottom: 10pt;`
    }

    const formatReferences = (referencesList) => {
      if (!referencesList) return []

      // Handle object with content property
      let refsToProcess = referencesList

      if (typeof referencesList === "object" && referencesList.content) {
        refsToProcess = referencesList.content
      }

      if (Array.isArray(refsToProcess)) {
        return refsToProcess
          .map((ref) => {
            // Handle object references properly
            if (typeof ref === "object" && ref !== null) {
              return ref.text || ref.content || JSON.stringify(ref)
            }
            return String(ref)
          })
          .filter((ref) => ref && ref.trim().length > 10 && !ref.includes("[object Object]"))
          .map((ref, index) => ({
            number: index + 1,
            text: ref.trim(),
            style: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1} ${TEXT_ALIGN_JUSTIFY} margin-bottom: 6pt; text-indent: -0.5in; padding-left: 0.5in;`,
          }))
      }

      // If it's a string, split by common reference separators
      if (typeof refsToProcess === "string" && refsToProcess.trim().length > 0) {
        return refsToProcess
          .split(/\n|\d+\./)
          .filter((ref) => ref.trim().length > 10)
          .map((ref, index) => ({
            number: index + 1,
            text: ref.trim(),
            style: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1} ${TEXT_ALIGN_JUSTIFY} margin-bottom: 6pt; text-indent: -0.5in; padding-left: 0.5in;`,
          }))
      }

      return []
    }

    function formatAffiliations(authorsList) {
      if (!authorsList || !Array.isArray(authorsList)) return []

      const affiliations = []
      const uniqueAffiliations = new Set()

      authorsList.forEach((author) => {
        if (typeof author === "object" && author.affiliation && !uniqueAffiliations.has(author.affiliation)) {
          uniqueAffiliations.add(author.affiliation)
          affiliations.push({
            number: affiliations.length + 1,
            text: author.affiliation,
            style: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${TEXT_ALIGN_CENTER} margin-bottom: 6pt;`,
          })
        }
      })

      return affiliations
    }




    const cleanAbstractContent = (abstractContent) => {
      if (!abstractContent) return ""

      const cleaned = abstractContent
        // Remove department/affiliation text that got mixed in
        .replace(/Department.*?Kenya/gi, "")
        // Remove date patterns
        .replace(/\b\d{1,2}\s+(July|August|September|October|November|December)\s+\d{4}/gi, "")
        // Remove author names that got mixed in
        .replace(/Nancy Njeri Mungai.*?Yusuf Kibet/gi, "")
        // Remove DOI patterns
        .replace(/DOI:.*?(?=\n|$)/gi, "")
        // Remove publication info patterns
        .replace(/Received:.*?Published:.*?(?=\n|$)/gi, "")
        // .replace(/^(Department|School|University|College|Institute).*?(?=This paper|The study|This research)/s, "")
        // Remove duplicate abstract content that appears after title
        // Remove duplicated abstract that appears before the actual ABSTRACT heading
        .replace(/^Department.*?(?=ABSTRACT)/is, "")


        .trim()

      // Validate abstract content
      if (
        cleaned.length < 100 ||
        cleaned.match(/^[a-z]{1,5}$/i) ||
        cleaned === "hn" ||
        cleaned.includes("[object Object]")
      ) {
        return ""
      }

      return cleaned
    }

    // Create formatted document structure
    const formattedDocument = {
      metadata: {
        ...metadata,
        formatSpecs,
        formattedAt: new Date().toISOString(),
        originalFilename,
        formatting: "Academic Standard - IJRISS Template",

      },
      header: {
        content: formatSpecs.header.content,
        issn: formatSpecs.header.issn,
        doi: formatSpecs.header.doi,
        volume: formatSpecs.header.volume,
        editable: true,
        style: `${TEXT_ALIGN_CENTER} ${FONT_SIZE_10} ${FONT_FAMILY_TIMES} ${FONT_WEIGHT_BOLD} border-bottom: 1px solid #000; padding-bottom: 10pt; margin-bottom: 20pt;`,
      },
      title: {
        text: title,
        style: `font-size: 18pt; ${FONT_WEIGHT_BOLD} ${TEXT_ALIGN_CENTER} ${FONT_FAMILY_TIMES} margin: 20pt 0; line-height: 1.2;`,
      },
      authors: {
        text: formatAuthors(authorsForFormatting),
        style: `${FONT_SIZE_12} ${FONT_WEIGHT_BOLD} ${TEXT_ALIGN_CENTER} ${FONT_FAMILY_TIMES} margin: 10pt 0;`,
      },
      affiliations: affiliationsFinal,
      abstract: {
        heading: abstract?.heading || "ABSTRACT",
        content: cleanAbstractContent(abstract?.content || ""),
        style: {
          heading: `${FONT_SIZE_14} ${FONT_WEIGHT_BOLD} text-transform: uppercase; ${FONT_FAMILY_TIMES} text-align: left; margin-top: 20pt; margin-bottom: 10pt;`,
          content: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1} ${TEXT_ALIGN_JUSTIFY} margin-bottom: 15pt;`,
        },
      },
      keywords: {
        heading: keywords?.heading || "Keywords",
        content: Array.isArray(keywords?.content) ? keywords.content.join(", ") : keywords?.content || "",
        style: {
          heading: `${FONT_SIZE_12} ${FONT_WEIGHT_BOLD} ${FONT_FAMILY_TIMES} display: inline; margin-right: 5pt;`,
          content: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} font-style: italic; display: inline;`,
        },
      },
      sections: formatSections(sections),
      references: {
        heading: references?.heading || "REFERENCES",
        content: formatReferences(references),
        style: {
          heading: `${FONT_SIZE_14} ${FONT_WEIGHT_BOLD} text-transform: uppercase; ${FONT_FAMILY_TIMES} text-align: left; margin-top: 20pt; margin-bottom: 10pt;`,
          content: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${LINE_HEIGHT_1}`,
        },
      },                                                                          
      footer: {
        content: `Page {pageNumber}                                               www.rsisinternational.org`,
        style: `${FONT_SIZE_10} ${FONT_FAMILY_TIMES} ${TEXT_ALIGN_CENTER} border-top: 1px solid #000; padding-top: 5pt ; margin-top: 5pt;`,
      },
      publicationInfo: publicationInfo || {
        doi: `https://dx.doi.org/10.47772/IJRISS.2025.908000330`,
        received: "29 July 2025",
        accepted: "08 August 2025",
        published: "09 September 2025",
        style: `${FONT_SIZE_12} ${FONT_FAMILY_TIMES} ${TEXT_ALIGN_CENTER} margin: 15pt 0;`,
      },
      startPageNumber: startPage,

    }

    console.log("[v0] Academic formatting completed successfully")
    return formattedDocument
  } catch (error) {
    console.error("[v0] Formatting error:", error.message)
    throw new Error(`Academic formatting failed: ${error.message}`)
  }
}

export const applyLatexFormatting = (content) => {
  if (!content) return ""

  return (
    content
      // Preserve paragraph indentation
      .replace(/^\s+/gm, (match) => "&nbsp;".repeat(match.length))
      // Handle citations
      .replace(/\[(\d+)\]/g, "<sup>$1</sup>")
      // Handle emphasis
      .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
      .replace(/\*(.*?)\*/g, "<em>$1</em>")
      // Handle equations (basic)
      .replace(/\$\$(.*?)\$\$/g, '<div class="equation">$1</div>')
      .replace(/\$(.*?)\$/g, '<span class="inline-equation">$1</span>')
  )
}

export const formatCitations = (content) => {
  if (!content) return ""

  // Convert various citation formats to numbered format
  return content
    .replace(/$$([^)]+),\s*(\d{4})$$/g, "[$1, $2]") // (Author, Year) -> [Author, Year]
    .replace(/([A-Z][a-z]+\s+et\s+al\.,?\s*\d{4})/g, "[$1]") // Author et al., Year -> [Author et al., Year]
}

