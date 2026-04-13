// src/module/docx-engine.ts - docx 读写引擎

import {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, UnderlineType, PageOrientation,
    type IRunOptions, type IParagraphOptions,
    Header,
} from "docx";

export type DocxContent = {
    paragraphs: DocxParagraph[];
    images?: DocxImage[];
};

export type DocxParagraph = {
    text: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    heading?: number; // 1-5
    alignment?: "left" | "center" | "right";
    fontSize?: number; // in points
    fontColor?: string; // hex
    fontName?: string;
};

export type DocxImage = {
    id: string;
    width: number; // in EMU (English Metric Units)
    height: number; // in EMU
    data: string;   // base64 encoded image bytes (without data URI prefix)
    altText?: string;
};

export type DocxDocument = {
    content: DocxContent;
    filePath?: string;
};

/**
 * 从 ArrayBuffer 读取 docx 并提取文本内容（含样式、图片）
 * docx 本质是 ZIP 文件，需要解压后解析 XML
 */
export const readDocx = async (buffer: ArrayBuffer): Promise<DocxDocument> => {
    const bytes = new Uint8Array(buffer);

    // Check if this looks like a ZIP file (PK magic bytes)
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4B) {
        console.warn("readDocx: Not a valid ZIP/DOCX file (bad magic bytes)");
        return { content: { paragraphs: [], images: [] } };
    }

    // Extract document.xml
    let documentXml: Uint8Array | null = null;
    try {
        documentXml = await extractZipEntry(bytes, "word/document.xml");
    } catch (e) {
        console.error("readDocx: Failed to extract document.xml:", e);
    }

    if (!documentXml || documentXml.length === 0) {
        console.warn("readDocx: document.xml not found or empty");
        return { content: { paragraphs: [], images: [] } };
    }

    // Try to extract relationships for image detection
    let relsXml: Uint8Array | null = null;
    try {
        relsXml = await extractZipEntry(bytes, "word/_rels/document.xml.rels");
    } catch {
        relsXml = null; // non-critical, ignore
    }

    // Extract content types
    let contentTypesXml: Uint8Array | null = null;
    try {
        contentTypesXml = await extractZipEntry(bytes, "[Content_Types].xml");
    } catch {
        contentTypesXml = null;
    }

    const xmlStr = new TextDecoder("utf-8", { fatal: false }).decode(documentXml);
    if (!xmlStr || xmlStr.trim().length === 0) {
        console.warn("readDocx: document.xml is empty after decoding");
        return { content: { paragraphs: [] } };
    }

    let parsedDoc: DocxDocument;
    try {
        parsedDoc = parseDocxXml(xmlStr);
    } catch (e) {
        console.error("readDocx: Failed to parse XML:", e);
        parsedDoc = { content: { paragraphs: [] } };
    }

    // Extract images if present (non-critical, best-effort)
    let images: DocxImage[] = [];
    if (relsXml && contentTypesXml) {
        try {
            const relsStr = new TextDecoder("utf-8", { fatal: false }).decode(relsXml);
            const ctStr = new TextDecoder("utf-8", { fatal: false }).decode(contentTypesXml);
            images = await extractImages(bytes, relsStr, ctStr);
        } catch (e) {
            console.warn("readDocx: Image extraction failed (non-critical):", e);
        }
    }

    return {
        content: {
            ...parsedDoc.content,
            images,
        },
        filePath: undefined,
    };
};

/**
 * Parse docx XML content into structured paragraphs with full style support
 * Handles: bold, italic, underline, strike, font size, color, heading, alignment
 * Handles: inline images (drawings)
 */
function parseDocxXml(xml: string): DocxDocument {
    const paragraphs: DocxParagraph[] = [];

    if (!xml || typeof xml !== "string") return { content: { paragraphs: [] } };

    // Normalize XML - remove namespace prefixes for easier matching
    let normalized: string;
    try {
        normalized = normalizeDocxXml(xml);
    } catch {
        normalized = xml; // fallback to raw xml
    }

    // Split by paragraph closing tags
    const pParts = normalized.split(/<\/w:p>/);

    for (const part of pParts) {
        const trimmed = part.trim();
        if (!trimmed || !trimmed.includes("<w:p")) continue;

        // --- Paragraph-level properties ---
        // Heading style from pPr/pStyle - handle multiple patterns
        let headingMatch = trimmed.match(/<w:pStyle[^>]*w:val="[^"]*Heading(\d)[^"]*"/i);
        if (!headingMatch) {
            // Try alternative pattern where val comes before Heading
            headingMatch = trimmed.match(/<w:pStyle[^>]*val="[^"]*Heading(\d)[^"]*"[^>]*>/i);
        }
        if (!headingMatch) {
            // Try pattern with just Heading number
            headingMatch = trimmed.match(/<w:pStyle[^>]+Heading(\d)/i);
        }
        const heading = headingMatch ? parseInt(headingMatch[1] ?? "1") : undefined;
        if (heading && isNaN(heading)) continue;

        // Alignment
        const alignMatch = trimmed.match(/<w:jc[^>]*?:?val="?(center|right|left|start|end|both|distributed)"?/i);
        let alignment: "left" | "center" | "right" | undefined;
        if (alignMatch) {
            const val = alignMatch[1];
            if (val === "center") alignment = "center";
            else if (val === "right" || val === "end") alignment = "right";
            else alignment = "left";
        }

        // --- Run-level extraction ---
        // Split by run closing tags
        const rParts = trimmed.split(/<\/w:r>/);
        const runs: { text: string; styles: Partial<DocxParagraph> }[] = [];

        for (const rPart of rParts) {
            if (!rPart || !rPart.includes("<w:r")) continue;

            const runStyles: Partial<DocxParagraph> = {};

            // Extract run properties block - be more flexible with matching
            let rPrStart = rPart.indexOf("<w:rPr");
            // Handle self-closing tag <w:rPr/>
            if (rPrStart !== -1) {
                const nextCharIdx = rPrStart + 5;
                // Check if it's a valid opening or self-closing
                if (rPart[nextCharIdx] === "/" || rPart.substring(nextCharIdx, nextCharIdx + 2).includes("/>")) {
                    // Self-closing, no styles inside
                    rPrStart = -1;
                }
            }
            let rPrEnd = rPart.indexOf("</w:rPr>");
            if (rPrStart !== -1 && rPrEnd > rPrStart) {
                const rPr = rPart.substring(rPrStart, rPrEnd + 7); // include </w:rPr>

                // Bold
                if (/<w:b[\s>]/i.test(rPr)) runStyles.bold = true;
                else if (/<w:b[^>]*w:val="false"/i.test(rPr)) runStyles.bold = false;

                // Italic
                if (/<w:i[\s>]/i.test(rPr)) runStyles.italic = true;

                // Underline
                if (/<w:u[\s>]/i.test(rPr)) runStyles.underline = true;

                // Strike / Double-strike-through
                if (/<w:strike[\s>]|<w:strike\/>/i.test(rPr)) runStyles.strike = true;
                if (/<w:dstrike/i.test(rPr)) runStyles.strike = true;

                // Font size (in half-points) - also check w:szCs for complex script
                let szMatch = rPr.match(/<w:sz[^>]*?:?val="?(\d+)"?/i);
                if (!szMatch) szMatch = rPr.match(/<w:szCs[^>]*?:?val="?(\d+)"?/i);
                if (szMatch) {
                    const parsedSize = parseInt(szMatch[1]);
                    if (!isNaN(parsedSize)) runStyles.fontSize = parsedSize / 2;
                }

                // Font color
                const colorMatch = rPr.match(/<w:color[^>]*?:?val="?([0-9A-Fa-f]{6}|auto)"?/i);
                if (colorMatch && colorMatch[1] !== "auto") runStyles.fontColor = "#" + colorMatch[1].toLowerCase();

                // Font name - multiple attributes possible
                let fontMatch = rPr.match(/<w:rFonts[^>]*?:?ascii="([^"]+)"/i);
                if (!fontMatch) fontMatch = rPr.match(/<w:rFonts[^>]*?:?hAnsi="([^"]+)"/i);
                if (fontMatch) runStyles.fontName = fontMatch[1];
            }

            // Extract all <w:t> text nodes in this run (handle multiple per run, preserve spaces)
            let runText = "";
            const tRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
            let tMatch: RegExpExecArray | null;
            while ((tMatch = tRegex.exec(rPart)) !== null) {
                if (tMatch[1] != null) runText += tMatch[1];
            }
            // Also try <w:delText> for tracked deletions (treat as normal text)
            const dtRegex = /<w:delText(?:\s[^>]*)?>([^<]*)<\/w:delText>/g;
            let dtMatch: RegExpExecArray | null;
            while ((dtMatch = dtRegex.exec(rPart)) !== null) {
                if (dtMatch[1] != null) runText += dtMatch[1];
            }

            if (runText.length > 0) {
                runs.push({ text: runText, styles: runStyles });
            }
        }

        // Merge runs into a single paragraph
        if (runs.length > 0) {
            const fullText = runs.map((r) => r.text).join("");
            // Use first non-empty style set as the paragraph's dominant style
            const dominantStyle = runs.find((r) => Object.keys(r.styles).length > 0)?.styles || {};
            paragraphs.push({
                text: fullText,
                ...dominantStyle,
                heading,
                alignment,
            });
        }
    }

    return { content: { paragraphs } };
}

/** Normalize docx XML by stripping namespaces for simpler regex matching */
function normalizeDocxXml(xml: string): string {
    // Only strip xmlns declarations, keep w: prefixes so regexes with <w:... still work
    return xml.replace(/xmlns(:\w+)?="[^"]*"/g, "");
}

/**
 * Extract embedded images from the DOCX archive
 */
async function extractImages(zipBytes: Uint8Array, relsXml: string, _contentTypesXml: string): Promise<DocxImage[]> {
    const images: DocxImage[] = [];

    // Parse relationships to find image targets
    const imgRels: { id: string; target: string }[] = [];
    const relRegex = /<Relationship\s+Id="([^"]*)"\s+Type="[^"]*image[^"]*"\s+Target="([^"]*)"/gi;
    let match: RegExpExecArray | null;
    while ((match = relRegex.exec(relsXml)) !== null) {
        imgRels.push({ id: match[1], target: match[2] });
    }

    // For each relationship, extract the image data
    for (const rel of imgRels) {
        // Target is like "media/image1.png" - needs word/ prefix for zip path
        const zipPath = "word/" + rel.target;
        const imageData = await extractZipEntry(zipBytes, zipPath);
        if (imageData) {
            // Convert to base64
            const b64 = arrayBufferToBase64(imageData.buffer as ArrayBuffer || imageData);
            images.push({
                id: rel.id,
                width: 4000000, // default ~100mm
                height: 3000000, // default ~75mm
                data: b64,
                altText: rel.target.split("/").pop() || "image",
            });
        }
    }

    return images;
}

/**
 * Extract a file from a ZIP archive ( ArrayBuffer )
 * Minimal ZIP parser - handles STORED and DEFLATE entries
 * Also handles entries with data descriptor (bit 3 of general purpose flag)
 */
async function extractZipEntry(data: Uint8Array, targetPath: string): Promise<Uint8Array | null> {
    let offset = 0;
    const view = new DataView(data.buffer);
    const len = data.length;

    while (offset < len - 4) {
        // Check for local file header signature (0x04034b50)
        const sig = view.getUint32(offset, true);
        if (sig !== 0x04034b50) {
            // Try scanning forward for next signature (handle corrupted or non-standard zips)
            offset++;
            continue;
        }

        // General purpose bit flag
        const gpFlag = view.getUint16(offset + 6, true);
        // Compression method
        const compressionMethod = view.getUint16(offset + 8, true);

        // For entries with data descriptor, sizes are in the descriptor after data
        // We need to check if sizes are zero (data descriptor present) or actual values
        let compressedSize = view.getUint32(offset + 18, true);
        let uncompressedSize = view.getUint32(offset + 22, true);

        const fileNameLen = view.getUint16(offset + 26, true);
        const extraFieldLen = view.getUint16(offset + 28, true);

        const fileNameBytes = data.slice(offset + 30, offset + 30 + fileNameLen);
        const fileName = new TextDecoder("utf-8", { fatal: false }).decode(fileNameBytes).replace(/\0.*$/, "");

        const dataOffset = offset + 30 + fileNameLen + extraFieldLen;

        // If data descriptor flag is set (bit 3), we may have 0/0 for sizes
        // In this case we find the next header or end of file
        if ((gpFlag & (1 << 3)) !== 0 && compressedSize === 0 && uncompressedSize === 0) {
            // Search for next signature (local file header 0x04034b50 or central dir 0x02014b50)
            let searchOffset = dataOffset;
            while (searchOffset < len - 4) {
                const nextSig = view.getUint32(searchOffset, true);
                if (nextSig === 0x04034b50 || nextSig === 0x02014b50) {
                    break;
                }
                searchOffset++;
            }
            compressedSize = searchOffset - dataOffset;
            // For deflate with unknown uncompressed size, use compressed size as estimate
            uncompressedSize = Math.max(compressedSize * 4, 65536); // rough upper bound
        }

        // Safety bounds check
        if (dataOffset < 0 || dataOffset > len) break;
        const safeCompressedSize = Math.min(compressedSize, len - dataOffset);
        if (safeCompressedSize <= 0) return null;

        const compressedData = data.slice(dataOffset, dataOffset + safeCompressedSize);

        if (fileName === targetPath || fileName.replace(/\\/g, "/") === targetPath) {
            if (compressionMethod === 0) {
                // STORED - no compression
                return new Uint8Array(compressedData);
            } else if (compressionMethod === 8) {
                // DEFLATE
                try {
                    return await inflateData(compressedData, Math.max(uncompressedSize, safeCompressedSize));
                } catch (e) {
                    console.error("Failed to inflate entry:", fileName, e);
                    return null;
                }
            }
            return null; // unsupported compression method
        }

        // Move past this entry's data
        const nextOffset = dataOffset + compressedSize;
        if (nextOffset >= len) break;
        offset = nextOffset;
    }

    return null;
}

async function inflateData(data: Uint8Array, expectedSize: number): Promise<Uint8Array> {
    // Try Bun's built-in decompress first (most reliable in Bun runtime)
    try {
        // Bun supports Buffer-like operations on Uint8Array
        const result = await Bun.write(
            new Uint8Array(expectedSize),
            new Response(data).body! as unknown as BodyInit,
        );
        if (result && result instanceof Uint8Array && result.length > 0) {
            return result;
        }
    } catch {
        // fall through to other methods
    }

    // Method 2: Use DecompressionStream if available
    try {
        if (typeof DecompressionStream !== "undefined") {
            const ds = new DecompressionStream("deflate-raw");
            const writer = ds.writable.getWriter();
            await writer.write(data as unknown as BufferSource);
            await writer.close();

            const reader = ds.readable.getReader();
            const chunks: Uint8Array[] = [];

            while (true) {
                const { done, value } = await reader.read();
                if (done) break;
                chunks.push(value);
            }

            if (chunks.length > 0) {
                const totalLen = chunks.reduce((sum, c) => sum + c.length, 0);
                const result = new Uint8Array(totalLen);
                let pos = 0;
                for (const chunk of chunks) {
                    result.set(chunk, pos);
                    pos += chunk.length;
                }
                return result;
            }
        }
    } catch {
        // fall through
    }

    // Method 3: Manual fallback using Bun's zlib via dynamic import
    try {
        // In some Bun versions we can use node:zlib
        const zlib = await import("node:zlib");
        const buf = Buffer.from(data);
        const decompressed = zlib.inflateRawSync(buf);
        return new Uint8Array(decompressed.buffer, decompressed.byteOffset, decompressed.byteLength);
    } catch {
        // All methods failed
        throw new Error("Unable to decompress data: no suitable decompression method available");
    }
}

/**
 * 将 DocxDocument 写入 ArrayBuffer（保留样式，兼容 WPS/Word 预览）
 */
export const writeDocx = async (doc: DocxDocument): Promise<ArrayBuffer> => {
    const paragraphs: Paragraph[] = doc.content.paragraphs.map((p) => {
        const runOpts: IRunOptions = {
            text: p.text,
            bold: p.bold,
            italics: p.italic,
            underline: p.underline ? { type: UnderlineType.SINGLE } : undefined,
            strike: p.strike,
            size: p.fontSize ? Math.round(p.fontSize * 2) : undefined, // docx uses half-points
            color: p.fontColor,
            font: p.fontName ? { name: p.fontName } : undefined,
        };

        const paraOpts: IParagraphOptions = {
            heading: p.heading
                ? (p.heading as unknown as typeof HeadingLevel.HEADING_1)
                : undefined,
            alignment: p.alignment === "center"
                ? AlignmentType.CENTER
                : p.alignment === "right"
                    ? AlignmentType.RIGHT
                    : AlignmentType.LEFT,
        };

        return new Paragraph({
            ...paraOpts,
            children: [new TextRun(runOpts)],
        });
    });

    // Ensure at least one empty paragraph for valid document structure
    const sectionChildren = paragraphs.length > 0
        ? paragraphs
        : [new Paragraph({ children: [new TextRun("")] })];

    // Add images after paragraphs if present
    if (doc.content.images && doc.content.images.length > 0) {
        // Note: ImageRun in docx library requires buffer input; we'd need to decode base64
        // For now, we include image metadata but skip actual embedding in write path
        // This can be enhanced later with proper image embedding support
    }

    const d = new Document({
        creator: "xlsx-cli",
        description: "Created by xlsx-cli",
        title: doc.filePath || "Untitled",
        sections: [{
            properties: {
                page: {
                    margin: {
                        top: 1440,     // 1 inch (twips)
                        right: 1440,
                        bottom: 1440,
                        left: 1440,
                    },
                    size: {
                        orientation: PageOrientation.PORTRAIT,
                        width: 12240,  // Letter width (8.5 inches)
                        height: 15840, // Letter height (11 inches)
                    },
                },
            },
            children: sectionChildren,
        }],
    });

    const buf = await Packer.toBuffer(d);
    return buf.buffer as ArrayBuffer;
};

/**
 * 添加段落
 */
export const addParagraph = (doc: DocxDocument, paragraph: DocxParagraph): DocxDocument => {
    return {
        ...doc,
        content: {
            paragraphs: [...doc.content.paragraphs, paragraph],
        },
    };
};

/**
 * 更新段落
 */
export const updateParagraph = (doc: DocxDocument, index: number, paragraph: DocxParagraph): DocxDocument => {
    const paragraphs = [...doc.content.paragraphs];
    if (index >= 0 && index < paragraphs.length) {
        paragraphs[index] = paragraph;
    }
    return { ...doc, content: { paragraphs } };
};

/**
 * 删除段落
 */
export const deleteParagraph = (doc: DocxDocument, index: number): DocxDocument => {
    const paragraphs = [...doc.content.paragraphs];
    if (index >= 0 && index < paragraphs.length) {
        paragraphs.splice(index, 1);
    }
    return { ...doc, content: { paragraphs } };
};

// --- helpers ---

function arrayBufferToBase64(buffer: ArrayBuffer): string {
    const bytes = new Uint8Array(buffer);
    let binary = "";
    const chunkSize = 8192;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode(...bytes.subarray(i, Math.min(i + chunkSize, bytes.length)));
    }
    return btoa(binary);
}
