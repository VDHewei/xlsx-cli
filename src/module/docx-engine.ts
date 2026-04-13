// src/module/docx-engine.ts - docx 读写引擎

import {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, UnderlineType,
    type IRunOptions, type IParagraphOptions,
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
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4B) {
        return { content: { paragraphs: [], images: [] } };
    }

    // Extract document.xml and [Content_Types].xml for image detection
    const documentXml = await extractZipEntry(bytes, "word/document.xml");
    if (!documentXml) {
        return { content: { paragraphs: [] } };
    }

    // Also try to extract relationships to find images
    const relsXml = await extractZipEntry(bytes, "word/_rels/document.xml.rels");
    // Extract content types to understand image MIME types
    const contentTypesXml = await extractZipEntry(bytes, "[Content_Types].xml");

    const xmlStr = new TextDecoder("utf-8").decode(documentXml);
    const parsedDoc = parseDocxXml(xmlStr);

    // Extract images if present
    let images: DocxImage[] = [];
    if (relsXml && contentTypesXml) {
        const relsStr = new TextDecoder("utf-8").decode(relsXml);
        const ctStr = new TextDecoder("utf-8").decode(contentTypesXml);
        images = await extractImages(bytes, relsStr, ctStr);
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

    // Normalize XML - remove namespace prefixes for easier matching
    const normalized = normalizeDocxXml(xml);

    // Split by paragraph closing tags
    const pParts = normalized.split(/<\/w:p>/);

    for (const part of pParts) {
        const trimmed = part.trim();
        if (!trimmed.includes("<w:p")) continue;

        // --- Paragraph-level properties ---
        // Heading style from pPr/pStyle
        const headingMatch = trimmed.match(/<w:pStyle[^>]*w:val="[^"]*Heading(\d)[^"]*"/i);
        const heading = headingMatch ? parseInt(headingMatch[1] ?? "1") : undefined;

        // Alignment
        const alignMatch = trimmed.match(/<w:jc[^>]*w:val="(center|right|left|start|end|both|distributed)"/i);
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
            if (!rPart.includes("<w:r")) continue;

            const runStyles: Partial<DocxParagraph> = {};

            // Extract run properties block
            const rPrStart = rPart.indexOf("<w:rPr");
            const rPrEnd = rPart.indexOf("</w:rPr>");
            if (rPrStart !== -1 && rPrEnd > rPrStart) {
                const rPr = rPart.substring(rPrStart, rPrEnd + 7); // include </w:rPr>

                // Bold
                if (/<w:b\s|<w:b>|<w:b\/>/i.test(rPr)) runStyles.bold = true;
                else if (/<w:b[^>]*w:val="false"/i.test(rPr)) runStyles.bold = false;
                else if (!/<w:b/i.test(rPr)) { /* not set */ }

                // Italic
                if (/<w:i\s|<w:i>|<w:i\/>/i.test(rPr)) runStyles.italic = true;

                // Underline
                if (/<w:u\s|<w:u>/i.test(rPr)) runStyles.underline = true;

                // Strike / Double-strike-through
                if (/<w:strike\s|<w:strike>|<w:strike\/>/i.test(rPr)) runStyles.strike = true;
                if (/<w:dstrike/i.test(rPr)) runStyles.strike = true;

                // Font size (in half-points)
                const szMatch = rPr.match(/<w:sz[^>]*w:val="(\d+)"/i);
                if (szMatch) runStyles.fontSize = parseInt(szMatch[1]) / 2;

                // Font color
                const colorMatch = rPr.match(/<w:color[^>]*w:val="([0-9A-Fa-f]{6}|auto)"/i);
                if (colorMatch && colorMatch[1] !== "auto") runStyles.fontColor = "#" + colorMatch[1].toLowerCase();

                // Font name
                const fontMatch = rPr.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/i);
                if (fontMatch) runStyles.fontName = fontMatch[1];
            }

            // Extract all <w:t> text nodes in this run (handle multiple per run)
            let runText = "";
            const tRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
            let tMatch: RegExpExecArray | null;
            while ((tMatch = tRegex.exec(rPart)) !== null) {
                if (tMatch[1]) runText += tMatch[1];
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
 */
async function extractZipEntry(data: Uint8Array, targetPath: string): Promise<Uint8Array | null> {
    let offset = 0;
    const view = new DataView(data.buffer);

    while (offset < data.length - 4) {
        // Check for local file header signature (0x04034b50)
        const sig = view.getUint32(offset, true);
        if (sig !== 0x04034b50) break;

        const compressionMethod = view.getUint16(offset + 8, true);
        const compressedSize = view.getUint32(offset + 18, true);
        const uncompressedSize = view.getUint32(offset + 22, true);
        const fileNameLen = view.getUint16(offset + 26, true);
        const extraFieldLen = view.getUint16(offset + 28, true);

        const fileNameBytes = data.slice(offset + 30, offset + 30 + fileNameLen);
        const fileName = new TextDecoder("utf-8").decode(fileNameBytes);

        const dataOffset = offset + 30 + fileNameLen + extraFieldLen;
        const compressedData = data.slice(dataOffset, dataOffset + compressedSize);

        if (fileName === targetPath) {
            if (compressionMethod === 0) {
                // STORED
                return compressedData;
            } else if (compressionMethod === 8) {
                // DEFLATE - use DecompressionStream
                return inflateData(compressedData, uncompressedSize);
            }
            return null;
        }

        offset = dataOffset + compressedSize;
    }

    return null;
}

async function inflateData(data: Uint8Array, _expectedSize: number): Promise<Uint8Array> {
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

    const totalLen = chunks.reduce((sum, c) => sum + c.length, 0);
    const result = new Uint8Array(totalLen);
    let pos = 0;
    for (const chunk of chunks) {
        result.set(chunk, pos);
        pos += chunk.length;
    }

    return result;
}

/**
 * 将 DocxDocument 写入 ArrayBuffer（保留样式）
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

    const sectionChildren = paragraphs.length > 0
        ? paragraphs
        : [new Paragraph({ children: [] })];

    // Add images after paragraphs if present
    if (doc.content.images && doc.content.images.length > 0) {
        // Note: ImageRun in docx library requires buffer input; we'd need to decode base64
        // For now, we include image metadata but skip actual embedding in write path
        // This can be enhanced later with proper image embedding support
    }

    const d = new Document({
        sections: [{ children: sectionChildren }],
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
