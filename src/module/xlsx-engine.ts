// src/module/xlsx-engine.ts - xlsx 读写引擎

import * as XLSX from "xlsx";
import type { CellRule } from "../config/user-settings.ts";

export type SheetData = {
    name: string;
    cells: CellData[][];
    merges: XLSX.Range[];
    colWidths?: number[];
    rowHeights?: number[]; // in points
};

export type CellStyle = {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    fontSize?: number;
    fontName?: string;
    fontColor?: string; // ARGB hex
    fgColor?: string;   // background fill color
    bgColor?: string;   // alias for fgColor
    hAlign?: "left" | "center" | "right";
    vAlign?: "top" | "middle" | "bottom";
    wrapText?: boolean;
};

export type CellData = {
    v: string | number | boolean | null;
    t?: string; // cell type: s/n/b/d/z
    f?: string; // formula
    dropdown?: string[]; // dropdown values from rules
    isMerged?: boolean;
    style?: CellStyle;
};

export type XlsxDocument = {
    sheets: SheetData[];
    activeSheet: number;
    filePath?: string;
};

/**
 * 读取 xlsx 文件，返回结构化数据（含完整样式）
 */
export const readXlsx = (buffer: ArrayBuffer): XlsxDocument => {
    const wb = XLSX.read(buffer, { type: "array", cellStyles: true, cellNF: true });
    const sheets: SheetData[] = wb.SheetNames.map((name) => {
        const ws = wb.Sheets[name]!;
        const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
        const merges = ws["!merges"] || [];
        // Row heights from !rows
        const rowsInfo = ws["!rows"] || [];
        const rowHeights: number[] | undefined = rowsInfo.length > 0
            ? rowsInfo.map((r) => {
                // RowInfo has hpt (points), wpx (pixels), hch (characters), customHeight flag
                if (!r) return 0;
                return r.hpt || 0;
            }).filter((v): v is number => v > 0)
            : undefined;
        const cells: CellData[][] = [];

        for (let r = range.s.r; r <= range.e.r; r++) {
            const row: CellData[] = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                const addr = XLSX.utils.encode_cell({ r, c });
                const cell = ws[addr];
                const isMerged = merges.some(
                    (m) => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c
                );

                let style: CellStyle | undefined;
                if (cell && cell.s) {
                    style = extractCellStyle(cell.s);
                }

                row.push({
                    v: cell ? cell.v : null,
                    t: cell?.t,
                    f: cell?.f || undefined,
                    isMerged,
                    style,
                });
            }
            cells.push(row);
        }

        return {
            name,
            cells,
            merges,
            colWidths: ws["!cols"]?.map((col) => col?.wch || 10),
            rowHeights,
        };
    });

    return { sheets, activeSheet: 0 };
};

/**
 * 从 SheetJS 的 cell style 对象中提取我们需要的样式属性
 * SheetJS 的 s 属性格式: { font: { name, sz, color:{ rgb }, bold, italic, underline, ... },
 *                            fill: { fgColor:{ rgb }, patternType }, alignment: { horizontal, vertical, wrapText } }
 */
function extractCellStyle(s: XLSX.CellObject["s"] & Record<string, unknown>): CellStyle | undefined {
    if (!s) return undefined;
    const cs: CellStyle = {};
    const font = s.font as { bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; name?: string; sz?: number; color?: { rgb?: string; theme?: number } } | undefined;
    const fill = s.fill as { fgColor?: { rgb?: string; argb?: string; theme?: number }; patternType?: string } | undefined;
    const align = s.alignment as { horizontal?: string; vertical?: string; wrapText?: boolean } | undefined;

    if (font) {
        if (font.bold) cs.bold = true;
        if (font.italic) cs.italic = true;
        if (font.underline) cs.underline = true;
        if (font.strike) cs.strike = true;
        if (font.sz) cs.fontSize = font.sz;
        if (font.name) cs.fontName = font.name;
        if (font.color?.rgb) {
            // Handle ARGB format (e.g., "FFFF0000" for red)
            let color = font.color.rgb;
            if (color.length === 8 && color.startsWith("FF")) {
                color = color.substring(2); // strip FF alpha
            } else if (color.length === 8) {
                color = "#" + color;
            }
            cs.fontColor = color.startsWith("#") ? color : "#" + color;
        }
    }

    if (fill?.fgColor) {
        let bg = fill.fgColor.rgb || fill.fgColor.argb || "";
        if (bg.length === 8 && bg.startsWith("FF")) {
            bg = bg.substring(2);
        } else if (bg.length === 8) {
            bg = "#" + bg;
        }
        if (bg) cs.fgColor = bg.startsWith("#") ? bg : "#" + bg;
    }

    if (align) {
        if (align.horizontal === "center" || align.horizontal === "centerContinuous") cs.hAlign = "center";
        else if (align.horizontal === "right" || align.horizontal === "end" ||
                 align.horizontal === "distributed") cs.hAlign = "right";
        else if (align.horizontal === "left" || align.horizontal === "start" ||
                 align.horizontal === "justify") cs.hAlign = "left";

        if (align.vertical === "top") cs.vAlign = "top";
        else if (align.vertical === "center") cs.vAlign = "middle";
        else if (align.vertical === "bottom") cs.vAlign = "bottom";

        if (align.wrapText) cs.wrapText = true;
    }

    return Object.keys(cs).length > 0 ? cs : undefined;
}

/**
 * 将 XlsxDocument 写入 ArrayBuffer（保留样式）
 */
export const writeXlsx = (doc: XlsxDocument): ArrayBuffer => {
    const wb = XLSX.utils.book_new();

    for (const sheet of doc.sheets) {
        const ws: XLSX.WorkSheet = {};

        if (sheet.cells.length > 0) {
            const endR = sheet.cells.length - 1;
            const endC = Math.max(...sheet.cells.map((r) => r.length)) - 1;

            for (let r = 0; r < sheet.cells.length; r++) {
                for (let c = 0; c < sheet.cells[r].length; c++) {
                    const addr = XLSX.utils.encode_cell({ r, c });
                    const cellData = sheet.cells[r][c];
                    if (cellData.v !== null && cellData.v !== undefined && cellData.v !== "") {
                        const newCell: XLSX.CellObject = {
                            v: cellData.v,
                            t: (cellData.t as "s" | "n" | "b") || inferType(cellData.v),
                        };
                        if (cellData.f) newCell.f = cellData.f;
                        if (cellData.style) {
                            newCell.s = buildCellStyle(cellData.style);
                        }
                        ws[addr] = newCell;
                    }
                }
            }

            ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: endR, c: Math.max(endC, 0) } });
        }

        if (sheet.merges) {
            ws["!merges"] = sheet.merges;
        }
        if (sheet.colWidths) {
            ws["!cols"] = sheet.colWidths.map((w) => ({ wch: w }));
        }
        if (sheet.rowHeights && sheet.rowHeights.length > 0) {
            ws["!rows"] = sheet.rowHeights.map((h) => ({ hpt: h }));
        }

        XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    }

    const out = XLSX.write(wb, { type: "array", bookType: "xlsx", cellStyles: true });
    return out as ArrayBuffer;
};

/**
 * 将 CellStyle 转换回 SheetJS 的 cell style 对象格式
 */
function buildCellStyle(cs: CellStyle): Record<string, unknown> {
    const font: Record<string, unknown> = {};
    if (cs.bold) font.bold = true;
    if (cs.italic) font.italic = true;
    if (cs.underline) font.underline = true;
    if (cs.strike) font.strike = true;
    if (cs.fontSize) font.sz = cs.fontSize;
    if (cs.fontName) font.name = cs.fontName;
    if (cs.fontColor) {
        // Convert hex back to ARGB
        let rgb = cs.fontColor.startsWith("#") ? cs.fontColor.substring(1) : cs.fontColor;
        if (rgb.length === 6) rgb = "FF" + rgb;
        font.color = { rgb };
    }

    const fill: Record<string, unknown> = {};
    if (cs.fgColor || cs.bgColor) {
        let bg = (cs.fgColor || cs.bgColor || "").startsWith("#")
            ? (cs.fgColor || cs.bgColor || "").substring(1)
            : (cs.fgColor || cs.bgColor || "");
        if (bg.length === 6) bg = "FF" + bg;
        fill.fgColor = { rgb: bg };
        fill.patternType = "solid";
    }

    const alignment: Record<string, unknown> = {};
    if (cs.hAlign) alignment.horizontal = cs.hAlign;
    if (cs.vAlign) alignment.vertical = cs.vAlign === "middle" ? "center" : cs.vAlign;
    if (cs.wrapText) alignment.wrapText = true;

    const s: Record<string, unknown> = {};
    if (Object.keys(font).length > 0) s.font = font;
    if (Object.keys(fill).length > 0) s.fill = fill;
    if (Object.keys(alignment).length > 0) s.alignment = alignment;

    return s;
}

/**
 * 根据用户规则为单元格生成下拉列表
 */
export const applyRulesToSheet = (sheet: SheetData, rules: CellRule[]): SheetData => {
    const newCells = sheet.cells.map((row) => row.map((cell) => ({ ...cell })));

    for (const rule of rules) {
        switch (rule.type) {
            case "cell": {
                if (!rule.applyTo) continue;
                const range = parseCellRef(rule.applyTo);
                if (range && newCells[range.r]?.[range.c]) {
                    newCells[range.r][range.c].dropdown = rule.values;
                }
                break;
            }
            case "mergeCell": {
                if (!rule.applyTo) continue;
                const mRange = parseRangeRef(rule.applyTo);
                if (mRange) {
                    for (let r = mRange.s.r; r <= mRange.e.r; r++) {
                        for (let c = mRange.s.c; c <= mRange.e.c; c++) {
                            if (newCells[r]?.[c]) {
                                newCells[r][c].dropdown = rule.values;
                            }
                        }
                    }
                }
                break;
            }
            case "rowCell": {
                if (!rule.range) continue;
                const rowRange = parseRowRange(rule.range);
                if (rowRange) {
                    for (let r = rowRange.start; r <= rowRange.end; r++) {
                        if (newCells[r]) {
                            for (let c = 0; c < newCells[r].length; c++) {
                                newCells[r][c].dropdown = rule.values;
                            }
                        }
                    }
                }
                break;
            }
            case "alias": {
                if (!rule.applyTo) continue;
                const ref = parseCellRef(rule.applyTo);
                if (ref && newCells[ref.r]?.[ref.c]) {
                    newCells[ref.r][ref.c].dropdown = rule.values;
                }
                break;
            }
        }
    }

    return { ...sheet, cells: newCells };
};

/**
 * 合并单元格
 */
export const mergeCells = (sheet: SheetData, startR: number, startC: number, endR: number, endC: number): SheetData => {
    const newMerge: XLSX.Range = { s: { r: startR, c: startC }, e: { r: endR, c: endC } };
    const merges = [...(sheet.merges || []), newMerge];

    const cells = sheet.cells.map((row) => row.map((cell) => ({ ...cell, isMerged: false })));
    for (let r = startR; r <= endR; r++) {
        for (let c = startC; c <= endC; c++) {
            if (cells[r]?.[c]) {
                cells[r][c].isMerged = true;
                if (r !== startR || c !== startC) {
                    cells[r][c].v = null;
                }
            }
        }
    }

    return { ...sheet, cells, merges };
};

/**
 * 取消合并单元格
 */
export const unmergeCells = (sheet: SheetData, startR: number, startC: number): SheetData => {
    const merges = (sheet.merges || []).filter(
        (m) => !(m.s.r === startR && m.s.c === startC)
    );
    const cells = sheet.cells.map((row) => row.map((cell) => ({ ...cell, isMerged: false })));
    return { ...sheet, cells, merges };
};

/**
 * 添加行
 */
export const addRow = (sheet: SheetData, index?: number): SheetData => {
    const newRow: CellData[] = sheet.cells[0]?.map(() => ({ v: null })) || [{ v: null }];
    const cells = [...sheet.cells];
    if (index !== undefined && index >= 0 && index < cells.length) {
        cells.splice(index, 0, newRow);
    } else {
        cells.push(newRow);
    }
    return { ...sheet, cells };
};

/**
 * 添加列
 */
export const addCol = (sheet: SheetData, index?: number): SheetData => {
    const cells = sheet.cells.map((row) => {
        const newRow = [...row];
        if (index !== undefined && index >= 0 && index < newRow.length) {
            newRow.splice(index, 0, { v: null });
        } else {
            newRow.push({ v: null });
        }
        return newRow;
    });
    return { ...sheet, cells };
};

/**
 * 删除行
 */
export const deleteRow = (sheet: SheetData, index: number): SheetData => {
    if (index < 0 || index >= sheet.cells.length) return sheet;
    const cells = [...sheet.cells];
    cells.splice(index, 1);
    return { ...sheet, cells };
};

/**
 * 删除列
 */
export const deleteCol = (sheet: SheetData, index: number): SheetData => {
    const cells = sheet.cells.map((row) => {
        const newRow = [...row];
        if (index >= 0 && index < newRow.length) {
            newRow.splice(index, 1);
        }
        return newRow;
    });
    return { ...sheet, cells };
};

/**
 * 更新单元格值
 */
export const updateCell = (sheet: SheetData, r: number, c: number, value: string | number | boolean | null): SheetData => {
    const cells = sheet.cells.map((row) => row.map((cell) => ({ ...cell })));
    // expand if needed
    while (cells.length <= r) {
        cells.push(cells[0]?.map(() => ({ v: null })) || [{ v: null }]);
    }
    while (cells[r].length <= c) {
        cells[r].push({ v: null });
    }
    cells[r][c] = { ...cells[r][c], v: value };
    return { ...sheet, cells };
};

// --- helpers ---

function inferType(v: string | number | boolean): "s" | "n" | "b" {
    if (typeof v === "number") return "n";
    if (typeof v === "boolean") return "b";
    return "s";
}

function parseCellRef(ref: string): { r: number; c: number } | null {
    const match = ref.match(/^([A-Z]+)(\d+)$/i);
    if (!match || !match[1] || !match[2]) return null;
    const col = colToNum(match[1].toUpperCase());
    const row = parseInt(match[2]) - 1;
    return { r: row, c: col };
}

function parseRangeRef(ref: string): { s: { r: number; c: number }; e: { r: number; c: number } } | null {
    const parts = ref.split(":");
    if (parts.length !== 2) return null;
    const s = parseCellRef(parts[0]!);
    const e = parseCellRef(parts[1]!);
    if (!s || !e) return null;
    return { s, e };
}

function parseRowRange(range: string): { start: number; end: number } | null {
    const match = range.match(/^(\d+)\s*-\s*(\d+)$/);
    if (!match || !match[1] || !match[2]) return null;
    return { start: parseInt(match[1]) - 1, end: parseInt(match[2]) - 1 };
}

function colToNum(col: string): number {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
        num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num - 1;
}
