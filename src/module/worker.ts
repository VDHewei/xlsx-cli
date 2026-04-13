import { assetsMap } from "./generated-assets.ts";
import { readXlsx, writeXlsx, applyRulesToSheet, mergeCells, unmergeCells, addRow, addCol, deleteRow, deleteCol, updateCell, type XlsxDocument, type SheetData } from "./xlsx-engine.ts";
import { readDocx, writeDocx, addParagraph, updateParagraph, deleteParagraph, type DocxDocument } from "./docx-engine.ts";
import { loadSettings, saveSettings, type UserSettings } from "@config/user-settings.ts";
import { messages, type Lang } from "@i18n/index.ts";
import { App } from "@config/app.ts";

const supportedLangs: Lang[] = ["zh-CN", "zh-TW", "en", "ja"];

/**
 * Detect system language and map to supported lang, fallback to "en"
 */
function detectSystemLang(savedLang?: Lang): Lang {
    if (savedLang && supportedLangs.includes(savedLang)) return savedLang;
    try {
        const navLang = typeof navigator !== "undefined" ? navigator.language : undefined;
        const envLang = process.env.LANG || process.env.LC_ALL || process.env.LC_MESSAGES || navLang || "";
        const normalized = envLang.replace(/_/g, "-").split(".")[0] ?? "en";
        const langPrefix = normalized.split("-")[0] ?? "en";
        const map: Record<string, Lang> = { zh: "zh-CN", en: "en", ja: "ja" };
        const mapped = map[langPrefix] ?? (supportedLangs.includes(langPrefix as Lang) ? langPrefix as Lang : undefined);
        return mapped ?? "en";
    } catch {
        return "en";
    }
}

const baseHTML = getFrontendHTML();

type Asset = { mime: string; data: string };

type AssetResult = {
    buffer: Buffer;
    asset: Asset;
};

const getAssetsData = (name: string): AssetResult | undefined => {
    const assets = assetsMap[name] as Asset;
    if (assets === undefined || assets.data === undefined || assets.data === null) {
        return undefined;
    }
    let base64Data: string;
    const values = assets.data.split(",");
    if (values.length >= 2 && values[1] !== undefined) {
        base64Data = values[1];
        return {
            asset: assets,
            buffer: Buffer.from(base64Data, "base64"),
        };
    }
    return undefined;
};

// In-memory state for the current session
let currentXlsx: XlsxDocument | null = null;
let currentDocx: DocxDocument | null = null;

const startServer = () => {
    const server = Bun.serve({
        port: 0,
        fetch(req: Request) {
            const url = new URL(req.url);
            const { pathname } = url;

            // Static assets (favicon, icons etc)
            if (pathname.endsWith("favicon.ico")) {
                const fav = getAssetsData("favicon.ico");
                if (fav !== undefined) {
                    return new Response(fav.buffer as unknown as BodyInit, { headers: { "Content-Type": "image/x-icon" } });
                }
            }
            if (assetsMap[pathname]) {
                const assets = getAssetsData(pathname);
                if (assets !== undefined) {
                    return new Response(assets.buffer as unknown as BodyInit, { headers: { "Content-Type": assets.asset.mime } });
                }
            }

            // API routes
            if (pathname.startsWith("/api/")) {
                return handleApi(req, pathname, url);
            }

            // SPA fallback
            return new Response(baseHTML, { headers: { "Content-Type": "text/html; charset=utf-8" } });
        },
    });
    console.log(`URL: ${server.url}`);
    self.postMessage({ type: "PORT", port: server.port });
    return server;
};

async function handleApi(req: Request, pathname: string, url: URL): Promise<Response> {
    const method = req.method;
    const json = () => req.json();

    try {
        // --- Settings API ---
        if (pathname === "/api/settings" && method === "GET") {
            return jsonOk(loadSettings());
        }
        if (pathname === "/api/settings" && method === "POST") {
            const body = (await json()) as UserSettings;
            saveSettings(body);
            return jsonOk({ ok: true });
        }

        // --- i18n API ---
        if (pathname === "/api/i18n" && method === "GET") {
            const reqLang = (url.searchParams.get("lang") || "") as Lang;
            const settings = loadSettings();
            const lang = supportedLangs.includes(reqLang) ? reqLang : detectSystemLang(settings.lang);
            const msgs = messages[lang] || messages["en"]!;
            return jsonOk({ lang, messages: msgs });
        }

        // --- XLSX API ---
        if (pathname === "/api/xlsx/open" && method === "POST") {
            const body = (await json()) as { data: string; fileName?: string };
            const buffer = Buffer.from(body.data, "base64").buffer as ArrayBuffer;
            currentXlsx = readXlsx(buffer);
            currentXlsx.filePath = body.fileName;
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/save" && method === "POST") {
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const buf = writeXlsx(currentXlsx);
            const b64 = Buffer.from(buf).toString("base64");
            return jsonOk({ data: b64, fileName: currentXlsx.filePath });
        }
        if (pathname === "/api/xlsx/cell" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; r: number; c: number; value: string | number | boolean | null };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const si = body.sheetIndex;
            const sheet = currentXlsx.sheets[si];
            if (sheet) {
                currentXlsx.sheets[si] = updateCell(sheet, body.r, body.c, body.value);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/apply-rules" && method === "POST") {
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const settings = loadSettings();
            currentXlsx.sheets = currentXlsx.sheets.map((s: SheetData) => applyRulesToSheet(s, settings.rules));
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/merge" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; startR: number; startC: number; endR: number; endC: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = mergeCells(s, body.startR, body.startC, body.endR, body.endC);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/unmerge" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; startR: number; startC: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = unmergeCells(s, body.startR, body.startC);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/add-row" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; index?: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = addRow(s, body.index);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/add-col" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; index?: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = addCol(s, body.index);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/delete-row" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; index: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = deleteRow(s, body.index);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/delete-col" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; index: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = deleteCol(s, body.index);
            }
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/add-sheet" && method === "POST") {
            const body = (await json()) as { name: string };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const newSheet: SheetData = { name: body.name || "Sheet", cells: [[]], merges: [] };
            currentXlsx.sheets.push(newSheet);
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/delete-sheet" && method === "POST") {
            const body = (await json()) as { sheetIndex: number };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            currentXlsx.sheets.splice(body.sheetIndex, 1);
            currentXlsx.activeSheet = Math.min(currentXlsx.activeSheet, currentXlsx.sheets.length - 1);
            return jsonOk(currentXlsx);
        }
        if (pathname === "/api/xlsx/rename-sheet" && method === "POST") {
            const body = (await json()) as { sheetIndex: number; name: string };
            if (!currentXlsx) return jsonErr("No xlsx document open");
            const s = currentXlsx.sheets[body.sheetIndex];
            if (s) {
                currentXlsx.sheets[body.sheetIndex] = { ...s, name: body.name };
            }
            return jsonOk(currentXlsx);
        }

        // --- DOCX API ---
        if (pathname === "/api/docx/open" && method === "POST") {
            const body = (await json()) as { data: string; fileName?: string };
            const buffer = Buffer.from(body.data, "base64").buffer as ArrayBuffer;
            currentDocx = await readDocx(buffer);
            currentDocx.filePath = body.fileName;
            return jsonOk(currentDocx);
        }
        if (pathname === "/api/docx/save" && method === "POST") {
            if (!currentDocx) return jsonErr("No docx document open");
            const buf = await writeDocx(currentDocx);
            const b64 = Buffer.from(buf).toString("base64");
            return jsonOk({ data: b64, fileName: currentDocx.filePath });
        }
        if (pathname === "/api/docx/add-paragraph" && method === "POST") {
            const body = (await json()) as { paragraph: { text: string; bold?: boolean; italic?: boolean; underline?: boolean; heading?: number } };
            if (!currentDocx) return jsonErr("No docx document open");
            currentDocx = addParagraph(currentDocx, body.paragraph);
            return jsonOk(currentDocx);
        }
        if (pathname === "/api/docx/update-paragraph" && method === "POST") {
            const body = (await json()) as { index: number; paragraph: { text: string; bold?: boolean; italic?: boolean; underline?: boolean; heading?: number } };
            if (!currentDocx) return jsonErr("No docx document open");
            currentDocx = updateParagraph(currentDocx, body.index, body.paragraph);
            return jsonOk(currentDocx);
        }
        if (pathname === "/api/docx/delete-paragraph" && method === "POST") {
            const body = (await json()) as { index: number };
            if (!currentDocx) return jsonErr("No docx document open");
            currentDocx = deleteParagraph(currentDocx, body.index);
            return jsonOk(currentDocx);
        }

        return new Response("Not Found", { status: 404 });
    } catch (e) {
        console.error("API error:", e);
        return jsonErr((e as Error).message || "Internal server error");
    }
}

function jsonOk(data: unknown) {
    return new Response(JSON.stringify({ ok: true, data }), {
        headers: { "Content-Type": "application/json" },
    });
}

function jsonErr(msg: string) {
    return new Response(JSON.stringify({ ok: false, error: msg }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
    });
}

function getFrontendHTML(): string {
    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>xlsx-cli</title>
<link rel="icon" href="favicon.ico" type="image/x-icon">
<style>
:root {
  --bg: #f8f9fa;
  --surface: #ffffff;
  --border: #e0e0e0;
  --text: #1a1a1a;
  --text-secondary: #666;
  --primary: #4a6cf7;
  --primary-hover: #3b5de7;
  --danger: #e74c3c;
  --success: #27ae60;
  --radius: 8px;
  --shadow: 0 1px 3px rgba(0,0,0,0.08);
  --font: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  --nav-height: 48px;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: var(--font); background: var(--bg); color: var(--text); height: 100vh; overflow: hidden; }

/* Navigation */
.nav { height: var(--nav-height); background: var(--surface); border-bottom: 1px solid var(--border); display: flex; align-items: center; padding: 0 16px; gap: 8px; }
.nav-title { font-weight: 600; font-size: 15px; margin-right: auto; }
.nav-btn { background: none; border: 1px solid transparent; cursor: pointer; padding: 6px 12px; border-radius: var(--radius); font-size: 13px; color: var(--text); display: flex; align-items: center; gap: 4px; }
.nav-btn:hover { background: var(--bg); }
.nav-btn.active { background: var(--primary); color: white; border-color: var(--primary); }

/* Icon helpers */
.icon { display: inline-block; width: 16px; height: 16px; vertical-align: middle; }
.icon svg { width: 16px; height: 16px; fill: currentColor; }

/* Pages */
.page { display: none; height: calc(100vh - var(--nav-height)); overflow-y: auto; padding: 24px; }
.page.active { display: block; }

/* Home page */
.home { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 24px; }
.home-title { font-size: 28px; font-weight: 700; }
.home-subtitle { color: var(--text-secondary); font-size: 14px; }
.home-actions { display: flex; gap: 16px; }
.btn { padding: 10px 24px; border: 1px solid var(--border); background: var(--surface); border-radius: var(--radius); cursor: pointer; font-size: 14px; display: flex; align-items: center; gap: 8px; transition: all 0.15s; color: var(--text); }
.btn:hover { border-color: var(--primary); color: var(--primary); }
.btn-primary { background: var(--primary); color: white; border-color: var(--primary); }
.btn-primary:hover { background: var(--primary-hover); }
.btn-sm { padding: 4px 10px; font-size: 12px; }
.btn-danger { color: var(--danger); }
.btn-danger:hover { background: #fef2f2; border-color: var(--danger); }
.btn-success { color: var(--success); }
.btn-success:hover { background: #f0fdf4; border-color: var(--success); }

/* Drop zone */
.drop-zone { border: 2px dashed var(--border); border-radius: var(--radius); padding: 40px 60px; text-align: center; color: var(--text-secondary); font-size: 13px; cursor: pointer; transition: all 0.2s; }
.drop-zone:hover, .drop-zone.dragover { border-color: var(--primary); color: var(--primary); background: rgba(74,108,247,0.03); }
.drop-zone.hidden { display: none; }

/* XLSX Editor */
.xlsx-editor { display: flex; flex-direction: column; height: 100%; }
.xlsx-toolbar { display: flex; align-items: center; gap: 6px; padding: 8px 0; flex-wrap: wrap; }
.sheet-tabs { display: flex; gap: 2px; padding: 8px 0; border-bottom: 1px solid var(--border); }
.sheet-tab { padding: 6px 14px; font-size: 12px; cursor: pointer; border: 1px solid transparent; border-bottom: none; border-radius: var(--radius) var(--radius) 0 0; background: transparent; }
.sheet-tab:hover { background: var(--bg); }
.sheet-tab.active { background: var(--surface); border-color: var(--border); font-weight: 600; }
.sheet-tab-add { background: none; border: 1px dashed var(--border); }
.sheet-tab-add:hover { border-color: var(--primary); color: var(--primary); }
.spreadsheet-container { flex: 1; overflow: auto; border: 1px solid var(--border); border-radius: var(--radius); }
.spreadsheet { border-collapse: collapse; font-size: 12px; width: max-content; min-width: 100%; table-layout: fixed; }
.spreadsheet th, .spreadsheet td { border: 1px solid var(--border); padding: 0; min-width: 80px; height: 28px; position: relative; vertical-align: top; }
.spreadsheet th { background: var(--bg); font-weight: 500; position: sticky; top: 0; z-index: 1; min-width: 40px; width: 40px; color: var(--text-secondary); font-size: 11px; text-align: center; }
.spreadsheet th.corner { position: sticky; left: 0; z-index: 2; }
.spreadsheet th.row-header { position: sticky; left: 0; z-index: 1; }
.spreadsheet td { cursor: cell; }
.spreadsheet td:focus { outline: 2px solid var(--primary); outline-offset: -1px; z-index: 1; position: relative; }
/* Cell value - supports rich styles */
.spreadsheet td .cell-value { padding: 2px 6px; min-height: 24px; outline: none; white-space: pre-wrap; word-break: break-word; overflow-wrap: break-word; line-height: inherit; }
.spreadsheet td .cell-value[contenteditable=true] { background: #fff; cursor: text; }
.spreadsheet td.merged { background: #f0f4ff; }
.cell-select { width: 100%; height: 100%; border: none; background: transparent; font-size: 12px; padding: 2px 4px; cursor: pointer; }
.cell-select:focus { outline: none; box-shadow: inset 0 0 0 1px var(--primary); }

/* Cell editor popup */
.cell-editor-popup { position: fixed; z-index: 100; background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: 0 4px 12px rgba(0,0,0,0.12); padding: 12px; min-width: 240px; }
.cell-editor-popup .popup-title { font-size: 12px; font-weight: 600; margin-bottom: 8px; color: var(--text-secondary); }
.cell-editor-popup input { width: 100%; border: 1px solid var(--border); border-radius: 4px; padding: 6px 8px; font-size: 13px; margin-bottom: 8px; }
.cell-editor-popup .popup-actions { display: flex; gap: 6px; justify-content: flex-end; }

/* DOCX Editor */
.docx-editor { max-width: 800px; margin: 0 auto; }
.docx-toolbar { display: flex; gap: 6px; padding: 8px 0; border-bottom: 1px solid var(--border); margin-bottom: 16px; }
.docx-content { border: 1px solid var(--border); border-radius: var(--radius); min-height: 400px; padding: 20px; background: var(--surface); }
.docx-paragraph { padding: 4px 0; min-height: 24px; border-bottom: 1px solid transparent; line-height: 1.6; }
.docx-paragraph:hover { background: rgba(74,108,247,0.03); }
.docx-paragraph:focus { outline: 1px solid var(--primary); border-radius: 2px; }
.docx-paragraph h1 { font-size: 24px; font-weight: bold; margin: 12px 0 4px; }
.docx-paragraph h2 { font-size: 20px; font-weight: bold; margin: 10px 0 4px; }
.docx-paragraph h3 { font-size: 16px; font-weight: bold; margin: 8px 0 4px; }
.docx-paragraph[contenteditable=true] { cursor: text; }
/* Inline style support for docx paragraphs */
.docx-bold { font-weight: bold; }
.docx-italic { font-style: italic; }
.docx-underline { text-decoration: underline; }
.docx-strike { text-decoration: line-through; }
.docx-image { max-width: 100%; height: auto; display: block; margin: 8px 0; border: 1px solid #ddd; border-radius: 4px; }

/* Settings */
.settings { max-width: 600px; margin: 0 auto; }
.settings-section { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 16px; margin-bottom: 16px; }
.settings-section h3 { font-size: 14px; font-weight: 600; margin-bottom: 12px; }
.settings-row { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; }
.settings-row label { font-size: 13px; min-width: 100px; color: var(--text-secondary); }
.settings-row select, .settings-row input { flex: 1; border: 1px solid var(--border); border-radius: 4px; padding: 6px 8px; font-size: 13px; background: var(--surface); }
.rule-item { display: flex; gap: 8px; align-items: center; padding: 6px 0; border-bottom: 1px solid var(--border); font-size: 12px; flex-wrap: wrap; }
.rule-item:last-child { border-bottom: none; }
.rule-item select { width: 100px; flex-shrink: 0; border: 1px solid var(--border); border-radius: 4px; padding: 4px 6px; font-size: 12px; }
.rule-item input { flex: 1; min-width: 60px; border: 1px solid var(--border); border-radius: 4px; padding: 4px 6px; font-size: 12px; }
.fn-item { display: flex; flex-direction: column; gap: 4px; padding: 8px; border: 1px solid var(--border); border-radius: var(--radius); margin-bottom: 8px; }
.fn-item input { border: 1px solid var(--border); border-radius: 4px; padding: 4px 6px; font-size: 12px; }
.fn-item textarea { border: 1px solid var(--border); border-radius: 4px; padding: 6px; font-size: 11px; font-family: monospace; min-height: 60px; resize: vertical; }
.fn-item-header { display: flex; justify-content: space-between; align-items: center; }

/* Toast */
.toast { position: fixed; bottom: 20px; right: 20px; background: #333; color: white; padding: 10px 18px; border-radius: var(--radius); font-size: 13px; z-index: 1000; animation: slideIn 0.2s; }
.toast.success { background: var(--success); }
.toast.error { background: var(--danger); }
@keyframes slideIn { from { transform: translateY(10px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #ccc; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #999; }
</style>
</head>
<body>

<div class="nav">
  <span class="nav-title" data-i18n="app.title">xlsx-cli</span>
  <button class="nav-btn active" data-page="home" data-i18n="nav.home">首页</button>
  <button class="nav-btn" data-page="settings" data-i18n="nav.settings">设置</button>
</div>

<!-- Home Page -->
<div class="page active" id="page-home">
  <div class="home">
    <div class="drop-zone" id="dropZone">
      <div class="home-title" data-i18n="home.title">xlsx-cli</div>
      <div class="home-subtitle" data-i18n="home.subtitle">轻量级 Office 文件编辑器</div>
      <div style="margin-top: 16px;" data-i18n="home.dragHint">或拖拽文件到此处</div>
    </div>
    <div class="home-actions">
      <button class="btn" id="btnOpenXlsx" onclick="fileInput.click()">
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
        <span data-i18n="home.openXlsx">打开 XLSX</span>
      </button>
      <button class="btn" id="btnOpenDocx" onclick="fileInputDocx.click()">
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="10" y1="9" x2="8" y2="9"/></svg>
        <span data-i18n="home.openDocx">打开 DOCX</span>
      </button>
    </div>
  </div>
</div>

<!-- XLSX Editor Page -->
<div class="page" id="page-xlsx">
  <div class="xlsx-editor">
    <div class="xlsx-toolbar">
      <button class="btn btn-sm" onclick="saveXlsx()"><span data-i18n="xlsx.save">保存 XLSX</span></button>
      <button class="btn btn-sm" onclick="downloadXlsx()"><span data-i18n="xlsx.export">导出</span></button>
      <span style="flex:1"></span>
      <button class="btn btn-sm" onclick="applyRulesToXlsx()"><span data-i18n="xlsx.applyRule">应用规则</span></button>
      <button class="btn btn-sm" onclick="addXlsxRow()"><span data-i18n="xlsx.addRow">添加行</span></button>
      <button class="btn btn-sm" onclick="addXlsxCol()"><span data-i18n="xlsx.addCol">添加列</span></button>
      <button class="btn btn-sm" onclick="mergeSelectedCells()"><span data-i18n="xlsx.mergeCell">合并单元格</span></button>
      <button class="btn btn-sm btn-danger" onclick="goHome()"><span data-i18n="common.back">返回</span></button>
    </div>
    <div class="sheet-tabs" id="sheetTabs"></div>
    <div class="spreadsheet-container" id="spreadsheetContainer"></div>
  </div>
</div>

<!-- DOCX Editor Page -->
<div class="page" id="page-docx">
  <div class="docx-editor">
    <div class="docx-toolbar">
      <button class="btn btn-sm btn-primary" onclick="saveDocx()"><span data-i18n="docx.save">保存 DOCX</span></button>
      <button class="btn btn-sm" onclick="downloadDocx()"><span data-i18n="xlsx.export">导出</span></button>
      <span style="flex:1"></span>
      <button class="btn btn-sm" onclick="addDocxParagraph()"><span data-i18n="common.add">添加</span></button>
      <button class="btn btn-sm btn-danger" onclick="goHome()"><span data-i18n="common.back">返回</span></button>
    </div>
    <div class="docx-content" id="docxContent"></div>
  </div>
</div>

<!-- Settings Page -->
<div class="page" id="page-settings">
  <div class="settings">
    <h2 style="margin-bottom:16px;font-size:18px;" data-i18n="settings.title">设置</h2>

    <div class="settings-section">
      <h3 data-i18n="settings.lang">语言</h3>
      <div class="settings-row">
        <label data-i18n="settings.lang">语言</label>
        <select id="settingLang" onchange="changeLang(this.value)">
          <option value="zh-CN" data-i18n="settings.lang.zh-CN">简体中文</option>
          <option value="zh-TW" data-i18n="settings.lang.zh-TW">繁體中文</option>
          <option value="en" data-i18n="settings.lang.en">English</option>
          <option value="ja" data-i18n="settings.lang.ja">日本語</option>
        </select>
      </div>
    </div>

    <div class="settings-section">
      <h3 data-i18n="settings.api">API 设置</h3>
      <div class="settings-row">
        <label data-i18n="settings.api.host">Host 地址</label>
        <input type="text" id="settingApiHost" placeholder="http://localhost:3000">
      </div>
    </div>

    <div class="settings-section">
      <h3 data-i18n="settings.rules">常用规则</h3>
      <div id="rulesContainer"></div>
      <div style="margin-top:8px;display:flex;gap:8px;">
        <select id="newRuleType">
          <option value="cell">cell</option>
          <option value="mergeCell">mergeCell</option>
          <option value="rowCell">rowCell</option>
          <option value="alias">alias</option>
        </select>
        <button class="btn btn-sm" onclick="addRule()"><span data-i18n="common.add">添加</span></button>
      </div>
    </div>

    <div class="settings-section">
      <h3 data-i18n="settings.customFunctions">自定义函数</h3>
      <div id="fnContainer"></div>
      <div style="margin-top:8px;">
        <button class="btn btn-sm" onclick="addCustomFunction()"><span data-i18n="settings.customFunctions.add">添加函数</span></button>
      </div>
    </div>

    <div style="margin-top:16px;display:flex;gap:8px;">
      <button class="btn btn-primary" onclick="saveSettings()"><span data-i18n="common.save">保存</span></button>
      <button class="btn" onclick="loadAndRenderSettings()"><span data-i18n="common.cancel">取消</span></button>
    </div>
  </div>
</div>

<input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" onchange="handleFileSelect(event,'xlsx')">
<input type="file" id="fileInputDocx" accept=".docx" style="display:none" onchange="handleFileSelect(event,'docx')">

<script>
// ===== State =====
let currentPage = "home";
let xlsxData = null;
let xlsxActiveSheet = 0;
let docxData = null;
let i18nMessages = {};
let currentLang = "zh-CN";
let selectedCells = []; // for merge: [{r,c}]

// ===== Navigation =====
document.querySelectorAll("[data-page]").forEach(btn => {
  btn.addEventListener("click", () => navigateTo(btn.dataset.page));
});

function navigateTo(page) {
  currentPage = page;
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  document.getElementById("page-" + page)?.classList.add("active");
  document.querySelectorAll("[data-page]").forEach(b => {
    b.classList.toggle("active", b.dataset.page === page);
  });
  if (page === "settings") loadAndRenderSettings();
}

function goHome() {
  navigateTo("home");
}

// ===== i18n =====
async function loadI18n(lang) {
  try {
    const res = await fetch("/api/i18n?lang=" + lang);
    const json = await res.json();
    if (json.ok) {
      i18nMessages = json.data.messages;
      currentLang = lang;
      applyI18n();
    }
  } catch(e) { console.error("i18n error", e); }
}

function t(key) {
  return i18nMessages[key] || key;
}

function applyI18n() {
  document.querySelectorAll("[data-i18n]").forEach(el => {
    const key = el.dataset.i18n;
    if (i18nMessages[key]) {
      if (el.tagName === "INPUT" || el.tagName === "SELECT") {
        // skip, handled separately
      } else {
        el.textContent = i18nMessages[key];
      }
    }
  });
}

function changeLang(lang) {
  loadI18n(lang);
  const settings = getSettingsFromForm();
  settings.lang = lang;
  fetch("/api/settings", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify(settings) });
}

// ===== File handling =====
function handleFileSelect(event, type) {
  const file = event.target.files[0];
  if (!file) return;
  openFile(file, type);
  event.target.value = "";
}

function openFile(file, type) {
  const reader = new FileReader();
  reader.onload = function(e) {
    // Use readAsDataURL result or fallback to chunked btoa
    let base64;
    if (e.target.result && typeof e.target.result === "string" && e.target.result.startsWith("data:")) {
      base64 = e.target.result.split(",")[1];
    } else {
      const bytes = new Uint8Array(e.target.result);
      const chunks = [];
      for (let i = 0; i < bytes.length; i += 8192) {
        chunks.push(String.fromCharCode.apply(null, bytes.subarray(i, i + 8192)));
      }
      base64 = btoa(chunks.join(""));
    }
    if (type === "xlsx") {
      fetch("/api/xlsx/open", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({ data: base64, fileName: file.name }) })
        .then(r => r.json()).then(json => {
          if (json.ok) { xlsxData = json.data; xlsxActiveSheet = 0; renderXlsx(); navigateTo("xlsx"); }
          else showToast(json.error || "Failed to open xlsx", "error");
        }).catch(err => showToast("Error: " + err.message, "error"));
    } else if (type === "docx") {
      fetch("/api/docx/open", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({ data: base64, fileName: file.name }) })
        .then(r => r.json()).then(json => {
          if (json.ok) { docxData = json.data; renderDocx(); navigateTo("docx"); }
          else showToast(json.error || "Failed to open docx", "error");
        }).catch(err => showToast("Error: " + err.message, "error"));
    }
  };
  reader.readAsArrayBuffer(file);
}

// Drag and drop
const dropZone = document.getElementById("dropZone");
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "xlsx" || ext === "xls") openFile(file, "xlsx");
  else if (ext === "docx") openFile(file, "docx");
  else showToast("Unsupported file type", "error");
});

// ===== XLSX Rendering =====
function renderXlsx() {
  if (!xlsxData || !xlsxData.sheets.length) return;
  const sheet = xlsxData.sheets[xlsxActiveSheet];
  if (!sheet) return;

  // Sheet tabs
  const tabsEl = document.getElementById("sheetTabs");
  tabsEl.innerHTML = "";
  xlsxData.sheets.forEach((s, i) => {
    const tab = document.createElement("button");
    tab.className = "sheet-tab" + (i === xlsxActiveSheet ? " active" : "");
    tab.textContent = s.name;
    tab.onclick = () => { xlsxActiveSheet = i; renderXlsx(); };
    tabsEl.appendChild(tab);
  });
  // Add sheet button
  const addTab = document.createElement("button");
  addTab.className = "sheet-tab sheet-tab-add";
  addTab.textContent = "+";
  addTab.onclick = async () => {
    const name = prompt("Sheet name:", "Sheet" + (xlsxData.sheets.length + 1));
    if (!name) return;
    const res = await fetch("/api/xlsx/add-sheet", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({name}) });
    const json = await res.json();
    if (json.ok) { xlsxData = json.data; xlsxActiveSheet = xlsxData.sheets.length - 1; renderXlsx(); }
  };
  tabsEl.appendChild(addTab);

  // Spreadsheet
  const container = document.getElementById("spreadsheetContainer");
  const rows = sheet.cells || [];
  const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const merges = sheet.merges || [];
  const colWidths = sheet.colWidths || [];
  const rowHeights = sheet.rowHeights || [];

  let html = '<table class="spreadsheet"><thead><tr><th class="corner"></th>';
  for (let c = 0; c < maxCols; c++) {
    const w = colWidths[c] ? Math.max(colWidths[c] * 8, 40) : 80; // convert char width approx
    html += '<th style="width:' + w + 'px">' + colName(c) + '</th>';
  }
  html += '</tr></thead><tbody>';

  for (let r = 0; r < rows.length; r++) {
    const rh = rowHeights[r];
    const rowStyle = rh ? ' style="height:' + (rh * 1.3) + 'px"' : '';
    html += '<tr' + rowStyle + '><th class="row-header">' + (r + 1) + '</th>';
    for (let c = 0; c < maxCols; c++) {
      const cell = rows[r][c] || {};
      const isMerged = cell.isMerged || false;
      // Check if this is the top-left of a merge
      const merge = merges.find(m => m.s.r === r && m.s.c === c);
      const rowspan = merge ? (merge.e.r - merge.s.r + 1) : 1;
      const colspan = merge ? (merge.e.c - merge.s.c + 1) : 1;
      if (isMerged && !merge) continue; // skip non-top-left merged cells

      const val = cell.v !== null && cell.v !== undefined ? cell.v : "";
      const dropdown = cell.dropdown;
      // Build inline style from cell.style object
      const cellStyle = buildCellStyle(cell.style);
      let cellContent;
      if (dropdown && dropdown.length > 0) {
        cellContent = '<select class="cell-select" onchange="onCellChange(' + r + ',' + c + ',this.value)" style="' + cellStyle + '">' +
          '<option value="">' + t('xlsx.dropdown') + '</option>' +
          dropdown.map(v => '<option value="' + escHtml(v) + '"' + (String(val) === String(v) ? ' selected' : '') + '>' + escHtml(v) + '</option>').join('') +
          '</select>';
      } else {
        cellContent = '<div class="cell-value" contenteditable="true" onfocus="onCellFocus(this,' + r + ',' + c + ')" onblur="onCellBlur(this,' + r + ',' + c + ')" oninput="onCellInput(this,' + r + ',' + c + ')" style="' + cellStyle + '">' + escHtml(String(val)) + '</div>';
      }

      html += '<td' + (isMerged ? ' class="merged"' : '') +
        (rowspan > 1 ? ' rowspan="' + rowspan + '"' : '') +
        (colspan > 1 ? ' colspan="' + colspan + '"' : '') +
        ' data-r="' + r + '" data-c="' + c + '">' + cellContent + '</td>';
    }
    html += '</tr>';
  }
  html += '</tbody></table>';
  container.innerHTML = html;
  selectedCells = [];
}

/**
 * Build CSS inline style string from CellStyle data
 */
function buildCellStyle(style) {
  if (!style || typeof style !== "object") return "";
  const parts = [];
  if (style.bold) parts.push("font-weight:bold");
  if (style.italic) parts.push("font-style:italic");
  if (style.underline) parts.push("text-decoration:underline");
  if (style.strike) parts.push("text-decoration:line-through");
  if (style.fontSize) parts.push("font-size:" + style.fontSize + "px");
  if (style.fontName) parts.push('font-family:"' + escHtml(style.fontName) + '"');
  if (style.fontColor) parts.push("color:" + style.fontColor);
  if (style.fgColor || style.bgColor) parts.push("background-color:" + (style.fgColor || style.bgColor));
  if (style.hAlign) parts.push("text-align:" + style.hAlign);
  if (style.vAlign === "top") parts.push("vertical-align:top");
  else if (style.vAlign === "bottom") parts.push("vertical-align:bottom");
  else if (style.vAlign === "middle") parts.push("vertical-align:middle");
  if (style.wrapText) parts.push("white-space:pre-wrap");
  return parts.join(";") + ";";
}

function colName(c) {
  let name = "";
  while (c >= 0) { name = String.fromCharCode(65 + (c % 26)) + name; c = Math.floor(c / 26) - 1; }
  return name;
}

function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

let cellChangeTimer = null;
function onCellInput(el, r, c) {
  clearTimeout(cellChangeTimer);
  cellChangeTimer = setTimeout(() => {
    const val = el.textContent.trim();
    const num = Number(val);
    const sendVal = val === "" ? null : (isNaN(num) ? val : num);
    fetch("/api/xlsx/cell", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ sheetIndex: xlsxActiveSheet, r, c, value: sendVal }) })
      .then(res => res.json()).then(json => { if (json.ok) xlsxData = json.data; });
  }, 300);
}

function onCellChange(r, c, value) {
  const num = Number(value);
  const sendVal = value === "" ? null : (isNaN(num) ? value : num);
  fetch("/api/xlsx/cell", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ sheetIndex: xlsxActiveSheet, r, c, value: sendVal }) })
    .then(res => res.json()).then(json => { if (json.ok) xlsxData = json.data; });
}

function onCellFocus(el, r, c) {
  // Track selection for merge
  const idx = selectedCells.findIndex(s => s.r === r && s.c === c);
  if (idx === -1) selectedCells.push({r, c});
}

function onCellBlur(el, r, c) {
  // Remove from selection on blur
  setTimeout(() => {
    selectedCells = selectedCells.filter(s => !(s.r === r && s.c === c));
  }, 200);
}

// ===== XLSX Actions =====
async function saveXlsx() {
  if (!xlsxData) return;
  showToast(t('common.save') + "...");
}

async function downloadXlsx() {
  if (!xlsxData) return;
  const res = await fetch("/api/xlsx/save", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({}) });
  const json = await res.json();
  if (json.ok) {
    const blob = base64ToBlob(json.data.data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    downloadBlob(blob, json.data.fileName || "output.xlsx");
    showToast(t('xlsx.export') + " ✓", "success");
  }
}

async function applyRulesToXlsx() {
  const res = await fetch("/api/xlsx/apply-rules", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({}) });
  const json = await res.json();
  if (json.ok) { xlsxData = json.data; renderXlsx(); showToast(t('xlsx.applyRule') + " ✓", "success"); }
}

async function addXlsxRow() {
  if (!xlsxData) return;
  const res = await fetch("/api/xlsx/add-row", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ sheetIndex: xlsxActiveSheet }) });
  const json = await res.json();
  if (json.ok) { xlsxData = json.data; renderXlsx(); }
}

async function addXlsxCol() {
  if (!xlsxData) return;
  const res = await fetch("/api/xlsx/add-col", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ sheetIndex: xlsxActiveSheet }) });
  const json = await res.json();
  if (json.ok) { xlsxData = json.data; renderXlsx(); }
}

async function mergeSelectedCells() {
  if (!xlsxData || selectedCells.length < 2) {
    showToast("Select at least 2 cells to merge", "error");
    return;
  }
  const rows = selectedCells.map(s => s.r);
  const cols = selectedCells.map(s => s.c);
  const startR = Math.min(...rows), endR = Math.max(...rows);
  const startC = Math.min(...cols), endC = Math.max(...cols);
  const res = await fetch("/api/xlsx/merge", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ sheetIndex: xlsxActiveSheet, startR, startC, endR, endC }) });
  const json = await res.json();
  if (json.ok) { xlsxData = json.data; renderXlsx(); showToast(t('xlsx.mergeCell') + " ✓", "success"); }
}

// ===== DOCX Rendering =====
function renderDocx() {
  if (!docxData) return;
  const container = document.getElementById("docxContent");
  const paragraphs = docxData.content.paragraphs || [];
  const images = docxData.content.images || [];

  let html = "";

  // Render images first (if they exist)
  if (images && images.length > 0) {
    html += '<div style="margin:12px 0;padding:8px;background:#f8f9fa;border:1px solid #e0e0e0;border-radius:6px;">';
    html += '<strong style="display:block;margin-bottom:6px;color:#666;font-size:12px;">' + t('docx.images' || 'Images') + ' (' + images.length + ')</strong>';
    for (const img of images) {
      const wPx = Math.round(img.width / 9525); // EMU to px (approximate, 1 inch=914400 EMU, 96dpi => 1 inch=96px => 9525 EMU/px)
      const hPx = Math.round(img.height / 9525);
      html += '<img class="docx-image" src="data:image/png;base64,' + escHtml(img.data) + '" width="' + Math.min(wPx, 600) + '" alt="' + escHtml(img.altText || "image") + '" style="max-width:100%;height:auto;border:1px solid #ddd;border-radius:4px;">';
    }
    html += '</div>';
  }

  // Render paragraphs with styles
  html += paragraphs.map((p, i) => {
    const tag = p.heading ? "h" + p.heading : "p";
    const styleParts = [];

    // Build inline style from paragraph properties
    if (p.bold) styleParts.push("font-weight:bold");
    if (p.italic) styleParts.push("font-style:italic");
    if (p.underline) styleParts.push("text-decoration:underline");
    if (p.strike) styleParts.push("text-decoration:line-through");
    if (p.fontSize) styleParts.push("font-size:" + Math.max(p.fontSize, 10) + "px");
    if (p.fontColor) styleParts.push("color:" + p.fontColor);
    if (p.alignment === "center") styleParts.push("text-align:center");
    else if (p.alignment === "right") styleParts.push("text-align:right");

    const inlineStyle = styleParts.length > 0 ? ' style="' + styleParts.join(";") + '"' : '';

    return '<div class="docx-paragraph" contenteditable="true" data-index="' + i + '" onblur="onDocxBlur(this,' + i + ')">' +
      '<' + tag + inlineStyle + '>' + escHtml(p.text || "") + '</' + tag + '></div>';
  }).join("");

  // Show empty state message
  if (!html.trim()) {
    html = '<div style="padding:40px;text-align:center;color:#999;font-size:14px;">' +
      escHtml(t('docx.empty') || 'No content found. This document may be empty or use unsupported formatting.') +
      '</div>';
  }

  container.innerHTML = html;

  // If there are images, show a toast notification
  if (images && images.length > 0) {
    showToast(t('docx.imageCount')?.replace('{n}', String(images.length)) || ('Found ' + images.length + ' image(s)'), "success", 5000);
  }
}

function onDocxBlur(el, index) {
  if (!docxData) return;
  const tag = el.querySelector("h1,h2,h3,h4,h5");
  let heading;
  if (tag) heading = parseInt(tag.tagName[1]);
  const text = el.textContent || "";
  const style = (tag || el).style;
  docxData.content.paragraphs[index] = {
    text,
    bold: style.fontWeight === "bold" || parseInt(style.fontWeight) >= 700,
    italic: style.fontStyle === "italic",
    underline: style.textDecoration.includes("underline"),
    strike: style.textDecoration.includes("line-through"),
    fontSize: style.fontSize ? parseFloat(style.fontSize) : undefined,
    fontColor: style.color || undefined,
    heading,
  };
  // Sync to server
  fetch("/api/docx/update-paragraph", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ index, paragraph: docxData.content.paragraphs[index] }) })
    .then(r => r.json()).then(json => { if (json.ok) docxData = json.data; });
}

async function addDocxParagraph() {
  if (!docxData) return;
  const res = await fetch("/api/docx/add-paragraph", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ paragraph: { text: "" } }) });
  const json = await res.json();
  if (json.ok) { docxData = json.data; renderDocx(); }
}

async function saveDocx() {
  showToast(t('common.save') + "...");
}

async function downloadDocx() {
  if (!docxData) return;
  const res = await fetch("/api/docx/save", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({}) });
  const json = await res.json();
  if (json.ok) {
    const blob = base64ToBlob(json.data.data, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    downloadBlob(blob, json.data.fileName || "output.docx");
    showToast(t('xlsx.export') + " ✓", "success");
  }
}

// ===== Settings =====
async function loadAndRenderSettings() {
  const res = await fetch("/api/settings");
  const json = await res.json();
  if (!json.ok) return;
  const s = json.data;

  document.getElementById("settingLang").value = s.lang || "zh-CN";
  document.getElementById("settingApiHost").value = s.apiHost || "";
  currentLang = s.lang || "zh-CN";

  // Rules
  const rulesEl = document.getElementById("rulesContainer");
  rulesEl.innerHTML = (s.rules || []).map((rule, i) => {
    return '<div class="rule-item">' +
      '<select onchange="updateRuleType(' + i + ',this.value)">' +
        '<option value="cell"' + (rule.type==="cell"?" selected":"") + '>cell</option>' +
        '<option value="mergeCell"' + (rule.type==="mergeCell"?" selected":"") + '>mergeCell</option>' +
        '<option value="rowCell"' + (rule.type==="rowCell"?" selected":"") + '>rowCell</option>' +
        '<option value="alias"' + (rule.type==="alias"?" selected":"") + '>alias</option>' +
      '</select>' +
      '<input placeholder="Name" value="' + escHtml(rule.name||"") + '" onchange="updateRuleField(' + i + ',\\'name\\',this.value)">' +
      '<input placeholder="Values (comma sep)" value="' + escHtml((rule.values||[]).join(",")) + '" onchange="updateRuleValues(' + i + ',this.value)">' +
      '<input placeholder="Apply/Range" value="' + escHtml(rule.applyTo||rule.range||"") + '" onchange="updateRuleApply(' + i + ',this.value)">' +
      '<button class="btn btn-sm btn-danger" onclick="removeRule(' + i + ')">✕</button>' +
    '</div>';
  }).join("");

  // Custom functions
  const fnEl = document.getElementById("fnContainer");
  fnEl.innerHTML = (s.customFunctions || []).map((fn, i) => {
    return '<div class="fn-item"><div class="fn-item-header">' +
      '<input placeholder="Function name" value="' + escHtml(fn.name||"") + '" onchange="updateFnField(' + i + ',\\'name\\',this.value)">' +
      '<button class="btn btn-sm btn-danger" onclick="removeFn(' + i + ')">✕</button>' +
    '</div><textarea placeholder="JS code" onchange="updateFnField(' + i + ',\\'code\\',this.value)">' + escHtml(fn.code||"") + '</textarea></div>';
  }).join("");
}

function getSettingsFromForm() {
  const rules = [];
  document.querySelectorAll(".rule-item").forEach((el, i) => {
    const selects = el.querySelectorAll("select");
    const inputs = el.querySelectorAll("input");
    const type = selects[0]?.value || "cell";
    const name = inputs[0]?.value || "";
    const values = (inputs[1]?.value || "").split(",").map(v => v.trim()).filter(Boolean);
    const applyTo = inputs[2]?.value || "";
    const rule = { name, type, values };
    if (type === "rowCell") rule.range = applyTo;
    else rule.applyTo = applyTo;
    rules.push(rule);
  });

  const customFunctions = [];
  document.querySelectorAll(".fn-item").forEach(el => {
    const input = el.querySelector("input");
    const textarea = el.querySelector("textarea");
    customFunctions.push({ name: input?.value || "", code: textarea?.value || "" });
  });

  return {
    lang: document.getElementById("settingLang").value,
    apiHost: document.getElementById("settingApiHost").value,
    rules,
    customFunctions,
  };
}

async function saveSettings() {
  const settings = getSettingsFromForm();
  const res = await fetch("/api/settings", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify(settings) });
  const json = await res.json();
  if (json.ok) {
    showToast(t('settings.saved'), "success");
    loadI18n(settings.lang);
  }
}

// Rules inline editing
function updateRuleType(i, val) { /* type changed, will be picked up on save */ }
function updateRuleField(i, field, val) { /* name changed */ }
function updateRuleValues(i, val) { /* values changed */ }
function updateRuleApply(i, val) { /* applyTo/range changed */ }

function addRule() {
  const type = document.getElementById("newRuleType").value;
  const s = getSettingsFromForm();
  s.rules.push({ name: "", type, values: [], applyTo: "", range: "" });
  // Re-render
  const rulesEl = document.getElementById("rulesContainer");
  const html = '<div class="rule-item">' +
    '<select onchange="updateRuleType(' + (s.rules.length-1) + ',this.value)">' +
      '<option value="cell"' + (type==="cell"?" selected":"") + '>cell</option>' +
      '<option value="mergeCell"' + (type==="mergeCell"?" selected":"") + '>mergeCell</option>' +
      '<option value="rowCell"' + (type==="rowCell"?" selected":"") + '>rowCell</option>' +
      '<option value="alias"' + (type==="alias"?" selected":"") + '>alias</option>' +
    '</select>' +
    '<input placeholder="Name" value="">' +
    '<input placeholder="Values (comma sep)" value="">' +
    '<input placeholder="Apply/Range" value="">' +
    '<button class="btn btn-sm btn-danger" onclick="this.parentElement.remove()">✕</button>' +
  '</div>';
  rulesEl.insertAdjacentHTML("beforeend", html);
}

function removeRule(i) {
  // find and remove the i-th rule item
  const items = document.querySelectorAll(".rule-item");
  if (items[i]) items[i].remove();
}

function addCustomFunction() {
  const fnEl = document.getElementById("fnContainer");
  fnEl.insertAdjacentHTML("beforeend",
    '<div class="fn-item"><div class="fn-item-header">' +
    '<input placeholder="Function name" value="">' +
    '<button class="btn btn-sm btn-danger" onclick="this.closest(\\'.fn-item\\').remove()">✕</button>' +
    '</div><textarea placeholder="JS code"></textarea></div>'
  );
}

function removeFn(i) {
  const items = document.querySelectorAll(".fn-item");
  if (items[i]) items[i].remove();
}

// ===== Helpers =====
function base64ToBlob(b64, mime) {
  const bytes = atob(b64);
  const buf = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) buf[i] = bytes.charCodeAt(i);
  return new Blob([buf], { type: mime });
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function showToast(msg, type, duration) {
  const toast = document.createElement("div");
  toast.className = "toast" + (type ? " " + type : "");
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), duration || 2500);
}

// ===== Init =====
// Load i18n based on server-detected system language
fetch("/api/i18n").then(r => r.json()).then(json => {
  if (json.ok) {
    i18nMessages = json.data.messages;
    currentLang = json.data.lang;
    applyI18n();
    document.getElementById("settingLang").value = currentLang;
  }
});
// Also load settings to sync
fetch("/api/settings").then(r => r.json()).then(json => {
  if (json.ok) {
    document.getElementById("settingLang").value = json.data.lang || currentLang;
    document.getElementById("settingApiHost").value = json.data.apiHost || "";
  }
});
</script>
</body>
</html>`;
}

if (!import.meta.main) {
    startServer();
}
