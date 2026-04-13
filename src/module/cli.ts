// src/cli.ts - 命令行模式

import { createCommand } from "commander";
import { readXlsx, writeXlsx } from "./xlsx-engine.ts";
import { readDocx, writeDocx } from "./docx-engine.ts";
import { App } from "@config/app.ts";
import { readFileSync, writeFileSync } from "node:fs";

export type CliMode = "cli" | "ui";

export const createProgram = () => {
    const program = createCommand();

    program
        .name("xlsx-cli")
        .description("Lightweight Office file editor CLI")
        .version(App.version);

    // xlsx 命令
    program
        .command("xlsx")
        .description("XLSX file operations")
        .argument("<file>", "XLSX file path")
        .option("-o, --output <file>", "Output file path")
        .option("--info", "Show file info")
        .action((file: string, opts: { output?: string; info?: boolean }) => {
            const data = readFileSync(file);
            const doc = readXlsx(data.buffer as ArrayBuffer);

            if (opts.info) {
                console.log("=== XLSX File Info ===");
                console.log("Sheets:", doc.sheets.length);
                for (const sheet of doc.sheets) {
                    console.log(`  "${sheet.name}": ${sheet.cells.length} rows x ${sheet.cells[0]?.length || 0} cols`);
                    if (sheet.merges.length > 0) {
                        console.log(`  Merges: ${sheet.merges.length}`);
                    }
                }
                return;
            }

            // Read mode: print content
            console.log(`=== ${file} ===`);
            for (const sheet of doc.sheets) {
                console.log(`\n[Sheet: ${sheet.name}]`);
                for (const row of sheet.cells) {
                    const vals = row.map((c) => String(c.v ?? "")).join("\t");
                    console.log(vals);
                }
            }
        });

    // xlsx convert 命令
    program
        .command("xlsx-convert")
        .description("Convert XLSX to CSV/JSON")
        .argument("<file>", "XLSX file path")
        .option("-f, --format <format>", "Output format: csv, json", "csv")
        .option("-o, --output <file>", "Output file path")
        .action((file: string, opts: { format: string; output?: string }) => {
            const data = readFileSync(file);
            const doc = readXlsx(data.buffer as ArrayBuffer);

            if (opts.format === "json") {
                const output = JSON.stringify(doc.sheets, null, 2);
                if (opts.output) {
                    writeFileSync(opts.output, output);
                    console.log(`Written to ${opts.output}`);
                } else {
                    console.log(output);
                }
            } else {
                // CSV
                const lines: string[] = [];
                for (const sheet of doc.sheets) {
                    lines.push(`# Sheet: ${sheet.name}`);
                    for (const row of sheet.cells) {
                        lines.push(row.map((c) => {
                            const v = String(c.v ?? "");
                            return v.includes(",") || v.includes('"') ? `"${v.replace(/"/g, '""')}"` : v;
                        }).join(","));
                    }
                    lines.push("");
                }
                const output = lines.join("\n");
                if (opts.output) {
                    writeFileSync(opts.output, output);
                    console.log(`Written to ${opts.output}`);
                } else {
                    console.log(output);
                }
            }
        });

    // docx 命令
    program
        .command("docx")
        .description("DOCX file operations")
        .argument("<file>", "DOCX file path")
        .option("--info", "Show file info")
        .action(async (file: string, opts: { info?: boolean }) => {
            const data = readFileSync(file);
            const doc = await readDocx(data.buffer as ArrayBuffer);

            if (opts.info) {
                console.log("=== DOCX File Info ===");
                console.log("Paragraphs:", doc.content.paragraphs.length);
                return;
            }

            console.log(`=== ${file} ===`);
            for (const p of doc.content.paragraphs) {
                const prefix = p.heading ? `H${p.heading}: ` : "";
                const style = [p.bold ? "[B]" : "", p.italic ? "[I]" : "", p.underline ? "[U]" : ""].filter(Boolean).join(" ");
                console.log(`${style}${prefix}${p.text}`);
            }
        });

    return program;
};

/**
 * 检测执行模式：如果有任何参数则为 CLI 模式，否则为 UI 模式。
 */
export const detectMode = (argv: string[]): CliMode => {
    const args = argv.slice(2);
    return args.length > 0 ? "cli" : "ui";
};

/**
 * 解析命令行参数并执行。
 * 返回 true 表示有命令被匹配并执行（CLI 模式），false 表示应进入 UI 模式。
 */
export const runCli = (argv: string[]): boolean => {
    const program = createProgram();
    program.parse(argv, { from: "user" });
    return program.args.length > 0;
};
