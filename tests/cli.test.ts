// tests/cli.test.ts
import { describe, it, expect } from "bun:test";
import { detectMode, createProgram, type CliMode } from "@module/cli.ts";

describe("CLI mode detection", () => {
    it("should detect UI mode with no args", () => {
        expect(detectMode(["bun", "src/index.ts"])).toBe("ui");
    });

    it("should detect CLI mode with xlsx command", () => {
        expect(detectMode(["bun", "src/index.ts", "xlsx", "test.xlsx"])).toBe("cli");
    });

    it("should detect CLI mode with --version", () => {
        expect(detectMode(["bun", "src/index.ts", "--version"])).toBe("cli");
    });

    it("should detect CLI mode with --help", () => {
        expect(detectMode(["bun", "src/index.ts", "--help"])).toBe("cli");
    });

    it("should detect CLI mode with xlsx-convert", () => {
        expect(detectMode(["bun", "src/index.ts", "xlsx-convert", "test.xlsx"])).toBe("cli");
    });

    it("should detect CLI mode with docx command", () => {
        expect(detectMode(["bun", "src/index.ts", "docx", "test.docx"])).toBe("cli");
    });

    it("should detect CLI mode with any option", () => {
        expect(detectMode(["bun", "src/index.ts", "--anything"])).toBe("cli");
    });
});

describe("Commander program", () => {
    it("should create a program with expected commands", () => {
        const program = createProgram();
        const commands = program.commands.map(c => c.name());
        expect(commands).toContain("xlsx");
        expect(commands).toContain("xlsx-convert");
        expect(commands).toContain("docx");
    });

    it("should have version option", () => {
        const program = createProgram();
        expect(program.version()).toBeDefined();
    });
});
