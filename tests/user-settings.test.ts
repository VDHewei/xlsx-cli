// tests/user-settings.test.ts
import { describe, it, expect, beforeAll, afterAll, afterEach } from "bun:test";
import { existsSync, unlinkSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import type { UserSettings } from "@config/user-settings.ts";

const TEST_DIR = join(import.meta.dir, ".tmp-test-settings");

describe("user-settings", () => {
    let originalAppData: string | undefined;
    let originalHome: string | undefined;
    let originalXdg: string | undefined;

    beforeAll(() => {
        originalAppData = process.env.APPDATA;
        originalHome = process.env.HOME;
        originalXdg = process.env.XDG_CONFIG_HOME;
        // Remove APPDATA and HOME to force XDG_CONFIG_HOME
        delete process.env.APPDATA;
        delete process.env.HOME;
        process.env.XDG_CONFIG_HOME = TEST_DIR;
        mkdirSync(TEST_DIR, { recursive: true });
    });

    afterAll(() => {
        if (originalAppData !== undefined) process.env.APPDATA = originalAppData;
        else delete process.env.APPDATA;
        if (originalHome !== undefined) process.env.HOME = originalHome;
        else delete process.env.HOME;
        if (originalXdg !== undefined) process.env.XDG_CONFIG_HOME = originalXdg;
        else delete process.env.XDG_CONFIG_HOME;
        try { rmSync(TEST_DIR, { recursive: true }); } catch { /* ignore */ }
    });

    afterEach(() => {
        const settingsFile = join(TEST_DIR, "xlsx-cli", "settings.json");
        try { unlinkSync(settingsFile); } catch { /* ignore */ }
    });

    describe("loadSettings", () => {
        it("should return default settings when no file exists", async () => {
            const { loadSettings, defaultSettings } = await import("../src/config/user-settings.ts");
            const settings = loadSettings();
            expect(settings.lang).toBe(defaultSettings.lang);
            expect(settings.apiHost).toBe(defaultSettings.apiHost);
            expect(settings.rules).toEqual([]);
            expect(settings.customFunctions).toEqual([]);
        });

        it("should load saved settings", async () => {
            const { saveSettings, loadSettings } = await import("../src/config/user-settings.ts");
            const custom: UserSettings = {
                lang: "en",
                apiHost: "http://example.com",
                rules: [{ name: "Test", type: "cell", values: ["a", "b"], applyTo: "A1" }],
                customFunctions: [{ name: "testFn", code: "return 1" }],
            };
            saveSettings(custom);

            const loaded = loadSettings();
            expect(loaded.lang).toBe("en");
            expect(loaded.apiHost).toBe("http://example.com");
            expect(loaded.rules).toHaveLength(1);
            expect(loaded.rules[0].name).toBe("Test");
            expect(loaded.customFunctions).toHaveLength(1);
            expect(loaded.customFunctions[0].name).toBe("testFn");
        });

        it("should handle corrupted settings file gracefully", async () => {
            const { loadSettings, defaultSettings } = await import("../src/config/user-settings.ts");
            mkdirSync(join(TEST_DIR, "xlsx-cli"), { recursive: true });
            writeFileSync(join(TEST_DIR, "xlsx-cli", "settings.json"), "not json{{{");

            const settings = loadSettings();
            expect(settings.lang).toBe(defaultSettings.lang);
        });
    });

    describe("saveSettings", () => {
        it("should create settings file", async () => {
            const { saveSettings, defaultSettings } = await import("../src/config/user-settings.ts");
            const settings: UserSettings = { ...defaultSettings, lang: "ja" };
            saveSettings(settings);

            const settingsFile = join(TEST_DIR, "xlsx-cli", "settings.json");
            expect(existsSync(settingsFile)).toBe(true);
        });
    });
});
