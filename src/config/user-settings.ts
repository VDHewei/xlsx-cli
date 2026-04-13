// src/config/user-settings.ts - 用户配置管理模块

import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import type { Lang } from "../i18n/index.ts";

export type RuleType = "mergeCell" | "cell" | "rowCell" | "alias";

export type CellRule = {
    name: string;
    type: RuleType;
    values: string[];
    range?: string; // for rowCell: e.g. "1-10"
    applyTo?: string; // cell ref, e.g. "A1" or "A1:B5"
};

export type CustomFunction = {
    name: string;
    code: string;
};

export type UserSettings = {
    lang: Lang;
    apiHost: string;
    rules: CellRule[];
    customFunctions: CustomFunction[];
};

const getSettingsDir = (): string => {
    return join(
        process.env.APPDATA ||
        process.env.XDG_CONFIG_HOME ||
        join(process.env.HOME || process.cwd(), ".config"),
        "xlsx-cli"
    );
};

const getSettingsFile = (): string => {
    return join(getSettingsDir(), "settings.json");
};

const defaultSettings: UserSettings = {
    lang: "zh-CN",
    apiHost: "http://localhost:3000",
    rules: [],
    customFunctions: [],
};

export const loadSettings = (): UserSettings => {
    try {
        const settingsFile = getSettingsFile();
        if (existsSync(settingsFile)) {
            const raw = readFileSync(settingsFile, "utf-8");
            const parsed = JSON.parse(raw) as Partial<UserSettings>;
            return {
                lang: parsed.lang ?? defaultSettings.lang,
                apiHost: parsed.apiHost ?? defaultSettings.apiHost,
                rules: Array.isArray(parsed.rules) ? parsed.rules : defaultSettings.rules,
                customFunctions: Array.isArray(parsed.customFunctions)
                    ? parsed.customFunctions
                    : defaultSettings.customFunctions,
            };
        }
    } catch {
        // ignore parse errors
    }
    return { ...defaultSettings };
};

export const saveSettings = (settings: UserSettings): void => {
    try {
        const dir = getSettingsDir();
        mkdirSync(dir, { recursive: true });
        writeFileSync(getSettingsFile(), JSON.stringify(settings, null, 2), "utf-8");
    } catch (e) {
        console.error("Failed to save settings:", e);
    }
};

export { defaultSettings, getSettingsDir, getSettingsFile };
