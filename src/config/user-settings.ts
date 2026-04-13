// src/config/user-settings.ts - 用户配置管理模块

import {existsSync, mkdirSync, readFileSync, writeFileSync} from "node:fs";
import {join} from "node:path";
import {includes, type Lang, languages} from "@i18n/index.ts";

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
    defaultSaveDir: string;
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

const getSystemLang = (): Lang => {
    // 默认获取 操作系统本地语言, 如果 系统语言 不在支持范围内，默认使用 English
    const systemLang = process.env.LANG || process.env.LANGUAGE || process.env.LC_MESSAGES;
    if (systemLang!==undefined && systemLang!==null && systemLang && includes(systemLang)) {
        return systemLang as Lang;
    }
    const nav = typeof navigator !== "undefined" ? navigator.language : undefined;
    console.log("navigator.language:", nav);
    if(nav!==undefined && nav!==null && nav && includes(nav)){
        return nav as Lang;
    }
    return "en" as Lang;
};

const defaultSettings: UserSettings = {
    lang: getSystemLang(),
    apiHost: "http://localhost:3000",
    defaultSaveDir: "",
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
                defaultSaveDir: parsed.defaultSaveDir ?? defaultSettings.defaultSaveDir,
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
