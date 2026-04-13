// tests/i18n.test.ts
import { describe, it, expect } from "bun:test";
import { messages, type Lang } from "../src/i18n/index.ts";

describe("i18n", () => {
    const langs: Lang[] = ["zh-CN", "zh-TW", "en", "ja"];

    describe("all languages have required keys", () => {
        const requiredKeyPatterns = [
            "app.title",
            "common.save",
            "common.cancel",
            "nav.home",
            "nav.settings",
            "home.title",
            "home.openXlsx",
            "home.openDocx",
            "xlsx.editor",
            "xlsx.save",
            "xlsx.mergeCell",
            "xlsx.dropdown",
            "xlsx.manualInput",
            "docx.editor",
            "docx.save",
            "settings.title",
            "settings.lang",
            "settings.api",
            "settings.rules",
            "settings.customFunctions",
            "settings.saved",
        ];

        for (const lang of langs) {
            describe(`lang: ${lang}`, () => {
                it("should exist", () => {
                    expect(messages[lang]).toBeDefined();
                });

                for (const key of requiredKeyPatterns) {
                    it(`should have key "${key}"`, () => {
                        expect(messages[lang][key]).toBeDefined();
                        expect(messages[lang][key]).toBeTruthy();
                    });
                }
            });
        }
    });

    describe("language options", () => {
        it("should have 4 languages", () => {
            expect(Object.keys(messages)).toHaveLength(4);
        });

        it("should have consistent lang option keys", () => {
            for (const lang of langs) {
                expect(messages[lang]["settings.lang.zh-CN"]).toBeDefined();
                expect(messages[lang]["settings.lang.zh-TW"]).toBeDefined();
                expect(messages[lang]["settings.lang.en"]).toBeDefined();
                expect(messages[lang]["settings.lang.ja"]).toBeDefined();
            }
        });
    });
});
