import { describe, expect, it } from "vitest";
import { generateLocales, mergeLocaleObjects, resolveLocalePath } from "../core/localeGenerator";

describe("localeGenerator", () => {
  it("resolves locale paths from lang and module templates", () => {
    expect(resolveLocalePath("locales/{lang}/index.js", "zh-CN")).toBe("locales/zh-CN/index.js");
    expect(resolveLocalePath("src\\i18n\\{lang}.ts", "en-US")).toBe("src/i18n/en-US.ts");
    expect(resolveLocalePath("locales/{lang}/{module}.json", "zh-CN", "votings")).toBe("locales/zh-CN/votings.json");
    expect(resolveLocalePath("locales/{lang}.json", "zh-CN", "votings")).toBe("locales/zh-CN.json");
  });

  it("generates nested locale objects from configured columns", () => {
    const rows = [
      { key: "votings.tasks.title", cn: "标题", en: "Title" },
      { key: "common.ok", cn: "确定", en: "OK" }
    ];

    expect(generateLocales(rows, "key", { "zh-CN": "cn", "en-US": "en" }, { splitByModule: false, keyStyle: "nested" })).toEqual({
      "zh-CN": {
        "": {
          votings: { tasks: { title: "标题" } },
          common: { ok: "确定" }
        }
      },
      "en-US": {
        "": {
          votings: { tasks: { title: "Title" } },
          common: { ok: "OK" }
        }
      }
    });
  });

  it("skips empty language cells so fallback languages can resolve missing keys", () => {
    const rows = [
      { key: "room-mic-new-users-after-owner", cn: "房主之后的新用户", en: "" },
      { key: "follow-anchors", cn: "关注主播", en: "Follow anchors" }
    ];

    expect(generateLocales(rows, "key", { "zh-CN": "cn", "en-US": "en" }, { splitByModule: false, keyStyle: "nested" })).toEqual({
      "zh-CN": {
        "": {
          "room-mic-new-users-after-owner": "房主之后的新用户",
          "follow-anchors": "关注主播"
        }
      },
      "en-US": {
        "": {
          "follow-anchors": "Follow anchors"
        }
      }
    });
  });

  it("generates module split locale objects", () => {
    const rows = [{ key: "votings.tasks.other-room-stay-time", cn: "停留时间" }];

    expect(generateLocales(rows, "key", { "zh-CN": "cn" }, { splitByModule: true, keyStyle: "nested" })).toEqual({
      "zh-CN": {
        votings: {
          tasks: {
            "other-room-stay-time": "停留时间"
          }
        }
      }
    });
  });

  it("skips key-only helper rows when splitting by key prefix", () => {
    const rows = [
      { key: "epr 绿洲探险", zh: "" },
      { key: "epr.diceReward", zh: "累計獎勵" },
      { key: "epr.taskReward", zh: "獲取骰子" }
    ];

    expect(generateLocales(rows, "key", { zh: "zh" }, { moduleSplitMode: "keyPrefix", keyStyle: "nested" })).toEqual({
      zh: {
        epr: {
          diceReward: "累計獎勵",
          taskReward: "獲取骰子"
        }
      }
    });
  });

  it("generates modules from no-prefix section rows", () => {
    const rows = [
      { key: "大胃王", cn: "", en: "" },
      { key: "title", cn: "大胃王", en: "Big Eater" },
      { key: "cheese", cn: "奶酪", en: "Cheese" },
      { key: "商城", cn: "", en: "" },
      { key: "title", cn: "商城", en: "Shop" }
    ];

    expect(
      generateLocales(rows, "key", { "zh-CN": "cn", "en-US": "en" }, { moduleSplitMode: "sectionRow", keyStyle: "nested" })
    ).toEqual({
      "zh-CN": {
        大胃王: {
          title: "大胃王",
          cheese: "奶酪"
        },
        商城: {
          title: "商城"
        }
      },
      "en-US": {
        大胃王: {
          title: "Big Eater",
          cheese: "Cheese"
        },
        商城: {
          title: "Shop"
        }
      }
    });
  });

  it("can use key prefix as module name for section-row grouped content", () => {
    const rows = [
      { key: "epr 绿洲探险", zh: "" },
      { key: "epr.diceReward", zh: "累計獎勵" },
      { key: "epr.taskReward", zh: "獲取骰子" }
    ];

    expect(
      generateLocales(rows, "key", { zh: "zh" }, {
        moduleSplitMode: "sectionRow",
        moduleNameSource: "keyPrefix",
        keyStyle: "nested"
      })
    ).toEqual({
      zh: {
        epr: {
          diceReward: "累計獎勵",
          taskReward: "獲取骰子"
        }
      }
    });
  });

  it("replaces module names from section rows", () => {
    const rows = [
      { key: "大胃王", ar: "" },
      { key: "title", ar: "بيج آكل" }
    ];

    expect(
      generateLocales(rows, "key", { ar: "ar" }, { moduleSplitMode: "sectionRow", keyStyle: "nested", moduleNameReplacements: { 大胃王: "gachaguess" } })
    ).toEqual({
      ar: {
        gachaguess: {
          title: "بيج آكل"
        }
      }
    });
  });

  it("filters replaced section modules by original names", () => {
    const rows = [
      { key: "大胃王", ar: "" },
      { key: "title", ar: "بيج آكل" },
      { key: "商城", ar: "" },
      { key: "title", ar: "متجر" }
    ];

    expect(
      generateLocales(rows, "key", { ar: "ar" }, {
        moduleSplitMode: "sectionRow",
        keyStyle: "nested",
        moduleFilter: ["大胃王"],
        moduleNameReplacements: { 大胃王: "gachaguess" }
      })
    ).toEqual({
      ar: {
        gachaguess: {
          title: "بيج آكل"
        }
      }
    });
  });

  it("wraps configured language values with spaces", () => {
    const rows = [{ key: "title", ar: "مرحبا", ur: "سلام", en: "Hello" }];

    expect(
      generateLocales(rows, "key", { ar: "ar", ur: "ur", en: "en" }, {
        splitByModule: false,
        keyStyle: "nested",
        spaceWrappedLanguages: ["ar", "ur"]
      })
    ).toEqual({
      ar: { "": { title: " مرحبا " } },
      ur: { "": { title: " سلام " } },
      en: { "": { title: "Hello" } }
    });
  });

  it("supports flat key style", () => {
    const rows = [{ key: "votings.tasks.title", en: "Title" }];

    expect(generateLocales(rows, "key", { "en-US": "en" }, { splitByModule: false, keyStyle: "flat" })).toEqual({
      "en-US": {
        "": {
          "votings.tasks.title": "Title"
        }
      }
    });
  });

  it("filters imports by module names", () => {
    const rows = [
      { key: "base.title", en: "Base" },
      { key: "agency.title", en: "Agency" },
      { key: "other.title", en: "Other" }
    ];

    expect(
      generateLocales(rows, "key", { "en-US": "en" }, { splitByModule: false, keyStyle: "nested", moduleFilter: ["base", "agency"] })
    ).toEqual({
      "en-US": {
        "": {
          base: { title: "Base" },
          agency: { title: "Agency" }
        }
      }
    });
  });

  it("ignores modules by name", () => {
    const rows = [
      { key: "base.title", en: "Base" },
      { key: "agency.title", en: "Agency" },
      { key: "debug.title", en: "Debug" }
    ];

    expect(
      generateLocales(rows, "key", { "en-US": "en" }, { splitByModule: false, keyStyle: "nested", ignoredModuleFilter: ["debug"] })
    ).toEqual({
      "en-US": {
        "": {
          base: { title: "Base" },
          agency: { title: "Agency" }
        }
      }
    });
  });

  it("merges with overwrite strategy", () => {
    const result = mergeLocaleObjects({ common: { ok: "Old" } }, { common: { ok: "OK" }, bye: "Bye" }, "overwrite");
    expect(result.mergedLocale).toEqual({ common: { ok: "OK" }, bye: "Bye" });
    expect(result.addedKeys).toEqual(["bye"]);
    expect(result.overwrittenKeys).toEqual(["common.ok"]);
    expect(result.skippedKeys).toEqual([]);
  });

  it("merges with skip strategy", () => {
    const result = mergeLocaleObjects({ hello: "Old" }, { hello: "Hello", bye: "Bye" }, "skip");
    expect(result.mergedLocale).toEqual({ hello: "Old", bye: "Bye" });
    expect(result.addedKeys).toEqual(["bye"]);
    expect(result.overwrittenKeys).toEqual([]);
    expect(result.skippedKeys).toEqual(["hello"]);
  });

  it("can remove keys missing from incoming locale while overwriting matching keys", () => {
    const result = mergeLocaleObjects(
      { hello: "Old", stale: "Remove me", nested: { keep: "Old", stale: "Remove me" } },
      { hello: "Hello", nested: { keep: "New" } },
      "overwrite",
      "remove"
    );

    expect(result.mergedLocale).toEqual({ hello: "Hello", nested: { keep: "New" } });
    expect(result.overwrittenKeys).toEqual(["hello", "nested.keep"]);
  });

  it("can remove keys missing from incoming locale while skipping matching keys", () => {
    const result = mergeLocaleObjects(
      { hello: "Old", stale: "Remove me", nested: { keep: "Old", stale: "Remove me" } },
      { hello: "Hello", nested: { keep: "New" } },
      "skip",
      "remove"
    );

    expect(result.mergedLocale).toEqual({ hello: "Old", nested: { keep: "Old" } });
    expect(result.skippedKeys).toEqual(["hello", "nested.keep"]);
  });
});
