import { describe, expect, it } from "vitest";
import { detectEol, formatLocale, formatLocaleForSnapshot } from "../core/localeFileWriter";

describe("localeFileWriter", () => {
  it("formats json output", async () => {
    await expect(formatLocale({ hello: "Hello" }, "json")).resolves.toContain('"hello": "Hello"');
  });

  it("formats xcstrings output", async () => {
    const output = await formatLocale({ hello: { "en-US": "Hello", "zh-CN": "你好" } }, "xcstrings", {
      sourceLanguage: "en-US"
    });

    expect(output).toContain('"sourceLanguage": "en-US"');
    expect(output).toContain('"hello"');
    expect(output).toContain('"zh-CN"');
    expect(output).toContain('"value": "你好"');
  });

  it("formats ts output with const assertion", async () => {
    await expect(formatLocale({ hello: "Hello" }, "ts")).resolves.toContain("as const");
  });

  it("can quote js and ts object property names", async () => {
    const output = await formatLocale({ gachaguess: { title: "Big Eater" } }, "ts", { quoteObjectProperties: true });

    expect(output).toContain('"gachaguess": {');
    expect(output).toContain('"title": "Big Eater"');
  });

  it("keeps crlf and trailing newline options", async () => {
    const output = await formatLocale({ hello: "Hello" }, "json", { eol: "crlf", ensureTrailingNewline: true });

    expect(output).toContain("\r\n");
    expect(output.endsWith("\r\n")).toBe(true);
    expect(detectEol(output)).toBe("crlf");
  });

  it("can remove the trailing newline", async () => {
    const output = await formatLocale({ hello: "Hello" }, "json", { eol: "lf", ensureTrailingNewline: false });

    expect(output.endsWith("\n")).toBe(false);
  });

  it("replaces only the exported object in complex ts files", async () => {
    const content = `const win = window as any;
// 阿语
export const languages = {
  "gachaguess": {
    "title": "old"
  }
};

win.languages.ar = languages;
`;
    const output = await formatLocaleForSnapshot(
      { gachaguess: { title: "new", cheese: "جبن" } },
      "ts",
      {
        exists: true,
        content,
        locale: { gachaguess: { title: "old" } },
        eol: "lf",
        objectRange: { start: content.indexOf("{"), end: content.indexOf("};") + 1, exportKind: "named", exportName: "languages" }
      },
      { eol: "lf", ensureTrailingNewline: true, quoteObjectProperties: true }
    );

    expect(output).toContain("const win = window as any;");
    expect(output).toContain("// 阿语");
    expect(output).toContain("win.languages.ar = languages;");
    expect(output).toContain('"cheese": "جبن"');
    expect(output).not.toContain('"title": "old"');
  });
});
