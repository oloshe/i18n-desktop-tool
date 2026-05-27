import { describe, expect, it } from "vitest";
import { parseLocaleContent } from "../core/localeFileReader";

describe("localeFileReader", () => {
  it("parses json locale files", () => {
    expect(parseLocaleContent('{ "hello": "Hello" }', "json")).toEqual({ hello: "Hello" });
  });

  it("parses js default export object", () => {
    expect(parseLocaleContent('export default { hello: "Hello", count: 1 };', "js")).toEqual({
      hello: "Hello",
      count: "1"
    });
  });

  it("parses ts default export object with as const", () => {
    expect(parseLocaleContent('export default { hello: "Hello" } as const;', "ts")).toEqual({ hello: "Hello" });
  });

  it("parses named exported locale objects and ignores surrounding code", () => {
    const source = `
const win = window as any;
// 阿语
export const languages = {
  gachaguess: {
    title: "مرحبا",
    rules: { header: "قواعد" }
  }
};

win.languages.ar = languages;
`;

    expect(parseLocaleContent(source, "ts")).toEqual({
      gachaguess: {
        title: "مرحبا",
        rules: { header: "قواعد" }
      }
    });
  });
});
