import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { parseExcelBuffer, readExcelWorkbookInfo, rowsFromExcelBuffer } from "../core/excelParser";

describe("excelParser", () => {
  it("reads workbook sheet names without parsing a sheet matrix", () => {
    const info = readExcelWorkbookInfo(createWorkbookBuffer(), "Data");

    expect(info.sheetNames).toEqual(["Other", "Data"]);
    expect(info.activeSheetName).toBe("Data");
  });

  it("reads headers from selected sheet and configured header row", () => {
    const buffer = createWorkbookBuffer();

    const preview = parseExcelBuffer(buffer, { sheetName: "Data", skipRows: 2, headerRow: 2 });

    expect(preview.activeSheetName).toBe("Data");
    expect(preview.headers).toEqual(["key", "cn", "en"]);
    expect(preview.rows).toEqual([{ key: "hello", cn: "你好", en: "Hello" }]);
  });

  it("keeps one-cell rows for section based module parsing", () => {
    const rows = rowsFromExcelBuffer(createWorkbookBuffer(), { sheetName: "Data", skipRows: 2, headerRow: 2 });

    expect(rows).toHaveLength(2);
    expect(rows[0].key).toBe("hello");
    expect(rows[1].key).toBe("module-only");
  });

  it("can read rows from multiple selected sheets", () => {
    const rows = rowsFromExcelBuffer(createWorkbookBuffer(), { sheetNames: ["Other", "Data"], skipRows: 2, headerRow: 2 });

    expect(rows.map((row) => row.key)).toEqual(["other-hello", "hello", "module-only"]);
  });

  it("supports default language mapping with per-sheet overrides", () => {
    const rows = rowsFromExcelBuffer(createMismatchedWorkbookBuffer(), {
      sheetNames: ["Default", "Special"],
      skipRows: 1,
      headerRow: 1,
      keyColumn: "key",
      languageColumns: { "zh-CN": "cn", "en-US": "en" },
      sheetColumnOverrides: {
        Special: {
          languageColumns: {
            "zh-CN": "zh_CN_text",
            "en-US": "en_US_text"
          }
        }
      }
    });

    expect(rows).toEqual([
      { key: "home.title", cn: "榛樿涓枃", en: "Default English" },
      { key: "settings.title", cn: "鐗规畩涓枃", en: "Special English" }
    ]);
  });
});

function createWorkbookBuffer(): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([["metadata"], ["key", "cn", "en"], ["other-hello", "鍏朵粬", "Other"]]),
    "Other"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["metadata"],
      ["key", "cn", "en"],
      ["hello", "你好", "Hello"],
      ["module-only", "", ""],
      ["", "", ""]
    ]),
    "Data"
  );
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

function createMismatchedWorkbookBuffer(): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["key", "cn", "en"],
      ["home.title", "榛樿涓枃", "Default English"]
    ]),
    "Default"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["key", "zh_CN_text", "en_US_text"],
      ["settings.title", "鐗规畩涓枃", "Special English"]
    ]),
    "Special"
  );
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}
