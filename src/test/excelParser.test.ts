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
});

function createWorkbookBuffer(): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["other"]]), "Other");
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
