import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { formatExcelHealthCheckReport, inspectExcelBuffer } from "../core/excelHealthCheck";

describe("excelHealthCheck", () => {
  it("reports duplicate keys, empty key rows and missing translations with sheet row numbers", () => {
    const result = inspectExcelBuffer(createWorkbookBuffer(), "key", { ar_AE: "ar", en_US: "en" }, {
      sheetNames: ["x", "y"],
      headerRow: 1,
      skipRows: 1
    });

    expect(result.duplicateKeys).toEqual([
      {
        key: "home.title",
        locations: [
          { sheetName: "x", rowNumber: 2 },
          { sheetName: "y", rowNumber: 2 }
        ]
      }
    ]);
    expect(result.emptyKeyRows).toEqual([{ sheetName: "x", rowNumber: 4 }]);
    expect(result.missingTranslations).toContainEqual({ sheetName: "x", rowNumber: 3, key: "home.subtitle", lang: "ar_AE" });
  });

  it("formats a report for copying", () => {
    const report = formatExcelHealthCheckReport({
      duplicateKeys: [{ key: "home.title", locations: [{ sheetName: "x", rowNumber: 567 }, { sheetName: "y", rowNumber: 123 }] }],
      emptyKeyRows: [{ sheetName: "x", rowNumber: 111 }],
      missingTranslations: [{ sheetName: "a", rowNumber: 442, key: "xxx", lang: "ar_AE" }]
    });

    expect(report).toContain("home.title: x sheet 567 行，y sheet 123 行");
    expect(report).toContain("x sheet 111 行");
    expect(report).toContain("sheet: a 442 行 key: xxx 缺少 ar_AE 翻译");
  });
});

function createWorkbookBuffer(): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["key", "ar", "en"],
      ["home.title", "عنوان", "Title"],
      ["home.subtitle", "", "Subtitle"],
      ["", "空key", "Empty key"],
      ["module-only", "", ""]
    ]),
    "x"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["key", "ar", "en"],
      ["home.title", "عنوان2", "Title 2"]
    ]),
    "y"
  );
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}
