import * as XLSX from "xlsx";
import type { LanguageColumns, SheetColumnOverrides } from "./types";

export interface ExcelHealthCheckOptions {
  sheetName?: string;
  sheetNames?: string[];
  skipRows?: number;
  headerRow?: number;
  sheetColumnOverrides?: SheetColumnOverrides;
}

export interface ExcelHealthCheckRowRef {
  sheetName: string;
  rowNumber: number;
}

export interface ExcelDuplicateKeyIssue {
  key: string;
  locations: ExcelHealthCheckRowRef[];
}

export interface ExcelMissingTranslationIssue extends ExcelHealthCheckRowRef {
  key: string;
  lang: string;
}

export interface ExcelHealthCheckResult {
  duplicateKeys: ExcelDuplicateKeyIssue[];
  emptyKeyRows: ExcelHealthCheckRowRef[];
  missingTranslations: ExcelMissingTranslationIssue[];
}

export function inspectExcelBuffer(
  buffer: ArrayBuffer,
  keyColumn: string,
  languageColumns: LanguageColumns,
  options: ExcelHealthCheckOptions = {}
): ExcelHealthCheckResult {
  if (!keyColumn) throw new Error("请配置 key 列。");
  const languages = Object.entries(languageColumns).filter(([lang, column]) => lang.trim() && column.trim());
  if (languages.length === 0) throw new Error("请至少配置一个语言列映射。");

  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetNames = selectedSheetNames(workbook.SheetNames, options);
  const skipRows = Math.max(0, options.skipRows ?? 0);
  const headerRow = Math.max(1, options.headerRow ?? 1);
  const dataStartIndex = Math.max(skipRows, headerRow);
  const duplicateMap = new Map<string, ExcelHealthCheckRowRef[]>();
  const emptyKeyRows: ExcelHealthCheckRowRef[] = [];
  const missingTranslations: ExcelMissingTranslationIssue[] = [];

  sheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: false
    });
    const headers = getHeaders(matrix, headerRow, sheetName);
    const actualKeyColumn = options.sheetColumnOverrides?.[sheetName]?.keyColumn?.trim() || keyColumn;
    const keyIndex = getColumnIndex(headers, actualKeyColumn, sheetName);
    const languageIndexes = languages.map(([lang, column]) => ({
      lang,
      index: getColumnIndex(
        headers,
        options.sheetColumnOverrides?.[sheetName]?.languageColumns?.[lang]?.trim() || column,
        sheetName
      )
    }));

    matrix.slice(dataStartIndex).forEach((row, offset) => {
      const rowNumber = dataStartIndex + offset + 1;
      const values = headers.map((_, index) => normalizeCell(row[index]));
      if (!values.some(Boolean)) return;

      const key = values[keyIndex]?.trim() ?? "";
      if (!key) {
        emptyKeyRows.push({ sheetName, rowNumber });
        return;
      }

      const filledLanguageCount = languageIndexes.filter(({ index }) => values[index]?.trim()).length;
      if (filledLanguageCount === 0) return;

      duplicateMap.set(key, [...(duplicateMap.get(key) ?? []), { sheetName, rowNumber }]);
      languageIndexes.forEach(({ lang, index }) => {
        if (!values[index]?.trim()) {
          missingTranslations.push({ sheetName, rowNumber, key, lang });
        }
      });
    });
  });

  return {
    duplicateKeys: Array.from(duplicateMap.entries())
      .filter(([, locations]) => locations.length > 1)
      .map(([key, locations]) => ({ key, locations })),
    emptyKeyRows,
    missingTranslations
  };
}

export function formatExcelHealthCheckReport(result: ExcelHealthCheckResult): string {
  const lines = ["体检结果：", "重复key列表："];
  if (result.duplicateKeys.length === 0) {
    lines.push("无");
  } else {
    result.duplicateKeys.forEach((issue) => {
      lines.push(`${issue.key}: ${formatLocations(issue.locations)}`);
    });
  }

  lines.push("", "空key行：");
  if (result.emptyKeyRows.length === 0) {
    lines.push("无");
  } else {
    result.emptyKeyRows.forEach((location) => {
      lines.push(`${location.sheetName} sheet ${location.rowNumber} 行`);
    });
  }

  lines.push("", "个别语言翻译缺失：");
  if (result.missingTranslations.length === 0) {
    lines.push("无");
  } else {
    result.missingTranslations.forEach((issue) => {
      lines.push(`sheet: ${issue.sheetName} ${issue.rowNumber} 行 key: ${issue.key} 缺少 ${issue.lang} 翻译`);
    });
  }

  return lines.join("\n");
}

function selectedSheetNames(allSheetNames: string[], options: ExcelHealthCheckOptions): string[] {
  if (allSheetNames.length === 0) throw new Error("Excel 文件没有可读取的 sheet。");
  const names = options.sheetNames?.filter((name) => allSheetNames.includes(name)) ?? [];
  if (names.length > 0) return Array.from(new Set(names));
  if (options.sheetName && allSheetNames.includes(options.sheetName)) return [options.sheetName];
  return [allSheetNames[0]];
}

function getHeaders(matrix: unknown[][], headerRow: number, sheetName: string): string[] {
  const headerIndex = headerRow - 1;
  const row = matrix[headerIndex];
  if (!row) throw new Error(`Sheet "${sheetName}" 表头行超出 Excel 内容范围。`);
  const headers = row.map(normalizeCell);
  if (!headers.some(Boolean)) throw new Error(`Sheet "${sheetName}" 表头行没有有效表头。`);
  return headers;
}

function getColumnIndex(headers: string[], column: string, sheetName: string): number {
  const index = headers.indexOf(column);
  if (index === -1) throw new Error(`Sheet "${sheetName}" 缺少列 "${column}"。`);
  return index;
}

function normalizeCell(value: unknown): string {
  if (value == null) return "";
  return String(value).trim();
}

function formatLocations(locations: ExcelHealthCheckRowRef[]): string {
  return locations.map((location) => `${location.sheetName} sheet ${location.rowNumber} 行`).join("，");
}
