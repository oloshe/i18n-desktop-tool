import * as XLSX from "xlsx";
import type { ExcelPreview, LanguageColumns, SheetColumnOverrides } from "./types";

const PREVIEW_LIMIT = 20;

export interface ExcelReadOptions {
  sheetName?: string;
  sheetNames?: string[];
  skipRows?: number;
  headerRow?: number;
  keyColumn?: string;
  languageColumns?: LanguageColumns;
  sheetColumnOverrides?: SheetColumnOverrides;
}

export interface ExcelWorkbookInfo {
  sheetNames: string[];
  activeSheetName: string;
}

function normalizeCell(value: unknown): string {
  if (value == null) return "";
  return String(value).trim();
}

export function readExcelWorkbookInfo(buffer: ArrayBuffer, sheetName?: string): ExcelWorkbookInfo {
  const workbook = XLSX.read(buffer, { type: "array", bookSheets: true });
  const sheetNames = workbook.SheetNames;

  if (sheetNames.length === 0) {
    throw new Error("Excel 文件没有可读取的 sheet。");
  }

  return {
    sheetNames,
    activeSheetName: sheetName && sheetNames.includes(sheetName) ? sheetName : sheetNames[0]
  };
}

function readSheetMatrix(buffer: ArrayBuffer, sheetName?: string, sheetRows?: number) {
  const workbook = XLSX.read(buffer, { type: "array", sheetRows });
  const sheetNames = workbook.SheetNames;

  if (sheetNames.length === 0) {
    throw new Error("Excel 文件没有可读取的 sheet。");
  }

  const activeSheetName = sheetName && workbook.Sheets[sheetName] ? sheetName : sheetNames[0];
  const sheet = workbook.Sheets[activeSheetName];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: "",
    raw: false
  });

  if (matrix.length === 0) {
    throw new Error(`Sheet "${activeSheetName}" 是空的。`);
  }

  return { sheetNames, activeSheetName, matrix };
}

function selectedSheetNames(allSheetNames: string[], options: ExcelReadOptions): string[] {
  const names = options.sheetNames?.filter((name) => allSheetNames.includes(name)) ?? [];
  if (names.length > 0) return Array.from(new Set(names));
  if (options.sheetName && allSheetNames.includes(options.sheetName)) return [options.sheetName];
  return [allSheetNames[0]];
}

function getHeaders(matrix: unknown[][], headerRow: number): string[] {
  const headerIndex = headerRow - 1;
  if (headerIndex < 0 || headerIndex >= matrix.length) {
    throw new Error("表头行超出 Excel 内容范围。");
  }

  const headers = matrix[headerIndex].map(normalizeCell).filter(Boolean);
  if (headers.length === 0) {
    throw new Error("表头行没有有效表头。");
  }

  const headerSet = new Set(headers);
  if (headers.length !== headerSet.size) {
    throw new Error("Excel 表头存在重复列名，请先调整后再导入。");
  }

  return headers;
}

function rowToRecord(row: unknown[], headers: string[]): Record<string, string> {
  const record: Record<string, string> = {};
  headers.forEach((header, index) => {
    record[header] = normalizeCell(row[index]);
  });
  return record;
}

function isImportableRow(row: unknown[], headers: string[]): boolean {
  const values = headers.map((_, index) => normalizeCell(row[index]));
  const filledCount = values.filter(Boolean).length;
  return filledCount > 1;
}

function hasFilledCell(row: unknown[], headers: string[]): boolean {
  return headers.some((_, index) => normalizeCell(row[index]));
}

export function parseExcelBuffer(buffer: ArrayBuffer, options: ExcelReadOptions = {}): ExcelPreview {
  const skipRows = Math.max(0, options.skipRows ?? 0);
  const headerRow = Math.max(1, options.headerRow ?? 1);
  const previewRows = Math.max(skipRows, headerRow) + PREVIEW_LIMIT * 4;
  const { sheetNames, activeSheetName, matrix } = readSheetMatrix(buffer, options.sheetName, previewRows);
  const headers = getHeaders(matrix, headerRow);
  const dataStartIndex = Math.max(skipRows, headerRow);

  const rows = matrix
    .slice(dataStartIndex)
    .filter((row) => isImportableRow(row, headers))
    .slice(0, PREVIEW_LIMIT)
    .map((row) => rowToRecord(row, headers));

  return { sheetNames, activeSheetName, headerRow, skipRows, headers, rows };
}

export function rowsFromExcelBuffer(buffer: ArrayBuffer, options: ExcelReadOptions = {}): Array<Record<string, string>> {
  const skipRows = Math.max(0, options.skipRows ?? 0);
  const headerRow = Math.max(1, options.headerRow ?? 1);
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetNames = workbook.SheetNames;

  if (sheetNames.length === 0) {
    throw new Error("Excel 鏂囦欢娌℃湁鍙鍙栫殑 sheet銆?");
  }

  return selectedSheetNames(sheetNames, options).flatMap((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: false
    });
    const headers = getHeaders(matrix, headerRow);
    const dataStartIndex = Math.max(skipRows, headerRow);
    const columnConfig = resolveSheetColumnConfig(
      sheetName,
      options.keyColumn,
      options.languageColumns,
      options.sheetColumnOverrides
    );

    return matrix
      .slice(dataStartIndex)
      .filter((row) => hasFilledCell(row, headers))
      .map((row) => mapRowRecord(row, headers, columnConfig, sheetName));
  });
}

function resolveSheetColumnConfig(
  sheetName: string,
  keyColumn?: string,
  languageColumns?: LanguageColumns,
  sheetColumnOverrides?: SheetColumnOverrides
) {
  if (!keyColumn || !languageColumns) return null;
  const override = sheetColumnOverrides?.[sheetName];
  return {
    keyColumn,
    actualKeyColumn: override?.keyColumn?.trim() || keyColumn,
    languageColumns,
    actualLanguageColumns: Object.fromEntries(
      Object.entries(languageColumns).map(([lang, column]) => [lang, override?.languageColumns?.[lang]?.trim() || column])
    )
  };
}

function mapRowRecord(
  row: unknown[],
  headers: string[],
  columnConfig: ReturnType<typeof resolveSheetColumnConfig>,
  sheetName: string
): Record<string, string> {
  const record = rowToRecord(row, headers);
  if (!columnConfig) return record;

  const { keyColumn, actualKeyColumn, languageColumns, actualLanguageColumns } = columnConfig;
  if (!(actualKeyColumn in record)) {
    throw new Error(`Sheet "${sheetName}" 缺少 key 列 "${actualKeyColumn}"。`);
  }

  const mapped: Record<string, string> = {
    [keyColumn]: record[actualKeyColumn]
  };

  Object.entries(languageColumns).forEach(([lang, canonicalColumn]) => {
    const actualColumn = actualLanguageColumns[lang];
    if (!(actualColumn in record)) {
      throw new Error(`Sheet "${sheetName}" 缺少语言列 "${actualColumn}"。`);
    }
    mapped[canonicalColumn] = record[actualColumn];
  });

  return mapped;
}
