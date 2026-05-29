import type { LocaleObject, OutputFormat } from "./types";
import { parseXcstringsContent } from "./xcstrings";

export interface LocaleObjectRange {
  start: number;
  end: number;
  exportKind: "default" | "named";
  exportName?: string;
}

export function parseLocaleContent(content: string, format: OutputFormat): LocaleObject {
  const trimmed = content.trim();
  if (!trimmed) return {};

  if (format === "json") {
    return normalizeLocaleObject(JSON.parse(trimmed));
  }
  if (format === "xcstrings") {
    return parseXcstringsContent(content);
  }

  const objectRange = findLocaleObjectRange(content);
  return normalizeLocaleObject(parseObjectLiteral(content.slice(objectRange.start, objectRange.end)));
}

export function findLocaleObjectRange(content: string): LocaleObjectRange {
  const defaultExportIndex = content.indexOf("export default");
  if (defaultExportIndex !== -1) {
    return {
      ...findObjectAfter(content, defaultExportIndex + "export default".length),
      exportKind: "default"
    };
  }

  const namedExportPattern = /export\s+const\s+([A-Za-z_$][\w$]*)\s*=/g;
  const namedExport = namedExportPattern.exec(content);
  if (namedExport) {
    return {
      ...findObjectAfter(content, namedExport.index + namedExport[0].length),
      exportKind: "named",
      exportName: namedExport[1]
    };
  }

  throw new Error("JS/TS locale 文件需要包含 export default object 或 export const name = object。");
}

function findObjectAfter(content: string, startIndex: number): { start: number; end: number } {
  const objectStart = content.indexOf("{", startIndex);
  if (objectStart === -1) {
    throw new Error("未找到导出变量后的对象。");
  }

  let depth = 0;
  let inString: string | null = null;
  let escaped = false;
  for (let index = objectStart; index < content.length; index += 1) {
    const char = content[index];
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (char === "\\") {
        escaped = true;
      } else if (char === inString) {
        inString = null;
      }
      continue;
    }

    if (char === '"' || char === "'" || char === "`") {
      inString = char;
      continue;
    }
    if (char === "{") depth += 1;
    if (char === "}") depth -= 1;
    if (depth === 0) {
      return { start: objectStart, end: index + 1 };
    }
  }

  throw new Error("对象字面量没有正确闭合。");
}

function parseObjectLiteral(source: string): unknown {
  try {
    return Function(`"use strict"; return (${source});`)();
  } catch {
    throw new Error("JS/TS locale 文件解析失败，请确认导出的变量是对象字面量。");
  }
}

function normalizeLocaleObject(value: unknown): LocaleObject {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("locale 内容必须是对象。");
  }

  const locale: LocaleObject = {};
  Object.entries(value as Record<string, unknown>).forEach(([key, entry]) => {
    if (entry && typeof entry === "object" && !Array.isArray(entry)) {
      locale[key] = normalizeLocaleObject(entry);
    } else {
      locale[key] = entry == null ? "" : String(entry);
    }
  });
  return locale;
}
