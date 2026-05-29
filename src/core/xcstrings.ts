import type { LanguageColumns, LocaleObject } from "./types";

interface XcStringUnit {
  state?: string;
  value?: string;
}

interface XcLocalization {
  stringUnit?: XcStringUnit;
}

interface XcStringEntry {
  localizations?: Record<string, XcLocalization>;
}

interface XcStringsCatalog {
  sourceLanguage?: string;
  strings?: Record<string, XcStringEntry>;
  version?: string;
}

export function generateXcstringsLocale(
  rows: Array<Record<string, string>>,
  keyColumn: string,
  languageColumns: LanguageColumns
): LocaleObject {
  if (!keyColumn) throw new Error("请配置 key 列。");
  const languages = Object.entries(languageColumns).filter(([lang, column]) => lang.trim() && column.trim());
  if (languages.length === 0) throw new Error("请至少配置一个语言列映射。");

  const locale: LocaleObject = {};
  rows.forEach((row, index) => {
    if (!(keyColumn in row)) {
      throw new Error(`Excel 缺少 key 列 "${keyColumn}"。`);
    }

    const key = row[keyColumn]?.trim();
    if (!key) return;

    const localizations = (locale[key] as LocaleObject | undefined) ?? {};
    languages.forEach(([lang, column]) => {
      if (!(column in row)) {
        throw new Error(`Excel 缺少语言列 "${column}"。`);
      }

      const value = row[column]?.trim() ?? "";
      if (!value) return;
      if (Object.prototype.hasOwnProperty.call(localizations, lang)) {
        throw new Error(`Excel 第 ${index + 2} 行出现重复 key/lang "${key}" / "${lang}"。`);
      }
      localizations[lang] = value;
    });

    if (Object.keys(localizations).length > 0) {
      locale[key] = localizations;
    }
  });

  return locale;
}

export function parseXcstringsContent(content: string): LocaleObject {
  const trimmed = content.trim();
  if (!trimmed) return {};

  const catalog = JSON.parse(trimmed) as XcStringsCatalog;
  const locale: LocaleObject = {};
  Object.entries(catalog.strings ?? {}).forEach(([key, entry]) => {
    const localizations: LocaleObject = {};
    Object.entries(entry.localizations ?? {}).forEach(([lang, localization]) => {
      const value = localization.stringUnit?.value;
      if (value != null) localizations[lang] = String(value);
    });
    if (Object.keys(localizations).length > 0) locale[key] = localizations;
  });
  return locale;
}

export function formatXcstringsLocale(locale: LocaleObject, sourceLanguage: string): string {
  const strings: XcStringsCatalog["strings"] = {};

  Object.entries(locale).forEach(([key, value]) => {
    if (!isLocaleObject(value)) return;
    const localizations: Record<string, XcLocalization> = {};
    Object.entries(value).forEach(([lang, text]) => {
      if (isLocaleObject(text)) return;
      localizations[lang] = {
        stringUnit: {
          state: "translated",
          value: String(text)
        }
      };
    });
    strings[key] = { localizations };
  });

  return JSON.stringify(
    {
      sourceLanguage,
      strings,
      version: "1.0"
    },
    null,
    2
  );
}

function isLocaleObject(value: unknown): value is LocaleObject {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}
