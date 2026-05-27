import type { KeyStyle, LanguageColumns, LocaleObject, MergeStrategy, MissingKeyStrategy, ModuleNameSource, ModuleSplitMode } from "./types";

export function resolveLocalePath(template: string, lang: string, moduleName?: string): string {
  if (!template.trim()) {
    throw new Error("输出路径模板不能为空。");
  }
  if (!template.includes("{lang}")) {
    throw new Error('输出路径模板必须包含 "{lang}"。');
  }

  const normalized = template.split("\\").join("/");
  if (moduleName && normalized.includes("{module}")) {
    return normalized.split("{lang}").join(lang).split("{module}").join(moduleName);
  }

  const resolved = normalized.split("{lang}").join(lang);
  return resolved;
}

export function generateLocales(
  rows: Array<Record<string, string>>,
  keyColumn: string,
  languageColumns: LanguageColumns,
  options: {
    splitByModule?: boolean;
    moduleSplitMode?: ModuleSplitMode;
    moduleNameSource?: ModuleNameSource;
    keyStyle: KeyStyle;
    moduleFilter?: string[];
    ignoredModuleFilter?: string[];
    moduleNameReplacements?: Record<string, string>;
    spaceWrappedLanguages?: string[];
  }
): Record<string, Record<string, LocaleObject>> {
  if (!keyColumn) throw new Error("请配置 key 列。");
  const languages = Object.entries(languageColumns).filter(([lang, column]) => lang.trim() && column.trim());
  if (languages.length === 0) throw new Error("请至少配置一个语言列映射。");
  const moduleSplitMode = options.moduleSplitMode ?? (options.splitByModule ? "keyPrefix" : "none");
  const moduleNameSource = options.moduleNameSource ?? (moduleSplitMode === "sectionRow" ? "sectionRow" : "keyPrefix");

  const locales: Record<string, Record<string, LocaleObject>> = {};
  languages.forEach(([lang]) => {
    locales[lang] = {};
  });

  let sectionModuleName = "";

  rows.forEach((row, index) => {
    if (!(keyColumn in row)) {
      throw new Error(`Excel 缺少 key 列 "${keyColumn}"。`);
    }
    const rawKey = row[keyColumn]?.trim();
    if (!rawKey) return;

    if (isSectionModuleRow(row, keyColumn, languages)) {
      if (moduleSplitMode !== "sectionRow") return;
      sectionModuleName = rawKey;
      return;
    }

    const rawModuleName = getOutputModuleName(rawKey, moduleSplitMode, moduleNameSource, sectionModuleName);
    const filterModuleName = normalizeModuleName(rawModuleName, options.moduleNameReplacements);
    if (options.moduleFilter?.length && !matchesModuleFilter(options.moduleFilter, rawModuleName, filterModuleName)) return;
    if (matchesModuleFilter(options.ignoredModuleFilter, rawModuleName, filterModuleName)) return;

    languages.forEach(([lang, column]) => {
      if (!(column in row)) {
        throw new Error(`Excel 缺少语言列 "${column}"。`);
      }

      const rawValue = row[column]?.trim() ?? "";
      if (!rawValue) return;

      const { moduleName: rawOutputModuleName, localeKey } = splitModule(rawKey, moduleSplitMode, moduleNameSource, sectionModuleName);
      const moduleName = rawOutputModuleName ? normalizeModuleName(rawOutputModuleName, options.moduleNameReplacements) : rawOutputModuleName;
      const bucket = locales[lang][moduleName] ?? {};
      const value = transformLocaleValue(rawValue, lang, options);
      const inserted = setLocaleValue(bucket, localeKey, value, options.keyStyle);
      if (!inserted) {
        throw new Error(`Excel 第 ${index + 2} 行出现重复 key "${rawKey}"。`);
      }
      locales[lang][moduleName] = bucket;
    });
  });

  return locales;
}

export function mergeLocaleObjects(
  existingLocale: LocaleObject,
  incomingLocale: LocaleObject,
  strategy: MergeStrategy,
  missingKeyStrategy: MissingKeyStrategy = "keep"
) {
  const mergedLocale = missingKeyStrategy === "remove" ? cloneLocale(incomingLocale) : cloneLocale(existingLocale);
  const addedKeys: string[] = [];
  const overwrittenKeys: string[] = [];
  const skippedKeys: string[] = [];

  if (missingKeyStrategy === "remove") {
    mergeExistingValuesInto(mergedLocale, existingLocale, incomingLocale, strategy, [], overwrittenKeys, skippedKeys);
  } else {
    mergeInto(mergedLocale, incomingLocale, strategy, [], addedKeys, overwrittenKeys, skippedKeys);
  }

  return { mergedLocale, addedKeys, overwrittenKeys, skippedKeys };
}

function mergeExistingValuesInto(
  target: LocaleObject,
  existing: LocaleObject,
  incoming: LocaleObject,
  strategy: MergeStrategy,
  path: string[],
  overwrittenKeys: string[],
  skippedKeys: string[]
) {
  Object.entries(incoming).forEach(([key, incomingValue]) => {
    if (!Object.prototype.hasOwnProperty.call(existing, key)) return;
    const existingValue = existing[key];
    const nextPath = [...path, key];

    if (isLocaleObject(existingValue) && isLocaleObject(incomingValue)) {
      mergeExistingValuesInto(target[key] as LocaleObject, existingValue, incomingValue, strategy, nextPath, overwrittenKeys, skippedKeys);
      return;
    }

    const displayKey = nextPath.join(".");
    if (strategy === "skip") {
      target[key] = cloneValue(existingValue);
      skippedKeys.push(displayKey);
    } else {
      overwrittenKeys.push(displayKey);
    }
  });
}

function splitModule(rawKey: string, moduleSplitMode: ModuleSplitMode, moduleNameSource: ModuleNameSource, sectionModuleName: string) {
  if (moduleSplitMode === "none") return { moduleName: "", localeKey: rawKey };
  if (moduleSplitMode === "sectionRow" && moduleNameSource === "sectionRow") return { moduleName: sectionModuleName, localeKey: rawKey };
  const parts = rawKey.split(".").filter(Boolean);
  if (parts.length <= 1) return { moduleName: sectionModuleName || rawKey, localeKey: rawKey };
  return { moduleName: parts[0], localeKey: parts.slice(1).join(".") };
}

function getModuleName(rawKey: string): string {
  return rawKey.split(".").filter(Boolean)[0] ?? "";
}

function getOutputModuleName(
  rawKey: string,
  moduleSplitMode: ModuleSplitMode,
  moduleNameSource: ModuleNameSource,
  sectionModuleName: string
): string {
  if (moduleSplitMode === "none") return getModuleName(rawKey);
  if (moduleNameSource === "sectionRow") return sectionModuleName || getModuleName(rawKey);
  return getModuleName(rawKey) || sectionModuleName;
}

function normalizeModuleName(moduleName: string, replacements?: Record<string, string>): string {
  const replacement = replacements?.[moduleName]?.trim();
  return replacement || moduleName;
}

function matchesModuleFilter(filter: string[] | undefined, rawModuleName: string, normalizedModuleName: string): boolean {
  return Boolean(filter?.includes(rawModuleName) || filter?.includes(normalizedModuleName));
}

function transformLocaleValue(
  value: string,
  lang: string,
  options: {
    spaceWrappedLanguages?: string[];
  }
): string {
  if (!value || !options.spaceWrappedLanguages?.includes(lang)) return value;
  return ` ${value.trim()} `;
}

function isSectionModuleRow(row: Record<string, string>, keyColumn: string, languages: Array<[string, string]>): boolean {
  const languageColumns = new Set(languages.map(([, column]) => column));
  const filledEntries = Object.entries(row).filter(([, value]) => value.trim());
  return filledEntries.length === 1 && filledEntries[0][0] === keyColumn && !languageColumns.has(keyColumn);
}

function setLocaleValue(target: LocaleObject, rawKey: string, value: string, keyStyle: KeyStyle): boolean {
  if (keyStyle === "flat") {
    if (Object.prototype.hasOwnProperty.call(target, rawKey)) return false;
    target[rawKey] = value;
    return true;
  }

  const parts = rawKey.split(".").filter(Boolean);
  if (parts.length === 0) return true;
  let cursor = target;
  for (let index = 0; index < parts.length; index += 1) {
    const part = parts[index];
    const isLeaf = index === parts.length - 1;
    if (isLeaf) {
      if (Object.prototype.hasOwnProperty.call(cursor, part)) return false;
      cursor[part] = value;
      return true;
    }

    const next = cursor[part];
    if (next == null) {
      cursor[part] = {};
    } else if (typeof next !== "object") {
      return false;
    }
    cursor = cursor[part] as LocaleObject;
  }
  return true;
}

function mergeInto(
  target: LocaleObject,
  incoming: LocaleObject,
  strategy: MergeStrategy,
  path: string[],
  addedKeys: string[],
  overwrittenKeys: string[],
  skippedKeys: string[]
) {
  Object.entries(incoming).forEach(([key, value]) => {
    const nextPath = [...path, key];
    const existingValue = target[key];
    const hasExisting = Object.prototype.hasOwnProperty.call(target, key);

    if (isLocaleObject(value) && isLocaleObject(existingValue)) {
      mergeInto(existingValue, value, strategy, nextPath, addedKeys, overwrittenKeys, skippedKeys);
      return;
    }

    const displayKey = nextPath.join(".");
    if (hasExisting) {
      if (strategy === "overwrite") {
        target[key] = cloneValue(value);
        overwrittenKeys.push(displayKey);
      } else {
        skippedKeys.push(displayKey);
      }
      return;
    }

    target[key] = cloneValue(value);
    addedKeys.push(displayKey);
  });
}

function cloneLocale(locale: LocaleObject): LocaleObject {
  return JSON.parse(JSON.stringify(locale)) as LocaleObject;
}

function cloneValue(value: unknown) {
  return JSON.parse(JSON.stringify(value));
}

function isLocaleObject(value: unknown): value is LocaleObject {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}
