export type OutputFormat = "json" | "js" | "ts" | "xcstrings";
export type MergeStrategy = "overwrite" | "skip";
export type MissingKeyStrategy = "keep" | "remove";
export type KeyStyle = "nested" | "flat";
export type ModuleSplitMode = "none" | "keyPrefix" | "sectionRow";
export type ModuleNameSource = "keyPrefix" | "sectionRow";

export type LanguageColumns = Record<string, string>;
export type LocaleValue = string | LocaleObject;
export interface LocaleObject {
  [key: string]: LocaleValue;
}

export interface ExcelPreview {
  sheetNames: string[];
  activeSheetName: string;
  headerRow: number;
  skipRows: number;
  headers: string[];
  rows: Array<Record<string, string>>;
}

export interface ImportSettings {
  excelUrl?: string;
  projectRoot: string;
  keyColumn: string;
  languageColumns: LanguageColumns;
  outputPathTemplate: string;
  outputFormat: OutputFormat;
  mergeStrategy: MergeStrategy;
  missingKeyStrategy: MissingKeyStrategy;
  sheetName: string;
  sheetNames?: string[];
  skipRows: number;
  headerRow: number;
  moduleSplitMode?: ModuleSplitMode;
  moduleNameSource?: ModuleNameSource;
  splitByModule: boolean;
  keyStyle: KeyStyle;
  moduleFilter: string;
  ignoredModuleFilter: string;
  moduleNameReplacements: string;
  removeModulePrefix: boolean;
  quoteObjectProperties: boolean;
  spaceWrappedLanguages: string;
  ensureTrailingNewline: boolean;
}

export interface SavedProjectConfig extends ImportSettings {
  id: string;
  projectName: string;
  updatedAt: string;
}

export interface LocaleFilePlan {
  lang: string;
  moduleName: string;
  path: string;
  fileAction: "create" | "update";
  existingKeys: number;
  incomingKeys: number;
  addedKeys: string[];
  overwrittenKeys: string[];
  skippedKeys: string[];
  deletedKeys: string[];
  mergedLocale: LocaleObject;
  existingContent: string;
  nextContent: string;
  eol: "lf" | "crlf";
  error?: string;
}
