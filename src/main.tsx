import React, { useEffect, useMemo, useRef, useState } from "react";
import ReactDOM from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Alert,
  Box,
  Button,
  ButtonBase,
  Checkbox,
  CircularProgress,
  CssBaseline,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  FormControlLabel,
  MenuItem,
  Paper,
  Stack,
  TextField,
  ThemeProvider,
  Typography,
  createTheme
} from "@mui/material";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readDir, readTextFile, writeTextFile } from "@tauri-apps/plugin-fs";
import "./styles.css";
import {
  deleteProjectConfig,
  loadProjectConfigs,
  parseProjectConfigExport,
  saveProjectConfig,
  serializeProjectConfig
} from "./core/configStore";
import { generateLocales, mergeLocaleObjects, resolveLocalePath } from "./core/localeGenerator";
import {
  detectEol,
  formatLocaleForSnapshot,
  readLocaleFileSnapshot,
  resolveProjectFile,
  writeLocaleContent,
  type LocaleFileSnapshot
} from "./core/localeFileWriter";
import { findLocaleObjectRange, parseLocaleContent } from "./core/localeFileReader";
import { generateXcstringsLocale } from "./core/xcstrings";
import type {
  ExcelPreview,
  ImportSettings,
  KeyStyle,
  LanguageColumns,
  LocaleFilePlan,
  LocaleObject,
  MergeStrategy,
  MissingKeyStrategy,
  ModuleNameSource,
  ModuleSplitMode,
  OutputFormat,
  SavedProjectConfig,
  SheetColumnOverrides
} from "./core/types";
import {
  parseExcelBuffer,
  readExcelWorkbookInfo as readExcelWorkbookInfoFromBuffer,
  rowsFromExcelBuffer,
  type ExcelReadOptions,
  type ExcelWorkbookInfo
} from "./core/excelParser";
import {
  formatExcelHealthCheckReport,
  inspectExcelBuffer,
  type ExcelHealthCheckResult
} from "./core/excelHealthCheck";

const DEFAULT_LANGUAGE_COLUMNS: LanguageColumns = {
  "zh-CN": "中文",
  "en-US": "英文"
};

const theme = createTheme({
  palette: {
    mode: "dark",
    primary: { main: "#2dd4bf" },
    secondary: { main: "#38bdf8" },
    background: { default: "#0b1220", paper: "#0f172a" }
  },
  shape: {
    borderRadius: 12
  }
});

type StatusKind = "idle" | "reading" | "previewing" | "checking" | "importing" | "exporting";
type DiffRow =
  | { kind: "hunk"; header: string }
  | {
      kind: "row";
      rowKind: "equal" | "replace" | "delete" | "insert";
      oldLineNumber: number | null;
      newLineNumber: number | null;
      oldText: string;
      newText: string;
    };
type ExcelSource = { label: string; path?: string; bytes?: number[] };
type WebDirectoryHandle = FileSystemDirectoryHandle;
interface ImportSummary {
  changedFiles: number;
  lines: Array<{ lang: string; moduleName: string; added: number; modified: number; deleted: number }>;
}
interface LocaleExportModule {
  displayName: string;
  pathName: string;
}
interface LocaleExportPreview {
  headers: string[];
  rows: Array<Record<string, string>>;
  modules: string[];
  fileCount: number;
  missingFiles: string[];
}

function App() {
  const [excelPath, setExcelPath] = useState("");
  const [excelSource, setExcelSource] = useState<ExcelSource | null>(null);
  const [preview, setPreview] = useState<ExcelPreview | null>(null);
  const [webProjectDirectory, setWebProjectDirectory] = useState<WebDirectoryHandle | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [selectedSheetNames, setSelectedSheetNames] = useState<string[]>([]);
  const [skipRows, setSkipRows] = useState(0);
  const [headerRow, setHeaderRow] = useState(1);
  const [projectName, setProjectName] = useState("默认项目");
  const [projectRoot, setProjectRoot] = useState("");
  const [keyColumn, setKeyColumn] = useState("key");
  const [languageColumns, setLanguageColumns] = useState<LanguageColumns>(DEFAULT_LANGUAGE_COLUMNS);
  const [sheetColumnOverrides, setSheetColumnOverrides] = useState<SheetColumnOverrides>({});
  const [outputPathTemplate, setOutputPathTemplate] = useState("locales/{lang}/{module}.json");
  const [outputFormat, setOutputFormat] = useState<OutputFormat>("json");
  const [mergeStrategy, setMergeStrategy] = useState<MergeStrategy>("overwrite");
  const [missingKeyStrategy, setMissingKeyStrategy] = useState<MissingKeyStrategy>("keep");
  const [moduleSplitMode, setModuleSplitMode] = useState<ModuleSplitMode>("none");
  const [moduleNameSource, setModuleNameSource] = useState<ModuleNameSource>("keyPrefix");
  const [keyStyle, setKeyStyle] = useState<KeyStyle>("nested");
  const [moduleFilter, setModuleFilter] = useState("");
  const [exportModuleFilter, setExportModuleFilter] = useState("");
  const [ignoredModuleFilter, setIgnoredModuleFilter] = useState("");
  const [moduleNameReplacements, setModuleNameReplacements] = useState("");
  const [removeModulePrefix, setRemoveModulePrefix] = useState(false);
  const [quoteObjectProperties, setQuoteObjectProperties] = useState(false);
  const [spaceWrappedLanguages, setSpaceWrappedLanguages] = useState("");
  const [ensureTrailingNewline, setEnsureTrailingNewline] = useState(true);
  const [configs, setConfigs] = useState<SavedProjectConfig[]>([]);
  const [selectedConfigId, setSelectedConfigId] = useState("");
  const [plans, setPlans] = useState<LocaleFilePlan[]>([]);
  const [selectedPlanPath, setSelectedPlanPath] = useState("");
  const [message, setMessage] = useState("");
  const [status, setStatus] = useState<StatusKind>("idle");
  const [summary, setSummary] = useState<ImportSummary | null>(null);
  const [healthCheckResult, setHealthCheckResult] = useState<ExcelHealthCheckResult | null>(null);
  const [localeExportPreview, setLocaleExportPreview] = useState<LocaleExportPreview | null>(null);
  const previewRequestRef = useRef(0);
  const diffScrollerRef = useRef<HTMLDivElement | null>(null);
  const splitByModule = moduleSplitMode === "keyPrefix";
  const isXcstringsOutput = outputFormat === "xcstrings";

  useEffect(() => {
    setConfigs(loadProjectConfigs());
  }, []);

  useEffect(() => {
    const preventBrowserDrop = (event: DragEvent) => {
      event.preventDefault();
    };

    window.addEventListener("dragover", preventBrowserDrop);
    window.addEventListener("drop", preventBrowserDrop);

    let unlisten: (() => void) | undefined;
    if (isTauriRuntime()) {
      getCurrentWebview()
        .onDragDropEvent((event) => {
          if (event.payload.type !== "drop") return;
          const path = event.payload.paths.find((item) => /\.(xlsx|xls)$/i.test(item));
          if (!path) {
            setMessage("请拖入 .xlsx 或 .xls 文件。");
            return;
          }
          void loadExcelFilePath(path, "拖拽读取 Excel 失败");
        })
        .then((handler) => {
          unlisten = handler;
        })
        .catch((error) => {
          showError("注册窗口拖拽失败", error);
        });
    }

    return () => {
      window.removeEventListener("dragover", preventBrowserDrop);
      window.removeEventListener("drop", preventBrowserDrop);
      unlisten?.();
    };
  }, []);

  useEffect(() => {
    if (!excelSource) return;

    const requestId = previewRequestRef.current + 1;
    previewRequestRef.current = requestId;

    void (async () => {
      setStatus("reading");
      await yieldToUi();
      try {
        const nextPreview = await readExcelPreview(excelSource, { sheetName, skipRows, headerRow });
        if (previewRequestRef.current !== requestId) return;
        setPreview(nextPreview);
        setSheetName(nextPreview.activeSheetName);
        setSelectedSheetNames((current) => {
          const kept = current.filter((name) => nextPreview.sheetNames.includes(name));
          return kept.length > 0 ? kept : [nextPreview.activeSheetName];
        });
        setSheetColumnOverrides((current) =>
          Object.fromEntries(Object.entries(current).filter(([name]) => nextPreview.sheetNames.includes(name)))
        );
        if (nextPreview.headers.includes("key")) setKeyColumn("key");
        setPlans([]);
      } catch (error) {
        if (previewRequestRef.current === requestId) showError("刷新 Excel 预览失败", error);
      } finally {
        if (previewRequestRef.current === requestId) setStatus("idle");
      }
    })();
  }, [excelSource, sheetName, skipRows, headerRow]);

  const settings: ImportSettings = useMemo(
    () => ({
      excelUrl: isHttpUrl(excelPath) ? excelPath.trim() : "",
      projectRoot,
      keyColumn,
      languageColumns,
      sheetColumnOverrides,
      outputPathTemplate,
      outputFormat,
      mergeStrategy,
      missingKeyStrategy,
      sheetName,
      sheetNames: selectedSheetNames,
      skipRows,
      headerRow,
      moduleSplitMode,
      moduleNameSource,
      splitByModule,
      keyStyle,
      moduleFilter,
      ignoredModuleFilter,
      moduleNameReplacements,
      removeModulePrefix,
      quoteObjectProperties,
      spaceWrappedLanguages,
      ensureTrailingNewline
    }),
    [
      projectRoot,
      excelPath,
      keyColumn,
      languageColumns,
      sheetColumnOverrides,
      outputPathTemplate,
      outputFormat,
      mergeStrategy,
      missingKeyStrategy,
      sheetName,
      selectedSheetNames,
      skipRows,
      headerRow,
      moduleSplitMode,
      moduleNameSource,
      splitByModule,
      keyStyle,
      moduleFilter,
      ignoredModuleFilter,
      moduleNameReplacements,
      removeModulePrefix,
      quoteObjectProperties,
      spaceWrappedLanguages,
      ensureTrailingNewline
    ]
  );

  const diffPlans = useMemo(
    () =>
      plans.map((plan) => ({
        plan,
        diff: plan.error ? [] : createGitStyleDiff(plan.existingContent, plan.nextContent)
      })),
    [plans]
  );
  const selectedPlan = plans.find((plan) => plan.path === selectedPlanPath) ?? plans[0];
  const isWorking = status !== "idle";
  const canChooseProjectRoot = isTauriRuntime() || (window.isSecureContext && Boolean(window.showDirectoryPicker));

  async function chooseExcelFile() {
    setMessage("");
    if (!isTauriRuntime()) {
      const file = await chooseBrowserFile(".xlsx,.xls");
      if (!file) return;
      setStatus("reading");
      await yieldToUi();
      try {
        await loadExcelSource({ label: file.name, bytes: arrayBufferToBytes(await file.arrayBuffer()) });
      } catch (error) {
        showError("读取 Excel 失败", error);
        setStatus("idle");
      }
      return;
    }

    const selected = await open({
      multiple: false,
      filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
    });
    if (!selected || Array.isArray(selected)) return;

    await loadExcelFilePath(selected, "读取 Excel 失败");
  }

  async function loadExcelUrl() {
    const url = excelPath.trim();
    if (!url) {
      setMessage("请先输入 Google Sheet 链接。");
      return;
    }

    setStatus("reading");
    await yieldToUi();
    try {
      await loadExcelSource(await loadGoogleSheetSource(url));
    } catch (error) {
      showError("读取 Google Sheet 失败", error);
      setStatus("idle");
    }
  }

  async function loadExcelFilePath(path: string, errorPrefix: string) {
    setStatus("reading");
    await yieldToUi();
    try {
      await loadExcelSource({ label: path, path });
    } catch (error) {
      showError(errorPrefix, error);
      setStatus("idle");
    }
  }

  async function loadExcelSource(source: ExcelSource) {
    await yieldToUi();
    const info = await readExcelWorkbookInfo(source, sheetName);
    setExcelPath(source.label);
    setExcelSource(source);
    setPreview({
      sheetNames: info.sheetNames,
      activeSheetName: info.activeSheetName,
      headerRow,
      skipRows,
      headers: [],
      rows: []
    });
    setSheetName(info.activeSheetName);
    setSelectedSheetNames([info.activeSheetName]);
    setMessage("Excel 已读取。");
  }

  async function chooseProjectRoot() {
    if (!isTauriRuntime()) {
      try {
        const showDirectoryPicker = window.showDirectoryPicker;
        if (!window.isSecureContext || !showDirectoryPicker) {
          setWebProjectDirectory(null);
          setProjectRoot("");
          setMessage("当前浏览器环境不支持目录授权，可直接生成预览；执行导入时会下载文件。");
          return;
        }
        const handle = await showDirectoryPicker({ mode: "readwrite" });
        await ensureDirectoryPermission(handle);
        setWebProjectDirectory(handle);
        setProjectRoot(`[本机目录] ${handle.name}`);
        setPlans([]);
        setSelectedPlanPath("");
      } catch (error) {
        if (isAbortError(error)) return;
        showError("选择项目目录失败", error);
      }
      return;
    }

    const selected = await open({ directory: true, multiple: false });
    if (!selected || Array.isArray(selected)) return;
    setWebProjectDirectory(null);
    setProjectRoot(selected);
  }

  async function handleDrop(event: React.DragEvent<HTMLElement>) {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setMessage("请拖入 .xlsx 或 .xls 文件。");
      return;
    }
    setStatus("reading");
    await yieldToUi();
    try {
      const path = (file as File & { path?: string }).path;
      if (path) {
        await loadExcelSource({ label: path, path });
      } else {
        await loadExcelSource({ label: file.name, bytes: arrayBufferToBytes(await file.arrayBuffer()) });
      }
    } catch (error) {
      showError("拖拽读取 Excel 失败", error);
      setStatus("idle");
    }
  }

  function updateLanguage(index: number, field: "lang" | "column", value: string) {
    const entries = Object.entries(languageColumns);
    entries[index] = field === "lang" ? [value, entries[index][1]] : [entries[index][0], value];
    setLanguageColumns(Object.fromEntries(entries));
  }

  function addLanguageRow() {
    let index = Object.keys(languageColumns).length + 1;
    let lang = `lang-${index}`;
    while (lang in languageColumns) {
      index += 1;
      lang = `lang-${index}`;
    }
    setLanguageColumns({ ...languageColumns, [lang]: "" });
  }

  function removeLanguageRow(index: number) {
    const entries = Object.entries(languageColumns).filter((_, entryIndex) => entryIndex !== index);
    setLanguageColumns(Object.fromEntries(entries));
  }

  function addSheetOverride() {
    const nextSheetName = selectedSheetNames.find((name) => !(name in sheetColumnOverrides))
      ?? preview?.sheetNames.find((name) => !(name in sheetColumnOverrides));
    if (!nextSheetName) return;
    setSheetColumnOverrides((current) => ({ ...current, [nextSheetName]: { languageColumns: {} } }));
  }

  function renameSheetOverride(sheetNameToReplace: string, nextSheetName: string) {
    if (!nextSheetName || nextSheetName === sheetNameToReplace) return;
    setSheetColumnOverrides((current) => {
      const next = { ...current };
      const override = next[sheetNameToReplace];
      delete next[sheetNameToReplace];
      next[nextSheetName] = override;
      return next;
    });
  }

  function updateSheetOverride(sheetNameToUpdate: string, field: "keyColumn", value: string) {
    setSheetColumnOverrides((current) => ({
      ...current,
      [sheetNameToUpdate]: {
        ...current[sheetNameToUpdate],
        [field]: value
      }
    }));
  }

  function updateSheetOverrideLanguage(sheetNameToUpdate: string, lang: string, value: string) {
    setSheetColumnOverrides((current) => ({
      ...current,
      [sheetNameToUpdate]: {
        ...current[sheetNameToUpdate],
        languageColumns: {
          ...(current[sheetNameToUpdate]?.languageColumns ?? {}),
          [lang]: value
        }
      }
    }));
  }

  function removeSheetOverride(sheetNameToRemove: string) {
    setSheetColumnOverrides((current) => Object.fromEntries(Object.entries(current).filter(([name]) => name !== sheetNameToRemove)));
  }

  function applyConfig(config: SavedProjectConfig) {
    setProjectName(config.projectName);
    setExcelPath(config.excelUrl ?? "");
    setExcelSource(null);
    setPreview(null);
    setProjectRoot(config.projectRoot);
    setKeyColumn(config.keyColumn);
    setLanguageColumns(config.languageColumns);
    setSheetColumnOverrides(config.sheetColumnOverrides ?? {});
    setOutputPathTemplate(config.outputPathTemplate);
    setOutputFormat(config.outputFormat);
    setMergeStrategy(config.mergeStrategy);
    setMissingKeyStrategy(config.missingKeyStrategy ?? "keep");
    setSheetName(config.sheetName ?? "");
    setSelectedSheetNames(config.sheetNames?.length ? config.sheetNames : config.sheetName ? [config.sheetName] : []);
    setSkipRows(config.skipRows ?? 0);
    setHeaderRow(config.headerRow ?? 1);
    const legacyBigEaterConfig = !config.moduleSplitMode && (config.moduleFilter ?? "").trim() === "大胃王";
    const nextModuleSplitMode = config.moduleSplitMode ?? (legacyBigEaterConfig ? "sectionRow" : config.splitByModule ? "keyPrefix" : "none");
    setModuleSplitMode(nextModuleSplitMode);
    setModuleNameSource(config.moduleNameSource ?? (nextModuleSplitMode === "sectionRow" ? "sectionRow" : "keyPrefix"));
    setKeyStyle(config.keyStyle ?? "nested");
    setModuleFilter(config.moduleFilter ?? "");
    setIgnoredModuleFilter(config.ignoredModuleFilter ?? "");
    setModuleNameReplacements(config.moduleNameReplacements ?? (legacyBigEaterConfig ? "大胃王=gachaguess" : ""));
    setRemoveModulePrefix(config.removeModulePrefix ?? false);
    setQuoteObjectProperties(config.quoteObjectProperties ?? false);
    setSpaceWrappedLanguages(config.spaceWrappedLanguages ?? "");
    setEnsureTrailingNewline(config.ensureTrailingNewline ?? true);
    setPlans([]);
    setSelectedPlanPath("");
  }

  function handleOutputFormatChange(format: OutputFormat) {
    setOutputFormat(format);
    if (format === "xcstrings" && !/\.xcstrings$/i.test(outputPathTemplate.trim())) {
      setOutputPathTemplate("Localizable.xcstrings");
    }
  }

  function loadConfig(id: string) {
    const config = configs.find((item) => item.id === id);
    setSelectedConfigId(id);
    if (!config) return;
    applyConfig(config);
    setMessage(`已加载配置：${config.projectName}`);
  }

  function saveConfig() {
    try {
      const saved = saveProjectConfig(projectName, settings, selectedConfigId || undefined);
      setConfigs(loadProjectConfigs());
      setSelectedConfigId(saved.id);
      setMessage("配置已保存。");
    } catch (error) {
      showError("保存配置失败", error);
    }
  }

  async function copyConfig() {
    try {
      const config: SavedProjectConfig = {
        id: selectedConfigId || crypto.randomUUID(),
        projectName: projectName.trim() || "未命名项目",
        updatedAt: new Date().toISOString(),
        ...settings
      };
      await copyTextToClipboard(serializeProjectConfig(config));
      setMessage("配置已复制到剪贴板。");
    } catch (error) {
      showError("复制配置失败", error);
    }
  }

  function deleteConfig() {
    if (!selectedConfigId) return;
    deleteProjectConfig(selectedConfigId);
    setConfigs(loadProjectConfigs());
    setSelectedConfigId("");
    setMessage("配置已删除。");
  }

  async function exportConfig() {
    const targetConfig = saveProjectConfig(projectName, settings, selectedConfigId || undefined);
    if (!isTauriRuntime()) {
      downloadTextFile(`${targetConfig.projectName}.i18n-config.json`, serializeProjectConfig(targetConfig));
      setConfigs(loadProjectConfigs());
      setSelectedConfigId(targetConfig.id);
      setMessage("配置已导出。");
      return;
    }

    const path = await save({
      defaultPath: `${targetConfig.projectName}.i18n-config.json`,
      filters: [{ name: "i18n config", extensions: ["json"] }]
    });
    if (!path) return;

    try {
      await writeTextFile(path, serializeProjectConfig(targetConfig));
      setConfigs(loadProjectConfigs());
      setSelectedConfigId(targetConfig.id);
      setMessage("配置已导出。");
    } catch (error) {
      showError("导出配置失败", error);
    }
  }

  async function importConfig() {
    if (!isTauriRuntime()) {
      const file = await chooseBrowserFile(".json");
      if (!file) return;

      setStatus("importing");
      try {
        const imported = parseProjectConfigExport(await file.text());
        const saved = saveProjectConfig(imported.projectName, imported, imported.id);
        setConfigs(loadProjectConfigs());
        setSelectedConfigId(saved.id);
        applyConfig(saved);
        setMessage("配置已导入。");
      } catch (error) {
        showError("导入配置失败", error);
      } finally {
        setStatus("idle");
      }
      return;
    }

    const path = await open({
      multiple: false,
      filters: [{ name: "i18n config", extensions: ["json"] }]
    });
    if (!path || Array.isArray(path)) return;

    setStatus("importing");
    try {
      const imported = parseProjectConfigExport(await readTextFile(path));
      const saved = saveProjectConfig(imported.projectName, imported, imported.id);
      setConfigs(loadProjectConfigs());
      setSelectedConfigId(saved.id);
      applyConfig(saved);
      setMessage("配置已导入。");
    } catch (error) {
      showError("导入配置失败", error);
    } finally {
      setStatus("idle");
    }
  }

  async function buildPreviewPlans() {
    if (!excelSource) {
      setMessage("请先选择或拖入 Excel 文件。");
      return;
    }
    if (isTauriRuntime() && !projectRoot) {
      setMessage("请先选择项目目录。");
      return;
    }

    setStatus("previewing");
    setMessage("");
    try {
      const rows = await readExcelRows(excelSource, {
        sheetName,
        sheetNames: selectedSheetNames,
        skipRows,
        headerRow,
        keyColumn,
        languageColumns,
        sheetColumnOverrides
      });
      if (isXcstringsOutput) {
        const sourceLanguage = getSourceLanguage(languageColumns);
        const locale = generateXcstringsLocale(rows, keyColumn, languageColumns);
        const path = await resolveOutputFilePath(projectRoot, getXcstringsOutputPath(outputPathTemplate), webProjectDirectory);
        const snapshot = await readOutputLocaleSnapshot(path, outputFormat, webProjectDirectory, projectRoot);
        const merged = mergeLocaleObjects(snapshot.locale, locale, mergeStrategy, missingKeyStrategy);
        const deletedKeys = missingKeyStrategy === "remove" ? getDeletedLocaleKeys(snapshot.locale, merged.mergedLocale) : [];
        const nextContent = await formatLocaleForSnapshot(merged.mergedLocale, outputFormat, snapshot, {
          eol: snapshot.eol,
          ensureTrailingNewline,
          sourceLanguage
        });
        const nextPlans: LocaleFilePlan[] = [
          {
            lang: sourceLanguage,
            moduleName: "xcstrings",
            path,
            fileAction: snapshot.exists ? "update" : "create",
            existingKeys: countLocaleKeys(snapshot.locale),
            incomingKeys: countLocaleKeys(locale),
            existingContent: snapshot.content,
            nextContent,
            eol: snapshot.eol,
            deletedKeys,
            ...merged
          }
        ];
        const changedPlans = nextPlans.filter((plan) => plan.error || plan.existingContent !== plan.nextContent);
        setPlans(changedPlans);
        setSelectedPlanPath(changedPlans[0]?.path ?? "");
        setMessage(changedPlans.length > 0 ? "写入预览已生成。" : "写入预览已生成：没有文件变化。");
        return;
      }
      const locales = generateLocales(rows, keyColumn, languageColumns, {
        splitByModule,
        moduleSplitMode,
        moduleNameSource,
        keyStyle,
        moduleFilter: parseModuleFilter(moduleFilter),
        ignoredModuleFilter: parseModuleFilter(ignoredModuleFilter),
        moduleNameReplacements: parseModuleNameReplacements(moduleNameReplacements),
        spaceWrappedLanguages: parseModuleFilter(spaceWrappedLanguages)
      });
      const writes = new Map<string, { lang: string; path: string; locale: LocaleObject; modules: Set<string> }>();
      const nextPlans: LocaleFilePlan[] = [];
      const templateSplitsFiles = outputPathTemplate.includes("{module}");

      for (const [lang, modules] of Object.entries(locales)) {
        for (const [moduleName, locale] of Object.entries(modules)) {
          await yieldToUi();
          const relativePath = resolveLocalePath(outputPathTemplate, lang, templateSplitsFiles ? moduleName : undefined);
          const path = await resolveOutputFilePath(projectRoot, relativePath, webProjectDirectory);
          const localeForPath =
            moduleSplitMode !== "none" && !templateSplitsFiles && moduleName && !removeModulePrefix ? { [moduleName]: locale } : locale;
          const existing = writes.get(path);
          writes.set(path, {
            lang,
            path,
            locale: existing ? combineIncomingLocale(existing.locale, localeForPath) : localeForPath,
            modules: new Set([...(existing?.modules ?? []), moduleName || "root"])
          });
        }
      }

      for (const { lang, path, locale, modules } of writes.values()) {
        await yieldToUi();
        try {
          const snapshot = await readOutputLocaleSnapshot(path, outputFormat, webProjectDirectory, projectRoot);
          const merged = mergeLocaleObjects(snapshot.locale, locale, mergeStrategy, missingKeyStrategy);
          const deletedKeys = missingKeyStrategy === "remove" ? getDeletedLocaleKeys(snapshot.locale, merged.mergedLocale) : [];
          const nextContent = await formatLocaleForSnapshot(merged.mergedLocale, outputFormat, snapshot, {
            eol: snapshot.eol,
            quoteObjectProperties,
            ensureTrailingNewline
          });
          nextPlans.push({
            lang,
            moduleName: Array.from(modules).join(", "),
            path,
            fileAction: snapshot.exists ? "update" : "create",
            existingKeys: countLocaleKeys(snapshot.locale),
            incomingKeys: countLocaleKeys(locale),
            existingContent: snapshot.content,
            nextContent,
            eol: snapshot.eol,
            deletedKeys,
            ...merged
          });
        } catch (error) {
          nextPlans.push({
            lang,
            moduleName: Array.from(modules).join(", "),
            path,
            fileAction: "update",
            existingKeys: 0,
            incomingKeys: countLocaleKeys(locale),
            addedKeys: [],
            overwrittenKeys: [],
            skippedKeys: [],
            deletedKeys: [],
            mergedLocale: {},
            existingContent: "",
            nextContent: "",
            eol: "lf",
            error: error instanceof Error ? error.message : String(error)
          });
        }
      }

      const changedPlans = nextPlans.filter((plan) => plan.error || plan.existingContent !== plan.nextContent);
      setPlans(changedPlans);
      setSelectedPlanPath(changedPlans[0]?.path ?? "");
      setMessage(changedPlans.length > 0 ? "写入预览已生成。" : "写入预览已生成：没有文件变化。");
    } catch (error) {
      showError("生成预览失败", error);
    } finally {
      setStatus("idle");
    }
  }

  async function runExcelHealthCheck() {
    if (!excelSource) {
      setMessage("请先选择或拖入 Excel 文件。");
      return;
    }

    setStatus("checking");
    setMessage("");
    await yieldToUi();
    try {
      const result = await inspectExcelSource(excelSource, keyColumn, languageColumns, {
        sheetName,
        sheetNames: selectedSheetNames,
        skipRows,
        headerRow,
        sheetColumnOverrides
      });
      setHealthCheckResult(result);
      setMessage("Excel 体检完成。");
    } catch (error) {
      showError("Excel 体检失败", error);
    } finally {
      setStatus("idle");
    }
  }

  async function previewLocaleExport() {
    if (!webProjectDirectory && isTauriRuntime() && !projectRoot) {
      setMessage("请先选择项目目录。");
      return;
    }
    if (outputFormat !== "xcstrings" && !outputPathTemplate.includes("{lang}")) {
      setMessage('导出路径模板必须包含 "{lang}"。');
      return;
    }

    setStatus("exporting");
    setMessage("");
    await yieldToUi();
    try {
      const nextPreview = await buildLocaleExportPreview();
      setLocaleExportPreview(nextPreview);
      setMessage(`导出预览已生成：${nextPreview.rows.length} 行，${nextPreview.fileCount} 个文件。`);
    } catch (error) {
      showError("生成导出预览失败", error);
    } finally {
      setStatus("idle");
    }
  }

  async function buildLocaleExportPreview(): Promise<LocaleExportPreview> {
    const languages = Object.entries(languageColumns).filter(([lang, column]) => lang.trim() && column.trim());
    if (languages.length === 0) throw new Error("请至少配置一个语言列映射。");
    if (!keyColumn.trim()) throw new Error("请配置 key 列。");

    if (outputFormat === "xcstrings") {
      const path = await resolveOutputFilePath(projectRoot, getXcstringsOutputPath(outputPathTemplate), webProjectDirectory);
      const snapshot = await readOutputLocaleSnapshot(path, outputFormat, webProjectDirectory, projectRoot);
      const rows = buildXcstringsExportRows(snapshot.locale, keyColumn, languageColumns, parseModuleFilter(exportModuleFilter));
      return {
        headers: getLocaleExportHeaders(keyColumn, languageColumns),
        rows,
        modules: parseModuleFilter(exportModuleFilter),
        fileCount: snapshot.exists ? 1 : 0,
        missingFiles: snapshot.exists ? [] : [path]
      };
    }

    const templateSplitsFiles = outputPathTemplate.includes("{module}");
    const modules = await resolveLocaleExportModules(templateSplitsFiles);
    const localeByLanguage: Record<string, Record<string, LocaleObject>> = {};
    const missingFiles: string[] = [];
    let fileCount = 0;

    for (const [lang] of languages) {
      localeByLanguage[lang] = {};
      const languageModules = templateSplitsFiles ? modules : [{ displayName: "", pathName: "" }];
      for (const moduleName of languageModules) {
        await yieldToUi();
        const relativePath = resolveLocalePath(outputPathTemplate, lang, templateSplitsFiles ? moduleName.pathName : undefined);
        const path = await resolveOutputFilePath(projectRoot, relativePath, webProjectDirectory);
        const snapshot = await readOutputLocaleSnapshot(path, outputFormat, webProjectDirectory, projectRoot);
        if (!snapshot.exists) {
          missingFiles.push(path);
          continue;
        }
        fileCount += 1;
        localeByLanguage[lang][moduleName.displayName] = snapshot.locale;
      }
    }

    const rows = buildLocaleExportRows(localeByLanguage, keyColumn, languageColumns, {
      templateSplitsFiles,
      moduleFilter: parseModuleFilter(exportModuleFilter),
      removeModulePrefix
    });

    return {
      headers: getLocaleExportHeaders(keyColumn, languageColumns),
      rows,
      modules: templateSplitsFiles ? modules.map((moduleName) => moduleName.displayName) : parseModuleFilter(exportModuleFilter),
      fileCount,
      missingFiles
    };
  }

  function downloadLocaleExportCsv() {
    if (!localeExportPreview) return;
    const sheet = XLSX.utils.json_to_sheet(localeExportPreview.rows, { header: localeExportPreview.headers });
    const csv = XLSX.utils.sheet_to_csv(sheet);
    downloadBlobFile(`${getSafeDownloadName(projectName)}.csv`, new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" }));
  }

  function downloadLocaleExportXlsx() {
    if (!localeExportPreview) return;
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(localeExportPreview.rows, { header: localeExportPreview.headers });
    XLSX.utils.book_append_sheet(workbook, sheet, "i18n");
    const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
    downloadBlobFile(
      `${getSafeDownloadName(projectName)}.xlsx`,
      new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
    );
  }

  async function resolveLocaleExportModules(templateSplitsFiles: boolean): Promise<LocaleExportModule[]> {
    const replacements = parseModuleNameReplacements(moduleNameReplacements);
    const explicitModules = parseModuleFilter(exportModuleFilter);
    if (explicitModules.length > 0) {
      return explicitModules.map((moduleName) => ({
        displayName: moduleName,
        pathName: replacements[moduleName] || moduleName
      }));
    }
    if (!templateSplitsFiles) return [];

    const firstLanguage = Object.keys(languageColumns).find((lang) => lang.trim());
    if (!firstLanguage) throw new Error("请至少配置一个语言列映射。");

    const marker = "__MODULE__";
    const markerPath = resolveLocalePath(outputPathTemplate, firstLanguage, marker);
    const parts = normalizeRelativePath(markerPath).split("/");
    const markerPartIndex = parts.findIndex((part) => part.includes(marker));
    if (markerPartIndex === -1 || parts.slice(0, markerPartIndex).some((part) => part.includes(marker))) {
      throw new Error('当前路径模板无法自动发现模块，请在"只导出模块"中填写模块名。');
    }

    const [prefix, suffix] = parts[markerPartIndex].split(marker);
    const directory = parts.slice(0, markerPartIndex).join("/");
    const names = await listProjectDirectoryEntryNames(directory);
    const discovered = names
      .filter((name) => name.startsWith(prefix) && name.endsWith(suffix))
      .map((name) => name.slice(prefix.length, name.length - suffix.length))
      .filter(Boolean)
      .sort((left, right) => left.localeCompare(right));

    if (discovered.length === 0) {
      throw new Error('没有在路径模板对应目录发现模块文件，请在"只导出模块"中填写模块名。');
    }

    return discovered.map((moduleName) => ({ displayName: moduleName, pathName: moduleName }));
  }

  async function listProjectDirectoryEntryNames(relativePath: string): Promise<string[]> {
    const normalized = normalizeRelativePath(relativePath);
    if (webProjectDirectory) {
      const directory = await getBrowserDirectoryHandle(webProjectDirectory, normalized);
      const entries = (directory as unknown as { entries?: () => AsyncIterable<[string, FileSystemHandle]> }).entries?.();
      if (!entries) return [];
      const names: string[] = [];
      for await (const [name, handle] of entries) {
        if (handle.kind === "file") names.push(name);
      }
      return names;
    }

    if (!isTauriRuntime()) return [];
    const path = normalized ? await resolveProjectFile(projectRoot, normalized) : projectRoot;
    const entries = await readDir(path);
    return entries.filter((entry) => !("children" in entry) || !entry.children).map((entry) => entry.name).filter(Boolean);
  }

  async function runImport() {
    if (plans.length === 0) {
      setMessage("请先生成写入预览。");
      return;
    }
    const failed = plans.find((plan) => plan.error);
    if (failed) {
      setMessage("存在解析失败的文件，请处理后再导入。");
      return;
    }

    setStatus("importing");
    try {
      const importDirectory = isTauriRuntime() ? webProjectDirectory : await chooseWebDirectoryForImport();
      if (!isTauriRuntime() && !importDirectory) {
        await downloadPreviewFiles();
        setMessage("已下载预览文件。");
        setStatus("idle");
        return;
      }
      await Promise.all(plans.map((plan) => writeOutputContent(plan.path, plan.nextContent, importDirectory, projectRoot)));
      setSummary({
        changedFiles: plans.length,
        lines: plans.map((plan) => ({
          lang: plan.lang,
          moduleName: plan.moduleName,
          added: plan.addedKeys.length,
          modified: plan.overwrittenKeys.length,
          deleted: plan.deletedKeys.length
        }))
      });
      setMessage("导入完成。");
    } catch (error) {
      showError("写入失败", error);
    } finally {
      setStatus("idle");
    }
  }

  async function downloadPreviewFiles() {
    const downloadablePlans = plans.filter((plan) => !plan.error);
    if (downloadablePlans.length === 0) {
      setMessage("没有可下载的预览文件。");
      return;
    }

    if (downloadablePlans.length === 1) {
      const plan = downloadablePlans[0];
      downloadTextFile(getFileName(plan.path), plan.nextContent);
      return;
    }

    const files = downloadablePlans.map((plan) => ({
      path: getPreviewDownloadPath(plan.path, projectRoot),
      content: plan.nextContent
    }));
    downloadBlobFile("i18n-preview-files.zip", createZipBlob(files));
  }

  function showError(prefix: string, error: unknown) {
    setMessage(`${prefix}：${error instanceof Error ? error.message : String(error)}`);
  }

  function scrollToPlan(path: string) {
    setSelectedPlanPath(path);
    const target = diffScrollerRef.current?.querySelector(`[data-plan-path="${cssEscape(path)}"]`);
    target?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function closeSummary() {
    setSummary(null);
  }

  function handleDiffScroll() {
    const scroller = diffScrollerRef.current;
    if (!scroller) return;
    const sections = Array.from(scroller.querySelectorAll<HTMLElement>("[data-plan-path]"));
    const active = sections.find((section) => section.getBoundingClientRect().bottom > scroller.getBoundingClientRect().top + 72);
    const path = active?.dataset.planPath;
    if (path && path !== selectedPlanPath) setSelectedPlanPath(path);
  }

  return (
    <main className="shell" onDragOver={(event) => event.preventDefault()} onDrop={handleDrop}>
      <Box component="nav" className="menuBar">
        <TextField
          label="配置名称"
          value={projectName}
          onChange={(event) => setProjectName(event.target.value)}
          size="small"
        />
        <Button onClick={saveConfig} variant="outlined">保存配置</Button>
        <TextField
          label="读取项目配置"
          value={selectedConfigId}
          onChange={(event) => loadConfig(event.target.value)}
          select
          size="small"
        >
          <MenuItem value="">读取项目配置</MenuItem>
          {configs.map((config) => (
            <MenuItem key={config.id} value={config.id}>{config.projectName}</MenuItem>
          ))}
        </TextField>
        <Button onClick={importConfig} variant="outlined">导入配置</Button>
        <Button onClick={exportConfig} variant="outlined">导出配置</Button>
        <Button onClick={copyConfig} variant="outlined">复制配置</Button>
        <Button onClick={deleteConfig} disabled={!selectedConfigId} color="error" variant="outlined">删除配置</Button>
      </Box>
      <header className="topbar">
        <div>
          <Typography component="h1" variant="h4">多语言导入工具</Typography>
          <Typography color="text.secondary">从 Excel 生成并合并项目 locale 文件</Typography>
        </div>
        <Stack className="actions" direction="row" spacing={1}>
          <Button onClick={runExcelHealthCheck} disabled={isWorking} variant="outlined">Excel体检</Button>
          <Button onClick={previewLocaleExport} disabled={isWorking} variant="outlined">导出预览</Button>
          <Button onClick={buildPreviewPlans} disabled={isWorking} variant="outlined">生成预览</Button>
          <Button onClick={runImport} disabled={isWorking} variant="contained">执行导入</Button>
        </Stack>
      </header>

      {message && <Alert className="notice" severity="info">{message}</Alert>}

      <section className="grid">
        <div className="setupArea">
          <div className="setupLeft">
            <Paper className="panel" variant="outlined">
              <Typography component="h2" variant="h6">文件</Typography>
              <Stack spacing={1.5}>
                <Stack className="inline" direction="row" spacing={1}>
                  <TextField
                    label="Excel 文件"
                    value={excelPath}
                    onChange={(event) => setExcelPath(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter" && isGoogleSheetUrl(excelPath)) void loadExcelUrl();
                    }}
                    placeholder="选择、拖入 .xlsx/.xls，或粘贴 Google Sheet 链接"
                    size="small"
                    fullWidth
                  />
                  <Button onClick={chooseExcelFile} disabled={isWorking} variant="outlined">选择</Button>
                  <Button onClick={loadExcelUrl} disabled={isWorking || !isGoogleSheetUrl(excelPath)} variant="outlined">读取链接</Button>
                </Stack>
                {canChooseProjectRoot && (
                  <Stack className="inline" direction="row" spacing={1}>
                    <TextField
                      label="项目目录"
                      value={projectRoot}
                      placeholder="选择写入的项目根目录"
                      slotProps={{ input: { readOnly: true } }}
                      size="small"
                      fullWidth
                    />
                    <Button onClick={chooseProjectRoot} variant="outlined">选择</Button>
                  </Stack>
                )}
              </Stack>
            </Paper>

            <Paper className="panel" variant="outlined">
              <Typography component="h2" variant="h6">Excel 设置</Typography>
              <TextField
                label="Sheet"
                value={selectedSheetNames}
                onChange={(event) => {
                  const value = event.target.value;
                  const next = typeof value === "string" ? value.split(",") : value;
                  setSelectedSheetNames(next);
                  setSheetName(next[0] ?? "");
                }}
                disabled={!preview}
                select
                slotProps={{ select: { multiple: true } }}
                size="small"
                fullWidth
              >
                {!preview && <MenuItem value="">Sheet</MenuItem>}
                {preview?.sheetNames.map((name) => (
                  <MenuItem key={name} value={name}>{name}</MenuItem>
                ))}
              </TextField>
              <Stack className="split" direction="row" spacing={1}>
                <TextField
                  label="跳过前 n 行"
                  type="number"
                  value={skipRows}
                  onChange={(event) => setSkipRows(Number(event.target.value))}
                  slotProps={{ htmlInput: { min: 0 } }}
                  size="small"
                  fullWidth
                />
                <TextField
                  label="表头行"
                  type="number"
                  value={headerRow}
                  onChange={(event) => setHeaderRow(Number(event.target.value))}
                  slotProps={{ htmlInput: { min: 1 } }}
                  size="small"
                  fullWidth
                />
              </Stack>
            </Paper>

            <Paper className="panel languagePanel" variant="outlined">
              <div className="panelHead">
                <Typography component="h2" variant="h6">语言列映射</Typography>
                <Button onClick={addLanguageRow} variant="outlined">添加语言</Button>
              </div>
              <div className="languageRows">
                {Object.entries(languageColumns).map(([lang, column], index) => (
                  <div className="languageRow" key={index}>
                    <TextField
                      label="语言"
                      value={lang}
                      onChange={(event) => updateLanguage(index, "lang", event.target.value)}
                      placeholder="zh-CN"
                      size="small"
                    />
                    <span>{"->"}</span>
                    <TextField
                      label="Excel 列"
                      value={column}
                      onChange={(event) => updateLanguage(index, "column", event.target.value)}
                      select
                      size="small"
                    >
                      <MenuItem value="">选择 Excel 列</MenuItem>
                      {preview?.headers.map((header) => (
                        <MenuItem key={header} value={header}>{header}</MenuItem>
                      ))}
                      {!preview?.headers.includes(column) && column && <MenuItem value={column}>{column}</MenuItem>}
                    </TextField>
                    <Button onClick={() => removeLanguageRow(index)} color="error" variant="text">删除</Button>
                  </div>
                ))}
              </div>
            </Paper>

            <Paper className="panel overridePanel" variant="outlined">
              <div className="panelHead">
                <div>
                  <Typography component="h2" variant="h6">Sheet 列覆盖</Typography>
                  <Typography variant="body2" color="text.secondary">
                    默认使用通用映射，只有个别 sheet 列名不一致时才在这里覆盖。
                  </Typography>
                </div>
                <Button onClick={addSheetOverride} disabled={!preview || preview.sheetNames.length === 0} variant="outlined">
                  添加覆盖
                </Button>
              </div>
              {Object.entries(sheetColumnOverrides).length > 0 ? (
                <div className="overrideRows">
                  {Object.entries(sheetColumnOverrides).map(([overrideSheetName, override]) => (
                    <div className="overrideCard" key={overrideSheetName}>
                      <Stack className="split" direction="row" spacing={1}>
                        <TextField
                          label="Sheet"
                          value={overrideSheetName}
                          onChange={(event) => renameSheetOverride(overrideSheetName, event.target.value)}
                          select
                          size="small"
                          fullWidth
                        >
                          {(preview?.sheetNames ?? [overrideSheetName]).map((name) => (
                            <MenuItem key={name} value={name}>{name}</MenuItem>
                          ))}
                        </TextField>
                        <Button onClick={() => removeSheetOverride(overrideSheetName)} color="error" variant="text">
                          删除
                        </Button>
                      </Stack>
                      <TextField
                        label="key 列覆盖"
                        value={override.keyColumn ?? ""}
                        onChange={(event) => updateSheetOverride(overrideSheetName, "keyColumn", event.target.value)}
                        placeholder={`留空则沿用 ${keyColumn}`}
                        size="small"
                        fullWidth
                      />
                      <div className="overrideLanguageRows">
                        {Object.entries(languageColumns).map(([lang, column]) => (
                          <div className="languageRow compactLanguageRow" key={`${overrideSheetName}-${lang}`}>
                            <TextField label="语言" value={lang} size="small" slotProps={{ input: { readOnly: true } }} />
                            <span>{"->"}</span>
                            <TextField
                              label="Sheet 列"
                              value={override.languageColumns?.[lang] ?? ""}
                              onChange={(event) => updateSheetOverrideLanguage(overrideSheetName, lang, event.target.value)}
                              placeholder={`留空则沿用 ${column}`}
                              size="small"
                            />
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <Typography className="empty" color="text.secondary">
                  多个 sheet 表头一致时不需要配置，只有特殊 sheet 才补一层覆盖即可。
                </Typography>
              )}
            </Paper>
          </div>

          <div className="setupRight">
            <Paper className="panel configPanel" variant="outlined">
              <Typography component="h2" variant="h6">导入配置</Typography>
              <TextField label="key 列" value={keyColumn} onChange={(event) => setKeyColumn(event.target.value)} select size="small">
                <MenuItem value={keyColumn}>{keyColumn}</MenuItem>
                {preview?.headers.filter((header) => header !== keyColumn).map((header) => (
                  <MenuItem key={header} value={header}>{header}</MenuItem>
                ))}
              </TextField>
              <TextField
                label="输出路径模板"
                value={outputPathTemplate}
                onChange={(event) => setOutputPathTemplate(event.target.value)}
                helperText={`可用变量：{lang}、{module}，例如 {lang}/{module}.json`}
                size="small"
              />
              <Stack className="split" direction="row" spacing={1}>
                <TextField label="输出格式" value={outputFormat} onChange={(event) => handleOutputFormatChange(event.target.value as OutputFormat)} select size="small" fullWidth>
                  <MenuItem value="json">json</MenuItem>
                  <MenuItem value="js">js</MenuItem>
                  <MenuItem value="ts">ts</MenuItem>
                  <MenuItem value="xcstrings">xcstrings</MenuItem>
                </TextField>
                <TextField label="合并策略" value={mergeStrategy} onChange={(event) => setMergeStrategy(event.target.value as MergeStrategy)} select size="small" fullWidth>
                  <MenuItem value="overwrite">覆盖已有 key</MenuItem>
                  <MenuItem value="skip">跳过已有 key</MenuItem>
                </TextField>
                <TextField label="Excel 不存在的 key" value={missingKeyStrategy} onChange={(event) => setMissingKeyStrategy(event.target.value as MissingKeyStrategy)} select size="small" fullWidth>
                  <MenuItem value="keep">保留</MenuItem>
                  <MenuItem value="remove">删除</MenuItem>
                </TextField>
              </Stack>
              {!isXcstringsOutput && (
                <>
                  <Stack className="split" direction="row" spacing={1}>
                    <TextField label="key 风格" value={keyStyle} onChange={(event) => setKeyStyle(event.target.value as KeyStyle)} select size="small" fullWidth>
                      <MenuItem value="nested">nested key</MenuItem>
                      <MenuItem value="flat">平铺 key</MenuItem>
                    </TextField>
                    <TextField label="模块划分" value={moduleSplitMode} onChange={(event) => setModuleSplitMode(event.target.value as ModuleSplitMode)} select size="small" fullWidth>
                      <MenuItem value="none">不划分</MenuItem>
                      <MenuItem value="keyPrefix">按 key 第一段</MenuItem>
                      <MenuItem value="sectionRow">无前缀模块行</MenuItem>
                    </TextField>
                    <TextField label="模块名来源" value={moduleNameSource} onChange={(event) => setModuleNameSource(event.target.value as ModuleNameSource)} select size="small" fullWidth>
                      <MenuItem value="keyPrefix">按 key 第一段</MenuItem>
                      <MenuItem value="sectionRow">按模块行</MenuItem>
                    </TextField>
                  </Stack>
                  <TextField label="只导入模块" value={moduleFilter} onChange={(event) => setModuleFilter(event.target.value)} placeholder="例如：base, agency；留空则全部导入" size="small" />
                  <TextField
                    label="只导出模块"
                    value={exportModuleFilter}
                    onChange={(event) => setExportModuleFilter(event.target.value)}
                    placeholder="例如：base, agency；留空则按路径自动发现"
                    size="small"
                  />
                  <TextField label="忽略模块" value={ignoredModuleFilter} onChange={(event) => setIgnoredModuleFilter(event.target.value)} placeholder="例如：debug, deprecated" size="small" />
                  <TextField
                    label="模块名替换"
                    value={moduleNameReplacements}
                    onChange={(event) => setModuleNameReplacements(event.target.value)}
                    placeholder="例如：大胃王=gachaguess，每行一个"
                    multiline
                    minRows={3}
                    size="small"
                  />
                  <FormControlLabel
                    control={<Checkbox checked={removeModulePrefix} onChange={(event) => setRemoveModulePrefix(event.target.checked)} />}
                    label="移除模块前缀"
                  />
                  <FormControlLabel
                    control={
                      <Checkbox
                        checked={quoteObjectProperties}
                        onChange={(event) => setQuoteObjectProperties(event.target.checked)}
                        disabled={outputFormat === "json"}
                      />
                    }
                    label="JS/TS 属性名使用双引号"
                  />
                  <TextField
                    label="两边加空格语言"
                    value={spaceWrappedLanguages}
                    onChange={(event) => setSpaceWrappedLanguages(event.target.value)}
                    placeholder="例如：ar, ur；留空则不处理"
                    helperText="适用于 Cocos 阿语/乌尔都语等需要文本左右空格的项目。"
                    size="small"
                  />
                </>
              )}
              <FormControlLabel
                control={<Checkbox checked={ensureTrailingNewline} onChange={(event) => setEnsureTrailingNewline(event.target.checked)} />}
                label="文件末尾保留空行"
              />
            </Paper>
          </div>
        </div>

        <Paper className="panel wide" variant="outlined">
          <Typography component="h2" variant="h6">Excel 预览</Typography>
          {preview && preview.headers.length > 0 ? (
            <div className="tableWrap">
              <div className="meta">
                <span>Sheet：{preview.activeSheetName}</span>
                <span>跳过：{preview.skipRows} 行</span>
                <span>表头：第 {preview.headerRow} 行</span>
                <span>列：{preview.headers.join(", ")}</span>
              </div>
              <table>
                <thead>
                  <tr>{preview.headers.map((header) => <th key={header}>{header}</th>)}</tr>
                </thead>
                <tbody>
                  {preview.rows.map((row, index) => (
                    <tr key={index}>
                      {preview.headers.map((header) => <td key={header}>{row[header]}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <Typography className="empty" color="text.secondary">选择或拖入 Excel 后展示对应 sheet 的前 20 行有效数据。</Typography>
          )}
        </Paper>

        <Paper className="panel wide previewPanel" variant="outlined">
          <div className="panelHead">
            <Typography component="h2" variant="h6">写入预览</Typography>
            <Stack direction="row" spacing={1}>
              <Button onClick={downloadPreviewFiles} disabled={isWorking || plans.length === 0} variant="outlined">下载文件</Button>
              <Button onClick={buildPreviewPlans} disabled={isWorking} variant="outlined">刷新预览</Button>
            </Stack>
          </div>
          {plans.length > 0 ? (
            <div className="previewSplit">
              <div className="planList scrollList">
                {plans.map((plan) => (
                  <ButtonBase
                    className={`plan ${selectedPlan?.path === plan.path ? "selectedPlan" : ""}`}
                    key={`${plan.lang}-${plan.path}`}
                    onClick={() => scrollToPlan(plan.path)}
                  >
                    <span className="planMain">
                      <strong>{plan.lang}</strong>
                      <code title={plan.path}>{plan.path}</code>
                    </span>
                    {plan.error ? (
                      <span className="danger">{plan.error}</span>
                    ) : (
                      <span className="planStats">
                        {plan.fileAction === "create" ? "新增文件" : "编辑文件"} · 新增 {plan.addedKeys.length} · 修改{" "}
                        {plan.overwrittenKeys.length} · 删除 {plan.deletedKeys.length} · 跳过 {plan.skippedKeys.length} ·{" "}
                        {plan.eol.toUpperCase()}
                      </span>
                    )}
                  </ButtonBase>
                ))}
              </div>
              <div className="diffPanel" ref={diffScrollerRef} onScroll={handleDiffScroll}>
                {diffPlans.map(({ plan, diff }) => (
                  <section className="diffFile" data-plan-path={plan.path} key={`${plan.lang}-${plan.path}`}>
                    <div className="diffFileHead">
                      <strong>{getFileName(plan.path)}</strong>
                      <span>{plan.fileAction === "create" ? "新增文件" : "编辑文件"}</span>
                      <span>
                        add: {plan.addedKeys.length}, modify: {plan.overwrittenKeys.length}, delete: {plan.deletedKeys.length}, skip:{" "}
                        {plan.skippedKeys.length}
                      </span>
                    </div>
                    <div className="diffPath">{plan.path}</div>
                    {plan.error ? (
                      <p className="danger">{plan.error}</p>
                    ) : diff.length > 0 ? (
                      <div className="diffBody">
                        {diff.map((row, index) =>
                          row.kind === "hunk" ? (
                            <div className="diffHunkHeader" key={`${index}-${row.header}`}>
                              {row.header}
                            </div>
                          ) : (
                            <div
                              className={`diffRow ${row.rowKind}`}
                              key={`${index}-${row.oldLineNumber ?? "n"}-${row.newLineNumber ?? "n"}-${row.oldText}-${row.newText}`}
                            >
                              <span className="diffLineNo diffOldNo">{formatDiffLineNumber(row.oldLineNumber)}</span>
                              <code className="diffCell diffOldCell">{row.oldText || " "}</code>
                              <span className="diffLineNo diffNewNo">{formatDiffLineNumber(row.newLineNumber)}</span>
                              <code className="diffCell diffNewCell">{row.newText || " "}</code>
                            </div>
                          )
                        )}
                      </div>
                    ) : (
                      <Typography className="empty" color="text.secondary">没有内容变化。</Typography>
                    )}
                  </section>
                ))}
              </div>
            </div>
          ) : (
            <Box className="previewEmpty">
              <Typography color="text.secondary">生成预览后会展示将修改的文件、key 数量和完整 diff。</Typography>
            </Box>
          )}
        </Paper>
      </section>

      <StatusBar status={status} plans={plans.length} previewRows={preview?.rows.length ?? 0} />
      <OperationDialog status={status} />
      {summary && <ImportSummaryDialog summary={summary} onClose={closeSummary} />}
      {healthCheckResult && <ExcelHealthCheckDialog result={healthCheckResult} onClose={() => setHealthCheckResult(null)} />}
      {localeExportPreview && (
        <LocaleExportPreviewDialog
          preview={localeExportPreview}
          onClose={() => setLocaleExportPreview(null)}
          onDownloadCsv={downloadLocaleExportCsv}
          onDownloadXlsx={downloadLocaleExportXlsx}
        />
      )}
    </main>
  );
}

function StatusBar({ status, plans, previewRows }: { status: StatusKind; plans: number; previewRows: number }) {
  const text =
    status === "reading"
      ? "读取文件中..."
      : status === "previewing"
        ? "生成预览中..."
        : status === "importing"
          ? "导入中..."
          : status === "exporting"
            ? "导出预览中..."
          : "就绪";

  return (
    <footer className="statusBar">
      <span>{text}</span>
      <span>预览行：{previewRows}</span>
      <span>待写入文件：{plans}</span>
    </footer>
  );
}

function OperationDialog({ status }: { status: StatusKind }) {
  if (status === "idle") return null;
  if (status === "checking") {
    return (
      <Dialog open aria-live="polite" aria-labelledby="operation-dialog-title">
        <DialogContent className="compactDialog">
          <CircularProgress size={42} />
          <Box>
            <DialogTitle id="operation-dialog-title" sx={{ p: 0, mb: 1 }}>Excel体检</DialogTitle>
            <Typography color="text.secondary">正在检查重复 key、空 key 行和缺失翻译。</Typography>
          </Box>
        </DialogContent>
      </Dialog>
    );
  }

  const title =
    status === "reading"
      ? "导入 Excel"
      : status === "previewing"
        ? "生成写入预览"
        : status === "exporting"
          ? "生成导出预览"
          : "写入文件";
  const detail =
    status === "reading"
      ? "正在读取工作簿并解析表头。"
      : status === "previewing"
        ? "正在比较 Excel 内容和现有 locale 文件。"
        : status === "exporting"
          ? "正在读取当前项目的多语言文件并整理为表格。"
          : "正在把预览中的变更写入项目。";

  return (
    <Dialog open aria-live="polite" aria-labelledby="operation-dialog-title">
      <DialogContent className="compactDialog">
        <CircularProgress size={42} />
        <Box>
          <DialogTitle id="operation-dialog-title" sx={{ p: 0, mb: 1 }}>{title}</DialogTitle>
          <Typography color="text.secondary">{detail}</Typography>
        </Box>
      </DialogContent>
    </Dialog>
  );
}

function ImportSummaryDialog({ summary, onClose }: { summary: ImportSummary; onClose: () => void }) {
  return (
    <Dialog open onClose={onClose} maxWidth="sm" fullWidth aria-labelledby="import-summary-title">
      <DialogTitle id="import-summary-title">导入成功，修改了 {summary.changedFiles} 个文件</DialogTitle>
      <DialogContent>
        <div className="summaryList">
          {summary.lines.map((line, index) => (
            <div className="summaryRow" key={`${line.lang}-${line.moduleName}-${index}`}>
              <strong>{line.lang}</strong>
              <code>{line.moduleName}</code>
              <span>新增、修改、删除：{line.added} / {line.modified} / {line.deleted}</span>
            </div>
          ))}
        </div>
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose} variant="contained">完成</Button>
      </DialogActions>
    </Dialog>
  );
}

function ExcelHealthCheckDialog({ result, onClose }: { result: ExcelHealthCheckResult; onClose: () => void }) {
  const report = formatExcelHealthCheckReport(result);

  async function copyReport() {
    await copyTextToClipboard(report);
  }

  return (
    <Dialog open onClose={onClose} maxWidth="md" fullWidth aria-labelledby="excel-health-check-title">
      <DialogTitle id="excel-health-check-title">Excel体检结果</DialogTitle>
      <DialogContent>
        <Stack spacing={2}>
          <Stack direction="row" spacing={1}>
            <Typography>重复key：{result.duplicateKeys.length}</Typography>
            <Typography>空key行：{result.emptyKeyRows.length}</Typography>
            <Typography>缺失翻译：{result.missingTranslations.length}</Typography>
          </Stack>
          <TextField value={report} multiline minRows={16} fullWidth slotProps={{ input: { readOnly: true } }} />
        </Stack>
      </DialogContent>
      <DialogActions>
        <Button onClick={copyReport} variant="outlined">一键复制</Button>
        <Button onClick={onClose} variant="contained">完成</Button>
      </DialogActions>
    </Dialog>
  );
}

function LocaleExportPreviewDialog({
  preview,
  onClose,
  onDownloadCsv,
  onDownloadXlsx
}: {
  preview: LocaleExportPreview;
  onClose: () => void;
  onDownloadCsv: () => void;
  onDownloadXlsx: () => void;
}) {
  return (
    <Dialog open onClose={onClose} maxWidth="xl" fullWidth aria-labelledby="locale-export-preview-title">
      <DialogTitle id="locale-export-preview-title">多语言导出预览</DialogTitle>
      <DialogContent>
        <Stack spacing={2}>
          <Stack className="meta exportMeta" direction="row" spacing={2}>
            <span>行数：{preview.rows.length}</span>
            <span>文件：{preview.fileCount}</span>
            <span>模块：{preview.modules.length > 0 ? preview.modules.join(", ") : "全部"}</span>
            {preview.missingFiles.length > 0 && <span>缺失文件：{preview.missingFiles.length}</span>}
          </Stack>
          {preview.missingFiles.length > 0 && (
            <Alert severity="warning">
              有 {preview.missingFiles.length} 个文件不存在，已跳过。可检查路径模板、语言列配置或只导出模块。
            </Alert>
          )}
          <div className="tableWrap exportPreviewTable">
            <table>
              <thead>
                <tr>{preview.headers.map((header) => <th key={header}>{header}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.map((row, index) => (
                  <tr key={index}>
                    {preview.headers.map((header) => <td key={header}>{row[header]}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Stack>
      </DialogContent>
      <DialogActions>
        <Button onClick={onDownloadCsv} variant="outlined">下载 CSV</Button>
        <Button onClick={onDownloadXlsx} variant="outlined">下载 XLSX</Button>
        <Button onClick={onClose} variant="contained">完成</Button>
      </DialogActions>
    </Dialog>
  );
}

function getLocaleExportHeaders(keyColumn: string, languageColumns: LanguageColumns): string[] {
  return [keyColumn, ...Object.values(languageColumns).filter((column) => column.trim())];
}

function buildXcstringsExportRows(
  locale: LocaleObject,
  keyColumn: string,
  languageColumns: LanguageColumns,
  moduleFilter: string[]
): Array<Record<string, string>> {
  const rows: Array<Record<string, string>> = [];
  Object.entries(locale).forEach(([key, value]) => {
    const moduleName = key.split(".").filter(Boolean)[0] ?? "";
    if (moduleFilter.length > 0 && !moduleFilter.includes(moduleName)) return;
    const row: Record<string, string> = { [keyColumn]: key };
    Object.entries(languageColumns).forEach(([lang, column]) => {
      if (!column.trim()) return;
      row[column] = isLocaleObject(value) && !isLocaleObject(value[lang]) ? String(value[lang] ?? "") : "";
    });
    rows.push(row);
  });
  return rows;
}

function buildLocaleExportRows(
  localeByLanguage: Record<string, Record<string, LocaleObject>>,
  keyColumn: string,
  languageColumns: LanguageColumns,
  options: { templateSplitsFiles: boolean; moduleFilter: string[]; removeModulePrefix: boolean }
): Array<Record<string, string>> {
  const rowsByKey = new Map<string, Record<string, string>>();

  Object.entries(languageColumns).forEach(([lang, column]) => {
    if (!lang.trim() || !column.trim()) return;
    Object.entries(localeByLanguage[lang] ?? {}).forEach(([moduleName, locale]) => {
      flattenLocaleValues(locale).forEach(([localeKey, value]) => {
        const exportKey = getLocaleExportKey(moduleName, localeKey, options);
        const exportModuleName = moduleName || exportKey.split(".").filter(Boolean)[0] || "";
        if (options.moduleFilter.length > 0 && !options.moduleFilter.includes(exportModuleName)) return;
        const row = rowsByKey.get(exportKey) ?? { [keyColumn]: exportKey };
        row[column] = value;
        rowsByKey.set(exportKey, row);
      });
    });
  });

  return Array.from(rowsByKey.values()).map((row) => {
    Object.values(languageColumns).forEach((column) => {
      if (column.trim() && row[column] == null) row[column] = "";
    });
    return row;
  });
}

function getLocaleExportKey(
  moduleName: string,
  localeKey: string,
  options: { templateSplitsFiles: boolean; removeModulePrefix: boolean }
): string {
  if (!options.templateSplitsFiles || !moduleName || options.removeModulePrefix) return localeKey;
  return localeKey ? `${moduleName}.${localeKey}` : moduleName;
}

function flattenLocaleValues(locale: LocaleObject, prefix = ""): Array<[string, string]> {
  return Object.entries(locale).flatMap(([key, value]) => {
    const nextKey = prefix ? `${prefix}.${key}` : key;
    return isLocaleObject(value) ? flattenLocaleValues(value, nextKey) : [[nextKey, String(value ?? "")]];
  });
}

function getSafeDownloadName(value: string): string {
  return (value.trim() || "i18n-export").replace(/[\\/:*?"<>|]+/g, "-");
}

function parseModuleFilter(value: string): string[] {
  return value
    .split(/[,，\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function getSourceLanguage(languageColumns: LanguageColumns): string {
  return Object.keys(languageColumns).find((lang) => lang.trim())?.trim() ?? "en";
}

function getXcstringsOutputPath(template: string): string {
  const outputPath = template.trim();
  if (!outputPath) throw new Error("输出路径不能为空。");
  return outputPath.split("\\").join("/");
}

function parseModuleNameReplacements(value: string): Record<string, string> {
  return value
    .split(/\r?\n|[,，]/)
    .map((item) => item.trim())
    .filter(Boolean)
    .reduce<Record<string, string>>((result, item) => {
      const match = item.match(/^(.*?)\s*(?:=>|=|:|：)\s*(.*?)$/);
      if (!match) return result;
      const from = match[1].trim();
      const to = match[2].trim();
      if (from && to) result[from] = to;
      return result;
    }, {});
}

function countLocaleKeys(locale: LocaleObject): number {
  return Object.values(locale).reduce((count, value) => {
    if (value && typeof value === "object") return count + countLocaleKeys(value);
    return count + 1;
  }, 0);
}

function getDeletedLocaleKeys(before: LocaleObject, after: LocaleObject): string[] {
  const afterKeys = new Set(flattenLocaleKeys(after));
  return flattenLocaleKeys(before).filter((key) => !afterKeys.has(key));
}

function flattenLocaleKeys(locale: LocaleObject, prefix = ""): string[] {
  return Object.entries(locale).flatMap(([key, value]) => {
    const nextKey = prefix ? `${prefix}.${key}` : key;
    return isLocaleObject(value) ? flattenLocaleKeys(value, nextKey) : [nextKey];
  });
}

function combineIncomingLocale(target: LocaleObject, incoming: LocaleObject): LocaleObject {
  const combined: LocaleObject = { ...target };
  Object.entries(incoming).forEach(([key, value]) => {
    const existing = combined[key];
    if (isLocaleObject(existing) && isLocaleObject(value)) {
      combined[key] = combineIncomingLocale(existing, value);
    } else {
      combined[key] = value;
    }
  });
  return combined;
}

function isLocaleObject(value: unknown): value is LocaleObject {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

type DiffEdit =
  | { kind: "equal"; text: string; oldLineNumber: number; newLineNumber: number }
  | { kind: "delete"; text: string; oldLineNumber: number; newLineNumber: null }
  | { kind: "insert"; text: string; oldLineNumber: null; newLineNumber: number };

function createGitStyleDiff(before: string, after: string): DiffRow[] {
  if (before === after) return [];

  const beforeLines = splitLines(before);
  const afterLines = splitLines(after);
  const edits = buildLineEdits(beforeLines, afterLines);
  const rows = buildDiffRows(edits);
  return buildDiffHunks(rows);
}

function splitLines(text: string): string[] {
  return text.length === 0 ? [] : text.split(/\r?\n/);
}

function buildLineEdits(beforeLines: string[], afterLines: string[]): DiffEdit[] {
  const n = beforeLines.length;
  const m = afterLines.length;
  const max = n + m;
  const trace: Map<number, number>[] = [];
  let v = new Map<number, number>([[1, 0]]);

  for (let d = 0; d <= max; d += 1) {
    trace.push(new Map(v));
    for (let k = -d; k <= d; k += 2) {
      const left = v.get(k - 1) ?? Number.NEGATIVE_INFINITY;
      const right = v.get(k + 1) ?? Number.NEGATIVE_INFINITY;
      let x: number;
      if (k === -d || (k !== d && left < right)) {
        x = right;
      } else {
        x = left + 1;
      }
      let y = x - k;
      while (x < n && y < m && beforeLines[x] === afterLines[y]) {
        x += 1;
        y += 1;
      }
      v.set(k, x);
      if (x >= n && y >= m) {
        return backtrackLineEdits(trace, beforeLines, afterLines);
      }
    }
  }

  return [];
}

function backtrackLineEdits(trace: Map<number, number>[], beforeLines: string[], afterLines: string[]): DiffEdit[] {
  let x = beforeLines.length;
  let y = afterLines.length;
  const edits: DiffEdit[] = [];

  for (let d = trace.length - 1; d > 0; d -= 1) {
    const v = trace[d - 1];
    const k = x - y;
    const left = v.get(k - 1) ?? Number.NEGATIVE_INFINITY;
    const right = v.get(k + 1) ?? Number.NEGATIVE_INFINITY;
    const prevK = k === -d || (k !== d && left < right) ? k + 1 : k - 1;
    const prevX = v.get(prevK) ?? 0;
    const prevY = prevX - prevK;

    while (x > prevX && y > prevY) {
      edits.push({
        kind: "equal",
        text: beforeLines[x - 1] ?? "",
        oldLineNumber: x,
        newLineNumber: y
      });
      x -= 1;
      y -= 1;
    }

    if (x === prevX) {
      edits.push({
        kind: "insert",
        text: afterLines[prevY] ?? "",
        oldLineNumber: null,
        newLineNumber: prevY + 1
      });
    } else {
      edits.push({
        kind: "delete",
        text: beforeLines[prevX] ?? "",
        oldLineNumber: prevX + 1,
        newLineNumber: null
      });
    }

    x = prevX;
    y = prevY;
  }

  while (x > 0 && y > 0) {
    edits.push({
      kind: "equal",
      text: beforeLines[x - 1] ?? "",
      oldLineNumber: x,
      newLineNumber: y
    });
    x -= 1;
    y -= 1;
  }

  while (x > 0) {
    edits.push({
      kind: "delete",
      text: beforeLines[x - 1] ?? "",
      oldLineNumber: x,
      newLineNumber: null
    });
    x -= 1;
  }

  while (y > 0) {
    edits.push({
      kind: "insert",
      text: afterLines[y - 1] ?? "",
      oldLineNumber: null,
      newLineNumber: y
    });
    y -= 1;
  }

  return edits.reverse();
}

function buildDiffRows(edits: DiffEdit[]): DiffRow[] {
  const rows: DiffRow[] = [];
  let index = 0;

  while (index < edits.length) {
    const edit = edits[index];
    if (edit.kind === "equal") {
      rows.push({
        kind: "row",
        rowKind: "equal",
        oldLineNumber: edit.oldLineNumber,
        newLineNumber: edit.newLineNumber,
        oldText: edit.text,
        newText: edit.text
      });
      index += 1;
      continue;
    }

    const block: DiffEdit[] = [];
    while (index < edits.length && edits[index].kind !== "equal") {
      block.push(edits[index]);
      index += 1;
    }

    const deletes = block.filter((item): item is Extract<DiffEdit, { kind: "delete" }> => item.kind === "delete");
    const inserts = block.filter((item): item is Extract<DiffEdit, { kind: "insert" }> => item.kind === "insert");

    deletes.forEach((removed) => {
      rows.push({
        kind: "row",
        rowKind: "delete",
        oldLineNumber: removed.oldLineNumber,
        newLineNumber: null,
        oldText: removed.text,
        newText: ""
      });
    });

    inserts.forEach((added) => {
      rows.push({
        kind: "row",
        rowKind: "insert",
        oldLineNumber: null,
        newLineNumber: added.newLineNumber,
        oldText: "",
        newText: added.text
      });
    });
  }

  return rows;
}

function buildDiffHunks(rows: DiffRow[], contextLines = 3): DiffRow[] {
  const changeIndices = rows
    .map((row, index) => (row.kind === "row" && row.rowKind !== "equal" ? index : null))
    .filter((value): value is number => value !== null);

  if (changeIndices.length === 0) return [];

  const ranges: Array<{ start: number; end: number }> = [];
  for (const index of changeIndices) {
    const nextRange = { start: Math.max(0, index - contextLines), end: Math.min(rows.length - 1, index + contextLines) };
    const lastRange = ranges[ranges.length - 1];
    if (lastRange && nextRange.start <= lastRange.end + 1) {
      lastRange.end = Math.max(lastRange.end, nextRange.end);
    } else {
      ranges.push(nextRange);
    }
  }

  const hunks: DiffRow[] = [];
  for (const range of ranges) {
    const hunkRows = rows.slice(range.start, range.end + 1);
    hunks.push({ kind: "hunk", header: formatDiffHeader(hunkRows) });
    hunks.push(...hunkRows);
  }

  return hunks;
}

function formatDiffHeader(rows: DiffRow[]): string {
  const oldLineNumbers = rows.flatMap((row) => (row.kind === "row" && row.oldLineNumber !== null ? [row.oldLineNumber] : []));
  const newLineNumbers = rows.flatMap((row) => (row.kind === "row" && row.newLineNumber !== null ? [row.newLineNumber] : []));
  const oldStart = oldLineNumbers[0] ?? 0;
  const newStart = newLineNumbers[0] ?? 0;
  const oldCount = oldLineNumbers.length;
  const newCount = newLineNumbers.length;
  return `@@ -${formatDiffRange(oldStart, oldCount)} +${formatDiffRange(newStart, newCount)} @@`;
}

function formatDiffRange(start: number, count: number): string {
  return count <= 1 ? `${start}` : `${start},${count}`;
}

function formatDiffLineNumber(value: number | null): string {
  return value === null ? "" : String(value);
}

function getFileName(path: string): string {
  return path.split(/[\\/]/).filter(Boolean).pop() ?? path;
}

function getPreviewDownloadPath(path: string, projectRoot: string): string {
  const normalizedPath = path.split("\\").join("/");
  const normalizedRoot = projectRoot.split("\\").join("/").replace(/\/+$/, "");
  if (normalizedRoot && normalizedPath.startsWith(`${normalizedRoot}/`)) {
    return normalizeRelativePath(normalizedPath.slice(normalizedRoot.length + 1));
  }
  return normalizeRelativePath(normalizedPath);
}

function cssEscape(value: string): string {
  return globalThis.CSS?.escape(value) ?? value.replace(/["\\]/g, "\\$&");
}

function isTauriRuntime(): boolean {
  return "__TAURI_INTERNALS__" in window;
}

async function loadGoogleSheetSource(url: string): Promise<ExcelSource> {
  const exportUrl = toGoogleSheetExportUrl(url);
  const response = await fetch(exportUrl);
  if (!response.ok) {
    throw new Error(`Google Sheet 下载失败：${response.status}`);
  }

  const contentType = response.headers.get("content-type") ?? "";
  const buffer = await response.arrayBuffer();
  if (contentType.includes("text/html")) {
    throw new Error("Google Sheet 返回了网页内容，请确认链接有访问权限。");
  }

  return { label: url, bytes: arrayBufferToBytes(buffer) };
}

function isGoogleSheetUrl(value: string): boolean {
  return /https:\/\/docs\.google\.com\/spreadsheets\/d\/[^/]+/i.test(value.trim());
}

function isHttpUrl(value: string): boolean {
  return /^https?:\/\//i.test(value.trim());
}

function toGoogleSheetExportUrl(value: string): string {
  const url = new URL(value.trim());
  const match = url.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
  if (!match) {
    throw new Error("请输入有效的 Google Sheet 链接。");
  }
  const gid = url.searchParams.get("gid") ?? url.hash.match(/gid=(\d+)/)?.[1];
  const exportUrl = new URL(`https://docs.google.com/spreadsheets/d/${match[1]}/export`);
  exportUrl.searchParams.set("format", "xlsx");
  if (gid) exportUrl.searchParams.set("gid", gid);
  return exportUrl.toString();
}

async function chooseBrowserFile(accept: string): Promise<File | null> {
  if ("showOpenFilePicker" in window) {
    try {
      const showOpenFilePicker = window.showOpenFilePicker;
      if (!showOpenFilePicker) throw new Error("当前浏览器不支持文件选择。");
      const [handle] = await showOpenFilePicker({
        multiple: false,
        types: [{ description: "Files", accept: { "*/*": accept.split(",") } }]
      });
      return handle ? await handle.getFile() : null;
    } catch (error) {
      if (isAbortError(error)) return null;
      throw error;
    }
  }

  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = accept;
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.click();
  });
}

async function copyTextToClipboard(text: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();
  const copied = document.execCommand("copy");
  document.body.removeChild(textarea);
  if (!copied) throw new Error("当前浏览器不支持复制到剪贴板。");
}

async function ensureDirectoryPermission(handle: WebDirectoryHandle): Promise<void> {
  const queryPermission = handle.queryPermission?.bind(handle);
  const requestPermission = handle.requestPermission?.bind(handle);
  if (!queryPermission || !requestPermission) return;
  if ((await queryPermission({ mode: "readwrite" })) === "granted") return;
  if ((await requestPermission({ mode: "readwrite" })) !== "granted") {
    throw new Error("没有获得目录读写权限。");
  }
}

async function chooseWebDirectoryForImport(): Promise<WebDirectoryHandle | null> {
  const showDirectoryPicker = window.showDirectoryPicker;
  if (!window.isSecureContext || !showDirectoryPicker) return null;

  try {
    const handle = await showDirectoryPicker({ mode: "readwrite" });
    await ensureDirectoryPermission(handle);
    return handle;
  } catch (error) {
    if (isAbortError(error)) return null;
    throw error;
  }
}

async function resolveOutputFilePath(projectRoot: string, relativePath: string, directory: WebDirectoryHandle | null): Promise<string> {
  const normalized = normalizeRelativePath(relativePath);
  if (directory) return normalized;
  if (!isTauriRuntime()) return normalized;
  return resolveProjectFile(projectRoot, normalized);
}

async function readOutputLocaleSnapshot(
  path: string,
  format: OutputFormat,
  directory: WebDirectoryHandle | null,
  projectRoot: string
): Promise<LocaleFileSnapshot> {
  if (!directory && !isTauriRuntime()) {
    return { exists: false, content: "", locale: {}, eol: "lf" };
  }
  if (!directory) return readLocaleFileSnapshot(path, format);

  try {
    await ensureDirectoryPermission(directory);
    const handle = await getBrowserFileHandle(directory, path, false);
    if (!handle) return { exists: false, content: "", locale: {}, eol: "lf" };

    const content = await (await handle.getFile()).text();
    return createLocaleSnapshotFromContent(content, format);
  } catch (error) {
    if (isMissingBrowserFileError(error)) return { exists: false, content: "", locale: {}, eol: "lf" };
    throw error;
  }
}

async function writeOutputContent(
  path: string,
  content: string,
  directory: WebDirectoryHandle | null,
  projectRoot: string
): Promise<void> {
  if (!directory) {
    if (!isTauriRuntime()) {
      throw new Error("请先选择当前电脑上的项目目录。");
    }
    await writeLocaleContent(path, content);
    return;
  }

  await ensureDirectoryPermission(directory);
  const handle = await getBrowserFileHandle(directory, path, true);
  if (!handle) throw new Error(`无法创建文件：${path}`);
  const writable = await handle.createWritable();
  await writable.write(content);
  await writable.close();
}

function createLocaleSnapshotFromContent(content: string, format: OutputFormat): LocaleFileSnapshot {
  return {
    exists: true,
    content,
    locale: parseLocaleContent(content, format),
    eol: detectEol(content),
    objectRange: format === "json" || format === "xcstrings" || !content.trim() ? undefined : findLocaleObjectRange(content)
  };
}

async function getBrowserFileHandle(
  directory: WebDirectoryHandle,
  relativePath: string,
  create: boolean
): Promise<FileSystemFileHandle | null> {
  const parts = normalizeRelativePath(relativePath).split("/").filter(Boolean);
  if (parts.length === 0) return null;

  let cursor = directory;
  for (const part of parts.slice(0, -1)) {
    cursor = await cursor.getDirectoryHandle(part, { create });
  }

  return cursor.getFileHandle(parts[parts.length - 1], { create });
}

async function getBrowserDirectoryHandle(directory: WebDirectoryHandle, relativePath: string): Promise<WebDirectoryHandle> {
  const parts = normalizeRelativePath(relativePath).split("/").filter(Boolean);
  let cursor = directory;
  for (const part of parts) {
    cursor = await cursor.getDirectoryHandle(part, { create: false });
  }
  return cursor;
}

function normalizeRelativePath(path: string): string {
  const normalized = path.split("\\").join("/").replace(/^\/+/, "");
  if (normalized.split("/").some((part) => part === "..")) {
    throw new Error("输出路径不能包含 ..。");
  }
  return normalized;
}

function downloadTextFile(fileName: string, content: string): void {
  const blob = new Blob([content], { type: "application/json;charset=utf-8" });
  downloadBlobFile(fileName, blob);
}

function downloadBlobFile(fileName: string, blob: Blob): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function createZipBlob(files: Array<{ path: string; content: string }>): Blob {
  const encoder = new TextEncoder();
  const fileRecords: Array<{ name: Uint8Array; data: Uint8Array; crc: number; offset: number }> = [];
  const localParts: Uint8Array[] = [];
  let offset = 0;

  files.forEach((file) => {
    const name = encoder.encode(file.path);
    const data = encoder.encode(file.content);
    const crc = crc32(data);
    const header = createZipLocalHeader(name, data, crc);
    localParts.push(header, data);
    fileRecords.push({ name, data, crc, offset });
    offset += header.length + data.length;
  });

  const centralParts = fileRecords.map((record) => createZipCentralHeader(record));
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = createZipEndRecord(fileRecords.length, centralSize, offset);
  return new Blob([...localParts, ...centralParts, end].map(toBlobPart), { type: "application/zip" });
}

function toBlobPart(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.length);
  copy.set(bytes);
  return copy.buffer;
}

function createZipLocalHeader(name: Uint8Array, data: Uint8Array, crc: number): Uint8Array {
  const header = new Uint8Array(30 + name.length);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0x0800, true);
  view.setUint16(8, 0, true);
  setZipDateTime(view, 10);
  view.setUint32(14, crc, true);
  view.setUint32(18, data.length, true);
  view.setUint32(22, data.length, true);
  view.setUint16(26, name.length, true);
  header.set(name, 30);
  return header;
}

function createZipCentralHeader(record: { name: Uint8Array; data: Uint8Array; crc: number; offset: number }): Uint8Array {
  const header = new Uint8Array(46 + record.name.length);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0x0800, true);
  view.setUint16(10, 0, true);
  setZipDateTime(view, 12);
  view.setUint32(16, record.crc, true);
  view.setUint32(20, record.data.length, true);
  view.setUint32(24, record.data.length, true);
  view.setUint16(28, record.name.length, true);
  view.setUint32(42, record.offset, true);
  header.set(record.name, 46);
  return header;
}

function createZipEndRecord(fileCount: number, centralSize: number, centralOffset: number): Uint8Array {
  const header = new Uint8Array(22);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(8, fileCount, true);
  view.setUint16(10, fileCount, true);
  view.setUint32(12, centralSize, true);
  view.setUint32(16, centralOffset, true);
  return header;
}

function setZipDateTime(view: DataView, offset: number): void {
  const now = new Date();
  const time = (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2);
  const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  view.setUint16(offset, time, true);
  view.setUint16(offset + 2, date, true);
}

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

function isMissingBrowserFileError(error: unknown): boolean {
  return error instanceof DOMException && (error.name === "NotFoundError" || error.name === "NotAllowedError");
}

function yieldToUi(): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, 0));
}

function arrayBufferToBytes(buffer: ArrayBuffer): number[] {
  return Array.from(new Uint8Array(buffer));
}

function readExcelWorkbookInfo(source: ExcelSource, sheetName?: string): Promise<ExcelWorkbookInfo> {
  if (!isTauriRuntime()) {
    return Promise.resolve(readExcelWorkbookInfoFromBuffer(bytesToArrayBuffer(source.bytes ?? []), sheetName));
  }
  if (source.path) return invoke("read_excel_workbook_info", { path: source.path, sheetName });
  return invoke("read_excel_workbook_info_bytes", { bytes: source.bytes ?? [], sheetName });
}

function readExcelPreview(source: ExcelSource, options: ExcelReadOptions): Promise<ExcelPreview> {
  if (!isTauriRuntime()) {
    return Promise.resolve(parseExcelBuffer(bytesToArrayBuffer(source.bytes ?? []), options));
  }
  if (source.path) return invoke("preview_excel", { path: source.path, options });
  return invoke("preview_excel_bytes", { bytes: source.bytes ?? [], options });
}

function readExcelRows(source: ExcelSource, options: ExcelReadOptions): Promise<Array<Record<string, string>>> {
  if (!isTauriRuntime()) {
    return Promise.resolve(rowsFromExcelBuffer(bytesToArrayBuffer(source.bytes ?? []), options));
  }
  if (source.path) return invoke("rows_from_excel", { path: source.path, options });
  return invoke("rows_from_excel_bytes", { bytes: source.bytes ?? [], options });
}

function inspectExcelSource(
  source: ExcelSource,
  keyColumn: string,
  languageColumns: LanguageColumns,
  options: ExcelReadOptions
): Promise<ExcelHealthCheckResult> {
  if (!isTauriRuntime()) {
    return Promise.resolve(inspectExcelBuffer(bytesToArrayBuffer(source.bytes ?? []), keyColumn, languageColumns, options));
  }
  if (source.path) return invoke("inspect_excel", { path: source.path, keyColumn, languageColumns, options });
  return Promise.resolve(inspectExcelBuffer(bytesToArrayBuffer(source.bytes ?? []), keyColumn, languageColumns, options));
}

function bytesToArrayBuffer(bytes: number[]): ArrayBuffer {
  return new Uint8Array(bytes).buffer;
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <App />
    </ThemeProvider>
  </React.StrictMode>
);
