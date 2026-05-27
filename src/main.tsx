import React, { useEffect, useMemo, useRef, useState } from "react";
import ReactDOM from "react-dom/client";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readTextFile, writeTextFile } from "@tauri-apps/plugin-fs";
import "./styles.css";
import {
  deleteProjectConfig,
  loadProjectConfigs,
  parseProjectConfigExport,
  saveProjectConfig,
  serializeProjectConfig
} from "./core/configStore";
import { generateLocales, mergeLocaleObjects, resolveLocalePath } from "./core/localeGenerator";
import { formatLocaleForSnapshot, readLocaleFileSnapshot, resolveProjectFile, writeLocaleContent } from "./core/localeFileWriter";
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
  SavedProjectConfig
} from "./core/types";
import type { ExcelReadOptions, ExcelWorkbookInfo } from "./core/excelParser";

const DEFAULT_LANGUAGE_COLUMNS: LanguageColumns = {
  "zh-CN": "中文",
  "en-US": "英文"
};

type StatusKind = "idle" | "reading" | "previewing" | "importing";
type DiffLine = { kind: "same" | "added" | "removed"; prefix: string; text: string };
type ExcelSource = { label: string; path?: string; bytes?: number[] };
interface ImportSummary {
  changedFiles: number;
  lines: Array<{ lang: string; moduleName: string; added: number; modified: number; deleted: number }>;
}

function App() {
  const [excelPath, setExcelPath] = useState("");
  const [excelSource, setExcelSource] = useState<ExcelSource | null>(null);
  const [preview, setPreview] = useState<ExcelPreview | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [skipRows, setSkipRows] = useState(0);
  const [headerRow, setHeaderRow] = useState(1);
  const [projectName, setProjectName] = useState("默认项目");
  const [projectRoot, setProjectRoot] = useState("");
  const [keyColumn, setKeyColumn] = useState("key");
  const [languageColumns, setLanguageColumns] = useState<LanguageColumns>(DEFAULT_LANGUAGE_COLUMNS);
  const [outputPathTemplate, setOutputPathTemplate] = useState("locales/{lang}/{module}.json");
  const [outputFormat, setOutputFormat] = useState<OutputFormat>("json");
  const [mergeStrategy, setMergeStrategy] = useState<MergeStrategy>("overwrite");
  const [missingKeyStrategy, setMissingKeyStrategy] = useState<MissingKeyStrategy>("keep");
  const [moduleSplitMode, setModuleSplitMode] = useState<ModuleSplitMode>("none");
  const [moduleNameSource, setModuleNameSource] = useState<ModuleNameSource>("keyPrefix");
  const [keyStyle, setKeyStyle] = useState<KeyStyle>("nested");
  const [moduleFilter, setModuleFilter] = useState("");
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
  const previewRequestRef = useRef(0);
  const diffScrollerRef = useRef<HTMLDivElement | null>(null);
  const splitByModule = moduleSplitMode === "keyPrefix";

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
      projectRoot,
      keyColumn,
      languageColumns,
      outputPathTemplate,
      outputFormat,
      mergeStrategy,
      missingKeyStrategy,
      sheetName,
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
      keyColumn,
      languageColumns,
      outputPathTemplate,
      outputFormat,
      mergeStrategy,
      missingKeyStrategy,
      sheetName,
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
        diff: plan.error ? [] : createTextDiff(plan.existingContent, plan.nextContent)
      })),
    [plans]
  );
  const selectedPlan = plans.find((plan) => plan.path === selectedPlanPath) ?? plans[0];
  const isWorking = status !== "idle";

  async function chooseExcelFile() {
    setMessage("");
    const selected = await open({
      multiple: false,
      filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
    });
    if (!selected || Array.isArray(selected)) return;

    await loadExcelFilePath(selected, "读取 Excel 失败");
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
    setMessage("Excel 已读取。");
  }

  async function chooseProjectRoot() {
    const selected = await open({ directory: true, multiple: false });
    if (!selected || Array.isArray(selected)) return;
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

  function applyConfig(config: SavedProjectConfig) {
    setProjectName(config.projectName);
    setProjectRoot(config.projectRoot);
    setKeyColumn(config.keyColumn);
    setLanguageColumns(config.languageColumns);
    setOutputPathTemplate(config.outputPathTemplate);
    setOutputFormat(config.outputFormat);
    setMergeStrategy(config.mergeStrategy);
    setMissingKeyStrategy(config.missingKeyStrategy ?? "keep");
    setSheetName(config.sheetName ?? "");
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
      await navigator.clipboard.writeText(serializeProjectConfig(config));
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
    if (!projectRoot) {
      setMessage("请先选择项目目录。");
      return;
    }

    setStatus("previewing");
    setMessage("");
    try {
      const rows = await readExcelRows(excelSource, { sheetName, skipRows, headerRow });
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
          const path = await resolveProjectFile(projectRoot, relativePath);
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
          const snapshot = await readLocaleFileSnapshot(path, outputFormat);
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
      await Promise.all(plans.map((plan) => writeLocaleContent(plan.path, plan.nextContent)));
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
      <nav className="menuBar">
        <input value={projectName} onChange={(event) => setProjectName(event.target.value)} placeholder="配置名称" />
        <button onClick={saveConfig}>保存配置</button>
        <select value={selectedConfigId} onChange={(event) => loadConfig(event.target.value)}>
          <option value="">读取项目配置</option>
          {configs.map((config) => (
            <option key={config.id} value={config.id}>{config.projectName}</option>
          ))}
        </select>
        <button onClick={importConfig}>导入配置</button>
        <button onClick={exportConfig}>导出配置</button>
        <button onClick={copyConfig}>复制配置</button>
        <button onClick={deleteConfig} disabled={!selectedConfigId}>删除配置</button>
      </nav>
      <header className="topbar">
        <div>
          <h1>多语言导入工具</h1>
          <p>从 Excel 生成并合并项目 locale 文件</p>
        </div>
        <div className="actions">
          <button onClick={buildPreviewPlans} disabled={isWorking}>生成预览</button>
          <button className="primary" onClick={runImport} disabled={isWorking}>执行导入</button>
        </div>
      </header>

      {message && <div className="notice">{message}</div>}

      <section className="grid">
        <div className="setupArea">
          <div className="setupLeft">
            <div className="panel">
          <h2>1. 文件</h2>
          <label>
            Excel 文件
            <div className="inline">
              <input value={excelPath} readOnly placeholder="选择或拖入 .xlsx / .xls 文件" />
              <button onClick={chooseExcelFile} disabled={isWorking}>选择</button>
            </div>
          </label>
          <label>
            项目目录
            <div className="inline">
              <input value={projectRoot} readOnly placeholder="选择写入的项目根目录" />
              <button onClick={chooseProjectRoot}>选择</button>
            </div>
          </label>
            </div>

            <div className="panel">
          <h2>2. Excel 设置</h2>
          <label>
            Sheet
            <select value={sheetName} onChange={(event) => setSheetName(event.target.value)} disabled={!preview}>
              {preview?.sheetNames.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </label>
          <div className="split">
            <label>
              跳过前 n 行
              <input type="number" min={0} value={skipRows} onChange={(event) => setSkipRows(Number(event.target.value))} />
            </label>
            <label>
              表头行
              <input type="number" min={1} value={headerRow} onChange={(event) => setHeaderRow(Number(event.target.value))} />
            </label>
          </div>
            </div>

            <div className="panel languagePanel">
              <div className="panelHead">
                <h2>4. 语言列映射</h2>
                <button onClick={addLanguageRow}>添加语言</button>
              </div>
              <div className="languageRows">
                {Object.entries(languageColumns).map(([lang, column], index) => (
                  <div className="languageRow" key={index}>
                    <input value={lang} onChange={(event) => updateLanguage(index, "lang", event.target.value)} placeholder="zh-CN" />
                    <span>{"->"}</span>
                    <select value={column} onChange={(event) => updateLanguage(index, "column", event.target.value)}>
                      <option value="">选择 Excel 列</option>
                      {preview?.headers.map((header) => (
                        <option key={header} value={header}>{header}</option>
                      ))}
                      {!preview?.headers.includes(column) && column && <option value={column}>{column}</option>}
                    </select>
                    <button onClick={() => removeLanguageRow(index)}>删除</button>
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div className="setupRight">
            <div className="panel configPanel">
          <h2>3. 导入配置</h2>
          <label>
            key 列
            <select value={keyColumn} onChange={(event) => setKeyColumn(event.target.value)}>
              <option value={keyColumn}>{keyColumn}</option>
              {preview?.headers.filter((header) => header !== keyColumn).map((header) => (
                <option key={header} value={header}>{header}</option>
              ))}
            </select>
          </label>
          <label>
            输出路径模板
            <input value={outputPathTemplate} onChange={(event) => setOutputPathTemplate(event.target.value)} />
            <span className="hint">可用变量：{"{lang}"}、{"{module}"}，例如 {"{lang}/{module}.json"}</span>
          </label>
          <div className="split">
            <label>
              输出格式
              <select value={outputFormat} onChange={(event) => setOutputFormat(event.target.value as OutputFormat)}>
                <option value="json">json</option>
                <option value="js">js</option>
                <option value="ts">ts</option>
              </select>
            </label>
            <label>
              合并策略
              <select value={mergeStrategy} onChange={(event) => setMergeStrategy(event.target.value as MergeStrategy)}>
                <option value="overwrite">覆盖已有 key</option>
                <option value="skip">跳过已有 key</option>
              </select>
            </label>
            <label>
              Excel 不存在的 key
              <select value={missingKeyStrategy} onChange={(event) => setMissingKeyStrategy(event.target.value as MissingKeyStrategy)}>
                <option value="keep">保留</option>
                <option value="remove">删除</option>
              </select>
            </label>
          </div>
          <div className="split">
            <label>
              key 风格
              <select value={keyStyle} onChange={(event) => setKeyStyle(event.target.value as KeyStyle)}>
                <option value="nested">nested key</option>
                <option value="flat">平铺 key</option>
              </select>
            </label>
            <label>
              模块划分
              <select value={moduleSplitMode} onChange={(event) => setModuleSplitMode(event.target.value as ModuleSplitMode)}>
                <option value="none">不划分</option>
                <option value="keyPrefix">按 key 第一段</option>
                <option value="sectionRow">无前缀模块行</option>
              </select>
            </label>
            <label>
              模块名来源
              <select value={moduleNameSource} onChange={(event) => setModuleNameSource(event.target.value as ModuleNameSource)}>
                <option value="keyPrefix">按 key 第一段</option>
                <option value="sectionRow">按模块行</option>
              </select>
            </label>
          </div>
          <label>
            只导入模块
            <input value={moduleFilter} onChange={(event) => setModuleFilter(event.target.value)} placeholder="例如：base, agency；留空则全部导入" />
          </label>
          <label>
            忽略模块
            <input value={ignoredModuleFilter} onChange={(event) => setIgnoredModuleFilter(event.target.value)} placeholder="例如：debug, deprecated" />
          </label>
          <label>
            模块名替换
            <textarea
              value={moduleNameReplacements}
              onChange={(event) => setModuleNameReplacements(event.target.value)}
              placeholder="例如：大胃王=gachaguess，每行一个"
              rows={3}
            />
          </label>
          <label className="checkboxLabel">
            <input
              type="checkbox"
              checked={removeModulePrefix}
              onChange={(event) => setRemoveModulePrefix(event.target.checked)}
            />
            移除模块前缀
          </label>
          <label className="checkboxLabel">
            <input
              type="checkbox"
              checked={quoteObjectProperties}
              onChange={(event) => setQuoteObjectProperties(event.target.checked)}
              disabled={outputFormat === "json"}
            />
            JS/TS 属性名使用双引号
          </label>
          <label>
            两边加空格语言
            <input
              value={spaceWrappedLanguages}
              onChange={(event) => setSpaceWrappedLanguages(event.target.value)}
              placeholder="例如：ar, ur；留空则不处理"
            />
            <span className="hint">适用于 Cocos 阿语/乌尔都语等需要文本左右空格的项目。</span>
          </label>
          <label className="checkboxLabel">
            <input
              type="checkbox"
              checked={ensureTrailingNewline}
              onChange={(event) => setEnsureTrailingNewline(event.target.checked)}
            />
            文件末尾保留空行
          </label>
            </div>
          </div>
        </div>

        <div className="panel wide">
          <h2>5. Excel 预览</h2>
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
            <p className="empty">选择或拖入 Excel 后展示对应 sheet 的前 20 行有效数据。</p>
          )}
        </div>

        <div className="panel wide previewPanel">
          <div className="panelHead">
            <h2>6. 写入预览</h2>
            <button onClick={buildPreviewPlans} disabled={isWorking}>刷新预览</button>
          </div>
          {plans.length > 0 ? (
            <div className="previewSplit">
              <div className="planList scrollList">
                {plans.map((plan) => (
                  <button
                    className={`plan ${selectedPlan?.path === plan.path ? "selectedPlan" : ""}`}
                    key={`${plan.lang}-${plan.path}`}
                    onClick={() => scrollToPlan(plan.path)}
                  >
                    <strong>{plan.lang}</strong>
                    <code title={plan.path}>{plan.path}</code>
                    {plan.error ? (
                      <span className="danger">{plan.error}</span>
                    ) : (
                      <span className="planStats">
                        新增 {plan.addedKeys.length} · 修改 {plan.overwrittenKeys.length} · 删除 {plan.deletedKeys.length} · 跳过 {plan.skippedKeys.length} · {plan.eol.toUpperCase()}
                      </span>
                    )}
                  </button>
                ))}
              </div>
              <div className="diffPanel" ref={diffScrollerRef} onScroll={handleDiffScroll}>
                {diffPlans.map(({ plan, diff }) => (
                  <section className="diffFile" data-plan-path={plan.path} key={`${plan.lang}-${plan.path}`}>
                    <div className="diffFileHead">
                      <strong>{getFileName(plan.path)}</strong>
                      <span>
                        add: {plan.addedKeys.length}, modify: {plan.overwrittenKeys.length}, delete: {plan.deletedKeys.length}, skip:{" "}
                        {plan.skippedKeys.length}
                      </span>
                    </div>
                    <div className="diffPath">{plan.path}</div>
                    {plan.error ? (
                      <p className="danger">{plan.error}</p>
                    ) : diff.length > 0 ? (
                      <pre>
                        {diff.map((line, index) => (
                          <div className={line.kind} key={`${index}-${line.text}`}>
                            <span>{line.prefix}</span>
                            <code>{line.text || " "}</code>
                          </div>
                        ))}
                      </pre>
                    ) : (
                      <p className="empty">没有内容变化。</p>
                    )}
                  </section>
                ))}
              </div>
            </div>
          ) : (
            <p className="empty">生成预览后会展示将修改的文件、key 数量和完整 diff。</p>
          )}
        </div>
      </section>

      <StatusBar status={status} plans={plans.length} previewRows={preview?.rows.length ?? 0} />
      <OperationDialog status={status} />
      {summary && <ImportSummaryDialog summary={summary} onClose={closeSummary} />}
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

  const title =
    status === "reading" ? "导入 Excel" : status === "previewing" ? "生成写入预览" : "写入文件";
  const detail =
    status === "reading"
      ? "正在读取工作簿并解析表头。"
      : status === "previewing"
        ? "正在比较 Excel 内容和现有 locale 文件。"
        : "正在把预览中的变更写入项目。";

  return (
    <div className="modalScrim" role="status" aria-live="polite">
      <div className="materialDialog compactDialog">
        <div className="progressRing" />
        <div>
          <h2>{title}</h2>
          <p>{detail}</p>
        </div>
      </div>
    </div>
  );
}

function ImportSummaryDialog({ summary, onClose }: { summary: ImportSummary; onClose: () => void }) {
  return (
    <div className="modalScrim">
      <div className="materialDialog summaryDialog" role="dialog" aria-modal="true" aria-labelledby="import-summary-title">
        <h2 id="import-summary-title">导入成功，修改了 {summary.changedFiles} 个文件</h2>
        <div className="summaryList">
          {summary.lines.map((line, index) => (
            <div className="summaryRow" key={`${line.lang}-${line.moduleName}-${index}`}>
              <strong>{line.lang}</strong>
              <code>{line.moduleName}</code>
              <span>新增、修改、删除：{line.added} / {line.modified} / {line.deleted}</span>
            </div>
          ))}
        </div>
        <div className="dialogActions">
          <button className="primary" onClick={onClose}>完成</button>
        </div>
      </div>
    </div>
  );
}

function parseModuleFilter(value: string): string[] {
  return value
    .split(/[,，\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
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

function createTextDiff(before: string, after: string): DiffLine[] {
  if (!before && after) {
    return after.split(/\r?\n/).map((text) => ({ kind: "added", prefix: "+", text }));
  }
  const beforeLines = before.split(/\r?\n/);
  const afterLines = after.split(/\r?\n/);
  const maxLength = Math.max(beforeLines.length, afterLines.length);
  const diff: DiffLine[] = [];

  for (let index = 0; index < maxLength; index += 1) {
    const oldLine = beforeLines[index];
    const nextLine = afterLines[index];
    if (oldLine === nextLine) {
      diff.push({ kind: "same", prefix: " ", text: oldLine ?? "" });
      continue;
    }
    if (oldLine !== undefined) diff.push({ kind: "removed", prefix: "-", text: oldLine });
    if (nextLine !== undefined) diff.push({ kind: "added", prefix: "+", text: nextLine });
  }

  return diff;
}

function getFileName(path: string): string {
  return path.split(/[\\/]/).filter(Boolean).pop() ?? path;
}

function cssEscape(value: string): string {
  return globalThis.CSS?.escape(value) ?? value.replace(/["\\]/g, "\\$&");
}

function isTauriRuntime(): boolean {
  return "__TAURI_INTERNALS__" in window;
}

function yieldToUi(): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, 0));
}

function arrayBufferToBytes(buffer: ArrayBuffer): number[] {
  return Array.from(new Uint8Array(buffer));
}

function readExcelWorkbookInfo(source: ExcelSource, sheetName?: string): Promise<ExcelWorkbookInfo> {
  if (source.path) return invoke("read_excel_workbook_info", { path: source.path, sheetName });
  return invoke("read_excel_workbook_info_bytes", { bytes: source.bytes ?? [], sheetName });
}

function readExcelPreview(source: ExcelSource, options: ExcelReadOptions): Promise<ExcelPreview> {
  if (source.path) return invoke("preview_excel", { path: source.path, options });
  return invoke("preview_excel_bytes", { bytes: source.bytes ?? [], options });
}

function readExcelRows(source: ExcelSource, options: ExcelReadOptions): Promise<Array<Record<string, string>>> {
  if (source.path) return invoke("rows_from_excel", { path: source.path, options });
  return invoke("rows_from_excel_bytes", { bytes: source.bytes ?? [], options });
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
