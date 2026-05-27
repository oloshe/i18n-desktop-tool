import { exists, mkdir, readTextFile, writeTextFile } from "@tauri-apps/plugin-fs";
import { dirname, join } from "@tauri-apps/api/path";
import prettier from "prettier/standalone";
import babelPlugin from "prettier/plugins/babel";
import estreePlugin from "prettier/plugins/estree";
import typescriptPlugin from "prettier/plugins/typescript";
import type { LocaleObject, OutputFormat } from "./types";
import { findLocaleObjectRange, parseLocaleContent, type LocaleObjectRange } from "./localeFileReader";

export interface LocaleFileSnapshot {
  exists: boolean;
  content: string;
  locale: LocaleObject;
  eol: "lf" | "crlf";
  objectRange?: LocaleObjectRange;
}

export interface FormatLocaleOptions {
  eol?: "lf" | "crlf";
  ensureTrailingNewline?: boolean;
  quoteObjectProperties?: boolean;
}

export async function readExistingLocale(path: string, format: OutputFormat): Promise<LocaleObject> {
  return (await readLocaleFileSnapshot(path, format)).locale;
}

export async function readLocaleFileSnapshot(path: string, format: OutputFormat): Promise<LocaleFileSnapshot> {
  try {
    if (!(await exists(path))) {
      return { exists: false, content: "", locale: {}, eol: "lf" };
    }
    const content = await readTextFile(path);
    const objectRange = format === "json" || !content.trim() ? undefined : findLocaleObjectRange(content);
    return {
      exists: true,
      content,
      locale: parseLocaleContent(content, format),
      eol: detectEol(content),
      objectRange
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`读取 locale 文件失败：${message}`);
  }
}

export async function writeLocaleContent(path: string, content: string): Promise<void> {
  const directory = await dirname(path);
  if (!(await exists(directory))) {
    await mkdir(directory, { recursive: true });
  }
  await writeTextFile(path, content);
}

export async function writeLocaleFile(
  path: string,
  locale: LocaleObject,
  format: OutputFormat,
  options: FormatLocaleOptions = {}
): Promise<void> {
  const directory = await dirname(path);
  if (!(await exists(directory))) {
    await mkdir(directory, { recursive: true });
  }
  await writeTextFile(path, await formatLocale(locale, format, options));
}

export async function formatLocale(
  locale: LocaleObject,
  format: OutputFormat,
  options: FormatLocaleOptions = {}
): Promise<string> {
  const eol = options.eol ?? "lf";
  const ensureTrailingNewline = options.ensureTrailingNewline ?? true;
  const quoteObjectProperties = options.quoteObjectProperties ?? false;
  const json = JSON.stringify(locale, null, 2);
  const formatted =
    format === "json"
      ? await prettier.format(json, { parser: "json", plugins: [babelPlugin, estreePlugin] })
      : await prettier.format(`export default ${json}${format === "ts" ? " as const" : ""};\n`, {
          parser: format === "ts" ? "typescript" : "babel",
          quoteProps: quoteObjectProperties ? "preserve" : "as-needed",
          plugins: [babelPlugin, estreePlugin, typescriptPlugin]
        });

  return applyLineEndings(formatted, eol, ensureTrailingNewline);
}

export async function formatLocaleForSnapshot(
  locale: LocaleObject,
  format: OutputFormat,
  snapshot: LocaleFileSnapshot,
  options: FormatLocaleOptions = {}
): Promise<string> {
  if (format === "json" || !snapshot.exists || !snapshot.objectRange) {
    return formatLocale(locale, format, options);
  }

  const eol = options.eol ?? snapshot.eol;
  const ensureTrailingNewline = options.ensureTrailingNewline ?? true;
  const objectSource = await formatObjectLiteral(locale, format, options.quoteObjectProperties ?? false);
  const replaced = `${snapshot.content.slice(0, snapshot.objectRange.start)}${objectSource}${snapshot.content.slice(
    snapshot.objectRange.end
  )}`;

  return applyLineEndings(replaced, eol, ensureTrailingNewline);
}

async function formatObjectLiteral(locale: LocaleObject, format: OutputFormat, quoteObjectProperties: boolean): Promise<string> {
  const json = JSON.stringify(locale, null, 2);
  const formatted = await prettier.format(`const __locale = ${json};\n`, {
    parser: format === "ts" ? "typescript" : "babel",
    quoteProps: quoteObjectProperties ? "preserve" : "as-needed",
    plugins: [babelPlugin, estreePlugin, typescriptPlugin]
  });
  return formatted.replace(/^const __locale = /, "").replace(/;\s*$/, "");
}

export function detectEol(content: string): "lf" | "crlf" {
  return content.includes("\r\n") ? "crlf" : "lf";
}

export function applyLineEndings(content: string, eol: "lf" | "crlf", ensureTrailingNewline: boolean): string {
  const newline = eol === "crlf" ? "\r\n" : "\n";
  let normalized = content.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  normalized = ensureTrailingNewline ? `${normalized.replace(/\n*$/, "")}\n` : normalized.replace(/\n+$/, "");
  return normalized.replace(/\n/g, newline);
}

export async function resolveProjectFile(projectRoot: string, relativePath: string): Promise<string> {
  return join(projectRoot, relativePath);
}
