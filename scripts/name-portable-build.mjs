import { copyFile, mkdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageJsonPath = path.join(rootDir, "package.json");
const packageJson = JSON.parse(await readFile(packageJsonPath, "utf8"));

const appName = packageJson.name;
const version = packageJson.version;
const label = process.env.PORTABLE_LABEL || packageJson.portable?.label || "免安装版";
const sourceExe = path.join(rootDir, "src-tauri", "target", "release", `${appName}.exe`);
const outputDir = path.resolve(rootDir, process.env.PORTABLE_OUTPUT_DIR || packageJson.portable?.outputDir || "src-tauri/target/release");
const outputName = `${appName}-v${version}-${sanitizeFileName(label)}.exe`;
const outputExe = path.join(outputDir, outputName);

await mkdir(outputDir, { recursive: true });
await copyFile(sourceExe, outputExe);

console.log(`Portable build copied to: ${path.relative(rootDir, outputExe)}`);

function sanitizeFileName(value) {
  return String(value).trim().replace(/[<>:"/\\|?*\x00-\x1F]/g, "-");
}
