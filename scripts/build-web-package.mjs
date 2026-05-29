import { cp, mkdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(rootDir, "web-package");

await rm(outputDir, { recursive: true, force: true });
await mkdir(outputDir, { recursive: true });
await cp(path.join(rootDir, "dist"), path.join(outputDir, "dist"), { recursive: true });
await cp(path.join(rootDir, "scripts", "serve-web.mjs"), path.join(outputDir, "serve-web.mjs"));
await writeFile(
  path.join(outputDir, "start-web.bat"),
  `@echo off\r\nsetlocal\r\ncd /d "%~dp0"\r\nnode serve-web.mjs\r\npause\r\n`,
  "utf8"
);
await writeFile(
  path.join(outputDir, "start-web.command"),
  `#!/bin/zsh\ncd "$(dirname "$0")"\nnode serve-web.mjs\necho ""\necho "Press any key to close..."\nread -k 1\n`,
  { encoding: "utf8", mode: 0o755 }
);
await writeFile(
  path.join(outputDir, "README.txt"),
  [
    "Windows 双击 start-web.bat 启动局域网 Web 服务。",
    "macOS 双击 start-web.command 启动；如果提示无权限，请在终端执行：chmod +x start-web.command",
    "默认端口为 1420，可先设置 WEB_PORT 环境变量修改端口。",
    "浏览器端读写本地目录需要 Chrome/Edge 的目录权限授权。"
  ].join("\r\n"),
  "utf8"
);

console.log(`Web package created: ${path.relative(rootDir, outputDir)}`);
