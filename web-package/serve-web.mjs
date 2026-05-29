import { createReadStream, existsSync } from "node:fs";
import { stat } from "node:fs/promises";
import { createServer } from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(scriptDir, "..");
const siblingDist = path.join(scriptDir, "dist");
const defaultPackageDist = path.join(rootDir, "web-package", "dist");
const defaultDist = path.join(rootDir, "dist");
const webRoot = path.resolve(
  process.env.WEB_ROOT || (existsSync(siblingDist) ? siblingDist : existsSync(defaultPackageDist) ? defaultPackageDist : defaultDist)
);
const port = Number(process.env.WEB_PORT || process.env.PORT || 1420);
const host = process.env.WEB_HOST || "0.0.0.0";

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url || "/", `http://${request.headers.host || "localhost"}`);
    const filePath = await resolveFilePath(url.pathname);
    response.writeHead(200, {
      "Content-Type": getContentType(filePath),
      "Cache-Control": "no-cache"
    });
    createReadStream(filePath).pipe(response);
  } catch (error) {
    const status = error instanceof HttpError ? error.status : 500;
    const message = status === 404 ? "Not found" : error instanceof Error ? error.message : String(error);
    response.writeHead(status, { "Content-Type": "text/plain;charset=utf-8" });
    response.end(message);
  }
});

server.listen(port, host, () => {
  console.log(`Web root: ${webRoot}`);
  console.log(`Local:   http://127.0.0.1:${port}/`);
  getLanAddresses().forEach((address) => {
    console.log(`LAN:     http://${address}:${port}/`);
  });
});

async function resolveFilePath(urlPath) {
  const decodedPath = decodeURIComponent(urlPath);
  const normalizedPath = path.normalize(decodedPath).replace(/^(\.\.[/\\])+/, "");
  let filePath = path.join(webRoot, normalizedPath);

  if (!filePath.startsWith(webRoot)) {
    throw new HttpError(404);
  }

  if (await isDirectory(filePath)) {
    filePath = path.join(filePath, "index.html");
  }

  if (!existsSync(filePath)) {
    filePath = path.join(webRoot, "index.html");
  }

  if (!existsSync(filePath)) {
    throw new HttpError(404);
  }

  return filePath;
}

async function isDirectory(filePath) {
  try {
    return (await stat(filePath)).isDirectory();
  } catch {
    return false;
  }
}

function getLanAddresses() {
  return Object.values(os.networkInterfaces())
    .flatMap((items) => items ?? [])
    .filter((item) => item.family === "IPv4" && !item.internal)
    .map((item) => item.address);
}

function getContentType(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  return (
    {
      ".css": "text/css;charset=utf-8",
      ".html": "text/html;charset=utf-8",
      ".ico": "image/x-icon",
      ".js": "text/javascript;charset=utf-8",
      ".json": "application/json;charset=utf-8",
      ".map": "application/json;charset=utf-8",
      ".png": "image/png",
      ".svg": "image/svg+xml",
      ".webp": "image/webp"
    }[extension] || "application/octet-stream"
  );
}

class HttpError extends Error {
  constructor(status, message = `HTTP ${status}`) {
    super(message);
    this.status = status;
  }
}
