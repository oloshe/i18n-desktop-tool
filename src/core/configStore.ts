import type { ImportSettings, SavedProjectConfig } from "./types";

const STORAGE_KEY = "i18n-desktop-tool.configs";

export function loadProjectConfigs(): SavedProjectConfig[] {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveProjectConfig(projectName: string, settings: ImportSettings, existingId?: string): SavedProjectConfig {
  const configs = loadProjectConfigs();
  const now = new Date().toISOString();
  const config: SavedProjectConfig = {
    id: existingId ?? crypto.randomUUID(),
    projectName: projectName.trim() || "未命名项目",
    updatedAt: now,
    ...settings
  };

  const next = configs.some((item) => item.id === config.id)
    ? configs.map((item) => (item.id === config.id ? config : item))
    : [config, ...configs];

  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  return config;
}

export function deleteProjectConfig(id: string): void {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify(loadProjectConfigs().filter((config) => config.id !== id))
  );
}

export function serializeProjectConfig(config: SavedProjectConfig): string {
  return JSON.stringify({ version: 1, config }, null, 2);
}

export function parseProjectConfigExport(content: string): SavedProjectConfig {
  const parsed = JSON.parse(content);
  const config = parsed?.config ?? parsed;
  if (!config || typeof config !== "object") {
    throw new Error("配置文件格式不正确。");
  }
  if (!config.projectName || !config.projectRoot || !config.keyColumn) {
    throw new Error("配置文件缺少必要字段。");
  }
  return config as SavedProjectConfig;
}
