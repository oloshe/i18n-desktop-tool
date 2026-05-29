import { spawnSync } from "node:child_process";

const forwardedArgs = process.argv.slice(2);

if (process.platform !== "darwin") {
  console.error("macOS app bundles must be built on macOS. Please run this command on a Mac.");
  process.exit(1);
}

const portableResult = runTauri(["build", "--bundles", "app", "--no-sign", ...forwardedArgs]);
if (portableResult.status === 0) {
  process.exit(0);
}

console.warn("Building a standalone .app bundle failed. Falling back to the default macOS bundle targets.");
const bundleResult = runTauri(["build", "--no-sign", ...forwardedArgs]);
process.exit(bundleResult.status ?? 1);

function runTauri(args) {
  return spawnSync("npx", ["tauri", ...args], { stdio: "inherit" });
}
