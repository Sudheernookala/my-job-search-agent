// Copies ../data/jobs.json into public/ so the UI can load it.
// If the tracker has not run yet, writes an empty dataset instead of failing.
import { copyFileSync, existsSync, mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const src = resolve(here, "../../data/jobs.json");
const dest = resolve(here, "../public/jobs.json");

mkdirSync(dirname(dest), { recursive: true });
if (existsSync(src)) {
  copyFileSync(src, dest);
  console.log("Copied data/jobs.json -> public/jobs.json");
} else {
  writeFileSync(dest, JSON.stringify({ generated_at: null, retention_days: 30, jobs: [] }));
  console.log("data/jobs.json not found - wrote empty dataset");
}
