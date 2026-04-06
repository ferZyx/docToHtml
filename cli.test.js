import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const cliPath = path.join(__dirname, "cli.js");

test("cli --help prints usage", () => {
  const run = spawnSync(process.execPath, [cliPath, "--help"], {
    encoding: "utf-8",
  });

  assert.equal(run.status, 0);
  assert.equal(run.stdout.includes("Usage:"), true);
});

test("cli without args exits with error", () => {
  const run = spawnSync(process.execPath, [cliPath], {
    encoding: "utf-8",
  });

  assert.equal(run.status, 1);
  assert.equal(run.stderr.includes("Usage:"), true);
});
