#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { convertWordToHtml } from "./wordToHtmlConverter.js";

function printUsage(stream = process.stdout) {
  stream.write(
    [
      "Usage:",
      "  node-jsdoc-to-html <input.doc|input.docx> <output.html> [options]",
      "",
      "Options:",
      "  --timeout-ms <number>       Timeout for conversion commands",
      "  --intermediate-dir <path>   Directory for temporary .docx files",
      "  --keep-intermediate         Keep intermediate converted .docx",
      "  --no-self-contained         Disable pandoc --self-contained",
      "  -h, --help                  Show this help",
      "  -v, --version               Show package version",
      "",
    ].join("\n"),
  );
}

function parseArgs(argv) {
  if (argv.includes("-h") || argv.includes("--help")) {
    return { mode: "help" };
  }

  if (argv.includes("-v") || argv.includes("--version")) {
    return { mode: "version" };
  }

  if (argv.length < 2) {
    return { mode: "invalid", reason: "Input and output paths are required." };
  }

  const inputPath = argv[0];
  const outputPath = argv[1];

  const options = {
    keepIntermediateDocx: false,
    pandoc: {},
  };

  let i = 2;
  while (i < argv.length) {
    const arg = argv[i];

    if (arg === "--keep-intermediate") {
      options.keepIntermediateDocx = true;
      i += 1;
      continue;
    }

    if (arg === "--no-self-contained") {
      options.pandoc.selfContained = false;
      i += 1;
      continue;
    }

    if (arg === "--timeout-ms") {
      const value = argv[i + 1];
      if (!value) {
        return { mode: "invalid", reason: "--timeout-ms requires value" };
      }
      const timeoutMs = Number.parseInt(value, 10);
      if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
        return { mode: "invalid", reason: "--timeout-ms must be positive number" };
      }
      options.timeoutMs = timeoutMs;
      i += 2;
      continue;
    }

    if (arg === "--intermediate-dir") {
      const value = argv[i + 1];
      if (!value) {
        return { mode: "invalid", reason: "--intermediate-dir requires value" };
      }
      options.intermediateDir = value;
      i += 2;
      continue;
    }

    return { mode: "invalid", reason: `Unknown option: ${arg}` };
  }

  return {
    mode: "convert",
    inputPath,
    outputPath,
    options,
  };
}

async function readVersion() {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);
  const pkgPath = path.join(__dirname, "package.json");
  const pkgRaw = await fs.readFile(pkgPath, "utf-8");
  const pkg = JSON.parse(pkgRaw);
  return pkg.version || "unknown";
}

async function main() {
  const parsed = parseArgs(process.argv.slice(2));

  if (parsed.mode === "help") {
    printUsage();
    process.exit(0);
  }

  if (parsed.mode === "version") {
    const version = await readVersion();
    process.stdout.write(`${version}\n`);
    process.exit(0);
  }

  if (parsed.mode === "invalid") {
    process.stderr.write(`${parsed.reason}\n\n`);
    printUsage(process.stderr);
    process.exit(1);
  }

  try {
    const resultPath = await convertWordToHtml(parsed.inputPath, parsed.outputPath, parsed.options);
    process.stdout.write(`${resultPath}\n`);
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exit(1);
  }
}

main();
