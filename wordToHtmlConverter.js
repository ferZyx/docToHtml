import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

import { convertDocxToHtmlByMammoth } from "./wordToHtmlMammoth.js";

function mergeCommandOptions(base, scoped = {}) {
  return {
    ...base,
    ...scoped,
  };
}

function terminateProcess(child) {
  const pid = child.pid;
  if (!pid) {
    return;
  }

  if (process.platform === "win32") {
    spawn("taskkill", ["/pid", String(pid), "/t", "/f"], { stdio: "ignore" });
    return;
  }

  try {
    child.kill("SIGTERM");
  } catch {
    return;
  }

  setTimeout(() => {
    try {
      child.kill("SIGKILL");
    } catch {
      return;
    }
  }, 2_000);
}

async function assertFileExists(filePath) {
  await fs.access(filePath);
}

async function ensureDirectory(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

function basenameWithoutExt(filePath) {
  return path.basename(filePath, path.extname(filePath));
}

function formatCommandError(command, args, stderr, code) {
  const renderedArgs = args.map((arg) => `"${arg}"`).join(" ");
  const details = stderr.trim() ? `\n${stderr.trim()}` : "";
  return `Command failed: ${command} ${renderedArgs} (code: ${code ?? "null"})${details}`;
}

async function runCommand(command, args, options = {}) {
  const { cwd, timeoutMs = 120_000, logger } = options;

  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd, stdio: ["ignore", "pipe", "pipe"] });

    let stdout = "";
    let stderr = "";
    let timedOut = false;

    const timeout = setTimeout(() => {
      timedOut = true;
      terminateProcess(child);
    }, timeoutMs);

    child.stdout.on("data", (chunk) => {
      const msg = chunk.toString();
      stdout += msg;
      if (logger) {
        logger(msg);
      }
    });

    child.stderr.on("data", (chunk) => {
      const msg = chunk.toString();
      stderr += msg;
      if (logger) {
        logger(msg);
      }
    });

    child.on("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });

    child.on("close", (code) => {
      clearTimeout(timeout);

      if (timedOut) {
        reject(new Error(`Command timed out after ${timeoutMs}ms: ${command}`));
        return;
      }

      if (code !== 0) {
        reject(new Error(formatCommandError(command, args, stderr, code)));
        return;
      }

      resolve({ stdout, stderr });
    });
  });
}

function getLibreOfficeCandidates(binary) {
  if (binary) {
    return [binary];
  }

  if (process.platform === "win32") {
    return ["soffice.exe", "soffice", "libreoffice"];
  }

  return ["soffice", "libreoffice"];
}

async function runLibreOfficeWithFallback(args, options = {}) {
  const { binary, ...commandOptions } = options;
  const candidates = getLibreOfficeCandidates(binary);
  let lastError;

  for (const candidate of candidates) {
    try {
      await runCommand(candidate, args, commandOptions);
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(
    `Unable to run LibreOffice (${candidates.join(", ")}). ${String(lastError instanceof Error ? lastError.message : lastError)}`,
  );
}

export async function docToDocxByLibreOffice(docPath, outputDir, options = {}) {
  await assertFileExists(docPath);
  await ensureDirectory(outputDir);

  const args = [
    "--headless",
    "--convert-to",
    "docx",
    "--outdir",
    outputDir,
    docPath,
    ...(options.args || []),
  ];

  await runLibreOfficeWithFallback(args, options);

  const outputPath = path.join(outputDir, `${basenameWithoutExt(docPath)}.docx`);
  await assertFileExists(outputPath);
  return outputPath;
}

export async function wordToHtmlByPandoc(wordPath, htmlPath, options = {}) {
  await assertFileExists(wordPath);
  await ensureDirectory(path.dirname(htmlPath));

  const binary = options.binary || "pandoc";
  const selfContained = options.selfContained ?? true;
  const args = [
    "--from",
    "docx",
    "--to",
    "html5",
    wordPath,
    "-o",
    htmlPath,
    ...(selfContained ? ["--self-contained"] : []),
    ...(options.args || []),
  ];

  await runCommand(binary, args, options);
  await assertFileExists(htmlPath);
  return htmlPath;
}

export async function wordToHtmlByMammoth(wordPath, htmlPath, options = {}) {
  return convertDocxToHtmlByMammoth(wordPath, htmlPath, options);
}

function resolveTool(tool) {
  const selected = (tool || "mammoth").toLowerCase();
  if (selected !== "mammoth" && selected !== "pandoc") {
    throw new Error(`Unsupported tool: ${selected}. Supported tools: mammoth, pandoc`);
  }

  return selected;
}

export async function convertWordToHtml(inputPath, outputHtmlPath, options = {}) {
  await assertFileExists(inputPath);

  const ext = path.extname(inputPath).toLowerCase();
  if (ext !== ".doc" && ext !== ".docx") {
    throw new Error(`Unsupported file extension: ${ext}. Only .doc and .docx are supported.`);
  }

  const baseCommandOptions = {
    cwd: options.cwd,
    timeoutMs: options.timeoutMs,
    logger: options.logger,
  };

  const tempRoot = options.intermediateDir || (await fs.mkdtemp(path.join(os.tmpdir(), "word-to-html-")));
  const shouldCleanupTemp = !options.intermediateDir;
  const selectedTool = resolveTool(options.tool);

  if (selectedTool === "mammoth" && ext === ".docx" && options.timeoutMs) {
    throw new Error(
      "timeoutMs is not supported with tool 'mammoth' for .docx input. Use tool 'pandoc' for timeout-controlled .docx conversion.",
    );
  }

  try {
    const sourceDocx =
      ext === ".doc"
        ? await docToDocxByLibreOffice(
            inputPath,
            tempRoot,
            mergeCommandOptions(baseCommandOptions, options.libreOffice),
          )
        : inputPath;

    if (selectedTool === "pandoc") {
      return await wordToHtmlByPandoc(
        sourceDocx,
        outputHtmlPath,
        mergeCommandOptions(baseCommandOptions, options.pandoc),
      );
    }

    return await wordToHtmlByMammoth(
      sourceDocx,
      outputHtmlPath,
      {
        logger: baseCommandOptions.logger,
        ...(options.mammoth || {}),
      },
    );
  } finally {
    if (shouldCleanupTemp && !options.keepIntermediateDocx) {
      await fs.rm(tempRoot, { recursive: true, force: true });
    }
  }
}
