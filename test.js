import fs from "node:fs/promises";
import path from "node:path";

import { convertWordToHtml } from "./wordToHtmlConverter.js";

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function run() {
  const inputPath = path.join("tests", "word", "main.doc");
  const outputPath = path.join("tests", "html", "main.html");

  await fs.mkdir(path.dirname(outputPath), { recursive: true });

  try {
    await convertWordToHtml(inputPath, outputPath, {
      libreOffice: { timeoutMs: 180_000 },
      pandoc: { timeoutMs: 180_000 },
    });
  } catch (error) {
    console.error(String(error instanceof Error ? error.message : error));
    process.exit(1);
  }

  const outputExists = await exists(outputPath);
  if (!outputExists) {
    console.error("HTML output file was not created");
    process.exit(1);
  }

  console.log(`OK: ${outputPath}`);
}

run();
