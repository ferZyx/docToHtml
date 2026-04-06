export interface CommandOptions {
  cwd?: string;
  timeoutMs?: number;
  logger?: (message: string) => void;
}

export interface LibreOfficeOptions extends CommandOptions {
  args?: string[];
  binary?: string;
}

export interface PandocOptions extends CommandOptions {
  args?: string[];
  binary?: string;
  selfContained?: boolean;
}

export interface ConvertWordToHtmlOptions extends CommandOptions {
  libreOffice?: LibreOfficeOptions;
  pandoc?: PandocOptions;
  keepIntermediateDocx?: boolean;
  intermediateDir?: string;
}

export function docToDocxByLibreOffice(
  docPath: string,
  outputDir: string,
  options?: LibreOfficeOptions,
): Promise<string>;

export function wordToHtmlByPandoc(
  wordPath: string,
  htmlPath: string,
  options?: PandocOptions,
): Promise<string>;

export function convertWordToHtml(
  inputPath: string,
  outputHtmlPath: string,
  options?: ConvertWordToHtmlOptions,
): Promise<string>;
