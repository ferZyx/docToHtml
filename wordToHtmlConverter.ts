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

export interface MammothOptions {
  styleMap?: string[];
  logger?: (message: string) => void;
  maxDocxBytes?: number;
  maxDocumentXmlBytes?: number;
}

export interface ConvertWordToHtmlOptions extends CommandOptions {
  tool?: "mammoth" | "pandoc";
  libreOffice?: LibreOfficeOptions;
  pandoc?: PandocOptions;
  mammoth?: MammothOptions;
  keepIntermediateDocx?: boolean;
  intermediateDir?: string;
}

export {
  convertWordToHtml,
  docToDocxByLibreOffice,
  wordToHtmlByMammoth,
  wordToHtmlByPandoc,
} from "./wordToHtmlConverter.js";
