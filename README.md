# doc-to-html

Convert Microsoft Word documents (`.doc` and `.docx`) to HTML using:

- LibreOffice (`soffice` / `libreoffice`) for `.doc -> .docx`
- Mammoth (default) or Pandoc for `.docx -> .html`

## Installation

From npm (when published):

```bash
npm i doc-to-html
```

From GitHub repository:

```bash
npm i ferZyx/docToHtml
```

or explicit git URL:

```bash
npm i git+https://github.com/ferZyx/docToHtml.git
```

## Runtime

- Node.js `>= 18`
- ESM package (use `import`, not `require`)

## System dependencies

Make sure these tools are available in `PATH`:

- `soffice` or `libreoffice`
- `pandoc` (required only if you use `--tool pandoc`)

## Usage

```js
import { convertWordToHtml } from "doc-to-html";

await convertWordToHtml("/files/input.doc", "/files/output.html");
```

By default, converter uses `mammoth` tool for `.docx -> .html`.

## CLI

After installation:

```bash
doc-to-html /files/input.doc /files/output.html
```

Options:

```bash
doc-to-html --help
doc-to-html --version
doc-to-html input.doc output.html --timeout-ms 180000
doc-to-html input.doc output.html --intermediate-dir /tmp/word --keep-intermediate
doc-to-html input.docx output.html --no-self-contained
doc-to-html input.docx output.html --tool mammoth
doc-to-html input.docx output.html --tool pandoc
```

### Convert `.doc` to `.docx` only

```js
import { docToDocxByLibreOffice } from "doc-to-html";

const docxPath = await docToDocxByLibreOffice("/files/input.doc", "/files/out");
```

### Convert `.docx` to HTML only

```js
import { wordToHtmlByPandoc } from "doc-to-html";

const htmlPath = await wordToHtmlByPandoc("/files/input.docx", "/files/output.html");
```

## API

### `convertWordToHtml(inputPath, outputHtmlPath, options?)`

Converts `.doc` or `.docx` to HTML.

For `.doc` files:
1. converts to `.docx` via LibreOffice,
2. converts resulting `.docx` to HTML via selected tool (`mammoth` by default).

### Options

```ts
interface ConvertWordToHtmlOptions {
  tool?: "mammoth" | "pandoc";
  libreOffice?: {
    args?: string[];
    binary?: string;
    cwd?: string;
    timeoutMs?: number;
    logger?: (message: string) => void;
  };
  pandoc?: {
    args?: string[];
    binary?: string;
    selfContained?: boolean;
    cwd?: string;
    timeoutMs?: number;
    logger?: (message: string) => void;
  };
  mammoth?: {
    styleMap?: string[];
    cwd?: string;
    logger?: (message: string) => void;
    maxDocxBytes?: number;
    maxDocumentXmlBytes?: number;
  };
  keepIntermediateDocx?: boolean;
  intermediateDir?: string;
}
```

## Notes

- `.doc` support depends on LibreOffice availability.
- Conversion result quality depends on Pandoc/LibreOffice versions and source document complexity.
- `timeoutMs` applies to LibreOffice/Pandoc subprocesses.
- For `--tool mammoth` with direct `.docx` input, timeout is not supported.

## Development checks

```bash
npm test
npm run test:integration
```

`test:integration` requires installed Pandoc + LibreOffice and local test files.
