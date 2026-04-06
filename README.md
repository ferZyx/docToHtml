# doc-to-html

Convert Microsoft Word documents (`.doc` and `.docx`) to HTML using:

- LibreOffice (`soffice` / `libreoffice`) for `.doc -> .docx`
- Pandoc for `.docx -> .html`

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

- `pandoc`
- `soffice` or `libreoffice`

## Usage

```js
import { convertWordToHtml } from "doc-to-html";

await convertWordToHtml("/files/input.doc", "/files/output.html");
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
2. converts resulting `.docx` to HTML via Pandoc.

### Options

```ts
interface ConvertWordToHtmlOptions {
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
  keepIntermediateDocx?: boolean;
  intermediateDir?: string;
}
```

## Notes

- `.doc` support depends on LibreOffice availability.
- Conversion result quality depends on Pandoc/LibreOffice versions and source document complexity.

## Development checks

```bash
npm test
npm run test:integration
```

`test:integration` requires installed Pandoc + LibreOffice and local test files.
