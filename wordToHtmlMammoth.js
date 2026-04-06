import fs from "node:fs/promises";
import path from "node:path";

import mammoth from "mammoth";
import sanitizeHtml from "sanitize-html";
import { load } from "cheerio";
import JSZip from "jszip";

const ALLOWED_TAGS = [
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "p",
  "blockquote",
  "ul",
  "ol",
  "li",
  "table",
  "thead",
  "tbody",
  "tfoot",
  "tr",
  "th",
  "td",
  "a",
  "strong",
  "em",
  "u",
  "s",
  "span",
  "br",
  "img",
  "figure",
  "figcaption",
  "hr",
  "code",
  "pre",
  "sub",
  "sup",
];

const DEFAULT_MAX_DOCX_BYTES = 25 * 1024 * 1024;
const DEFAULT_MAX_DOCUMENT_XML_BYTES = 8 * 1024 * 1024;

function isAllowedImgSrc(src) {
  if (!src) {
    return false;
  }

  if (/^https?:\/\//i.test(src)) {
    return true;
  }

  return /^data:image\/(png|jpeg|jpg|gif|webp);base64,[a-z0-9+/=\s]+$/i.test(src);
}

async function assertDocxSize(docxPath, maxDocxBytes = DEFAULT_MAX_DOCX_BYTES) {
  const stat = await fs.stat(docxPath);
  if (stat.size > maxDocxBytes) {
    throw new Error(`DOCX is too large (${stat.size} bytes). Max allowed: ${maxDocxBytes} bytes`);
  }
}

function mapWordAlignmentToCss(value) {
  if (value === "center") {
    return "center";
  }

  if (value === "right" || value === "end") {
    return "right";
  }

  return null;
}

function toNumber(value) {
  if (!value) {
    return null;
  }

  const num = Number.parseInt(value, 10);
  return Number.isFinite(num) ? num : null;
}

function twipsToPx(twips) {
  if (!twips || twips === 0) {
    return null;
  }

  return Math.round((twips / 15) * 10) / 10;
}

function appendStyle($node, styleRule) {
  if (!styleRule) {
    return;
  }

  const current = $node.attr("style") || "";
  $node.attr("style", `${current}${styleRule}`);
}

function sanitizeConvertedHtml(html) {
  return sanitizeHtml(html, {
    allowedTags: ALLOWED_TAGS,
    allowedAttributes: {
      "*": ["style"],
      a: ["href", "target", "rel"],
      img: ["src", "alt", "width", "height"],
      td: ["colspan", "rowspan"],
      th: ["colspan", "rowspan"],
    },
    allowedSchemes: ["http", "https", "mailto"],
    allowedSchemesByTag: {
      img: ["http", "https", "data"],
    },
    transformTags: {
      img: (tagName, attribs) => {
        if (!isAllowedImgSrc(attribs.src)) {
          delete attribs.src;
        }

        return { tagName, attribs };
      },
    },
    disallowedTagsMode: "discard",
  });
}

function normalizeHtml(html) {
  const $ = load(html, { decodeEntities: false });

  $("script,style,meta,link").remove();

  $("a").each((_, a) => {
    const href = ($(a).attr("href") || "").trim();
    if (!href || href.startsWith("javascript:")) {
      $(a).replaceWith($(a).text());
      return;
    }

    $(a).attr("target", "_blank");
    $(a).attr("rel", "noopener noreferrer");
  });

  $("table").each((_, table) => {
    const $table = $(table);
    if (!$table.children("thead").length) {
      const firstRow = $table.find("tr").first();
      if (firstRow.length > 0) {
        const head = $("<thead></thead>");
        head.append(firstRow.clone());
        firstRow.remove();
        $table.prepend(head);
      }
    }

    if (!$table.children("tbody").length) {
      const body = $("<tbody></tbody>");
      $table.children("tr").each((__, row) => {
        body.append(row);
      });
      $table.append(body);
    }
  });

  const bodyHtml = $("body").html();
  return (bodyHtml || $.root().html() || "").trim();
}

function parseParagraphHint(paragraphXml) {
  const alignMatch = paragraphXml.match(/<w:jc\b[^>]*w:val="([^"]+)"/);
  const firstLineMatch = paragraphXml.match(/<w:ind\b[^>]*w:firstLine="([^"]+)"/);
  const hangingMatch = paragraphXml.match(/<w:ind\b[^>]*w:hanging="([^"]+)"/);
  const leftMatch = paragraphXml.match(/<w:ind\b[^>]*w:left="([^"]+)"/);
  const rightMatch = paragraphXml.match(/<w:ind\b[^>]*w:right="([^"]+)"/);

  const firstLineTwips = toNumber(firstLineMatch?.[1]);
  const hangingTwips = toNumber(hangingMatch?.[1]);

  return {
    align: alignMatch ? mapWordAlignmentToCss(alignMatch[1]) : null,
    hasTab: /<w:tab\b/.test(paragraphXml),
    firstLineTwips:
      firstLineTwips || hangingTwips
        ? (firstLineTwips || 0) - (hangingTwips || 0)
        : null,
    leftTwips: toNumber(leftMatch?.[1]),
    rightTwips: toNumber(rightMatch?.[1]),
  };
}

async function getDocxParagraphHints(docxPath, maxDocumentXmlBytes = DEFAULT_MAX_DOCUMENT_XML_BYTES) {
  const raw = await fs.readFile(docxPath);
  const zip = await JSZip.loadAsync(raw);
  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    return [];
  }

  const uncompressedSize = documentXmlFile?._data?.uncompressedSize;
  if (typeof uncompressedSize === "number" && uncompressedSize > maxDocumentXmlBytes) {
    throw new Error(
      `DOCX document.xml is too large (${uncompressedSize} bytes). Max allowed: ${maxDocumentXmlBytes} bytes`,
    );
  }

  const documentXml = await documentXmlFile.async("string");
  const paragraphRegex = /<w:p\b[\s\S]*?<\/w:p>/g;
  const paragraphs = documentXml.match(paragraphRegex) || [];
  return paragraphs.map(parseParagraphHint);
}

function applyTabSplit($p) {
  const html = $p.html() || "";
  if (!html.includes("\t")) {
    return;
  }

  const parts = html.split(/\t+/);
  if (parts.length < 2) {
    return;
  }

  const left = parts.shift()?.trim() || "";
  const right = parts.join(" ").trim();

  if (!left || !right) {
    return;
  }

  if (left.length > 120 || right.length > 120) {
    return;
  }

  appendStyle($p, "display:flex;justify-content:space-between;align-items:flex-end;gap:16px;");
  $p.html(`<span>${left}</span><span>${right}</span>`);
}

function applyLayoutHintsToHtml(html, paragraphHints) {
  if (!paragraphHints.length) {
    return html;
  }

  const $ = load(html, { decodeEntities: false });
  const paragraphs = $("p").toArray();

  paragraphs.forEach((paragraph, index) => {
    const hint = paragraphHints[index];
    if (!hint) {
      return;
    }

    const $p = $(paragraph);
    const inProtectedContainer = $p.parents("li,td,th,pre,code").length > 0;
    const plainText = $p.text().replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();

    const shouldApplyAlign =
      hint.align &&
      (hint.align === "center" || hint.align === "right") &&
      plainText.length > 0 &&
      plainText.length <= 180 &&
      !inProtectedContainer;

    if (shouldApplyAlign) {
      appendStyle($p, `text-align:${hint.align};`);
    }

    if (!inProtectedContainer) {
      const firstLinePx = twipsToPx(hint.firstLineTwips);
      const leftPx = twipsToPx(hint.leftTwips);
      const rightPx = twipsToPx(hint.rightTwips);

      appendStyle($p, firstLinePx ? `text-indent:${firstLinePx}px;` : "");
      appendStyle($p, leftPx ? `margin-left:${leftPx}px;` : "");
      appendStyle($p, rightPx ? `margin-right:${rightPx}px;` : "");
    }

    if (hint.hasTab && !inProtectedContainer) {
      applyTabSplit($p);
    }
  });

  const bodyHtml = $("body").html();
  return (bodyHtml || $.root().html() || "").trim();
}

export async function convertDocxToHtmlByMammoth(docxPath, htmlPath, options = {}) {
  await fs.access(docxPath);
  await assertDocxSize(docxPath, options.maxDocxBytes);
  await fs.mkdir(path.dirname(htmlPath), { recursive: true });

  if (options.timeoutMs) {
    throw new Error("timeoutMs is not supported for mammoth conversion. Use tool 'pandoc' if timeout control is required.");
  }

  const { value, messages } = await mammoth.convertToHtml(
    { path: docxPath },
    {
      includeDefaultStyleMap: true,
      styleMap: options.styleMap || [],
    },
  );

  if (!value || !value.trim()) {
    throw new Error("Mammoth conversion returned empty HTML");
  }

  const sanitized = sanitizeConvertedHtml(value);
  const normalized = normalizeHtml(sanitized);
  const hints = await getDocxParagraphHints(docxPath, options.maxDocumentXmlBytes);
  const resultHtml = applyLayoutHintsToHtml(normalized, hints);

  await fs.writeFile(htmlPath, resultHtml, "utf-8");

  if (options.logger && messages.length > 0) {
    for (const message of messages) {
      options.logger(`[mammoth] ${message.message}`);
    }
  }

  return htmlPath;
}
