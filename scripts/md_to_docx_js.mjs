#!/usr/bin/env node
/**
 * Self-contained Markdown to DOCX converter.
 *
 * Converts Markdown files to professional Word documents without Pandoc.
 * Uses `marked` for Markdown parsing and `docx` npm for DOCX generation.
 *
 * Usage:
 *     node md_to_docx_js.mjs input.md output.docx [--template template.docx]
 *
 * Dependencies (npm install -g):
 *     marked, docx, adm-zip
 */

import { readFileSync, writeFileSync, existsSync, mkdirSync } from "fs";
import { tmpdir } from "os";
import { join, resolve, dirname, extname, isAbsolute } from "path";
import { execSync } from "child_process";
import { Buffer } from "buffer";

// ---------------------------------------------------------------------------
// Dynamic imports for npm packages
// ---------------------------------------------------------------------------
let marked, docx, AdmZip;

try {
  marked = await import("marked");
} catch {
  console.log("Installing marked...");
  execSync("npm install -g marked", { stdio: "inherit" });
  marked = await import("marked");
}

try {
  docx = await import("docx");
} catch {
  console.log("Installing docx...");
  execSync("npm install -g docx", { stdio: "inherit" });
  docx = await import("docx");
}

try {
  AdmZip = (await import("adm-zip")).default;
} catch {
  console.log("Installing adm-zip...");
  execSync("npm install -g adm-zip", { stdio: "inherit" });
  AdmZip = (await import("adm-zip")).default;
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, LevelFormat,
  ExternalHyperlink, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak, PageNumber,
  TableOfContents, SectionType,
} = docx;

// ---------------------------------------------------------------------------
// Style defaults & configuration
// ---------------------------------------------------------------------------
const STYLE_DEFAULTS = {
  font_body: "Arial",
  font_heading: "Arial",
  font_code: "Consolas",
  font_size: "10.5",
  color_heading: "2D3B4D",
  color_body: "333333",
  table_header_bg: "D5E8F0",
  table_header_text: "2D3B4D",
  table_alt_row: "F2F2F2",
  table_border: "CCCCCC",
  table_border_size: "4",
  table_cell_margin: "28",
  table_font_size: "9.5",
  table_banded_rows: "true",
  code_bg: "F5F5F5",
  code_font_size: "9",
};

// Module-level style variables (mutable, set by applyStyleConfig)
let FONT_BODY = "Arial";
let FONT_HEADING = "Arial";
let FONT_CODE = "Consolas";
let FONT_SIZE_BODY = 21;    // half-points (10.5pt * 2)
let FONT_SIZE_TABLE = 19;   // half-points (9.5pt * 2)
let FONT_SIZE_CODE = 18;    // half-points (9pt * 2)
let COLOR_HEADING = "2D3B4D";
let COLOR_BODY = "333333";
let COLOR_LINK = "0563C1";
let TBL_HDR_BG = "D5E8F0";
let TBL_HDR_TEXT = "2D3B4D";
let TBL_ALT_ROW = "F2F2F2";
let TBL_BORDER = "CCCCCC";
let TBL_BORDER_SIZE = 4;    // eighth-points (OOXML w:sz unit)
let TBL_CELL_MARGIN = 28;   // twips
let TBL_BANDED_ROWS = true;
let CODE_BG = "F5F5F5";

const HEADING_SIZES = { 1: 40, 2: 32, 3: 28, 4: 24, 5: 22, 6: 21 }; // half-points
const PAGE_WIDTH_DXA = 9360; // usable width with 1" margins on letter

const DOCX_STYLE_RE = /<!--\s*docx-style\s*\n([\s\S]*?)-->/;

function parseDocxStyle(text) {
  const match = DOCX_STYLE_RE.exec(text);
  if (!match) return {};
  const config = {};
  for (const line of match[1].trim().split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf(":");
    if (idx === -1) continue;
    const key = trimmed.slice(0, idx).trim().toLowerCase();
    const value = trimmed.slice(idx + 1).trim().replace(/^["']|["']$/g, "");
    if (key in STYLE_DEFAULTS) config[key] = value;
  }
  return config;
}

function stripDocxStyleComment(text) {
  return text.replace(DOCX_STYLE_RE, "");
}

function parseStyleFile(filePath) {
  const config = {};
  const text = readFileSync(filePath, "utf-8");
  for (const line of text.trim().split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf(":");
    if (idx === -1) continue;
    const key = trimmed.slice(0, idx).trim().toLowerCase();
    const value = trimmed.slice(idx + 1).trim().replace(/^["']|["']$/g, "");
    if (key in STYLE_DEFAULTS) config[key] = value;
  }
  return config;
}

/**
 * Ensure template is usable by converting .dotx/.dotm to .docx if needed.
 *
 * Word .dotx templates are structurally identical to .docx files except for
 * a single content-type string in [Content_Types].xml.  This function patches
 * that string using adm-zip and writes a temporary .docx file.
 *
 * @param {string} templatePath - Path to the template file (.docx or .dotx)
 * @returns {string} Path to a usable .docx (original if already .docx, temp path if converted)
 */
function ensureDocx(templatePath) {
  const ext = templatePath.toLowerCase();
  if (!ext.endsWith(".dotx") && !ext.endsWith(".dotm")) {
    return templatePath;
  }

  const zip = new AdmZip(templatePath);
  const ctEntry = zip.getEntry("[Content_Types].xml");
  if (ctEntry) {
    let ct = ctEntry.getData().toString("utf-8");
    ct = ct.replace(
      "wordprocessingml.template.main+xml",
      "wordprocessingml.document.main+xml"
    );
    ct = ct.replace(
      "word.template.macroEnabledTemplate.main+xml",
      "word.document.macroEnabled.main+xml"
    );
    zip.updateFile("[Content_Types].xml", Buffer.from(ct, "utf-8"));
  }

  const tmpPath = templatePath.replace(/\.dot[xm]$/i, ".tmp.docx");
  zip.writeZip(tmpPath);
  console.log(`  Converted .dotx template to .docx: ${tmpPath}`);
  return tmpPath;
}


/**
 * Extract style configuration from a DOCX/DOTX template's actual formatting.
 *
 * Reads docDefaults, heading styles, table styles, and theme colors from
 * the template's XML to produce an object compatible with STYLE_DEFAULTS keys.
 * The result can be used as a base style layer that is then overridden by
 * --style / inline / CLI flags.
 *
 * Since the `docx` npm library can only CREATE new documents (not open existing
 * DOCX files), we read the DOCX as a ZIP archive and parse the XML with regex.
 *
 * @param {string} templatePath - Path to the DOCX/DOTX template file
 * @returns {Object} Style configuration dict compatible with STYLE_DEFAULTS keys
 */
function extractStyleFromTemplate(templatePath) {
  const config = {};

  // Convert .dotx/.dotm to .docx-compatible ZIP before reading
  const effectivePath = ensureDocx(templatePath);

  let zip;
  try {
    zip = new AdmZip(effectivePath);
  } catch (e) {
    console.log(`  Warning: could not open template as ZIP: ${e.message}`);
    return config;
  }

  // Read XML entries as strings
  let stylesXml = "";
  let themeXml = "";
  try {
    const stylesEntry = zip.getEntry("word/styles.xml");
    if (stylesEntry) stylesXml = stylesEntry.getData().toString("utf-8");
  } catch {}
  try {
    const themeEntry = zip.getEntry("word/theme/theme1.xml");
    if (themeEntry) themeXml = themeEntry.getData().toString("utf-8");
  } catch {}

  if (!stylesXml) return config;

  // --- Helper: extract attribute value from an XML fragment ---
  // Handles both w:val="X" and w:ascii="X" patterns with full namespace URIs
  function getAttr(xml, attrLocalName) {
    // Match ns:attrName="value" where ns prefix can vary, or no namespace
    const re = new RegExp(`(?:^|\\s)(?:\\w+:)?${attrLocalName}\\s*=\\s*"([^"]*)"`, "i");
    const m = re.exec(xml);
    return m ? m[1] : null;
  }

  // --- 1. Document defaults (fonts, body size, body color) ---
  // Find the docDefaults block
  const docDefaultsMatch = stylesXml.match(/<[^<]*:?docDefaults[^>]*>([\s\S]*?)<\/[^<]*:?docDefaults>/);
  if (docDefaultsMatch) {
    const docDefaults = docDefaultsMatch[1];

    // Find rPrDefault > rPr block
    const rPrDefaultMatch = docDefaults.match(/<[^<]*:?rPrDefault[^>]*>([\s\S]*?)<\/[^<]*:?rPrDefault>/);
    if (rPrDefaultMatch) {
      const rPrDefault = rPrDefaultMatch[1];

      // Find rPr block within rPrDefault
      const rPrMatch = rPrDefault.match(/<[^<]*:?rPr[^>]*>([\s\S]*?)<\/[^<]*:?rPr>/);
      if (rPrMatch) {
        const rPr = rPrMatch[1];

        // Font: <w:rFonts w:ascii="FontName" .../>
        const rFontsMatch = rPr.match(/<[^<]*:?rFonts\s[^>]*>/);
        if (rFontsMatch) {
          const asciiFont = getAttr(rFontsMatch[0], "ascii");
          if (asciiFont) {
            config.font_body = asciiFont;
            config.font_heading = asciiFont; // may be overridden below
          }
        }

        // Size: <w:sz w:val="HALFPTS"/>
        const szMatch = rPr.match(/<[^<]*:?sz\s[^>]*>/);
        if (szMatch) {
          const halfPts = getAttr(szMatch[0], "val");
          if (halfPts) {
            config.font_size = String(parseInt(halfPts, 10) / 2);
          }
        }

        // Body color from defaults: <w:color w:val="HEXCOLOR"/>
        const colorMatch = rPr.match(/<[^<]*:?color\s[^>]*>/);
        if (colorMatch) {
          const val = getAttr(colorMatch[0], "val");
          if (val && val.toLowerCase() !== "auto") {
            config.color_body = val;
          }
        }
      }
    }
  }

  // --- 2. Normal style (body color fallback) ---
  // Find <w:style w:type="paragraph" w:styleId="Normal"> block
  const normalStyleRe = /<[^<]*:?style\s[^>]*styleId\s*=\s*"Normal"[^>]*>([\s\S]*?)<\/[^<]*:?style>/;
  const normalMatch = stylesXml.match(normalStyleRe);
  if (normalMatch) {
    const normalBlock = normalMatch[1];
    // Look for rPr > color within this style
    const rPrMatch = normalBlock.match(/<[^<]*:?rPr[^>]*>([\s\S]*?)<\/[^<]*:?rPr>/);
    if (rPrMatch) {
      const colorMatch = rPrMatch[1].match(/<[^<]*:?color\s[^>]*>/);
      if (colorMatch) {
        const val = getAttr(colorMatch[0], "val");
        if (val && val.toLowerCase() !== "auto") {
          config.color_body = val;
        }
      }
    }
  }

  // --- 3. Theme colors (fallback for headings / body) ---
  let themeDk1 = null;
  if (themeXml) {
    // Find <a:dk1> element within clrScheme
    const dk1Match = themeXml.match(/<[^<]*:?dk1[^>]*>([\s\S]*?)<\/[^<]*:?dk1>/);
    if (dk1Match) {
      const dk1Content = dk1Match[1];
      // Try sysClr lastClr attribute first
      const sysClrMatch = dk1Content.match(/<[^<]*:?sysClr\s[^>]*>/);
      if (sysClrMatch) {
        const lastClr = getAttr(sysClrMatch[0], "lastClr");
        if (lastClr) themeDk1 = lastClr;
      }
      // Try srgbClr val attribute
      if (!themeDk1) {
        const srgbMatch = dk1Content.match(/<[^<]*:?srgbClr\s[^>]*>/);
        if (srgbMatch) {
          const val = getAttr(srgbMatch[0], "val");
          if (val) themeDk1 = val;
        }
      }
    }
  }

  // Apply theme fallbacks where no explicit color was found
  if (!config.color_body) {
    config.color_body = themeDk1 || "000000";
  }
  if (!config.color_heading) {
    config.color_heading = themeDk1 || "000000";
  }

  // --- 4. Heading styles (font, color override) ---
  // Check Heading 1, 2, 3 – use first available
  for (const level of [1, 2, 3]) {
    const headingRe = new RegExp(
      `<[^<]*:?style\\s[^>]*styleId\\s*=\\s*"Heading${level}"[^>]*>([\\s\\S]*?)<\\/[^<]*:?style>`
    );
    const headingMatch = stylesXml.match(headingRe);
    if (!headingMatch) continue;

    const headingBlock = headingMatch[1];
    const rPrMatch = headingBlock.match(/<[^<]*:?rPr[^>]*>([\s\S]*?)<\/[^<]*:?rPr>/);
    if (!rPrMatch) continue;

    const rPr = rPrMatch[1];

    // Explicit heading font
    const rFontsMatch = rPr.match(/<[^<]*:?rFonts\s[^>]*>/);
    if (rFontsMatch) {
      const asciiFont = getAttr(rFontsMatch[0], "ascii");
      if (asciiFont) {
        config.font_heading = asciiFont;
      }
    }

    // Explicit heading color
    const colorMatch = rPr.match(/<[^<]*:?color\s[^>]*>/);
    if (colorMatch) {
      const val = getAttr(colorMatch[0], "val");
      if (val && val.toLowerCase() !== "auto") {
        config.color_heading = val;
      }
    }

    break; // first available heading is enough
  }

  // --- 5. Table styles (header bg, header text color) ---
  // Find all table styles and look for firstRow tblStylePr
  const tableStyleRe = /<[^<]*:?style\s[^>]*type\s*=\s*"table"[^>]*>([\s\S]*?)<\/[^<]*:?style>/gi;
  let tableStyleMatch;
  let foundTableHeader = false;
  while ((tableStyleMatch = tableStyleRe.exec(stylesXml)) !== null && !foundTableHeader) {
    const tableBlock = tableStyleMatch[1];

    // Find tblStylePr elements with type="firstRow"
    const tblStylePrRe = /<[^<]*:?tblStylePr\s[^>]*type\s*=\s*"firstRow"[^>]*>([\s\S]*?)<\/[^<]*:?tblStylePr>/g;
    let tspMatch;
    while ((tspMatch = tblStylePrRe.exec(tableBlock)) !== null) {
      const tspContent = tspMatch[1];

      // Header background: <w:shd ... w:fill="HEXCOLOR"/>
      const shdMatch = tspContent.match(/<[^<]*:?shd\s[^>]*>/);
      if (shdMatch) {
        const fill = getAttr(shdMatch[0], "fill");
        if (fill && fill.toLowerCase() !== "auto" && fill.toLowerCase() !== "ffffff") {
          config.table_header_bg = fill;
        }
      }

      // Header text color: within rPr > color
      const rPrMatch = tspContent.match(/<[^<]*:?rPr[^>]*>([\s\S]*?)<\/[^<]*:?rPr>/);
      if (rPrMatch) {
        const colorMatch = rPrMatch[1].match(/<[^<]*:?color\s[^>]*>/);
        if (colorMatch) {
          const val = getAttr(colorMatch[0], "val");
          if (val) {
            config.table_header_text = val;
          }
        }
      }

      if (config.table_header_bg) {
        foundTableHeader = true;
        break;
      }
    }
  }

  // --- 6. Table paragraph style (font size) ---
  for (const styleName of ["Table1", "Table Contents", "Table Text"]) {
    const tsRe = new RegExp(
      `<[^<]*:?style\\s[^>]*styleId\\s*=\\s*"${styleName.replace(/\s/g, "")}"[^>]*>([\\s\\S]*?)<\\/[^<]*:?style>`
    );
    // Also try with space-delimited name in w:name attribute
    const tsNameRe = new RegExp(
      `<[^<]*:?style\\s[^>]*>([\\s\\S]*?)<\\/[^<]*:?style>`,
      "g"
    );

    // Try by styleId first (no spaces)
    let tsMatch = stylesXml.match(tsRe);

    // If not found by styleId, search by w:name attribute
    if (!tsMatch) {
      let candidate;
      while ((candidate = tsNameRe.exec(stylesXml)) !== null) {
        const fullTag = stylesXml.slice(candidate.index, candidate.index + 200);
        // Check if the opening tag has name="Table1" or name="Table Contents" etc.
        const nameAttrMatch = fullTag.match(/name\s*=\s*"([^"]*)"/i);
        if (nameAttrMatch && nameAttrMatch[1] === styleName) {
          tsMatch = candidate;
          break;
        }
      }
    }

    if (!tsMatch) continue;
    const tsBlock = tsMatch[1];

    // Find rPr > sz within this style
    const szMatch = tsBlock.match(/<[^<]*:?sz\s[^>]*>/);
    if (szMatch) {
      const halfPts = getAttr(szMatch[0], "val");
      if (halfPts) {
        config.table_font_size = String(parseInt(halfPts, 10) / 2);
        break;
      }
    }
  }

  return config;
}

function applyStyleConfig(overrides = {}) {
  const cfg = { ...STYLE_DEFAULTS, ...overrides };
  FONT_BODY = cfg.font_body;
  FONT_HEADING = cfg.font_heading;
  FONT_CODE = cfg.font_code;
  FONT_SIZE_BODY = Math.round(parseFloat(cfg.font_size) * 2);
  FONT_SIZE_TABLE = Math.round(parseFloat(cfg.table_font_size) * 2);
  FONT_SIZE_CODE = Math.round(parseFloat(cfg.code_font_size) * 2);
  COLOR_HEADING = cfg.color_heading;
  COLOR_BODY = cfg.color_body;
  TBL_HDR_BG = cfg.table_header_bg;
  TBL_HDR_TEXT = cfg.table_header_text;
  TBL_ALT_ROW = cfg.table_alt_row;
  TBL_BORDER = cfg.table_border;
  TBL_BORDER_SIZE = parseInt(cfg.table_border_size, 10);
  TBL_CELL_MARGIN = parseInt(cfg.table_cell_margin, 10);
  TBL_BANDED_ROWS = ["true", "yes", "1"].includes(cfg.table_banded_rows.toLowerCase());
  CODE_BG = cfg.code_bg;
}

// ---------------------------------------------------------------------------
// Mermaid helpers
// ---------------------------------------------------------------------------
function generateMermaidUrl(code) {
  try {
    const encoded = Buffer.from(code, "utf-8")
      .toString("base64url")
      .replace(/=+$/, "");
    const url = `https://mermaid.ink/img/${encoded}`;
    return url.length <= 2000 ? url : null;
  } catch {
    return null;
  }
}

function fetchUrl(url) {
  // Use curl for proper binary output (powershell .Content corrupts binary)
  try {
    const result = execSync(
      `curl -sL -o - "${url}"`,
      { maxBuffer: 10 * 1024 * 1024, timeout: 30000, encoding: "buffer" }
    );
    return result;
  } catch {
    // Fallback: powershell with byte-safe download to temp file
    try {
      const tmp = join(tmpdir(), `fetch_${Date.now()}.bin`);
      execSync(
        `powershell -Command "Invoke-WebRequest -Uri '${url}' -UseBasicParsing -OutFile '${tmp}'"`,
        { timeout: 30000, stdio: "pipe" }
      );
      if (existsSync(tmp)) {
        const data = readFileSync(tmp);
        try { execSync(`del "${tmp}"`, { stdio: "pipe", shell: true }); } catch {}
        return data;
      }
    } catch {}
    return null;
  }
}

function checkMermaidCli() {
  const paths = [
    "mmdc",
    join(process.env.APPDATA || "", "npm", "mmdc.cmd"),
    join(process.env.APPDATA || "", "npm", "mmdc"),
  ];
  for (const p of paths) {
    try {
      execSync(`"${p}" --version`, { timeout: 10000, stdio: "pipe" });
      return p;
    } catch {
      continue;
    }
  }
  return null;
}

function renderMermaid(code, outPath) {
  // 1. Try URL
  const url = generateMermaidUrl(code);
  if (url) {
    const data = fetchUrl(url);
    if (data && data.length > 100) {
      writeFileSync(outPath, data);
      console.log(`  Mermaid: downloaded via URL (${data.length} bytes)`);
      return true;
    }
  }

  // 2. Try local CLI
  const cli = checkMermaidCli();
  if (cli) {
    const mmdPath = outPath.replace(/\.png$/, ".mmd");
    writeFileSync(mmdPath, code, "utf-8");
    try {
      execSync(`"${cli}" -i "${mmdPath}" -o "${outPath}" -b transparent`, {
        timeout: 60000,
        stdio: "pipe",
      });
      try { execSync(`del "${mmdPath}"`, { stdio: "pipe", shell: true }); } catch {}
      if (existsSync(outPath)) {
        console.log(`  Mermaid: rendered locally`);
        return true;
      }
    } catch (e) {
      console.log(`  Mermaid CLI error: ${e.message}`);
    }
  }

  return false;
}

// ---------------------------------------------------------------------------
// YAML front-matter stripper
// ---------------------------------------------------------------------------
function stripFrontMatter(text) {
  return text.replace(/^---\s*\n[\s\S]*?\n---\s*\n/, "");
}

// ---------------------------------------------------------------------------
// Inline token renderer -> TextRun[]
// ---------------------------------------------------------------------------
function renderInline(tokens) {
  const runs = [];
  if (!tokens) return runs;
  if (typeof tokens === "string") {
    runs.push(new TextRun({ text: tokens, font: FONT_BODY, size: FONT_SIZE_BODY }));
    return runs;
  }
  for (const tok of tokens) {
    switch (tok.type) {
      case "text":
        runs.push(new TextRun({ text: tok.text || tok.raw || "", font: FONT_BODY, size: FONT_SIZE_BODY }));
        break;
      case "strong":
        runs.push(
          ...flattenInlineText(tok.tokens || []).map(
            (t) => new TextRun({ text: t, bold: true, font: FONT_BODY, size: FONT_SIZE_BODY })
          )
        );
        break;
      case "em":
        runs.push(
          ...flattenInlineText(tok.tokens || []).map(
            (t) => new TextRun({ text: t, italics: true, font: FONT_BODY, size: FONT_SIZE_BODY })
          )
        );
        break;
      case "del":
        runs.push(
          ...flattenInlineText(tok.tokens || []).map(
            (t) => new TextRun({ text: t, strike: true, font: FONT_BODY, size: FONT_SIZE_BODY })
          )
        );
        break;
      case "codespan":
        runs.push(
          new TextRun({
            text: tok.text || tok.raw || "",
            font: FONT_CODE,
            size: FONT_SIZE_CODE,
            shading: { type: ShadingType.CLEAR, fill: CODE_BG },
          })
        );
        break;
      case "link":
        runs.push(
          new ExternalHyperlink({
            children: [
              new TextRun({
                text: flattenText(tok.tokens || tok.text || tok.href),
                style: "Hyperlink",
                color: COLOR_LINK,
                underline: { type: "single" },
                font: FONT_BODY,
                size: FONT_SIZE_BODY,
              }),
            ],
            link: tok.href || "",
          })
        );
        break;
      case "image":
        // Images as inline are rare; handle as placeholder text
        runs.push(new TextRun({ text: `[Image: ${tok.text || tok.href}]`, italics: true, font: FONT_BODY, size: FONT_SIZE_BODY }));
        break;
      case "br":
        runs.push(new TextRun({ break: 1 }));
        break;
      case "space":
        break;
      default:
        // Fallback: render raw text
        runs.push(new TextRun({ text: tok.raw || tok.text || "", font: FONT_BODY, size: FONT_SIZE_BODY }));
    }
  }
  return runs;
}

function flattenInlineText(tokens) {
  const texts = [];
  for (const t of tokens) {
    if (t.text) texts.push(t.text);
    else if (t.raw) texts.push(t.raw);
    else if (t.tokens) texts.push(...flattenInlineText(t.tokens));
  }
  return texts.length ? texts : [""];
}

function flattenText(tokens) {
  if (typeof tokens === "string") return tokens;
  if (Array.isArray(tokens)) return tokens.map(flattenText).join("");
  if (tokens && typeof tokens === "object") {
    return flattenText(tokens.tokens || tokens.text || tokens.raw || "");
  }
  return String(tokens || "");
}

// ---------------------------------------------------------------------------
// Image dimension reader (PNG + JPEG)
// ---------------------------------------------------------------------------
function readImageDimensions(filePath) {
  try {
    const buf = readFileSync(filePath);

    // PNG: 89 50 4E 47 ... IHDR width(4) height(4) at offset 16,20
    if (buf.length >= 24 && buf[0] === 0x89 && buf[1] === 0x50) {
      return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
    }

    // JPEG: FF D8 ... scan for SOF0/SOF2 markers (FFC0/FFC2) which contain dimensions
    if (buf.length >= 4 && buf[0] === 0xFF && buf[1] === 0xD8) {
      let offset = 2;
      while (offset + 8 < buf.length) {
        if (buf[offset] !== 0xFF) { offset++; continue; }
        const marker = buf[offset + 1];
        // SOF0 (0xC0) or SOF2 (0xC2) — baseline or progressive
        if (marker === 0xC0 || marker === 0xC2) {
          const height = buf.readUInt16BE(offset + 5);
          const width = buf.readUInt16BE(offset + 7);
          return { width, height };
        }
        // Skip to next marker using segment length
        if (marker === 0xD8 || marker === 0xD9) { offset += 2; continue; }
        if (marker >= 0xD0 && marker <= 0xD7) { offset += 2; continue; }
        const segLen = buf.readUInt16BE(offset + 2);
        offset += 2 + segLen;
      }
    }
  } catch {}
  return null;
}

function scaleToFit(imgW, imgH) {
  // A4 with 1" margins: 6.27" x 9.69"
  // Width at 100% of content area; height at 85% for header/footer room
  // At 96 DPI: maxW = 6.27 * 96 = 602px, maxH = 8.24 * 96 = 791px
  const maxW = 602;
  const maxH = 791;
  const scale = Math.min(1.0, maxW / imgW, maxH / imgH);
  return { width: Math.round(imgW * scale), height: Math.round(imgH * scale) };
}

// ---------------------------------------------------------------------------
// Block token renderer -> docx elements
// ---------------------------------------------------------------------------
let mermaidCounter = 0;
let numberedListCounter = 0;
const numberingConfigs = [];

function renderBlock(token, inputDir, tmpDir) {
  switch (token.type) {
    case "heading":
      return renderHeading(token);
    case "paragraph":
      return renderParagraph(token, inputDir, tmpDir);
    case "code":
      return renderCodeBlock(token, inputDir, tmpDir);
    case "table":
      return renderTable(token);
    case "list":
      return renderList(token, 0);
    case "blockquote":
      return renderBlockquote(token, inputDir, tmpDir);
    case "hr":
      return [new Paragraph({ children: [new PageBreak()] })];
    case "html":
      if (token.raw && token.raw.trim()) {
        return [
          new Paragraph({
            children: [new TextRun({ text: token.raw.trim(), font: FONT_BODY, size: 18, color: "888888" })],
          }),
        ];
      }
      return [];
    case "space":
      return [];
    default:
      return [];
  }
}

function renderHeading(token) {
  const levels = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
  };
  return [
    new Paragraph({
      heading: levels[token.depth] || HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: flattenText(token.tokens || token.text),
          bold: true,
          font: FONT_HEADING,
          size: HEADING_SIZES[token.depth] || 24,
          color: COLOR_HEADING,
        }),
      ],
    }),
  ];
}

function renderParagraph(token, inputDir, tmpDir) {
  // Check if sole child is an image
  const tokens = token.tokens || [];
  if (tokens.length === 1 && tokens[0].type === "image") {
    return renderImageBlock(tokens[0], inputDir, tmpDir);
  }
  const runs = renderInline(tokens);
  return [new Paragraph({ children: runs, spacing: { after: 120 } })];
}

function renderImageBlock(token, inputDir, tmpDir) {
  const src = token.href || "";
  const imgPath = resolveImage(src, inputDir, tmpDir);
  if (imgPath) {
    try {
      const data = readFileSync(imgPath);
      const ext = extname(imgPath).replace(".", "").toLowerCase() || "png";
      const typeMap = { jpg: "jpg", jpeg: "jpg", png: "png", gif: "gif", bmp: "bmp", svg: "svg" };
      const dims = readImageDimensions(imgPath);
      const { width, height } = dims ? scaleToFit(dims.width, dims.height) : { width: 500, height: 350 };
      return [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new ImageRun({
              type: typeMap[ext] || "png",
              data,
              transformation: { width, height },
              altText: { title: token.text || "Image", description: token.text || "Image", name: "image" },
            }),
          ],
        }),
      ];
    } catch (e) {
      console.log(`  Image error: ${e.message}`);
    }
  }
  return [
    new Paragraph({
      children: [new TextRun({ text: `[Image: ${token.text || src}]`, italics: true, font: FONT_BODY, size: 21 })],
    }),
  ];
}

function resolveImage(src, inputDir, tmpDir) {
  if (src.startsWith("http://") || src.startsWith("https://")) {
    const data = fetchUrl(src);
    if (data) {
      const ext = extname(src).split("?")[0] || ".png";
      const p = join(tmpDir, `img_${Date.now()}${ext}`);
      writeFileSync(p, data);
      return p;
    }
    return null;
  }
  const p = isAbsolute(src) ? src : join(inputDir, src);
  return existsSync(p) ? p : null;
}

function renderCodeBlock(token, inputDir, tmpDir) {
  const lang = (token.lang || "").trim().toLowerCase();
  const code = token.text || "";

  if (lang === "mermaid") {
    return renderMermaidBlock(code, tmpDir);
  }

  // Render code lines as paragraphs with monospace font
  const lines = code.split("\n");
  const elements = [];
  for (const line of lines) {
    elements.push(
      new Paragraph({
        indent: { left: 360 },
        spacing: { before: 0, after: 0 },
        shading: { type: ShadingType.CLEAR, fill: CODE_BG },
        children: [
          new TextRun({
            text: line || " ",
            font: FONT_CODE,
            size: FONT_SIZE_CODE,
            color: "1A1A1A",
          }),
        ],
      })
    );
  }
  // Add spacing paragraph after code block
  elements.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: [] }));
  return elements;
}

function renderMermaidBlock(code, tmpDir) {
  mermaidCounter++;
  const outPath = join(tmpDir, `mermaid_${mermaidCounter}.png`);
  if (renderMermaid(code.trim(), outPath) && existsSync(outPath)) {
    try {
      const data = readFileSync(outPath);
      const dims = readImageDimensions(outPath);
      const { width, height } = dims ? scaleToFit(dims.width, dims.height) : { width: 500, height: 350 };
      // Detect actual image type from header (mermaid.ink often returns JPEG)
      const imgType = (data[0] === 0xFF && data[1] === 0xD8) ? "jpg" : "png";
      return [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new ImageRun({
              type: imgType,
              data,
              transformation: { width, height },
              altText: {
                title: `Mermaid Diagram ${mermaidCounter}`,
                description: `Mermaid Diagram ${mermaidCounter}`,
                name: "mermaid",
              },
            }),
          ],
        }),
      ];
    } catch (e) {
      console.log(`  Mermaid image error: ${e.message}`);
    }
  }
  return [
    new Paragraph({
      children: [
        new TextRun({
          text: `[Mermaid diagram ${mermaidCounter} could not be rendered]`,
          italics: true,
          color: "999999",
          font: FONT_BODY,
          size: FONT_SIZE_BODY,
        }),
      ],
    }),
  ];
}

function renderTable(token) {
  const headerRow = token.header || [];
  const bodyRows = token.rows || [];
  const ncols = headerRow.length || 1;
  const colWidth = Math.floor(PAGE_WIDTH_DXA / ncols);

  const border = { style: BorderStyle.SINGLE, size: TBL_BORDER_SIZE, color: TBL_BORDER };
  const cellBorders = { top: border, bottom: border, left: border, right: border };

  const tblMargins = { top: TBL_CELL_MARGIN, bottom: TBL_CELL_MARGIN, left: TBL_CELL_MARGIN, right: TBL_CELL_MARGIN };

  // Zero-spacing paragraph for tight cell content
  const cellParaSpacing = { before: 0, after: 0, line: 240 };

  // Header – centered both ways
  const hdrCells = headerRow.map((cell) => {
    const text = flattenText(cell.tokens || cell.text || "");
    return new TableCell({
      borders: cellBorders,
      width: { size: colWidth, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: TBL_HDR_BG },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: cellParaSpacing,
          children: [new TextRun({ text, bold: true, font: FONT_BODY, size: FONT_SIZE_TABLE, color: TBL_HDR_TEXT })],
        }),
      ],
    });
  });

  // Body – centered, banded rows only
  const rows = bodyRows.map((row, rIdx) => {
    const cells = row.map((cell) => {
      const runs = renderInline(cell.tokens || []);
      const shade = TBL_BANDED_ROWS && rIdx % 2 === 1 ? { type: ShadingType.CLEAR, fill: TBL_ALT_ROW } : undefined;
      return new TableCell({
        borders: cellBorders,
        width: { size: colWidth, type: WidthType.DXA },
        shading: shade,
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: cellParaSpacing, children: runs })],
      });
    });
    return new TableRow({ children: cells });
  });

  const columnWidths = Array(ncols).fill(colWidth);

  return [
    new Table({
      columnWidths,
      margins: tblMargins,
      rows: [new TableRow({ tableHeader: true, children: hdrCells }), ...rows],
    }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [] }),
  ];
}

function renderList(token, level) {
  const elements = [];
  const ordered = token.ordered;
  const bullets = ["\u2022", "\u25E6", "\u25AA"];
  let counter = token.start || 1;

  for (const item of token.items || []) {
    // Main text of list item
    const textTokens = [];
    const subElements = [];

    for (const child of item.tokens || []) {
      if (child.type === "text" || child.type === "paragraph") {
        const inline = child.tokens || [];
        // Add prefix
        const prefix = ordered ? `${counter}. ` : `${bullets[Math.min(level, bullets.length - 1)]} `;
        const runs = [
          new TextRun({ text: prefix, font: FONT_BODY, size: 21 }),
          ...renderInline(inline),
        ];
        elements.push(
          new Paragraph({
            indent: { left: 360 + level * 360, hanging: 240 },
            spacing: { before: 20, after: 20 },
            children: runs,
          })
        );
        if (ordered) counter++;
      } else if (child.type === "list") {
        elements.push(...renderList(child, level + 1));
      } else if (child.type === "code") {
        elements.push(...renderCodeBlock(child, "", ""));
      }
    }
  }

  return elements;
}

function renderBlockquote(token, inputDir, tmpDir) {
  const elements = [];
  for (const child of token.tokens || []) {
    if (child.type === "paragraph") {
      // Render inline tokens with italic + gray styling for blockquote
      const inlineTokens = child.tokens || [];
      const runs = [];
      for (const tok of inlineTokens) {
        const text = flattenText(tok.tokens || tok.text || tok.raw || "");
        runs.push(
          new TextRun({
            text,
            italics: true,
            color: "666666",
            font: FONT_BODY,
            size: FONT_SIZE_BODY,
            bold: tok.type === "strong" || undefined,
          })
        );
      }
      elements.push(
        new Paragraph({
          indent: { left: 720 },
          spacing: { before: 40, after: 40 },
          border: { left: { style: BorderStyle.SINGLE, size: 6, space: 8, color: "AAAAAA" } },
          children: runs,
        })
      );
    } else {
      // Non-paragraph children (code blocks, nested quotes, etc.)
      elements.push(...renderBlock(child, inputDir, tmpDir));
    }
  }
  return elements;
}

// ---------------------------------------------------------------------------
// Title page, TOC, footer builders
// ---------------------------------------------------------------------------
function buildTitleSection(title, date, preambleTokens) {
  const children = [];
  // Spacers to push title towards vertical centre
  for (let i = 0; i < 6; i++) {
    children.push(new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }));
  }
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: title, bold: true, font: FONT_HEADING, size: 56, color: COLOR_HEADING })],
    })
  );
  if (date) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240 },
        children: [new TextRun({ text: date, font: FONT_BODY, size: 28, color: "666666" })],
      })
    );
  }
  // Render preamble tokens (paragraphs between H1 and first ---) centred
  if (preambleTokens && preambleTokens.length > 0) {
    children.push(new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }));
    for (const tok of preambleTokens) {
      if (tok.type === "paragraph") {
        // Split multi-line paragraphs into separate centred lines
        const lines = (tok.raw || tok.text || "").split("\n").filter(l => l.trim());
        for (const line of lines) {
          const lineToks = marked.lexer(line);
          const inlineToks = (lineToks[0] && lineToks[0].tokens) ? lineToks[0].tokens : [];
          const runs = renderInline(inlineToks);
          children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: runs, spacing: { after: 60 } }));
        }
      }
    }
  }
  return {
    properties: {
      type: SectionType.NEXT_PAGE,
      page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
    },
    children,
  };
}

function buildTocSection() {
  return {
    properties: {
      type: SectionType.NEXT_PAGE,
      page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
    },
    children: [
      new Paragraph({
        children: [new TextRun({ text: "Table of Contents", bold: true, font: FONT_HEADING, size: 32, color: COLOR_HEADING })],
        spacing: { after: 200 },
      }),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
    ],
  };
}

function buildFooter(pagination, copyrightText) {
  const runs = [];
  if (copyrightText && pagination) {
    runs.push(new TextRun({ text: `${copyrightText}  |  Page `, font: FONT_BODY, size: 16, color: "AAAAAA" }));
    runs.push(new TextRun({ children: [PageNumber.CURRENT], font: FONT_BODY, size: 16, color: "AAAAAA" }));
  } else if (pagination) {
    runs.push(new TextRun({ text: "Page ", font: FONT_BODY, size: 16, color: "AAAAAA" }));
    runs.push(new TextRun({ children: [PageNumber.CURRENT], font: FONT_BODY, size: 16, color: "AAAAAA" }));
  } else if (copyrightText) {
    runs.push(new TextRun({ text: copyrightText, font: FONT_BODY, size: 16, color: "AAAAAA" }));
  }
  if (runs.length === 0) return undefined;
  return new Footer({
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: runs })],
  });
}

// ---------------------------------------------------------------------------
// Main conversion
// ---------------------------------------------------------------------------
async function convert(inputPath, outputPath, opts = {}) {
  inputPath = resolve(inputPath);
  outputPath = resolve(outputPath);
  const inputDir = dirname(inputPath);

  const { templatePath, title, date, toc, pagination = true, copyright, skipH1 = false } = opts;

  console.log(`Reading: ${inputPath}`);
  let content = readFileSync(inputPath, "utf-8");
  content = stripFrontMatter(content);

  // --- Style resolution (priority: defaults < template < --style < inline < CLI) ---
  applyStyleConfig();  // reset to defaults

  // Extract style from template (first pass)
  let templateStyle = {};
  if (templatePath && existsSync(templatePath)) {
    templateStyle = extractStyleFromTemplate(templatePath);
    if (Object.keys(templateStyle).length > 0) {
      console.log(`  Template style extracted: ${Object.keys(templateStyle).join(", ")}`);
    }
  }

  // Inline style from markdown <!-- docx-style ... --> comment
  const inlineStyle = parseDocxStyle(content);
  content = stripDocxStyleComment(content);

  // Merge: template < --style file < inline < CLI
  const mergedStyle = { ...templateStyle };
  // --style file overrides (from opts.styleFileOverrides, set by CLI parser)
  if (opts.styleFileOverrides) {
    Object.assign(mergedStyle, opts.styleFileOverrides);
  }
  // Inline docx-style comment overrides --style file
  Object.assign(mergedStyle, inlineStyle);
  // CLI flags have highest priority
  if (opts.cliStyleOverrides) {
    Object.assign(mergedStyle, opts.cliStyleOverrides);
  }
  if (Object.keys(mergedStyle).length > 0) {
    applyStyleConfig(mergedStyle);
    console.log(`  Style applied: ${Object.keys(mergedStyle).join(", ")}`);
  }

  // Temp dir
  const tmpDir = join(tmpdir(), `md2docx_${Date.now()}`);
  mkdirSync(tmpDir, { recursive: true });

  // Parse with marked
  const allTokens = marked.lexer(content);

  // --- Title page logic ---
  // Find first H1 and first HR positions
  let h1Idx = null;
  let h1Text = null;
  let firstHrIdx = null;
  for (let i = 0; i < allTokens.length; i++) {
    if (h1Idx === null && allTokens[i].type === "heading" && allTokens[i].depth === 1) {
      h1Idx = i;
      h1Text = allTokens[i].text;
    }
    if (allTokens[i].type === "hr") {
      firstHrIdx = i;
      break;
    }
  }

  // Determine effective title
  const effectiveTitle = title || h1Text;
  const sections = [];
  let bodyTokens;

  if (effectiveTitle && h1Idx !== null && firstHrIdx !== null) {
    // Preamble = tokens between H1 and first HR (excluding both)
    const preambleTokens = allTokens.slice(h1Idx + 1, firstHrIdx);

    // Body starts after the first HR
    bodyTokens = allTokens.slice(firstHrIdx + 1);

    if (!title) {
      // Auto-detect mode: H1 was used as title, already consumed
    } else if (skipH1) {
      // Explicit --title + --skip-h1: H1 is before HR, already excluded
    } else {
      // Explicit --title without --skip-h1: H1 stays in body
      bodyTokens = [allTokens[h1Idx], ...bodyTokens];
    }

    sections.push(buildTitleSection(effectiveTitle, date, preambleTokens));
    console.log(`  Title page: ${effectiveTitle}` + (date ? ` (${date})` : ""));
  } else if (effectiveTitle && h1Idx !== null) {
    // H1 found but no HR – use H1 as title, rest is body
    if (title && !skipH1) {
      bodyTokens = allTokens; // keep everything including H1
    } else {
      bodyTokens = allTokens.slice(h1Idx + 1); // skip H1
    }
    sections.push(buildTitleSection(effectiveTitle, date, []));
    console.log(`  Title page: ${effectiveTitle}` + (date ? ` (${date})` : ""));
  } else {
    bodyTokens = allTokens;
  }

  // Build document elements from body tokens
  const bodyChildren = [];
  for (const token of bodyTokens) {
    bodyChildren.push(...renderBlock(token, inputDir, tmpDir));
  }

  // Footer
  const footer = buildFooter(pagination, copyright);
  const footerObj = footer ? { default: footer } : undefined;

  // TOC section
  if (toc) {
    sections.push(buildTocSection());
    console.log("  TOC: inserted (update in Word with Ctrl+A, F9)");
  }

  // Body section
  sections.push({
    properties: {
      type: (effectiveTitle || toc) ? SectionType.NEXT_PAGE : undefined,
      page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [] })],
      }),
    },
    footers: footerObj,
    children: bodyChildren,
  });

  // Document styles
  const doc = new Document({
    features: { updateFields: toc },
    styles: {
      default: {
        document: {
          run: { font: FONT_BODY, size: FONT_SIZE_BODY, color: COLOR_BODY },
        },
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 40, bold: true, color: COLOR_HEADING, font: FONT_HEADING },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 },
        },
        {
          id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, color: COLOR_HEADING, font: FONT_HEADING },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 },
        },
        {
          id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, color: COLOR_HEADING, font: FONT_HEADING },
          paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 },
        },
        {
          id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, color: COLOR_HEADING, font: FONT_HEADING },
          paragraph: { spacing: { before: 120, after: 60 }, outlineLevel: 3 },
        },
        {
          id: "Heading5", name: "Heading 5", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 22, bold: true, color: COLOR_HEADING, font: FONT_HEADING },
          paragraph: { spacing: { before: 100, after: 40 }, outlineLevel: 4 },
        },
        {
          id: "Heading6", name: "Heading 6", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 21, bold: true, color: COLOR_HEADING, font: FONT_HEADING },
          paragraph: { spacing: { before: 80, after: 40 }, outlineLevel: 5 },
        },
      ],
    },
    sections,
  });

  // Pack and save
  const buffer = await Packer.toBuffer(doc);
  const outDir = dirname(outputPath);
  if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
  writeFileSync(outputPath, buffer);

  console.log(`Saved: ${outputPath}`);
  console.log(`  Mermaid diagrams processed: ${mermaidCounter}`);

  // Cleanup temp dir
  try {
    execSync(process.platform === "win32" ? `rmdir /s /q "${tmpDir}"` : `rm -rf "${tmpDir}"`, {
      stdio: "pipe",
    });
  } catch {}
}

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------
const args = process.argv.slice(2);
let inputFile = null;
let outputFile = null;
const opts = {};

for (let i = 0; i < args.length; i++) {
  if (args[i] === "--template" && i + 1 < args.length) {
    opts.templatePath = args[++i];
  } else if (args[i] === "--title" && i + 1 < args.length) {
    opts.title = args[++i];
  } else if (args[i] === "--date" && i + 1 < args.length) {
    opts.date = args[++i];
  } else if (args[i] === "--toc") {
    opts.toc = true;
  } else if (args[i] === "--no-pagination") {
    opts.pagination = false;
  } else if (args[i] === "--copyright" && i + 1 < args.length) {
    opts.copyright = args[++i];
  } else if (args[i] === "--skip-h1") {
    opts.skipH1 = true;
  } else if (args[i] === "--style" && i + 1 < args.length) {
    opts.styleFile = args[++i];
  } else if (args[i] === "--font-body" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.font_body = args[++i];
  } else if (args[i] === "--font-heading" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.font_heading = args[++i];
  } else if (args[i] === "--font-code" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.font_code = args[++i];
  } else if (args[i] === "--font-size" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.font_size = args[++i];
  } else if (args[i] === "--color-heading" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.color_heading = args[++i];
  } else if (args[i] === "--color-body" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.color_body = args[++i];
  } else if (args[i] === "--table-header-bg" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_header_bg = args[++i];
  } else if (args[i] === "--table-header-text" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_header_text = args[++i];
  } else if (args[i] === "--table-alt-row" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_alt_row = args[++i];
  } else if (args[i] === "--table-border" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_border = args[++i];
  } else if (args[i] === "--table-border-size" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_border_size = args[++i];
  } else if (args[i] === "--table-cell-margin" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_cell_margin = args[++i];
  } else if (args[i] === "--table-font-size" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_font_size = args[++i];
  } else if (args[i] === "--no-banded-rows") {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.table_banded_rows = "false";
  } else if (args[i] === "--code-bg" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.code_bg = args[++i];
  } else if (args[i] === "--code-font-size" && i + 1 < args.length) {
    opts._cli_style = opts._cli_style || {}; opts._cli_style.code_font_size = args[++i];
  } else if (!inputFile) {
    inputFile = args[i];
  } else if (!outputFile) {
    outputFile = args[i];
  }
}

if (!inputFile || !outputFile) {
  console.log(
    "Usage: node md_to_docx_js.mjs input.md output.docx [options]\n" +
    "\nOptions:\n" +
    "  --template FILE          Reference DOCX template\n" +
    "  --title TEXT              Add title page with this title\n" +
    "  --date TEXT               Date shown on title page\n" +
    "  --toc                    Insert Table of Contents after title page\n" +
    "  --no-pagination          Disable page numbers in footer\n" +
    "  --copyright TEXT         Copyright string in footer\n" +
    "  --skip-h1               Skip first H1 from body (use with --title)\n" +
    "  --style FILE             Style file (key: value format)\n" +
    "\nStyle options (override inline <!-- docx-style --> comment):\n" +
    "  --font-body FONT         Body font (default: Arial)\n" +
    "  --font-heading FONT      Heading font (default: Arial)\n" +
    "  --font-code FONT         Code font (default: Consolas)\n" +
    "  --font-size PT           Body text size in pt (default: 10.5)\n" +
    "  --color-heading HEX      Heading color (default: 2D3B4D)\n" +
    "  --color-body HEX         Body text color (default: 333333)\n" +
    "  --table-header-bg HEX    Table header background (default: D5E8F0)\n" +
    "  --table-header-text HEX  Table header text color (default: 2D3B4D)\n" +
    "  --table-alt-row HEX      Alternating row color (default: F2F2F2)\n" +
    "  --table-border HEX       Border color (default: CCCCCC)\n" +
    "  --table-border-size N    Border width in 1/8pt (default: 4)\n" +
    "  --table-cell-margin N    Cell margin in twips (default: 28)\n" +
    "  --table-font-size PT     Table text size in pt (default: 9.5)\n" +
    "  --no-banded-rows         Disable alternating row shading\n" +
    "  --code-bg HEX            Code background (default: F5F5F5)\n" +
    "  --code-font-size PT      Code text size in pt (default: 9)"
  );
  process.exit(1);
}

if (!existsSync(inputFile)) {
  console.error(`Error: Input file not found: ${inputFile}`);
  process.exit(1);
}

// Build style overrides: --style file and CLI flags separated for proper priority
if (opts.styleFile && existsSync(opts.styleFile)) {
  opts.styleFileOverrides = parseStyleFile(opts.styleFile);
}
if (opts._cli_style) {
  opts.cliStyleOverrides = opts._cli_style;
  delete opts._cli_style;
}

convert(inputFile, outputFile, opts).catch((e) => {
  console.error(`Error: ${e.message}`);
  process.exit(1);
});
