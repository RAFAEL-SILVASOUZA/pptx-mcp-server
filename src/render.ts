import { chromium, type Browser, type Page } from "playwright";
import { pathToFileURL, fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import path from "node:path";
import fs from "node:fs";
import os from "node:os";
import crypto from "node:crypto";

/**
 * Local UMD bundle of Mermaid (defines `window.mermaid`), shipped as a
 * dependency so diagram rendering works offline and with a pinned version —
 * no CDN fetch at conversion time.
 *
 * Resolved via `require.resolve`, not a hand-built relative path — npm/npx
 * routinely hoists a single-consumer dependency like this one up to a
 * parent `node_modules` (e.g. when this package is installed as someone
 * else's sole dependency via `npx`), so "mermaid is a sibling of dist/" is
 * not a safe assumption. `require.resolve` walks the real Node module
 * resolution algorithm and finds it wherever npm actually put it.
 */
const require = createRequire(import.meta.url);
const MERMAID_BUNDLE_PATH = path.join(
  path.dirname(require.resolve("mermaid/package.json")),
  "dist",
  "mermaid.min.js"
);

let browserPromise: Promise<Browser> | null = null;

function getBrowser(): Promise<Browser> {
  if (!browserPromise) {
    browserPromise = chromium.launch({ headless: true });
  }
  return browserPromise;
}

export async function closeBrowser(): Promise<void> {
  if (browserPromise) {
    const browser = await browserPromise;
    await browser.close();
    browserPromise = null;
  }
}

export type PptxItemType = "text" | "shape" | "image";

export interface ExtractedItem {
  type: PptxItemType;
  x: number;
  y: number;
  w: number;
  h: number;
  // text
  text?: string;
  fontFamily?: string;
  fontSizePx?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string | null;
  align?: "left" | "center" | "right" | "justify";
  // shape
  fill?: string | null;
  borderColor?: string | null;
  borderWidthPx?: number;
  borderRadiusPx?: number;
  opacity?: number;
  // image
  src?: string | null;
}

export interface ExtractedSlide {
  widthPx: number;
  heightPx: number;
  background: string | null;
  items: ExtractedItem[];
  /** Directory of the HTML file this slide was read from — used to resolve relative image paths. */
  sourceDir: string;
}

export interface ExtractResult {
  slides: ExtractedSlide[];
  warnings: string[];
}

/**
 * Loads one or more HTML files in a real browser (so CSS layout — flexbox,
 * grid, fonts, everything — is computed exactly like it would be on screen)
 * and reads back the rendered position/style of every element the author
 * explicitly marked with a `data-pptx` attribute. No screenshot is taken;
 * only geometry and computed style values are extracted, for mapping onto
 * native PPTX objects.
 *
 * Each file may contain one or more `[data-pptx-slide]` elements (the
 * classic single-file deck), or exactly one (the one-file-per-slide
 * pattern) — either way, slides are concatenated in the order the files are
 * given, then in DOM order within each file.
 */
export async function extractPresentation(htmlPaths: string[]): Promise<ExtractResult> {
  const browser = await getBrowser();
  const slides: ExtractedSlide[] = [];
  const warnings: string[] = [];

  for (const htmlPath of htmlPaths) {
    // Higher device scale factor so elements rasterized via screenshot
    // (currently just Mermaid diagrams) come out crisp, not blurry, when
    // scaled up inside PowerPoint. Doesn't affect layout/position reads,
    // which are always in CSS pixels regardless of device scale factor.
    const page = await browser.newPage({ deviceScaleFactor: 2 });
    try {
      const fileUrl = pathToFileURL(htmlPath).href;
      await page.goto(fileUrl, { waitUntil: "networkidle" });
      warnings.push(...(await renderMermaidDiagrams(page)));
      const fileSlides = await page.evaluate(extractInBrowser);
      if (fileSlides.length === 0) {
        warnings.push(
          `No element with [data-pptx-slide] was found in "${htmlPath}" — no slides were read from this file.`
        );
      }
      const sourceDir = path.dirname(htmlPath);
      for (const slide of fileSlides) {
        slides.push({ ...slide, sourceDir });
      }
    } finally {
      await page.close();
    }
  }

  return { slides, warnings };
}

/**
 * Finds every `[data-pptx-mermaid]` element on the page, renders its Mermaid
 * source into an SVG diagram (via a locally bundled copy of Mermaid, no
 * network access needed), rasterizes each rendered diagram to a temp PNG
 * sized to fill the element's own box, and rewrites the element to
 * `data-pptx="image"` with that PNG as a CSS `background-image`.
 *
 * This runs entirely before `extractInBrowser` — by the time extraction
 * happens, a Mermaid element looks like any other image element, so
 * `extractInBrowser` and `convert.ts` need no Mermaid-specific code at all.
 *
 * Returns warnings (e.g. invalid diagram source) in the same shape used
 * elsewhere in this module; never throws for a single bad diagram, so one
 * broken diagram doesn't take down the rest of the deck.
 */
async function renderMermaidDiagrams(page: Page): Promise<string[]> {
  const warnings: string[] = [];
  const count = await page.locator("[data-pptx-mermaid]").count();
  if (count === 0) return warnings;

  await page.addScriptTag({ path: MERMAID_BUNDLE_PATH });

  const ids = await page.evaluate(() => {
    const els = Array.from(document.querySelectorAll<HTMLElement>("[data-pptx-mermaid]"));
    return els.map((el, i) => {
      const id = `pptx-mmd-${i}`;
      el.setAttribute("data-pptx-mermaid-id", id);
      return id;
    });
  });

  const tempDir = path.join(os.tmpdir(), "pptx-mcp-mermaid");
  fs.mkdirSync(tempDir, { recursive: true });
  const runTag = crypto.randomUUID();

  for (const id of ids) {
    const selector = `[data-pptx-mermaid-id="${id}"]`;

    const renderError = await page.evaluate(
      async ({ selector, id }) => {
        const w = window as any;
        const el = document.querySelector(selector) as HTMLElement | null;
        if (!el) return `element "${id}" disappeared before rendering`;
        const source = (el.getAttribute("data-pptx-mermaid") || el.textContent || "").trim();
        el.textContent = "";
        try {
          w.mermaid.initialize({ startOnLoad: false, securityLevel: "loose" });
          const { svg } = await w.mermaid.render(`render-${id}`, source);
          el.innerHTML = svg;
          const svgEl = el.querySelector("svg") as SVGElement | null;
          if (svgEl) {
            // Let the diagram fill the author's box width-first, height
            // following the diagram's own aspect ratio (via its viewBox) —
            // the same "set one dimension, let the other follow" convention
            // already used for data-pptx="image" elements in this contract.
            svgEl.removeAttribute("height");
            (svgEl as unknown as HTMLElement).style.width = "100%";
            (svgEl as unknown as HTMLElement).style.height = "auto";
            (svgEl as unknown as HTMLElement).style.display = "block";
          }
          return null;
        } catch (err: any) {
          return err?.message || String(err);
        }
      },
      { selector, id }
    );

    if (renderError) {
      warnings.push(`A [data-pptx-mermaid] diagram failed to render and was skipped (${renderError}).`);
      continue;
    }

    const box = await page.locator(selector).boundingBox();
    if (!box || box.width <= 0 || box.height <= 0) {
      warnings.push(
        `A [data-pptx-mermaid] diagram rendered with zero size and was skipped — give its ` +
          `element an explicit width in CSS.`
      );
      continue;
    }

    const filePath = path.join(tempDir, `${runTag}-${id}.png`);
    await page.locator(selector).screenshot({ path: filePath });

    const fileUrl = pathToFileURL(filePath).href;
    await page.evaluate(
      ({ selector, fileUrl }) => {
        const el = document.querySelector(selector) as HTMLElement | null;
        if (!el) return;
        el.setAttribute("data-pptx", "image");
        el.style.backgroundImage = `url('${fileUrl}')`;
      },
      { selector, fileUrl }
    );
  }

  return warnings;
}

/**
 * Runs inside the page (serialized by Playwright) — must be self-contained,
 * no references to anything outside this function body.
 */
function extractInBrowser(): Array<Omit<ExtractedSlide, "sourceDir">> {
  function toHex(colorStr: string): { hex: string; alpha: number } | null {
    const m = colorStr.match(/rgba?\(([^)]+)\)/);
    if (!m) return null;
    const parts = m[1].split(",").map((s) => parseFloat(s.trim()));
    const [r, g, b, a = 1] = parts;
    if (a === 0) return null;
    const hex = [r, g, b]
      .map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0"))
      .join("");
    return { hex: hex.toUpperCase(), alpha: a };
  }

  const slideEls = Array.from(document.querySelectorAll<HTMLElement>("[data-pptx-slide]"));

  return slideEls.map((slideEl) => {
    const slideRect = slideEl.getBoundingClientRect();
    const slideStyle = getComputedStyle(slideEl);
    const bg = toHex(slideStyle.backgroundColor);

    const items = Array.from(slideEl.querySelectorAll<HTMLElement>("[data-pptx]")).map((el) => {
      const rect = el.getBoundingClientRect();
      const style = getComputedStyle(el);
      const type = el.getAttribute("data-pptx") as PptxItemType;

      const base = {
        type,
        x: rect.left - slideRect.left,
        y: rect.top - slideRect.top,
        w: rect.width,
        h: rect.height,
      };

      if (type === "text") {
        const color = toHex(style.color);
        const weight = parseInt(style.fontWeight, 10) || (style.fontWeight === "bold" ? 700 : 400);
        return {
          ...base,
          text: (el as HTMLElement).innerText,
          fontFamily: style.fontFamily.split(",")[0].replace(/['"]/g, "").trim(),
          fontSizePx: parseFloat(style.fontSize),
          bold: weight >= 600,
          italic: style.fontStyle === "italic",
          color: color ? color.hex : null,
          align: (["left", "center", "right", "justify"].includes(style.textAlign)
            ? style.textAlign
            : "left") as ExtractedItem["align"],
        };
      }

      if (type === "image") {
        let src: string | null = null;
        if (el instanceof HTMLImageElement) {
          src = el.currentSrc || el.src || null;
        } else if (style.backgroundImage && style.backgroundImage !== "none") {
          const m = style.backgroundImage.match(/url\(["']?(.*?)["']?\)/);
          src = m ? m[1] : null;
        }
        return { ...base, src };
      }

      // shape (default)
      const fill = toHex(style.backgroundColor);
      const borderWidthPx = parseFloat(style.borderTopWidth) || 0;
      const borderColor = borderWidthPx > 0 ? toHex(style.borderTopColor) : null;
      return {
        ...base,
        fill: fill ? fill.hex : null,
        opacity: fill ? fill.alpha : 1,
        borderColor: borderColor ? borderColor.hex : null,
        borderWidthPx,
        borderRadiusPx: parseFloat(style.borderTopLeftRadius) || 0,
      };
    });

    return {
      widthPx: slideRect.width,
      heightPx: slideRect.height,
      background: bg ? bg.hex : null,
      items: items as ExtractedItem[],
    };
  });
}
