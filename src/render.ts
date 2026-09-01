import { chromium, type Browser } from "playwright";
import { pathToFileURL } from "node:url";
import path from "node:path";

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
    const page = await browser.newPage();
    try {
      const fileUrl = pathToFileURL(htmlPath).href;
      await page.goto(fileUrl, { waitUntil: "networkidle" });
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
