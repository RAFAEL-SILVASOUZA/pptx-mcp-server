import * as PptxGenJSModule from "pptxgenjs";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { ExtractedItem, ExtractedSlide } from "./render.js";

// See the note in the old pptx.ts / prior commits: pptxgenjs's CJS build
// doesn't resolve cleanly as a default import under NodeNext, so resolve at
// runtime instead of relying on the (mis-typed) default import.
const PptxGenJS = (PptxGenJSModule as any).default ?? (PptxGenJSModule as any);

const PX_PER_INCH = 96;
const PX_TO_PT = 0.75; // 1px = 0.75pt at 96dpi

export interface ConvertResult {
  outputPath: string;
  slideCount: number;
  warnings: string[];
}

export async function convertHtmlToPptx(
  slides: ExtractedSlide[],
  outputPath: string
): Promise<ConvertResult> {
  const warnings: string[] = [];
  if (slides.length === 0) {
    warnings.push(
      "No element with [data-pptx-slide] was found in the HTML file(s) — the output .pptx has 0 slides."
    );
  }

  const pres: any = new PptxGenJS();
  const first = slides[0];
  const widthIn = first ? first.widthPx / PX_PER_INCH : 13.333;
  const heightIn = first ? first.heightPx / PX_PER_INCH : 7.5;
  pres.defineLayout({ name: "CUSTOM", width: widthIn, height: heightIn });
  pres.layout = "CUSTOM";

  for (let i = 0; i < slides.length; i++) {
    const extracted = slides[i];
    if (
      Math.abs(extracted.widthPx / PX_PER_INCH - widthIn) > 0.05 ||
      Math.abs(extracted.heightPx / PX_PER_INCH - heightIn) > 0.05
    ) {
      warnings.push(
        `Slide ${i + 1} is ${extracted.widthPx}x${extracted.heightPx}px, different from slide 1 ` +
          `(${first.widthPx}x${first.heightPx}px). All slides should share the same size; using ` +
          `slide 1's size for the whole deck, so slide ${i + 1} may look off.`
      );
    }

    const pptxSlide = pres.addSlide();
    if (extracted.background) {
      pptxSlide.background = { color: extracted.background };
    }

    for (const item of extracted.items) {
      if (item.w <= 0 || item.h <= 0) {
        warnings.push(
          `Slide ${i + 1}: a [data-pptx="${item.type}"] element has zero width/height and was skipped.`
        );
        continue;
      }
      try {
        await addItemToSlide(pptxSlide, item, extracted.sourceDir, warnings, i + 1);
      } catch (err: any) {
        warnings.push(
          `Slide ${i + 1}: a [data-pptx="${item.type}"] element failed to convert and was ` +
            `skipped (${err?.message || err}).`
        );
      }
    }
  }

  const resolved = path.resolve(outputPath);
  fs.mkdirSync(path.dirname(resolved), { recursive: true });
  await pres.writeFile({ fileName: resolved });

  return { outputPath: resolved, slideCount: slides.length, warnings };
}

async function addItemToSlide(
  pptxSlide: any,
  item: ExtractedItem,
  htmlDir: string,
  warnings: string[],
  slideNumber: number
): Promise<void> {
  const x = item.x / PX_PER_INCH;
  const y = item.y / PX_PER_INCH;
  const w = item.w / PX_PER_INCH;
  const h = item.h / PX_PER_INCH;

  if (item.type === "text") {
    if (!item.text || item.text.trim().length === 0) return;
    pptxSlide.addText(item.text, {
      x,
      y,
      w,
      h,
      fontFace: item.fontFamily || "Arial",
      fontSize: Math.max(1, Math.round((item.fontSizePx || 16) * PX_TO_PT)),
      bold: !!item.bold,
      italic: !!item.italic,
      color: item.color || "000000",
      align: item.align || "left",
      valign: "top",
      wrap: true,
      margin: 0,
    });
    return;
  }

  if (item.type === "image") {
    if (!item.src) {
      warnings.push(`Slide ${slideNumber}: an [data-pptx="image"] element has no resolvable image source and was skipped.`);
      return;
    }
    const data = await resolveImageAsDataUri(item.src, htmlDir);
    if (!data) {
      warnings.push(`Slide ${slideNumber}: could not load image "${item.src}" and it was skipped.`);
      return;
    }
    pptxSlide.addImage({ data, x, y, w, h });
    return;
  }

  // shape
  const hasRadius = (item.borderRadiusPx || 0) > 0;
  const shapeType = hasRadius ? "roundRect" : "rect";
  const opts: any = { x, y, w, h };
  if (item.fill) {
    opts.fill = { color: item.fill, transparency: Math.round((1 - (item.opacity ?? 1)) * 100) };
  } else {
    opts.fill = { color: "FFFFFF", transparency: 100 };
  }
  if (item.borderColor && item.borderWidthPx) {
    opts.line = { color: item.borderColor, width: Math.max(0.25, item.borderWidthPx * PX_TO_PT) };
  } else {
    opts.line = { type: "none" };
  }
  if (hasRadius) {
    const maxRadiusIn = Math.min(w, h) / 2;
    opts.rectRadius = Math.min(1, Math.min(maxRadiusIn, (item.borderRadiusPx || 0) / PX_PER_INCH));
  }
  pptxSlide.addShape(shapeType, opts);
}

async function resolveImageAsDataUri(src: string, htmlDir: string): Promise<string | null> {
  try {
    if (src.startsWith("data:")) {
      return isValidDataUri(src) ? src : null;
    }
    if (src.startsWith("file://")) {
      const filePath = fileURLToPath(src);
      return fileToDataUri(filePath);
    }
    if (src.startsWith("http://") || src.startsWith("https://")) {
      const res = await fetch(src);
      if (!res.ok) return null;
      const buf = Buffer.from(await res.arrayBuffer());
      const mime = res.headers.get("content-type") || guessMime(src);
      return `data:${mime};base64,${buf.toString("base64")}`;
    }
    // Bare relative/absolute filesystem path.
    const filePath = path.isAbsolute(src) ? src : path.join(htmlDir, src);
    return fileToDataUri(filePath);
  } catch {
    return null;
  }
}

function isValidDataUri(src: string): boolean {
  const m = src.match(/^data:[^;,]+;base64,([\s\S]*)$/);
  if (!m) return false;
  const payload = m[1];
  return payload.length > 0 && /^[A-Za-z0-9+/]*={0,2}$/.test(payload);
}

function fileToDataUri(filePath: string): string | null {
  if (!fs.existsSync(filePath)) return null;
  const buf = fs.readFileSync(filePath);
  return `data:${guessMime(filePath)};base64,${buf.toString("base64")}`;
}

function guessMime(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  switch (ext) {
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".gif":
      return "image/gif";
    case ".svg":
      return "image/svg+xml";
    case ".webp":
      return "image/webp";
    default:
      return "image/png";
  }
}
