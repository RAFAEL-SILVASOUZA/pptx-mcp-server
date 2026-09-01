import * as PptxGenJSModule from "pptxgenjs";
import path from "node:path";
import fs from "node:fs";

// pptxgenjs ships a CJS build with an ambient-style .d.ts; under NodeNext
// resolution the default export doesn't line up cleanly, so resolve it at
// runtime instead of relying on the (mis-typed) default import.
const PptxGenJS = (PptxGenJSModule as any).default ?? (PptxGenJSModule as any);

const PX_PER_INCH = 96;

/**
 * Assembles a list of full-slide PNG screenshots into a .pptx file.
 * Each image is placed full-bleed (0,0) covering the entire slide.
 */
export async function buildPptx(
  slidePngs: Buffer[],
  outputPath: string,
  width: number,
  height: number
): Promise<void> {
  const pres: any = new PptxGenJS();
  const widthIn = width / PX_PER_INCH;
  const heightIn = height / PX_PER_INCH;

  pres.defineLayout({ name: "CUSTOM", width: widthIn, height: heightIn });
  pres.layout = "CUSTOM";

  for (const png of slidePngs) {
    const slide = pres.addSlide();
    slide.addImage({
      data: `image/png;base64,${png.toString("base64")}`,
      x: 0,
      y: 0,
      w: widthIn,
      h: heightIn,
    });
  }

  const resolved = path.resolve(outputPath);
  fs.mkdirSync(path.dirname(resolved), { recursive: true });
  await pres.writeFile({ fileName: resolved });
}
