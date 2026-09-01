import { chromium, type Browser } from "playwright";

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

/**
 * Renders a full HTML document to a PNG screenshot at an exact pixel size.
 * The HTML is expected to fill the given viewport (e.g. a slide-sized <body>).
 */
export async function renderHtmlToPng(
  html: string,
  width: number,
  height: number,
  deviceScaleFactor = 1
): Promise<Buffer> {
  const browser = await getBrowser();
  const page = await browser.newPage({
    viewport: { width, height },
    deviceScaleFactor,
  });
  try {
    await page.setContent(html, { waitUntil: "networkidle" });
    const png = await page.screenshot({ type: "png", clip: { x: 0, y: 0, width, height } });
    return png;
  } finally {
    await page.close();
  }
}
