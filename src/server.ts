#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { renderHtmlToPng, closeBrowser } from "./render.js";
import { buildPptx } from "./pptx.js";
import { resolveOutputPath } from "./output-path.js";

const server = new McpServer({ name: "hjmcp-ppt", version: "0.1.0" });

const DEFAULT_WIDTH = 1280;
const DEFAULT_HEIGHT = 720;

server.registerTool(
  "preview_slide",
  {
    title: "Preview slide (design check only — does NOT save a file)",
    description:
      "DESIGN-ITERATION TOOL ONLY. Renders a single complete HTML document to a PNG image " +
      "and returns it inline for visual review — it does not write anything to disk and is " +
      "NOT how a presentation gets delivered. Use it only to check how one slide looks while " +
      "you're still tweaking its HTML/CSS. Once the user's slides are finalized, you MUST call " +
      "build_pptx to actually produce the .pptx file — that is the only tool that creates a " +
      "deliverable presentation file.",
    inputSchema: {
      html: z
        .string()
        .describe("Complete, self-contained HTML document (with inline <style>) for the slide"),
      width: z.number().int().positive().optional().describe("Slide width in px (default 1280)"),
      height: z.number().int().positive().optional().describe("Slide height in px (default 720)"),
    },
  },
  async ({ html, width, height }) => {
    // Rendered at 1x (not the 2x used for the final pptx) to keep the inline
    // base64 payload small — this is a quick visual check, not the deliverable.
    const png = await renderHtmlToPng(html, width ?? DEFAULT_WIDTH, height ?? DEFAULT_HEIGHT, 1);
    return {
      content: [{ type: "image", data: png.toString("base64"), mimeType: "image/png" }],
    };
  }
);

server.registerTool(
  "build_pptx",
  {
    title: "Build PPTX (creates the actual presentation file)",
    description:
      "THE DELIVERABLE TOOL. Renders a list of complete HTML documents (one per slide) and " +
      "assembles them into a single PowerPoint (.pptx) file written to disk, one image-backed " +
      "slide per HTML input, in order. Call this — not preview_slide — whenever the user asks " +
      "for a PPT/PowerPoint/presentation/apresentação/slide deck; the tool result gives you the " +
      "saved file path to report back to the user.",
    inputSchema: {
      slides: z
        .array(z.string())
        .min(1)
        .describe("Array of complete, self-contained HTML documents, one per slide, in order"),
      outputPath: z
        .string()
        .optional()
        .describe(
          "Where to save the .pptx. Absolute path recommended. If a bare filename or " +
            "nothing is given, the file is saved in the client's workspace root (if the " +
            "client shares one) or in a configured default folder otherwise."
        ),
      width: z.number().int().positive().optional().describe("Slide width in px (default 1280)"),
      height: z.number().int().positive().optional().describe("Slide height in px (default 720)"),
    },
  },
  async ({ slides, outputPath, width, height }) => {
    const w = width ?? DEFAULT_WIDTH;
    const h = height ?? DEFAULT_HEIGHT;
    const resolvedPath = await resolveOutputPath(server.server, outputPath);
    const pngs: Buffer[] = [];
    for (const html of slides) {
      pngs.push(await renderHtmlToPng(html, w, h));
    }
    await buildPptx(pngs, resolvedPath, w, h);
    return {
      content: [
        {
          type: "text",
          text: `PPTX gerado com ${slides.length} slide(s) em: ${resolvedPath}`,
        },
      ],
    };
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

const shutdown = async () => {
  await closeBrowser();
  process.exit(0);
};
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);

main().catch((err) => {
  console.error("Fatal error starting hjmcp-ppt server:", err);
  process.exit(1);
});
