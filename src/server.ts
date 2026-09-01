#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import fs from "node:fs";
import path from "node:path";
import { extractPresentation, closeBrowser } from "./render.js";
import { convertHtmlToPptx } from "./convert.js";
import { resolveOutputPath } from "./output-path.js";
import { AUTHORING_GUIDE } from "./guide.js";

const server = new McpServer({ name: "pptx-mcp-server", version: "0.2.0" });

server.registerTool(
  "get_pptx_authoring_guide",
  {
    title: "Get PPTX authoring guide (read this first)",
    description:
      "Returns the full guide for authoring a presentation this server can convert: the " +
      "HTML markup contract (data-pptx-slide / data-pptx=\"text\"|\"shape\"|\"image\"), how " +
      "positioning and styling map to native PowerPoint objects, what's unsupported, and " +
      "design guidance. Call this BEFORE writing any presentation HTML — the conversion is " +
      "markup-driven, so writing HTML without following this contract will produce an empty " +
      "or broken .pptx.",
    inputSchema: {},
  },
  async () => {
    return { content: [{ type: "text", text: AUTHORING_GUIDE }] };
  }
);

server.registerTool(
  "convert_html_to_pptx",
  {
    title: "Convert HTML to PPTX (creates the actual presentation file)",
    description:
      "THE DELIVERABLE TOOL. You must have already written a complete HTML file (following " +
      "get_pptx_authoring_guide's contract) to the workspace yourself — this tool does not " +
      "render, design, or screenshot anything. It loads that HTML file in a headless browser " +
      "only to read exact positions/styles of elements you marked with data-pptx attributes, " +
      "and writes a native, editable .pptx built from those elements (text boxes, shapes, " +
      "pictures — never a flattened image). Call this once the HTML is finished; the result " +
      "includes the saved file path and any warnings about elements that didn't convert " +
      "cleanly.",
    inputSchema: {
      htmlPath: z
        .string()
        .describe("Absolute path to the single HTML file to convert (written earlier by you)"),
      outputPath: z
        .string()
        .optional()
        .describe(
          "Where to save the .pptx. Absolute path recommended. If a bare filename or " +
            "nothing is given, the file is saved in the client's workspace root (if the " +
            "client shares one) or in a configured default folder otherwise."
        ),
    },
  },
  async ({ htmlPath, outputPath }) => {
    const resolvedHtmlPath = path.resolve(htmlPath);
    if (!fs.existsSync(resolvedHtmlPath)) {
      return {
        isError: true,
        content: [{ type: "text", text: `HTML file not found: ${resolvedHtmlPath}` }],
      };
    }

    const slides = await extractPresentation(resolvedHtmlPath);
    const resolvedOutputPath = await resolveOutputPath(
      server.server,
      outputPath,
      `apresentacao-${Date.now()}.pptx`
    );

    const result = await convertHtmlToPptx(slides, resolvedHtmlPath, resolvedOutputPath);

    const lines = [`PPTX gerado com ${result.slideCount} slide(s) em: ${result.outputPath}`];
    if (result.warnings.length > 0) {
      lines.push("", "Warnings:", ...result.warnings.map((w) => `- ${w}`));
    }

    return { content: [{ type: "text", text: lines.join("\n") }] };
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
  console.error("Fatal error starting pptx-mcp-server:", err);
  process.exit(1);
});
