import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import type { Server } from "@modelcontextprotocol/sdk/server/index.js";

const FALLBACK_DIR =
  process.env.PPT_MCP_OUTPUT_DIR && process.env.PPT_MCP_OUTPUT_DIR.trim().length > 0
    ? process.env.PPT_MCP_OUTPUT_DIR
    : path.join(os.homedir(), "Documents", "PPT-MCP");

/**
 * Resolves where a .pptx should be written.
 *
 * - If `requested` is an absolute path, it's used as-is.
 * - If `requested` is a bare filename (or omitted), the directory is taken from the
 *   MCP client's first declared "root" (its workspace/project folder) when the client
 *   supports that capability — otherwise falls back to the PPT_MCP_OUTPUT_DIR env var,
 *   or ~/Documents/PPT-MCP if that isn't set either.
 *
 * MCP servers run as their own process and don't inherit the calling agent's working
 * directory; "roots" is the protocol's mechanism for a client to share that folder.
 */
export async function resolveOutputPath(
  mcpServer: Server,
  requested: string | undefined
): Promise<string> {
  if (requested && path.isAbsolute(requested)) {
    return requested;
  }

  const filename = requested && requested.trim().length > 0
    ? requested
    : `apresentacao-${Date.now()}.pptx`;

  const baseDir = await resolveBaseDir(mcpServer);
  return path.join(baseDir, filename);
}

async function resolveBaseDir(mcpServer: Server): Promise<string> {
  const clientCapabilities = mcpServer.getClientCapabilities();
  if (!clientCapabilities?.roots) {
    return FALLBACK_DIR;
  }

  try {
    const { roots } = await mcpServer.listRoots(undefined, { timeout: 3000 });
    const firstFileRoot = roots.find((r) => r.uri.startsWith("file://"));
    if (firstFileRoot) {
      return fileURLToPath(firstFileRoot.uri);
    }
  } catch {
    // Client declared roots support but didn't answer usefully — fall back below.
  }

  return FALLBACK_DIR;
}
