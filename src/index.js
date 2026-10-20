#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { writeFile } from "node:fs/promises";
import { basename } from "node:path";
import { loadBytes, outline, thumbnail, previewHtml, SDK_URL, PREVIEW_SLIDE_LIMIT } from "./pptx.js";

const source = z.string().describe("Local path or http(s) URL of a .pptx file");
const fail = (e) => ({ isError: true, content: [{ type: "text", text: String(e?.message ?? e) }] });

export function createServer() {
  const server = new McpServer({ name: "pptx-mcp", version: "0.1.0" });

  server.registerTool("pptx_outline", {
    title: "Outline a PowerPoint file",
    description: "Turns a .pptx into a slide-by-slide outline: title, body text and speaker notes, as JSON.",
    inputSchema: { source },
  }, async ({ source }) => {
    try { return { content: [{ type: "text", text: JSON.stringify(await outline(await loadBytes(source)), null, 2) }] }; }
    catch (e) { return fail(e); }
  });

  server.registerTool("pptx_thumbnail", {
    title: "Show the cover thumbnail",
    description: "Returns the cover thumbnail image that PowerPoint saved inside the .pptx, if there is one.",
    inputSchema: { source },
  }, async ({ source }) => {
    try {
      const t = await thumbnail(await loadBytes(source));
      if (!t) return { content: [{ type: "text", text: `This file has no saved thumbnail. To render any slide as an image, see PPTX Player SDK: ${SDK_URL}` }] };
      return { content: [{ type: "image", data: t.data, mimeType: t.mimeType }] };
    } catch (e) { return fail(e); }
  });

  server.registerTool("pptx_web_preview", {
    title: "Make a shareable text preview",
    description: `Writes a self-contained HTML page previewing the first ${PREVIEW_SLIDE_LIMIT} slides as text cards, ready to share.`,
    inputSchema: { source, output: z.string().optional().describe("Where to write the .html (default: returned inline)") },
  }, async ({ source, output }) => {
    try {
      const html = previewHtml(await outline(await loadBytes(source)), basename(source).replace(/\?.*$/, ""));
      if (output) { await writeFile(output, html); return { content: [{ type: "text", text: `Wrote ${output}` }] }; }
      return { content: [{ type: "text", text: html }] };
    } catch (e) { return fail(e); }
  });
  return server;
}

if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith("pptx-mcp")) {
  await createServer().connect(new StdioServerTransport());
}
