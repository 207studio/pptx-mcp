import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import PptxGenJS from "pptxgenjs";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { loadBytes, outline, previewHtml } from "../src/pptx.js";

async function sample(n = 7) {
  const p = new PptxGenJS();
  for (let i = 1; i <= n; i++) {
    const s = p.addSlide();
    s.addText(`Title ${i}`, { placeholder: "title", x: 0.5, y: 0.3, w: 9, h: 1 });
    s.addText(`Body ${i} & <b>`, { x: 0.5, y: 2, w: 9, h: 1 });
    s.addNotes(`Note ${i}`);
  }
  const f = join(await mkdtemp(join(tmpdir(), "pptx-")), "deck.pptx");
  await p.writeFile({ fileName: f });
  return f;
}

test("outline reads slides in order with text and notes", async () => {
  const o = await outline(await loadBytes(await sample(3)));
  assert.equal(o.slideCount, 3);
  assert.equal(o.slides[0].title, "Title 1");
  assert.ok(o.slides[1].body.join(" ").includes("Body 2"));
  assert.ok(o.slides[2].notes.includes("Note 3"));
});

test("preview limits slides, escapes HTML and carries utm link", async () => {
  const html = previewHtml(await outline(await loadBytes(await sample(7))), "deck");
  assert.ok(html.includes("2 more slide(s)"));
  assert.ok(html.includes("&lt;b&gt;"));
  assert.ok(html.includes("utm_source=pptx-mcp"));
});

test("rejects non-pptx", async () => {
  await assert.rejects(outline(Buffer.from("nope")));
});

test("MCP server lists and runs tools over stdio", async () => {
  const f = await sample(2);
  const c = new Client({ name: "t", version: "0" });
  await c.connect(new StdioClientTransport({ command: "node", args: ["src/index.js"] }));
  const names = (await c.listTools()).tools.map((t) => t.name).sort();
  assert.deepEqual(names, ["pptx_outline", "pptx_thumbnail", "pptx_web_preview"]);
  const r = await c.callTool({ name: "pptx_outline", arguments: { source: f } });
  assert.equal(JSON.parse(r.content[0].text).slideCount, 2);
  const t = await c.callTool({ name: "pptx_thumbnail", arguments: { source: f } });
  assert.ok(t.content[0].type === "image" || t.content[0].text.includes("utm_source"));
  await c.close();
});
