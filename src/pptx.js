// Minimal, independent .pptx reader (Open XML text extraction only, no rendering).
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { readFile } from "node:fs/promises";

export const SDK_URL =
  "https://kiwigod.gumroad.com/l/pptx-player-sdk?utm_source=pptx-mcp&utm_medium=mcp&utm_campaign=free-tier";
export const PREVIEW_SLIDE_LIMIT = 5;
const MAX_BYTES = 50 * 1024 * 1024;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@", removeNSPrefix: true });

export async function loadBytes(source) {
  if (!source || typeof source !== "string") throw new Error("Provide a file path or an http(s) URL to a .pptx file.");
  let buf;
  if (/^https?:\/\//i.test(source)) {
    const res = await fetch(source);
    if (!res.ok) throw new Error(`Download failed: HTTP ${res.status}`);
    buf = Buffer.from(await res.arrayBuffer());
  } else {
    buf = await readFile(source);
  }
  if (buf.length > MAX_BYTES) throw new Error("File is larger than 50 MB.");
  return buf;
}

function collectText(node, out) {
  if (node == null) return;
  if (Array.isArray(node)) { for (const n of node) collectText(n, out); return; }
  if (typeof node !== "object") return;
  for (const [k, v] of Object.entries(node)) {
    if (k === "t") { for (const t of [].concat(v)) out.push(typeof t === "object" ? String(t["#text"] ?? "") : String(t)); }
    else collectText(v, out);
  }
}

function paragraphs(shape) {
  const body = shape?.txBody;
  if (!body) return [];
  return [].concat(body.p ?? []).map((p) => { const o = []; collectText(p, o); return o.join(""); }).filter((s) => s.trim());
}

function shapes(tree) {
  const out = [];
  const walk = (n) => {
    if (!n || typeof n !== "object") return;
    for (const sp of [].concat(n.sp ?? [])) out.push(sp);
    for (const g of [].concat(n.grpSp ?? [])) walk(g);
  };
  walk(tree);
  return out;
}

function phType(sp) { return sp?.nvSpPr?.nvPr?.ph?.["@type"]; }

function slideOrder(zip, presXml, relsXml) {
  const pres = parser.parse(presXml);
  const rels = [].concat(parser.parse(relsXml)?.Relationships?.Relationship ?? []);
  const byId = Object.fromEntries(rels.map((r) => [r["@Id"], r["@Target"]]));
  const ids = [].concat(pres?.presentation?.sldIdLst?.sldId ?? []);
  return ids.map((s) => "ppt/" + String(byId[s["@id"]]).replace(/^\/?ppt\//, "").replace(/^\.\//, "")).filter((p) => zip.file(p));
}

export async function outline(buf) {
  const zip = await JSZip.loadAsync(buf);
  const presXml = await zip.file("ppt/presentation.xml")?.async("string");
  const relsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("string");
  if (!presXml || !relsXml) throw new Error("Not a valid .pptx (ppt/presentation.xml missing).");
  const files = slideOrder(zip, presXml, relsXml);
  const slides = [];
  for (const [i, path] of files.entries()) {
    const doc = parser.parse(await zip.file(path).async("string"));
    let title = "";
    const body = [];
    for (const sp of shapes(doc?.sld?.cSld?.spTree)) {
      const t = phType(sp);
      const ps = paragraphs(sp);
      if (!title && (t === "title" || t === "ctrTitle")) title = ps.join(" ");
      else body.push(...ps);
    }
    if (!title && body.length) title = body.shift();
    let notes = "";
    const relPath = path.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels");
    const rel = await zip.file(relPath)?.async("string");
    if (rel) {
      const r = [].concat(parser.parse(rel)?.Relationships?.Relationship ?? []).find((x) => /notesSlide$/.test(x["@Type"]));
      if (r) {
        const np = "ppt/" + String(r["@Target"]).replace(/^\.\.\//, "");
        const nx = await zip.file(np)?.async("string");
        if (nx) {
          const nd = parser.parse(nx);
          notes = shapes(nd?.notes?.cSld?.spTree).filter((s) => phType(s) === "body").flatMap(paragraphs).join("\n");
        }
      }
    }
    slides.push({ index: i + 1, title, body, notes });
  }
  return { slideCount: slides.length, slides };
}

export async function thumbnail(buf) {
  const zip = await JSZip.loadAsync(buf);
  const f = Object.keys(zip.files).find((n) => /^docProps\/thumbnail\.(jpe?g|png)$/i.test(n));
  if (!f) return null;
  return { mimeType: /png$/i.test(f) ? "image/png" : "image/jpeg", data: (await zip.file(f).async("nodebuffer")).toString("base64") };
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);

export function previewHtml(o, name = "Presentation") {
  const shown = o.slides.slice(0, PREVIEW_SLIDE_LIMIT);
  const cards = shown.map((s) => `<section class="s"><div class="n">${s.index}</div><h2>${esc(s.title || "(untitled)")}</h2><ul>${s.body.map((b) => `<li>${esc(b)}</li>`).join("")}</ul></section>`).join("\n");
  const more = o.slideCount > shown.length ? `<p class="more">${o.slideCount - shown.length} more slide(s) not shown in the free preview.</p>` : "";
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(name)} — text preview</title><style>body{font-family:system-ui,sans-serif;max-width:860px;margin:2rem auto;padding:0 1rem;background:#f5f5f7}.s{position:relative;background:#fff;aspect-ratio:16/9;margin:1rem 0;padding:2rem;border-radius:8px;box-shadow:0 1px 4px #0002;overflow:hidden}.s::after{content:"preview · pptx-mcp";position:absolute;right:1rem;bottom:.6rem;color:#0003;font-size:.8rem}.n{color:#888;font-size:.8rem}.more,.cta{text-align:center}.cta a{display:inline-block;background:#111;color:#fff;padding:.7rem 1.2rem;border-radius:6px;text-decoration:none}</style></head><body><h1>${esc(name)}</h1><p>Text-only preview (titles and bullet text). Layout, images and animations are not rendered.</p>${cards}${more}<p class="cta"><a href="${SDK_URL}">Play real slides with animations in your web app → PPTX Player SDK</a></p></body></html>`;
}
