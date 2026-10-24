import { execFileSync } from "node:child_process";
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const output = join(root, "dist", "pptx-mcp.mcpb");
const mcpbRoot = join(root, "node_modules", "@anthropic-ai", "mcpb");
const mcpb = JSON.parse(readFileSync(join(mcpbRoot, "package.json"), "utf8"));
const cli = join(mcpbRoot, typeof mcpb.bin === "string" ? mcpb.bin : mcpb.bin.mcpb);
const stage = mkdtempSync(join(tmpdir(), "pptx-mcp-bundle-"));

try {
  for (const file of ["src", "package.json", "package-lock.json", "manifest.json", "README.md", "LICENSE", ".mcpbignore"]) {
    cpSync(join(root, file), join(stage, file), { recursive: true });
  }
  execFileSync(process.execPath, [process.env.npm_execpath, "ci", "--omit=dev", "--ignore-scripts", "--no-audit", "--no-fund"], {
    cwd: stage,
    stdio: "inherit",
  });
  mkdirSync(join(root, "dist"), { recursive: true });
  execFileSync(process.execPath, [cli, "pack", stage, output], { stdio: "inherit" });
} finally {
  rmSync(stage, { recursive: true, force: true });
}
