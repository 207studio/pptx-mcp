import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";

test("MCPB manifest matches the package and server entry point", () => {
  const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"));
  const manifest = JSON.parse(readFileSync(new URL("../manifest.json", import.meta.url), "utf8"));
  assert.equal(manifest.name, pkg.name.split("/").at(-1));
  assert.equal(manifest.version, pkg.version);
  assert.equal(manifest.server.entry_point, pkg.main);
  assert.equal(manifest.server.entry_point, pkg.bin[manifest.name]);
  assert.ok(existsSync(new URL(`../${manifest.server.entry_point}`, import.meta.url)));
  assert.equal(manifest.server.type, "node");
  assert.equal(manifest.server.mcp_config.command, "node");
  assert.deepEqual(manifest.server.mcp_config.args, ["${__dirname}/" + manifest.server.entry_point]);
  assert.equal(manifest.compatibility.runtimes.node, pkg.engines.node);
});
