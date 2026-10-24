# pptx-mcp

Read, outline and preview PowerPoint (.pptx) files from any AI agent. Free and open source (MIT) [Model Context Protocol](https://modelcontextprotocol.io) server by 207studio.

## What it does

| Tool | What you get |
|---|---|
| `pptx_outline` | A slide-by-slide outline of the deck as JSON: title, body text and speaker notes. |
| `pptx_thumbnail` | The cover thumbnail PowerPoint saved inside the file, when the file has one. |
| `pptx_web_preview` | A self-contained HTML page showing the first 5 slides as text cards, ready to share. |

`source` can be a local path or an http(s) URL. Files up to 50 MB.

## Install

```json
{
  "mcpServers": {
    "pptx": { "command": "npx", "args": ["-y", "github:207studio/pptx-mcp"] }
  }
}
```

Or clone and run `npm install && node src/index.js` (Node 18+).

## Bot ideas

- **Deck summarizer**: outline a .pptx you got by email and draft a reply.
- **Cover gallery**: pull the cover thumbnail of every deck in a folder.
- **Share preview**: turn a deck into one HTML page you can share as a link.

## Need real slides in your own app?

pptx-mcp reads text only. It does not render layouts, images, animations or transitions.

To play .pptx files inside your web app with layout, animations and transitions, use **[PPTX Player SDK](https://kiwigod.gumroad.com/l/pptx-player-sdk?utm_source=github&utm_medium=readme&utm_campaign=pptx-mcp)**. It is a one-time purchase with source included.

## Development

```
npm install
npm test
```

## MCPB bundle

```sh
npm ci
npm run bundle
```

Upload `dist/pptx-mcp.mcpb` to Smithery as "Local (MCPB Bundle)". The bundle includes production dependencies.

## License

MIT © 207studio
