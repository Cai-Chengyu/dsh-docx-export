# dsh-docx-export

**English** | [中文](README.zh-CN.md)

Export generated text as a **real Word document** (`.docx`) from any DeepSeek Harness session — not a renamed Markdown file, but a proper OOXML package with native heading styles, multilevel lists, tables, code blocks, quotes, hyperlinks, embedded pictures, a table-of-contents field and a page-numbered footer. Word, WPS, LibreOffice and Google Docs all open it as-is.

Zero runtime dependencies (only Node's built-in `zlib`), no `@deepseek-ai/*` imports, so it survives harness version drift. It installs as a bundle and mounts on the next start.

```
You: turn that analysis into a Word document for me
DSH: (calls export_docx, writes G:\work\analysis-20251006-1209.docx)
```

## Features

| Markdown you write | What lands in Word |
|---|---|
| `# Title` … `###### Title` | Heading 1–6 styles (template-friendly, TOC-ready) |
| Paragraphs, `**bold**`, `*italic*`, `~~strike~~`, `` `code` `` | Real character formatting, CJK `w:eastAsia` fonts included |
| `- item` / `1. step`, nested by indentation | True multilevel bullet and numbered lists (`numbering.xml`) |
| `- [x] done` / `- [ ] todo` | ☑ / ☐ checklist items |
| ` ```lang ` fences | Monospaced, shaded, bordered code block; blank lines preserved |
| `> quote` | Quote style with a left rule |
| GFM tables with `:--:` alignment | Real Word tables: bold shaded header, per-column alignment |
| `[text](https://…)` | Clickable external hyperlinks |
| `![alt](picture.png)` | Embedded picture, scaled to the text column (PNG/JPEG/GIF/BMP) |
| `---` | Horizontal rule |
| Leading `---\ntitle: …\n---` | YAML frontmatter is ignored |

Plus: `toc: true` inserts a real TOC field (filled in by the word processor on open), page numbers in the footer by default, and title/author/timestamps in the document properties.

## Install

Pick one route, then **restart DeepSeek Harness** — a new bundle mounts on the next start.

```powershell
# 1) npm (recommended once published)
dsh plugin --profile desktop add dsh-docx-export

# 2) straight from GitHub
dsh plugin --profile desktop add github:Cai-Chengyu/dsh-docx-export

# 3) a local directory or a packed tarball
dsh plugin --profile desktop add G:\path\to\dsh-docx-export
dsh plugin --profile desktop add .\dsh-docx-export-1.0.0.tgz
```

If `dsh` is not on `PATH`, call the CLI shipped with the app: `"<DeepSeek Harness>\resources\runtime\cli\bin\dsh.cmd"`. `scripts/install.ps1` (Windows) and `scripts/install.sh` (macOS, Linux) wrap that command and verify the bundle got registered.

Manual alternative: `pnpm add dsh-docx-export` inside the profile directory, then append `dsh-docx-export` to `dsh.profile.bundles` in that directory's `package.json`.

## Use

Just ask in the session — the model calls the tool for you:

> Export what you just wrote as a Word document titled "Q1 analysis", with a table of contents.

Or call `export_docx` directly:

```json
{
  "content": "# Title\n\nBody **bold**\n\n| col | value |\n|:--|--:|\n| a | 1 |\n",
  "title": "Q1 analysis",
  "outputPath": "G:\\dsh\\analysis.docx",
  "toc": true,
  "pageSize": "a4"
}
```

### Tool parameters

| Parameter | Type | Meaning |
|---|---|---|
| `content` | string | Document body, parsed as Markdown by default |
| `sourcePath` | string | Read the body from an existing `.md` / `.txt` file instead |
| `outputPath` | string | Target `.docx` (absolute, or relative to the session working directory). Defaults to `<outputDir>/<title>-<timestamp>.docx` |
| `title` / `subtitle` / `author` | string | Title block and document properties; `title` defaults to the first heading |
| `format` | `markdown` \| `text` | `text` keeps every line literal |
| `toc` | boolean | Insert a table-of-contents field |
| `pageSize` | `a4` \| `letter` \| `a5` | Paper size, A4 by default |
| `landscape` | boolean | Landscape orientation |
| `footer` | boolean | Page numbers in the footer, on by default |
| `overwrite` | boolean | On by default; `false` appends `-1`, `-2`, … instead of replacing |

The result text reports the absolute path plus paragraph, heading, table, image and size counts.

### CLI (usable outside DSH)

```bash
node bin/docx-export.mjs report.md -o report.docx --title "Quarterly report" --toc
cat notes.md | node bin/docx-export.mjs - -o notes.docx --text
```

`--page a4|letter|a5`, `--landscape`, `--no-footer`, `--font`, `--east-asia`, `--font-size` are available too.

## Configuration (optional)

Override defaults in the profile's `cordis.patch.yml`:

```yaml
- id: docx-export
  name: dsh-docx-export
  config:
    toolName: export_docx          # rename on collision
    outputDir: "D:/exports"        # default output directory ("" = session working directory)
    fontFamily: Calibri            # Latin body font
    eastAsiaFont: 微软雅黑          # East Asian body font
    monoFont: Consolas             # code font
    fontSize: 11                   # body size in points
    pageSize: a4
    toc: false                     # insert a TOC by default
    footer: true                   # page numbers by default
    allowOutsideWorkspace: false    # may the tool write outside the session workspace
    allowedRoots: []               # extra directories that are allowed
    maxContentChars: 400000        # per-export body limit
```

### Write scope (safe by default)

The tool reads and writes only inside the **session working directory** unless the profile says otherwise. An out-of-scope path is refused with an actionable message; grant specific locations with `allowedRoots`, or set `allowOutsideWorkspace: true`. Pictures are embedded only when they live inside the allowed scope.

## Development

```bash
node test/markdown.test.mjs   # Markdown reader (12 tests)
node test/docx.test.mjs       # OOXML writer + ZIP integrity (8 tests)
node test/plugin.test.mjs     # tool registration, path policy, real files (11 tests)
npm test                      # = node --test test/ (used by CI)
npm run demo                  # convert examples/sample.md to examples/sample.docx
```

The suite ships its own ZIP central-directory reader with CRC verification, an XML well-formedness checker and a miniature PNG encoder, so there are no dev dependencies either.

## Publishing and sharing

- Add the GitHub topic **`dsh-plugin`** so discovery sweeps pick the repository up.
- Submit it to the community catalog with a PR to `data/plugins.json` in [dshworks/awesome-dsh-plugins](https://github.com/dshworks/awesome-dsh-plugins) (fields: that repo's `data/schema.json`).
- Publish to npm with `npm publish`, after which `dsh plugin --profile <profile> add dsh-docx-export` installs it from the registry.
- Details: [docs/PUBLISHING.md](docs/PUBLISHING.md).

## Compatibility

- Node ≥ 20 (`node:zlib`, ESM).
- Any DSH from 0.1.7 on: the plugin declares no `@deepseek-ai/*` peer dependencies and imports no host package — it uses the runtime's raw JSON-Schema tool-registration path.
- Measured on DSH 0.2.0-rc.2 (Windows desktop): install → mount → generate → structural check → LibreOffice render.

## License

[MIT](LICENSE)
