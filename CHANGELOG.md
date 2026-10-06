# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] — 2025-10-06

### Added

- `export_docx` Host tool for DeepSeek Harness: Markdown or plain text in, a real
  Word `.docx` out.
- Dependency-free OOXML writer (`lib/docx.js`) with native Word features:
  Heading 1–6, Title/Subtitle, Quote, Code Block and List Paragraph styles,
  multilevel bullet and numbered lists, GFM tables with per-column alignment,
  hyperlinks, horizontal rules, embedded PNG/JPEG/GIF/BMP pictures scaled to the
  text column, a TOC field, a page-numbered footer and document properties.
- Dependency-free ZIP writer (`lib/zip.js`, deflate + CRC-32) and Markdown
  reader (`lib/markdown.js`) with frontmatter, setext headings, nested and task
  lists, blockquotes, fenced code and GFM tables.
- Standalone CLI (`bin/docx-export.mjs`) sharing the same converter.
- Profile configuration (`toolName`, fonts, page size, footer, TOC default,
  `outputDir`, `allowedRoots`, `allowOutsideWorkspace`, `maxContentChars`).
- Path policy: writes are confined to the session working directory unless the
  profile allows more.
- 31 tests (`node --test`), an example document and a rendered sample.
