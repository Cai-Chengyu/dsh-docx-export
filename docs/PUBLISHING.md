# Publishing dsh-docx-export

Everything below runs from the repository root. `git` and `npm` must be on your
`PATH` (on Windows use `npm.cmd` if script execution is restricted).

## 0. Pre-flight

```bash
node --check index.js && node --check lib/docx.js && node --check lib/markdown.js && node --check lib/zip.js
node test/markdown.test.mjs && node test/docx.test.mjs && node test/plugin.test.mjs
npm pack --dry-run          # prints the exact file list that would be published
```

## 1. Push to GitHub

The repository URLs already point at `github.com/Cai-Chengyu/dsh-docx-export`
(`package.json` → `repository`, `bugs`, `homepage`, both READMEs and this file).
Forking or renaming the repository means updating those five places.

```bash
git init -b main
git add .
git commit -m "dsh-docx-export 1.0.0: export generated text as a Word document"
git remote add origin https://github.com/Cai-Chengyu/dsh-docx-export.git
git push -u origin main
```

Then, on the repository page: **About → Topics** → add `dsh-plugin`
(plus `dsh`, `word`, `docx`, `markdown`, `deepseek-harness`). The topic is what
the DSH discovery sweeps look for, so it is worth adding even before an npm
release exists.

Tag the release so `dsh plugin add github:…` users can pin a version:

```bash
git tag v1.0.0 && git push origin v1.0.0
```

## 2. Publish to npm

```bash
npm login
npm publish --access public
```

The package has no dependencies and no build step, so the published tarball is
exactly the repository files listed in `package.json#files`.

Install from the registry afterwards:

```bash
dsh plugin --profile desktop add dsh-docx-export
```

## 3. Install from GitHub / locally (no npm release needed)

```bash
dsh plugin --profile desktop add github:Cai-Chengyu/dsh-docx-export
dsh plugin --profile desktop add /absolute/path/to/dsh-docx-export
dsh plugin --profile desktop add ./dsh-docx-export-1.0.0.tgz   # after npm pack
```

## 4. Submit to the community catalog

The catalog is a JSON registry, not a README list: open a PR against
[`dshworks/awesome-dsh-plugins`](https://github.com/dshworks/awesome-dsh-plugins)
that adds one entry to `data/plugins.json` (copy a neighbour, then check it
against `data/schema.json`). Run the repository's own validators from its root:

```bash
node scripts/validate.mjs
node scripts/render.mjs
```

Useful fields to get right:

- `name` / `repo` / `description` — honest one-liners; no keyword stuffing.
- `npm` — only once `npm view dsh-docx-export version` answers.
- `verifiedAgainst` — the DSH version you actually ran, e.g. `0.2.0-rc.2`.
- `status: verified` requires `evidence` pointing at the file that proves the
  install path, here `package.json#dsh.bundle.patch`.

The catalog also sweeps the `dsh-plugin` topic automatically, so a tagged,
documented repository is usually discovered without a PR.

## 5. Release checklist for the next version

1. Bump `version` in `package.json`, add a `CHANGELOG.md` section.
2. Re-run the tests and `npm pack --dry-run`.
3. `git commit`, `git tag vX.Y.Z`, `git push --follow-tags`.
4. `npm publish`.
5. In the catalog entry, bump `lastVerified` / `verifiedAgainst`.
