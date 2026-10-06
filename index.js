/**
 * dsh-docx-export — a Host plugin for DeepSeek Harness that turns generated
 * text into a real Word document.
 *
 * The plugin registers one tool, `export_docx`, and keeps every host-facing
 * dependency duck-typed: it imports no `@deepseek-ai/*` package, so the same
 * file mounts in a bundled profile, in a profile-installed bundle, and in a
 * checkout, across harness versions.
 *
 * File policy: the tool writes inside the session working directory by
 * default. Paths outside it are refused unless the profile config allows them
 * (`allowOutsideWorkspace: true`) or names them (`allowedRoots: [...]`).
 */
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { buildDocx } from './lib/docx.js'

export const name = 'docx-export'
export const inject = ['tools']

/** Profile-side defaults; every key can be overridden in cordis.patch.yml. */
export const DEFAULT_CONFIG = {
  toolName: 'export_docx',
  outputDir: '',
  fontFamily: 'Calibri',
  eastAsiaFont: '\u5fae\u8f6f\u96c5\u9ed1',
  monoFont: 'Consolas',
  fontSize: 11,
  pageSize: 'a4',
  toc: false,
  footer: true,
  allowOutsideWorkspace: false,
  allowedRoots: [],
  maxContentChars: 400_000,
}

const TOOL_DESCRIPTION = [
  'Export text as a Word document (.docx). Use whenever the user asks to save, export, download or hand over generated content as a Word file',
  '(导出/保存为 Word、docx、文档、报告).',
  'Pass the complete document body as `content` in Markdown: headings (#, ##), **bold**, *italic*, `code`, fenced code blocks, bullet and numbered lists,',
  'task lists, > quotes, GFM tables, --- rules and ![alt](picture.png) images all become native Word features (styles, numbering, tables, embedded pictures).',
  'Write the file into the session working directory unless the user names another location. Returns the absolute path of the written file.',
].join(' ')

const PARAMETERS = {
  type: 'object',
  properties: {
    content: {
      type: 'string',
      description:
        'The complete document body. Markdown by default (tables, lists, code blocks and local images included); pass format="text" for literal plain text.',
    },
    sourcePath: {
      type: 'string',
      description:
        'Read the document body from this existing .md/.txt file instead of `content`. Relative paths resolve against the session working directory.',
    },
    outputPath: {
      type: 'string',
      description:
        'Where to write the .docx. Absolute, or relative to the session working directory. Defaults to <outputDir>/<title>-<timestamp>.docx. A missing .docx extension is added.',
    },
    title: {
      type: 'string',
      description:
        'Document title: rendered as the opening title block and stored in the document properties. Defaults to the first heading of the content.',
    },
    subtitle: { type: 'string', description: 'Optional subtitle line under the title.' },
    author: { type: 'string', description: 'Optional author, shown under the title and stored in the document properties.' },
    format: {
      type: 'string',
      enum: ['markdown', 'text'],
      description: 'How to read `content`. Defaults to "markdown"; "text" keeps every line literal.',
    },
    toc: { type: 'boolean', description: 'Insert a table-of-contents field (Word fills it in on open).' },
    pageSize: { type: 'string', enum: ['a4', 'letter', 'a5'], description: 'Paper size. Defaults to A4.' },
    landscape: { type: 'boolean', description: 'Landscape orientation.' },
    footer: { type: 'boolean', description: 'Page numbers in the footer. Defaults to true.' },
    overwrite: {
      type: 'boolean',
      description: 'Overwrite an existing file at outputPath. Defaults to true; false appends -1, -2, ... instead.',
    },
  },
  required: [],
  additionalProperties: false,
}

const INVALID_FILENAME = /[<>:"/\\|?*\u0000-\u001f]/g

/** Turn a title into a safe, readable file name stem. */
export function fileStem(title, fallback = 'document') {
  const cleaned = String(title ?? '')
    .replace(INVALID_FILENAME, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
  const stem = cleaned.slice(0, 60).trim()
  return stem.length > 0 ? stem : fallback
}

/** `20250203-040506`, local time, stable across platforms. */
export function timestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0')
  return (
    `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-` +
    `${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`
  )
}

function isInside(target, root) {
  const rel = relative(resolve(root), resolve(target))
  return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel))
}

/** Human-readable byte size for the tool result line. */
function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

/** First ATX heading of a Markdown body, used when no title was given. */
function firstHeading(markdown) {
  const match = /^[ \t]{0,3}#{1,6}[ \t]+(.+?)[ \t]*#*[ \t]*$/m.exec(markdown)
  return match ? match[1].trim() : ''
}

/** Add `.docx` when the caller gave no (or an unusable) extension. */
function withDocxExtension(target) {
  return extname(target).toLowerCase() === '.docx' ? target : `${target}.docx`
}

function uniquePath(target) {
  if (!existsSync(target)) return target
  const ext = extname(target)
  const stem = target.slice(0, target.length - ext.length)
  for (let i = 1; i < 1000; i += 1) {
    const candidate = `${stem}-${i}${ext}`
    if (!existsSync(candidate)) return candidate
  }
  return `${stem}-${Date.now()}${ext}`
}

function describeStats(stats, bytes) {
  const parts = [`${stats.paragraphs} paragraphs`]
  if (stats.headings > 0) parts.push(`${stats.headings} headings`)
  if (stats.lists > 0) parts.push(`${stats.lists} lists`)
  if (stats.tables > 0) parts.push(`${stats.tables} tables`)
  if (stats.codeBlocks > 0) parts.push(`${stats.codeBlocks} code blocks`)
  if (stats.images > 0) parts.push(`${stats.images} images`)
  parts.push(formatSize(bytes))
  return parts.join(', ')
}

/**
 * Cordis entry point.
 *
 * @param {object} ctx host context (only `ctx.tools.register` is used)
 * @param {object} [config] profile config from cordis.patch.yml
 */
export function apply(ctx, config = {}) {
  const settings = { ...DEFAULT_CONFIG, ...(config ?? {}) }
  const toolName = typeof settings.toolName === 'string' && settings.toolName.trim() !== ''
    ? settings.toolName.trim()
    : DEFAULT_CONFIG.toolName
  const defaultOutputDir = settings.outputDir ? resolve(String(settings.outputDir)) : process.cwd()

  /** Roots the tool may read from and write to. */
  const roots = () => {
    if (settings.allowOutsideWorkspace === true) return null
    const configured = Array.isArray(settings.allowedRoots) ? settings.allowedRoots : []
    return [process.cwd(), ...configured.filter((item) => typeof item === 'string' && item !== '').map((item) => resolve(item))]
  }

  const assertAllowed = (target, action) => {
    const allowed = roots()
    if (allowed === null) return
    if (allowed.some((root) => isInside(target, root))) return
    throw new Error(
      `Refusing to ${action} "${target}": it is outside the session workspace (${allowed.join(', ')}). ` +
        'Choose a path inside it, or allow more locations with `allowOutsideWorkspace: true` / `allowedRoots` in the plugin config.',
    )
  }

  const resolveInputPath = (value) => {
    const target = isAbsolute(String(value)) ? resolve(String(value)) : resolve(process.cwd(), String(value))
    assertAllowed(target, 'read')
    return target
  }

  const definition = {
    name: toolName,
    description: TOOL_DESCRIPTION,
    parameters: PARAMETERS,
    // A large document is parsed and deflated synchronously in this process.
    timeoutMs: 120_000,
    isConcurrencySafe: () => true,
    presentCall: (args) => ({
      card: 'generic',
      title: args?.outputPath ? `Word: ${args.outputPath}` : 'Word document',
      kind: 'execute',
      rawInput: args,
    }),
    async execute(args, exec) {
      const params = args ?? {}
      if (exec?.signal?.aborted) throw new Error('Cancelled before the document was written.')

      let body = typeof params.content === 'string' ? params.content : ''
      let baseDir = process.cwd()
      if (typeof params.sourcePath === 'string' && params.sourcePath.trim() !== '') {
        const source = resolveInputPath(params.sourcePath)
        if (!existsSync(source) || !statSync(source).isFile()) {
          throw new Error(`sourcePath does not exist or is not a file: ${source}`)
        }
        body = readFileSync(source, 'utf8')
        baseDir = dirname(source)
      }
      if (body.trim() === '') {
        throw new Error('Nothing to export: pass `content`, or a `sourcePath` that holds the document text.')
      }
      if (body.length > settings.maxContentChars) {
        throw new Error(
          `Document body is ${body.length} characters, over the ${settings.maxContentChars} limit. ` +
            'Split it into several exports, or raise `maxContentChars` in the plugin config.',
        )
      }

      const format = params.format === 'text' ? 'text' : 'markdown'
      const title = typeof params.title === 'string' && params.title.trim() !== ''
        ? params.title.trim()
        : format === 'markdown'
          ? firstHeading(body)
          : ''

      const requested = typeof params.outputPath === 'string' && params.outputPath.trim() !== ''
        ? params.outputPath.trim()
        : join(defaultOutputDir, `${fileStem(title)}-${timestamp()}.docx`)
      const absolute = isAbsolute(requested) ? resolve(requested) : resolve(process.cwd(), requested)
      const target = withDocxExtension(absolute)
      assertAllowed(target, 'write')
      const finalPath = params.overwrite === false ? uniquePath(target) : target

      const loadImage = (src) => {
        try {
          if (/^(https?|data):/i.test(src)) return null
          const candidate = isAbsolute(src) ? src : resolve(baseDir, src)
          if (!existsSync(candidate) || !statSync(candidate).isFile()) return null
          const allowed = roots()
          if (allowed !== null && !allowed.some((root) => isInside(candidate, root))) return null
          const ext = extname(candidate).slice(1).toLowerCase()
          return { data: readFileSync(candidate), ext }
        } catch {
          return null
        }
      }

      const { buffer, stats } = await buildDocx(body, {
        title,
        subtitle: typeof params.subtitle === 'string' ? params.subtitle : '',
        author: typeof params.author === 'string' ? params.author : '',
        fontFamily: settings.fontFamily,
        eastAsiaFont: settings.eastAsiaFont,
        monoFont: settings.monoFont,
        fontSize: Number(settings.fontSize),
        pageSize: typeof params.pageSize === 'string' ? params.pageSize : settings.pageSize,
        landscape: params.landscape === true,
        toc: params.toc === true || (params.toc === undefined && settings.toc === true),
        footer: params.footer === undefined ? settings.footer !== false : params.footer !== false,
        plainText: format === 'text',
        loadImage,
      })

      if (exec?.signal?.aborted) throw new Error('Cancelled before the document was written.')
      mkdirSync(dirname(finalPath), { recursive: true })
      writeFileSync(finalPath, buffer)

      const overwritten = params.overwrite === false && finalPath !== target ? ` (kept the existing ${target})` : ''
      return (
        `Wrote Word document: ${finalPath}\n` +
        `${describeStats(stats, buffer.length)}${overwritten}` +
        (stats.images === 0 && /!\[[^\]]*\]\(/.test(body) ? '\nNote: pictures were referenced but not embedded (missing file or unsupported format).' : '')
      )
    },
  }

  try {
    ctx.tools.register(definition)
  } catch (error) {
    console.error(`[docx-export] ${toolName} registration skipped: ${error}`)
  }
}

export { buildDocx }
export default { name, inject, apply }
