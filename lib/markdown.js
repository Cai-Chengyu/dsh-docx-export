/**
 * Markdown reader for the Word exporter.
 *
 * The parser is deliberately small and dependency-free: it understands the
 * constructs that survive a trip into Word (headings, paragraphs, emphasis,
 * links, lists, quotes, fenced code, GFM tables, rules, local images) and
 * degrades everything else to text instead of guessing.
 */

/** Named HTML entities worth decoding; numeric escapes are handled by rule. */
const ENTITIES = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: '\u00a0',
  hellip: '\u2026',
  mdash: '\u2014',
  ndash: '\u2013',
  copy: '\u00a9',
  reg: '\u00ae',
  trade: '\u2122',
  laquo: '\u00ab',
  raquo: '\u00bb',
}

function decodeEntities(text) {
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (match, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10)
      if (Number.isFinite(code) && code > 0 && code <= 0x10ffff) {
        try {
          return String.fromCodePoint(code)
        } catch {
          return match
        }
      }
      return match
    }
    const named = ENTITIES[body.toLowerCase()]
    return named === undefined ? match : named
  })
}

const IMAGE_ONLY = /^!\[[^\]]*\]\([^)\s]+(?:\s+"[^"]*")?\)$/

/** Inline token table, ordered so longer emphasis forms win over shorter ones. */
const INLINE_TOKEN = new RegExp(
  [
    '(?<code>`+)(?<codeText>[\\s\\S]*?)\\k<code>',
    '!\\[(?<imgAlt>[^\\]]*)\\]\\((?<imgSrc>[^)\\s]+)(?:\\s+"[^"]*")?\\)',
    '\\[(?<linkText>[^\\]]*)\\]\\((?<linkHref>[^)\\s]+)(?:\\s+"[^"]*")?\\)',
    '\\*\\*\\*(?<boldItalic>[\\s\\S]+?)\\*\\*\\*',
    '\\*\\*(?<bold>[\\s\\S]+?)\\*\\*',
    '__(?<bold2>[\\s\\S]+?)__',
    '~~(?<strike>[\\s\\S]+?)~~',
    '(?<![\\w*])\\*(?<italic>[^*\\n]+?)\\*(?![\\w*])',
    '(?<![\\w_])_(?<italic2>[^_\\n]+?)_(?![\\w_])',
    '<(?<autoHref>(?:https?|mailto):[^>\\s]+)>',
    '\\\\(?<escaped>[\\\\`*_{}\\[\\]()#+\\-.!~>|])',
  ].join('|'),
  'g',
)

/**
 * Split inline Markdown into styled runs.
 *
 * @param {string} text
 * @param {{ bold?: boolean, italic?: boolean, strike?: boolean }} [style]
 * @returns {Array<Record<string, unknown>>}
 */
export function parseInline(text, style = {}) {
  const runs = []
  const push = (run) => {
    const previous = runs[runs.length - 1]
    if (
      previous &&
      run.type === 'text' &&
      previous.type === 'text' &&
      previous.bold === run.bold &&
      previous.italic === run.italic &&
      previous.strike === run.strike &&
      previous.code === run.code &&
      previous.href === run.href
    ) {
      previous.text += run.text
      return
    }
    runs.push(run)
  }
  const base = {
    bold: style.bold === true,
    italic: style.italic === true,
    strike: style.strike === true,
  }

  let last = 0
  INLINE_TOKEN.lastIndex = 0
  // The pattern is shared between recursive calls, so the scan position is
  // restored explicitly before every step instead of trusting `lastIndex`.
  for (let match = INLINE_TOKEN.exec(text); match !== null; ) {
    if (match.index > last) push({ type: 'text', text: decodeEntities(text.slice(last, match.index)), ...base })
    last = match.index + match[0].length
    const groups = match.groups ?? {}
    if (groups.codeText !== undefined) {
      push({ type: 'text', text: groups.codeText, ...base, code: true, href: undefined })
    } else if (groups.imgSrc !== undefined) {
      push({ type: 'image', src: groups.imgSrc, alt: groups.imgAlt ?? '', ...base })
    } else if (groups.linkHref !== undefined) {
      const label = groups.linkText === '' ? groups.linkHref : groups.linkText
      for (const run of parseInline(label, base)) push({ ...run, href: groups.linkHref })
    } else if (groups.boldItalic !== undefined) {
      for (const run of parseInline(groups.boldItalic, { ...base, bold: true, italic: true })) push(run)
    } else if (groups.bold !== undefined || groups.bold2 !== undefined) {
      for (const run of parseInline(groups.bold ?? groups.bold2, { ...base, bold: true })) push(run)
    } else if (groups.strike !== undefined) {
      for (const run of parseInline(groups.strike, { ...base, strike: true })) push(run)
    } else if (groups.italic !== undefined || groups.italic2 !== undefined) {
      for (const run of parseInline(groups.italic ?? groups.italic2, { ...base, italic: true })) push(run)
    } else if (groups.autoHref !== undefined) {
      const href = groups.autoHref.startsWith('mailto:') ? groups.autoHref : groups.autoHref
      push({ type: 'text', text: groups.autoHref.replace(/^mailto:/, ''), ...base, href })
    } else if (groups.escaped !== undefined) {
      push({ type: 'text', text: groups.escaped, ...base })
    }
    INLINE_TOKEN.lastIndex = last
    match = INLINE_TOKEN.exec(text)
  }
  if (last < text.length) push({ type: 'text', text: decodeEntities(text.slice(last)), ...base })
  return runs
}

/** Remove inline HTML that Word would otherwise print verbatim. */
function stripHtml(text) {
  return text.replace(/<\/?[a-zA-Z][^>]*>/g, '')
}

const RE_HEADING = /^(#{1,6})(?:\s+(.*?))?\s*#*\s*$/
const RE_FENCE = /^\s{0,3}(`{3,}|~{3,})\s*([^`\s]*)\s*$/
const RE_HR = /^\s{0,3}([-*_])(?:\s*\1){2,}\s*$/
const RE_QUOTE = /^\s{0,3}>\s?(.*)$/
const RE_ITEM = /^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/
const RE_TASK = /^\[([ xX])\]\s+(.*)$/
const RE_TABLE_DIVIDER = /^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?\s*$/
const RE_SETEXT = /^\s{0,3}(=+|-+)\s*$/

function splitTableRow(line) {
  let text = line.trim()
  if (text.startsWith('|')) text = text.slice(1)
  if (text.endsWith('|') && !text.endsWith('\\|')) text = text.slice(0, -1)
  const cells = []
  let current = ''
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i]
    if (char === '\\' && text[i + 1] === '|') {
      current += '|'
      i += 1
    } else if (char === '|') {
      cells.push(current.trim())
      current = ''
    } else {
      current += char
    }
  }
  cells.push(current.trim())
  return cells
}

function alignmentOf(cell) {
  const left = cell.startsWith(':')
  const right = cell.endsWith(':')
  if (left && right) return 'center'
  if (right) return 'right'
  if (left) return 'left'
  return null
}

/** Split a document into `---` frontmatter and the body that follows it. */
function splitFrontmatter(source) {
  const normalized = source.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n')
  if (!/^---\n/.test(normalized)) return { frontmatter: '', body: normalized }
  const end = normalized.indexOf('\n---', 3)
  if (end === -1) return { frontmatter: '', body: normalized }
  const after = normalized.indexOf('\n', end + 1)
  return { frontmatter: normalized.slice(4, end), body: after === -1 ? '' : normalized.slice(after + 1) }
}

/**
 * Parse Markdown into block nodes.
 *
 * @param {string} source
 * @returns {{ frontmatter: string, blocks: Array<Record<string, unknown>> }}
 */
export function parseMarkdown(source) {
  const { frontmatter, body } = splitFrontmatter(String(source ?? ''))
  const lines = body.split('\n')
  const blocks = []

  let index = 0
  while (index < lines.length) {
    const line = lines[index]

    if (line.trim() === '') {
      index += 1
      continue
    }

    const fence = RE_FENCE.exec(line)
    if (fence) {
      const marker = fence[1][0]
      const lang = (fence[2] ?? '').trim()
      const collected = []
      index += 1
      while (index < lines.length) {
        const closing = RE_FENCE.exec(lines[index])
        if (closing && closing[1][0] === marker && closing[1].length >= fence[1].length) break
        collected.push(lines[index])
        index += 1
      }
      index += 1
      blocks.push({ type: 'code', lang, text: collected.join('\n') })
      continue
    }

    const heading = RE_HEADING.exec(line)
    if (heading && !RE_HR.test(line)) {
      blocks.push({ type: 'heading', level: heading[1].length, text: stripHtml(heading[2] ?? '') })
      index += 1
      continue
    }

    if (index + 1 < lines.length && RE_SETEXT.test(lines[index + 1]) && line.trim() !== '') {
      const level = lines[index + 1].trim().startsWith('=') ? 1 : 2
      blocks.push({ type: 'heading', level, text: stripHtml(line.trim()) })
      index += 2
      continue
    }

    if (RE_HR.test(line)) {
      blocks.push({ type: 'hr' })
      index += 1
      continue
    }

    if (RE_QUOTE.test(line)) {
      const collected = []
      while (index < lines.length && (RE_QUOTE.test(lines[index]) || (collected.length > 0 && lines[index].trim() !== '' && !RE_ITEM.test(lines[index]) && !RE_HEADING.test(lines[index])))) {
        const quoted = RE_QUOTE.exec(lines[index])
        collected.push(quoted ? quoted[1] : lines[index])
        index += 1
      }
      blocks.push({ type: 'quote', blocks: parseMarkdown(collected.join('\n')).blocks })
      continue
    }

    if (line.includes('|') && index + 1 < lines.length && RE_TABLE_DIVIDER.test(lines[index + 1])) {
      const header = splitTableRow(line)
      const align = splitTableRow(lines[index + 1]).map(alignmentOf)
      index += 2
      const rows = []
      while (index < lines.length && lines[index].includes('|') && lines[index].trim() !== '') {
        rows.push(splitTableRow(lines[index]))
        index += 1
      }
      const width = Math.max(header.length, ...rows.map((row) => row.length), align.length)
      const pad = (row) => Array.from({ length: width }, (_, i) => row[i] ?? '')
      blocks.push({ type: 'table', align: pad(align).map((value) => value || null), header: pad(header), rows: rows.map(pad) })
      continue
    }

    if (RE_ITEM.test(line)) {
      const list = parseList(lines, index)
      blocks.push(list.node)
      index = list.index
      continue
    }

    // Paragraph: gather until a blank line or the start of another block.
    const collected = []
    while (index < lines.length) {
      const current = lines[index]
      if (current.trim() === '') break
      if (collected.length > 0) {
        if (RE_HEADING.test(current) || RE_FENCE.test(current) || RE_ITEM.test(current)) break
        if (RE_QUOTE.test(current) && !RE_QUOTE.test(collected[0])) break
        if (RE_HR.test(current)) break
        if (current.includes('|') && index + 1 < lines.length && RE_TABLE_DIVIDER.test(lines[index + 1])) break
      }
      collected.push(current)
      index += 1
    }

    const text = collected.map((value) => value.trim()).join('\n')
    if (IMAGE_ONLY.test(text.trim())) {
      const [run] = parseInline(text.trim())
      blocks.push({ type: 'image', src: run.src, alt: run.alt })
      continue
    }
    blocks.push({ type: 'paragraph', text: stripHtml(text) })
  }

  return { frontmatter, blocks }
}

/** Parse one list (with nested children) starting at `start`. */
function parseList(lines, start) {
  const first = RE_ITEM.exec(lines[start])
  const ordered = /\d/.test(first[2])
  const baseIndent = first[1].replace(/\t/g, '    ').length
  const items = []
  let index = start
  const startNumber = ordered ? Number.parseInt(first[2], 10) : null

  while (index < lines.length) {
    const match = RE_ITEM.exec(lines[index])
    if (!match) break
    const indent = match[1].replace(/\t/g, '    ').length
    if (indent < baseIndent) break
    if (indent > baseIndent) {
      // A deeper marker belongs to the previous item as a nested list.
      if (items.length === 0) break
      const nested = parseList(lines, index)
      items[items.length - 1].blocks.push(nested.node)
      index = nested.index
      continue
    }
    const isOrdered = /\d/.test(match[2])
    if (isOrdered !== ordered) break

    let text = match[3]
    let checked = null
    const task = RE_TASK.exec(text)
    if (task) {
      checked = task[1].toLowerCase() === 'x'
      text = task[2]
    }
    const item = { text: stripHtml(text.trim()), checked, blocks: [] }
    index += 1

    // Continuation lines: indented, and not a new marker or a nested list.
    while (index < lines.length) {
      const current = lines[index]
      if (current.trim() === '') break
      const nextItem = RE_ITEM.exec(current)
      const indentOfLine = /^\s*/.exec(current)[0].replace(/\t/g, '    ').length
      if (nextItem && indentOfLine <= baseIndent) break
      if (indentOfLine > baseIndent && !nextItem && !RE_FENCE.test(current)) {
        item.text += `\n${stripHtml(current.trim())}`
        index += 1
        continue
      }
      break
    }
    items.push(item)
  }

  return { node: { type: 'list', ordered, start: startNumber ?? 1, items }, index }
}
