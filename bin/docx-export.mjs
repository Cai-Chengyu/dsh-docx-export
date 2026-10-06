#!/usr/bin/env node
/**
 * dsh-docx-export CLI — the same converter the DSH tool uses, for terminals,
 * scripts and CI.
 *
 *   dsh-docx-export report.md -o report.docx --title "季度报告" --toc
 *   cat notes.md | dsh-docx-export - -o notes.docx
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, extname, isAbsolute, resolve } from 'node:path'
import { buildDocx } from '../lib/docx.js'

const USAGE = `dsh-docx-export — convert Markdown or text into a Word .docx

Usage:
  dsh-docx-export <input.md|-> [-o output.docx] [options]

Options:
  -o, --out <file>        Output path (default: <input>-<timestamp>.docx)
  -t, --title <text>      Document title (default: first heading)
      --subtitle <text>   Subtitle line
      --author <text>     Author / creator
      --text              Treat the input as literal plain text
      --toc               Insert a table-of-contents field
      --no-footer         Omit the page-numbered footer
      --page <size>       a4 | letter | a5 (default: a4)
      --landscape         Landscape orientation
      --font <name>       Latin body font (default: Calibri)
      --east-asia <name>  East Asian body font (default: 微软雅黑)
      --font-size <pt>    Body size in points (default: 11)
  -h, --help              Show this help
`

function parseArgs(argv) {
  const options = { positional: [] }
  const takesValue = new Set(['-o', '--out', '-t', '--title', '--subtitle', '--author', '--page', '--font', '--east-asia', '--font-size'])
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]
    if (arg === '-h' || arg === '--help') options.help = true
    else if (arg === '--text') options.text = true
    else if (arg === '--toc') options.toc = true
    else if (arg === '--landscape') options.landscape = true
    else if (arg === '--no-footer') options.footer = false
    else if (takesValue.has(arg)) {
      const value = argv[i + 1]
      if (value === undefined) throw new Error(`${arg} needs a value`)
      i += 1
      if (arg === '-o' || arg === '--out') options.out = value
      else if (arg === '-t' || arg === '--title') options.title = value
      else if (arg === '--subtitle') options.subtitle = value
      else if (arg === '--author') options.author = value
      else if (arg === '--page') options.page = value
      else if (arg === '--font') options.font = value
      else if (arg === '--east-asia') options.eastAsia = value
      else if (arg === '--font-size') options.fontSize = Number(value)
    } else if (arg.startsWith('-') && arg !== '-') throw new Error(`unknown option ${arg}`)
    else options.positional.push(arg)
  }
  return options
}

function timestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0')
  return (
    `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-` +
    `${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`
  )
}

function firstHeading(markdown) {
  const match = /^[ \t]{0,3}#{1,6}[ \t]+(.+?)[ \t]*#*[ \t]*$/m.exec(markdown)
  return match ? match[1].trim() : ''
}

async function main() {
  const options = parseArgs(process.argv.slice(2))
  if (options.help || options.positional.length === 0) {
    process.stdout.write(USAGE)
    process.exitCode = options.help ? 0 : 1
    return
  }

  const input = options.positional[0]
  const body = input === '-' ? readFileSync(0, 'utf8') : readFileSync(resolve(input), 'utf8')
  const plainText = options.text === true || extname(input).toLowerCase() === '.txt'
  const title = options.title ?? (plainText ? '' : firstHeading(body))

  const outArg = options.out ?? (input === '-' ? `document-${timestamp()}.docx` : `${input.replace(/\.[^.]*$/, '')}-${timestamp()}.docx`)
  const target = extname(outArg).toLowerCase() === '.docx' ? outArg : `${outArg}.docx`
  const absolute = isAbsolute(target) ? resolve(target) : resolve(process.cwd(), target)

  const loadImage = (src) => {
    try {
      if (/^(https?|data):/i.test(src)) return null
      const candidate = isAbsolute(src) ? src : resolve(input === '-' ? process.cwd() : dirname(resolve(input)), src)
      return { data: readFileSync(candidate), ext: extname(candidate).slice(1).toLowerCase() }
    } catch {
      return null
    }
  }

  const { buffer, stats } = await buildDocx(body, {
    title,
    subtitle: options.subtitle ?? '',
    author: options.author ?? '',
    toc: options.toc === true,
    footer: options.footer !== false,
    pageSize: options.page ?? 'a4',
    landscape: options.landscape === true,
    fontFamily: options.font ?? 'Calibri',
    eastAsiaFont: options.eastAsia ?? '\u5fae\u8f6f\u96c5\u9ed1',
    fontSize: Number.isFinite(options.fontSize) ? options.fontSize : 11,
    plainText,
    loadImage,
  })

  writeFileSync(absolute, buffer)
  process.stdout.write(
    `${absolute}\n${stats.paragraphs} paragraphs, ${stats.headings} headings, ${stats.tables} tables, ` +
      `${stats.images} images, ${buffer.length} bytes\n`,
  )
}

main().catch((error) => {
  process.stderr.write(`dsh-docx-export: ${error.message}\n`)
  process.exitCode = 1
})
