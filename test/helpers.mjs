/** Read a ZIP archive written by lib/zip.js, for assertions in tests. */
import { deflateSync, inflateRawSync } from 'node:zlib'
import { crc32 } from '../lib/zip.js'

export function readZip(buffer) {
  const entries = new Map()
  // Locate the end-of-central-directory record (no comment is ever written).
  let eocd = buffer.length - 22
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) eocd -= 1
  if (eocd < 0) throw new Error('no end-of-central-directory record')
  const count = buffer.readUInt16LE(eocd + 10)
  let offset = buffer.readUInt32LE(eocd + 16)

  for (let i = 0; i < count; i += 1) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) throw new Error(`bad central header at ${offset}`)
    const method = buffer.readUInt16LE(offset + 10)
    const expectedCrc = buffer.readUInt32LE(offset + 16)
    const compressedSize = buffer.readUInt32LE(offset + 20)
    const uncompressedSize = buffer.readUInt32LE(offset + 24)
    const nameLength = buffer.readUInt16LE(offset + 28)
    const extraLength = buffer.readUInt16LE(offset + 30)
    const commentLength = buffer.readUInt16LE(offset + 32)
    const localOffset = buffer.readUInt32LE(offset + 42)
    const name = buffer.toString('utf8', offset + 46, offset + 46 + nameLength)

    const localNameLength = buffer.readUInt16LE(localOffset + 26)
    const localExtraLength = buffer.readUInt16LE(localOffset + 28)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength
    const payload = buffer.subarray(dataStart, dataStart + compressedSize)
    const data = method === 8 ? inflateRawSync(payload) : Buffer.from(payload)
    if (data.length !== uncompressedSize) throw new Error(`${name}: size mismatch`)
    if (crc32(data) !== expectedCrc) throw new Error(`${name}: CRC mismatch`)

    offset += 46 + nameLength + extraLength + commentLength
    entries.set(name, data)
  }
  return entries
}

const VOID_TAGS = new Set(['br', 'hr', 'img', 'link', 'meta', 'input'])

/**
 * Structural XML check: balanced, correctly nested tags with quoted
 * attributes. Enough to catch escaping mistakes in generated OOXML without
 * pulling in a parser dependency.
 */
export function assertWellFormedXml(xml, label = 'xml') {
  const stack = []
  const cleaned = xml
    .replace(/<\?[\s\S]*?\?>/g, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<!DOCTYPE[^>]*>/g, '')
  const tagPattern = /<(\/?)([A-Za-z_][\w.:-]*)((?:"[^"]*"|[^>"])*?)(\/?)>/g
  let cursor = 0
  let match = tagPattern.exec(cleaned)
  while (match !== null) {
    const between = cleaned.slice(cursor, match.index)
    if (between.includes('<')) throw new Error(`${label}: stray "<" in text near ${between.slice(0, 40)}`)
    if (/&(?!(?:#\d+|#x[0-9a-fA-F]+|[a-zA-Z][\w.-]*);)/.test(between)) {
      throw new Error(`${label}: unescaped "&" near ${between.slice(0, 40)}`)
    }
    const [, closing, rawName, attrs, selfClosing] = match
    const name = rawName.toLowerCase()
    if (!closing && !selfClosing && !VOID_TAGS.has(name)) stack.push(name)
    if (closing) {
      const open = stack.pop()
      if (open !== name) throw new Error(`${label}: </${name}> closes <${open ?? 'nothing'}>`)
    }
    for (const attr of attrs.matchAll(/([\w.:-]+)\s*=\s*"([^"]*)"/g)) {
      if (/[<>]/.test(attr[2])) throw new Error(`${label}: attribute ${attr[1]} needs escaping`)
    }
    cursor = match.index + match[0].length
    match = tagPattern.exec(cleaned)
  }
  if (stack.length > 0) throw new Error(`${label}: unclosed <${stack[stack.length - 1]}>`)
  return true
}

/** Minimal PNG encoder, used to exercise picture embedding end to end. */
export function makePng(width, height) {
  const chunk = (type, data) => {
    const length = Buffer.alloc(4)
    length.writeUInt32BE(data.length, 0)
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
    const crc = Buffer.alloc(4)
    crc.writeUInt32BE(crc32(body), 0)
    return Buffer.concat([length, body, crc])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // truecolor
  const raw = Buffer.alloc(height * (1 + width * 3))
  for (let y = 0; y < height; y += 1) {
    const rowStart = y * (1 + width * 3)
    raw[rowStart] = 0
    for (let x = 0; x < width; x += 1) {
      raw[rowStart + 1 + x * 3] = 0x2b
      raw[rowStart + 2 + x * 3] = 0x57
      raw[rowStart + 3 + x * 3] = 0x9a
    }
  }
  // zlib.deflateSync produces exactly the zlib stream PNG IDAT expects.
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ])
}
