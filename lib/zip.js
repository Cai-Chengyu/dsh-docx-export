/**
 * Minimal ZIP writer for OOXML packages.
 *
 * A .docx is a ZIP archive; Node ships deflate, so the whole writer is a few
 * hundred lines instead of a dependency. Only what an OOXML package needs is
 * implemented: store/deflate entries, UTF-8 names, no ZIP64, no encryption.
 */
import { deflateRawSync } from 'node:zlib'

const CRC_TABLE = (() => {
  const table = new Int32Array(256)
  for (let i = 0; i < 256; i += 1) {
    let value = i
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    }
    table[i] = value
  }
  return table
})()

/** CRC-32 of a buffer, as required by the ZIP entry headers. */
export function crc32(buffer) {
  let crc = -1
  for (let i = 0; i < buffer.length; i += 1) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buffer[i]) & 0xff]
  }
  return (crc ^ -1) >>> 0
}

/** MS-DOS date/time pair for a Date (ZIP stores local time, 2-second steps). */
function dosStamp(date) {
  const year = Math.max(1980, date.getFullYear())
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  }
}

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_EOCD = 0x06054b50
const FLAG_UTF8 = 0x0800

/**
 * Build a ZIP archive.
 *
 * @param {Array<{ name: string, data: Buffer | string, compress?: boolean }>} entries
 *   Entries in the order they should appear; `compress` defaults to true.
 * @param {{ now?: Date }} [options]
 * @returns {Buffer} the complete archive
 */
export function createZip(entries, options = {}) {
  const now = options.now ?? new Date()
  const stamp = dosStamp(now)
  const chunks = []
  const central = []
  let offset = 0

  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8')
    const raw = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(String(entry.data), 'utf8')
    const compress = entry.compress !== false
    const deflated = compress && raw.length > 0 ? deflateRawSync(raw, { level: 9 }) : null
    // Storing is smaller than deflating tiny or already-compressed payloads.
    const useDeflate = deflated !== null && deflated.length < raw.length
    const payload = useDeflate ? deflated : raw
    const method = useDeflate ? 8 : 0
    const crc = crc32(raw)

    const local = Buffer.alloc(30)
    local.writeUInt32LE(SIG_LOCAL, 0)
    local.writeUInt16LE(20, 4) // version needed
    local.writeUInt16LE(FLAG_UTF8, 6)
    local.writeUInt16LE(method, 8)
    local.writeUInt16LE(stamp.time, 10)
    local.writeUInt16LE(stamp.date, 12)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(payload.length, 18)
    local.writeUInt32LE(raw.length, 22)
    local.writeUInt16LE(name.length, 26)
    local.writeUInt16LE(0, 28)

    const header = Buffer.alloc(46)
    header.writeUInt32LE(SIG_CENTRAL, 0)
    header.writeUInt16LE(20, 4) // version made by
    header.writeUInt16LE(20, 6) // version needed
    header.writeUInt16LE(FLAG_UTF8, 8)
    header.writeUInt16LE(method, 10)
    header.writeUInt16LE(stamp.time, 12)
    header.writeUInt16LE(stamp.date, 14)
    header.writeUInt32LE(crc, 16)
    header.writeUInt32LE(payload.length, 20)
    header.writeUInt32LE(raw.length, 24)
    header.writeUInt16LE(name.length, 28)
    // extra/comment/disk/internal attrs stay zero
    header.writeUInt32LE(0, 38) // external attrs
    header.writeUInt32LE(offset, 42) // local header offset

    chunks.push(local, name, payload)
    central.push(header, name)
    offset += local.length + name.length + payload.length
  }

  const centralBuffer = Buffer.concat(central)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(SIG_EOCD, 0)
  end.writeUInt16LE(0, 4) // this disk
  end.writeUInt16LE(0, 6) // disk with central directory
  end.writeUInt16LE(entries.length, 8)
  end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(centralBuffer.length, 12)
  end.writeUInt32LE(offset, 16)
  end.writeUInt16LE(0, 20) // comment length

  return Buffer.concat([...chunks, centralBuffer, end])
}
