/**
 * WordprocessingML writer.
 *
 * Turns the Markdown block tree produced by `parseMarkdown` into an OOXML
 * package (a .docx) that Word, WPS, LibreOffice and Google Docs all open
 * natively: real heading styles, real numbering, real tables, embedded
 * pictures, a TOC field and a page-numbered footer.
 *
 * Everything here is data-in/buffer-out: no filesystem access, no globals.
 * The caller injects `loadImage`, so the same builder serves the DSH tool and
 * the CLI with whatever path policy each of them owns.
 */
import { parseInline, parseMarkdown } from './markdown.js'
import { createZip } from './zip.js'

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

const EMU_PER_PIXEL = 9525
const EMU_PER_TWIP = 635

/** Page geometry in twips (1/1440 inch), the unit WordprocessingML uses. */
const PAGE_SIZES = {
  a4: { width: 11906, height: 16838, margin: 1418 },
  letter: { width: 12240, height: 15840, margin: 1440 },
  a5: { width: 8391, height: 11906, margin: 1134 },
}

export const SUPPORTED_MEDIA = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
}

/** Escape text for XML character data. */
function esc(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // Control characters Word refuses to load; \t \n \r are handled as markup.
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
}

/** Escape a value for an XML attribute position. */
function escAttr(value) {
  return esc(value).replace(/"/g, '&quot;')
}

/** A `<w:t>` element, preserving the leading/trailing spaces authors typed. */
function textNode(value) {
  return `<w:t xml:space="preserve">${esc(value)}</w:t>`
}

/** Split run text into `<w:t>` and `<w:br/>` pieces on newlines. */
function textNodes(value) {
  const parts = String(value).split('\n')
  return parts.map((part, i) => (i === 0 ? textNode(part) : `<w:br/>${textNode(part)}`)).join('')
}

function attributes(map) {
  return Object.entries(map)
    .filter(([, value]) => value !== undefined && value !== null && value !== false)
    .map(([key, value]) => (value === true ? ` ${key}` : ` ${key}="${escAttr(value)}"`))
    .join('')
}

/** Image dimensions in pixels, read straight from the file header. */
export function imageSize(buffer, ext) {
  try {
    if (ext === 'png' && buffer.length > 24 && buffer.readUInt32BE(0) === 0x89504e47) {
      return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) }
    }
    if ((ext === 'jpg' || ext === 'jpeg') && buffer.length > 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
      let offset = 2
      while (offset + 9 < buffer.length) {
        if (buffer[offset] !== 0xff) {
          offset += 1
          continue
        }
        const marker = buffer[offset + 1]
        const length = buffer.readUInt16BE(offset + 2)
        const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
        if (isSof) return { width: buffer.readUInt16BE(offset + 7), height: buffer.readUInt16BE(offset + 5) }
        offset += 2 + length
      }
      return null
    }
    if (ext === 'gif' && buffer.length > 10 && buffer.toString('ascii', 0, 3) === 'GIF') {
      return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) }
    }
    if (ext === 'bmp' && buffer.length > 26 && buffer.toString('ascii', 0, 2) === 'BM') {
      return { width: Math.abs(buffer.readInt32LE(18)), height: Math.abs(buffer.readInt32LE(22)) }
    }
  } catch {
    return null
  }
  return null
}

function createContext(options) {
  return {
    options,
    body: [],
    relationships: [],
    media: [],
    hyperlinks: new Map(),
    imageCounter: 0,
    drawingCounter: 0,
    bookmarkCounter: 0,
    stats: { words: 0, paragraphs: 0, headings: 0, lists: 0, tables: 0, images: 0, codeBlocks: 0, links: 0 },
  }
}

function addRelationship(ctx, type, target, mode) {
  const id = `rId${ctx.relationships.length + 1}`
  ctx.relationships.push({ id, type, target, mode })
  return id
}

function hyperlinkId(ctx, href) {
  if (ctx.hyperlinks.has(href)) return ctx.hyperlinks.get(href)
  const id = addRelationship(
    ctx,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    href,
    'External',
  )
  ctx.hyperlinks.set(href, id)
  ctx.stats.links += 1
  return id
}

/** Run properties for the requested style flags. */
function runProps(ctx, flags = {}) {
  const { bold, italic, strike, code, color, size, font, underline, link } = flags
  const parts = []
  if (link) parts.push('<w:rStyle w:val="Hyperlink"/>')
  const latin = font ?? (code ? ctx.options.monoFont : null)
  if (latin || code) {
    parts.push(
      `<w:rFonts${attributes({
        'w:ascii': latin,
        'w:hAnsi': latin,
        'w:cs': latin,
      })}/>`,
    )
  }
  if (bold) parts.push('<w:b/><w:bCs/>')
  if (italic) parts.push('<w:i/><w:iCs/>')
  if (strike) parts.push('<w:strike/>')
  if (underline || link) parts.push('<w:u w:val="single"/>')
  if (color || link) parts.push(`<w:color w:val="${color ?? '0969DA'}"/>`)
  if (code) parts.push('<w:shd w:val="clear" w:color="auto" w:fill="F2F3F5"/>')
  if (size) parts.push(`<w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/>`)
  return parts.length > 0 ? `<w:rPr>${parts.join('')}</w:rPr>` : ''
}

function run(ctx, text, flags = {}) {
  if (text === '') return ''
  return `<w:r>${runProps(ctx, flags)}${textNodes(text)}</w:r>`
}

function drawingRun(ctx, image) {
  const { relId, name, widthEmu, heightEmu, alt } = image
  ctx.drawingCounter += 1
  const id = ctx.drawingCounter
  return (
    '<w:r><w:drawing>' +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${escAttr(name)}" descr="${escAttr(alt)}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="' +
    A_NS +
    '" noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${escAttr(name)}" descr="${escAttr(alt)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  )
}

/** Render inline runs (text, emphasis, code, links, images) into XML. */
function inlineXml(ctx, text, flags = {}) {
  const runs = parseInline(text)
  const out = []
  for (const item of runs) {
    if (item.type === 'image') {
      const embedded = embedImage(ctx, item.src, item.alt)
      out.push(embedded ?? run(ctx, item.alt || item.src, { ...flags, italic: true, color: '808080' }))
      continue
    }
    const text2 = item.text ?? ''
    if (text2 === '') continue
    ctx.stats.words += text2.trim() === '' ? 0 : text2.trim().split(/\s+/).length
    const merged = {
      bold: flags.bold || item.bold,
      italic: flags.italic || item.italic,
      strike: flags.strike || item.strike,
      code: item.code === true,
    }
    const xml = run(ctx, text2, item.href ? { ...merged, link: true } : merged)
    if (item.href) {
      out.push(`<w:hyperlink r:id="${hyperlinkId(ctx, item.href)}" w:history="1">${xml}</w:hyperlink>`)
    } else {
      out.push(xml)
    }
  }
  return out.join('')
}

/** Read, size and register a local image; returns null when it cannot be used. */
function embedImage(ctx, src, alt) {
  const loaded = ctx.options.loadImage ? ctx.options.loadImage(src) : null
  if (!loaded || !loaded.data || loaded.data.length === 0) return null
  const ext = String(loaded.ext ?? '').toLowerCase().replace(/^\./, '')
  const contentType = SUPPORTED_MEDIA[ext]
  if (!contentType) return null

  const name = `image${ctx.media.length + 1}.${ext}`
  const relId = addRelationship(
    ctx,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    `media/${name}`,
  )
  ctx.media.push({ name: `word/media/${name}`, data: loaded.data, contentType })

  const natural = imageSize(loaded.data, ext)
  const maxWidthEmu = ctx.options.contentWidthTwips * EMU_PER_TWIP
  let widthEmu = (natural?.width ?? 480) * EMU_PER_PIXEL
  let heightEmu = (natural?.height ?? 320) * EMU_PER_PIXEL
  if (widthEmu > maxWidthEmu) {
    const scale = maxWidthEmu / widthEmu
    widthEmu = Math.round(widthEmu * scale)
    heightEmu = Math.round(heightEmu * scale)
  }
  ctx.stats.images += 1
  return drawingRun(ctx, { relId, name, widthEmu, heightEmu, alt: alt || src })
}

function paragraph(props, content) {
  return `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}${content}</w:p>`
}

function headingBookmark(ctx, text) {
  ctx.bookmarkCounter += 1
  const name = `_Toc${String(ctx.bookmarkCounter).padStart(6, '0')}`
  return (
    `<w:bookmarkStart w:id="${ctx.bookmarkCounter}" w:name="${name}"/>` +
    `${inlineXml(ctx, text)}` +
    `<w:bookmarkEnd w:id="${ctx.bookmarkCounter}"/>`
  )
}

function tableXml(ctx, node) {
  const columns = node.header.length
  const width = Math.floor(ctx.options.contentWidthTwips / columns)
  const grid = node.header.map(() => `<w:gridCol w:w="${width}"/>`).join('')
  const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>`
  const tblPr =
    '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>' +
    `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>` +
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'

  const cell = (text, align, header) => {
    const jc = align ? `<w:jc w:val="${align}"/>` : ''
    const shd = header ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F3F5"/>' : ''
    const content = header && text.trim() === '' ? '' : inlineXml(ctx, text, header ? { bold: true } : {})
    return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shd}</w:tcPr>${paragraph(jc, content)}</w:tc>`
  }

  const rows = []
  rows.push(
    `<w:tr><w:trPr><w:tblHeader/></w:trPr>${node.header
      .map((text, i) => cell(text, node.align[i], true))
      .join('')}</w:tr>`,
  )
  for (const row of node.rows) {
    rows.push(`<w:tr>${row.map((text, i) => cell(text, node.align[i], false)).join('')}</w:tr>`)
  }
  ctx.stats.tables += 1
  ctx.stats.paragraphs += 1
  return `<w:tbl>${tblPr}<w:tblGrid>${grid}</w:tblGrid>${rows.join('')}</w:tbl>${paragraph('', '')}`
}

function codeXml(ctx, node) {
  ctx.stats.codeBlocks += 1
  const body = run(ctx, node.text.replace(/\n$/, '').replace(/\t/g, '    '), {
    code: true,
    size: Math.max(ctx.options.fontSize - 1, 6),
  })
  const props =
    '<w:pStyle w:val="CodeBlock"/>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="F6F8FA"/>' +
    '<w:pBdr><w:top w:val="single" w:sz="4" w:space="4" w:color="D8DEE4"/>' +
    '<w:left w:val="single" w:sz="4" w:space="4" w:color="D8DEE4"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="4" w:color="D8DEE4"/>' +
    '<w:right w:val="single" w:sz="4" w:space="4" w:color="D8DEE4"/></w:pBdr>' +
    `<w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>`
  ctx.stats.paragraphs += 1
  return paragraph(props, body)
}

function quoteXml(ctx, node) {
  const inner = blocksToXml(ctx, node.blocks, { quote: true })
  return inner
}

function blocksToXml(ctx, blocks, context = {}) {
  const out = []
  for (const node of blocks) {
    switch (node.type) {
      case 'heading': {
        ctx.stats.headings += 1
        ctx.stats.paragraphs += 1
        const level = Math.min(Math.max(node.level, 1), 6)
        out.push(paragraph(`<w:pStyle w:val="Heading${level}"/><w:keepNext/>`, headingBookmark(ctx, node.text)))
        break
      }
      case 'paragraph': {
        ctx.stats.paragraphs += 1
        const style = context.quote ? '<w:pStyle w:val="Quote"/>' : ''
        if (node.literal === true) {
          // Plain-text input keeps its characters: no emphasis is interpreted.
          ctx.stats.words += node.text.trim() === '' ? 0 : node.text.trim().split(/\s+/).length
          out.push(paragraph(style, run(ctx, node.text, {})))
        } else {
          out.push(paragraph(style, inlineXml(ctx, node.text)))
        }
        break
      }
      case 'code': {
        out.push(codeXml(ctx, node))
        break
      }
      case 'list': {
        ctx.stats.lists += 1
        out.push(listXml(ctx, node, context.level ?? 0, context.quote === true))
        break
      }
      case 'quote': {
        out.push(quoteXml(ctx, node))
        break
      }
      case 'table': {
        out.push(tableXml(ctx, node))
        break
      }
      case 'hr': {
        ctx.stats.paragraphs += 1
        out.push(
          paragraph(
            '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="C9D1D9"/></w:pBdr><w:spacing w:before="120" w:after="120"/>',
            '',
          ),
        )
        break
      }
      case 'image': {
        ctx.stats.paragraphs += 1
        const embedded = embedImage(ctx, node.src, node.alt)
        out.push(
          paragraph(
            '<w:jc w:val="center"/>',
            embedded ?? run(ctx, node.alt || node.src, { italic: true, color: '808080' }),
          ),
        )
        break
      }
      default:
        break
    }
  }
  return out.join('')
}

function listXml(ctx, node, level, quoted) {
  const numId = node.ordered ? 2 : 1
  const out = []
  for (const item of node.items) {
    const marker = item.checked === null || item.checked === undefined ? '' : item.checked ? '\u2611 ' : '\u2610 '
    const numPr = `<w:numPr><w:ilvl w:val="${Math.min(level, 8)}"/><w:numId w:val="${numId}"/></w:numPr>`
    const props = `<w:pStyle w:val="ListParagraph"/>${numPr}${quoted ? '<w:ind w:left="1440"/>' : ''}`
    ctx.stats.paragraphs += 1
    out.push(paragraph(props, (marker ? run(ctx, marker, {}) : '') + inlineXml(ctx, item.text)))
    for (const child of item.blocks ?? []) {
      if (child.type === 'list') out.push(listXml(ctx, child, level + 1, quoted))
      else out.push(blocksToXml(ctx, [child], { level: level + 1, quote: quoted }))
    }
  }
  return out.join('')
}

function tocXml(ctx) {
  ctx.stats.paragraphs += 1
  const heading = paragraph('<w:pStyle w:val="TOCHeading"/>', run(ctx, ctx.options.tocTitle, {}))
  const field =
    paragraph(
      '',
      '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        run(ctx, ctx.options.tocPlaceholder, { color: '808080', italic: true }) +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    )
  return heading + field
}

function titleXml(ctx) {
  const { title, subtitle, author } = ctx.options
  const out = []
  if (title) {
    ctx.stats.paragraphs += 1
    out.push(paragraph('<w:pStyle w:val="Title"/>', inlineXml(ctx, title)))
  }
  if (subtitle) {
    ctx.stats.paragraphs += 1
    out.push(paragraph('<w:pStyle w:val="Subtitle"/>', inlineXml(ctx, subtitle)))
  }
  if (author) {
    ctx.stats.paragraphs += 1
    out.push(paragraph('<w:jc w:val="right"/>', run(ctx, author, { italic: true, color: '595959' })))
  }
  return out.join('')
}

function stylesXml(options) {
  const { fontFamily, eastAsiaFont, monoFont, fontSize } = options
  const fonts = (latin, eastAsia) =>
    `<w:rFonts w:ascii="${escAttr(latin)}" w:hAnsi="${escAttr(latin)}" w:eastAsia="${escAttr(eastAsia)}" w:cs="${escAttr(latin)}"/>`
  const headingSizes = [26, 20, 16, 14, 12.5, 11.5]
  const headingStyles = headingSizes
    .map((size, i) => {
      const level = i + 1
      return (
        `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
        '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
        `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${level === 1 ? 360 : 240}" w:after="120"/>` +
        `<w:outlineLvl w:val="${i}"/></w:pPr>` +
        `<w:rPr>${fonts(fontFamily, eastAsiaFont)}<w:b/><w:bCs/><w:sz w:val="${Math.round(size * 2)}"/>` +
        `<w:szCs w:val="${Math.round(size * 2)}"/>${level === 1 ? '<w:color w:val="1F2328"/>' : ''}</w:rPr></w:style>`
      )
    })
    .join('')

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${W_NS}">` +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    `${fonts(fontFamily, eastAsiaFont)}<w:sz w:val="${Math.round(fontSize * 2)}"/><w:szCs w:val="${Math.round(fontSize * 2)}"/>` +
    '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:before="60" w:after="120" w:line="320" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>' +
    '<w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr>' +
    `<w:rPr>${fonts(fontFamily, eastAsiaFont)}<w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>` +
    '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/>' +
    '<w:next w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
    `<w:rPr>${fonts(fontFamily, eastAsiaFont)}<w:color w:val="57606A"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>` +
    '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/>' +
    '<w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>' +
    headingStyles +
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:left="480"/><w:spacing w:before="120" w:after="120"/>' +
    '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="D0D7DE"/></w:pBdr></w:pPr>' +
    `<w:rPr>${fonts(fontFamily, eastAsiaFont)}<w:i/><w:iCs/><w:color w:val="57606A"/></w:rPr></w:style>` +
    '<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/></w:pPr>' +
    `<w:rPr>${fonts(monoFont, eastAsiaFont)}<w:sz w:val="${Math.round((fontSize - 1) * 2)}"/></w:rPr></w:style>` +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:spacing w:before="0" w:after="60" w:line="300" w:lineRule="auto"/></w:pPr></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
    '<w:rPr><w:color w:val="0969DA"/><w:u w:val="single"/></w:rPr></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
    '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/></w:tblBorders></w:tblPr></w:style>' +
    '</w:styles>'
  )
}

function numberingXml() {
  const bulletGlyphs = ['\u2022', '\u25e6', '\u25aa']
  const level = (index, kind) => {
    const left = 720 + index * 360
    const numFmt = kind === 'bullet' ? 'bullet' : ['decimal', 'lowerLetter', 'lowerRoman'][index % 3]
    const text = kind === 'bullet' ? bulletGlyphs[index % bulletGlyphs.length] : `%${index + 1}.`
    const fonts = kind === 'bullet' ? '<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>' : ''
    return (
      `<w:lvl w:ilvl="${index}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/>` +
      `<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>` +
      (fonts ? `<w:rPr>${fonts}</w:rPr>` : '') +
      '</w:lvl>'
    )
  }
  const abstract = (id, kind) =>
    `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>` +
    Array.from({ length: 9 }, (_, i) => level(i, kind)).join('') +
    '</w:abstractNum>'
  const num = (id, abstractId) =>
    `<w:num w:numId="${id}"><w:abstractNumId w:val="${abstractId}"/></w:num>`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:numbering xmlns:w="${W_NS}">${abstract(0, 'bullet')}${abstract(1, 'decimal')}${num(1, 0)}${num(2, 1)}</w:numbering>`
  )
}

function footerXml(ctx) {
  const title = ctx.options.title ? `${esc(ctx.options.title)}    ` : ''
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
    '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>' +
    run(ctx, title, { size: ctx.options.fontSize - 2, color: '6E7781' }) +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    run(ctx, '1', { size: ctx.options.fontSize - 2, color: '6E7781' }) +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:p></w:ftr>'
  )
}

function settingsXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:settings xmlns:w="${W_NS}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>` +
    '<w:updateFields w:val="true"/><w:compat><w:compatSetting w:name="compatibilityMode" ' +
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
  )
}

function coreXml(options, now) {
  const stamp = now.toISOString().replace(/\.\d{3}Z$/, 'Z')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${esc(options.title)}</dc:title>` +
    `<dc:subject>${esc(options.subject ?? '')}</dc:subject>` +
    `<dc:creator>${esc(options.author)}</dc:creator>` +
    `<cp:lastModifiedBy>${esc(options.author)}</cp:lastModifiedBy>` +
    `<cp:keywords>${esc(options.keywords ?? '')}</cp:keywords>` +
    '<cp:revision>1</cp:revision>' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    '</cp:coreProperties>'
  )
}

function appXml(options, stats) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>dsh-docx-export</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
    `<Company>${esc(options.company ?? '')}</Company>` +
    '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>' +
    '<AppVersion>1.0</AppVersion>' +
    `<Words>${stats.words}</Words><Paragraphs>${stats.paragraphs}</Paragraphs>` +
    '</Properties>'
  )
}

function contentTypesXml(ctx) {
  const extensions = [...new Set(ctx.media.map((item) => item.name.split('.').pop()))]
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    extensions
      .map((ext) => `<Default Extension="${ext}" ContentType="${SUPPORTED_MEDIA[ext] ?? 'application/octet-stream'}"/>`)
      .join('') +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    (ctx.footerRelId
      ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      : '') +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>'
  )
}

/**
 * Plain-text input: blank lines separate paragraphs, single newlines are kept
 * as line breaks, and every character stays literal.
 */
function plainTextBlocks(text) {
  const lines = String(text ?? '').replace(/\r\n?/g, '\n').split('\n')
  const blocks = []
  let buffer = []
  const flush = () => {
    if (buffer.length > 0) {
      blocks.push({ type: 'paragraph', text: buffer.join('\n'), literal: true })
      buffer = []
    }
  }
  for (const line of lines) {
    if (line.trim() === '') flush()
    else buffer.push(line.replace(/\s+$/, ''))
  }
  flush()
  return blocks
}

function normalizeOptions(options = {}) {
  const pageKey = String(options.pageSize ?? 'a4').toLowerCase()
  const page = PAGE_SIZES[pageKey] ?? PAGE_SIZES.a4
  const landscape = options.landscape === true
  const merged = {
    title: options.title ?? '',
    subtitle: options.subtitle ?? '',
    author: options.author ?? '',
    subject: options.subject ?? '',
    keywords: options.keywords ?? '',
    company: options.company ?? '',
    fontFamily: options.fontFamily || 'Calibri',
    eastAsiaFont: options.eastAsiaFont || '\u5fae\u8f6f\u96c5\u9ed1',
    monoFont: options.monoFont || 'Consolas',
    fontSize: Number.isFinite(options.fontSize) ? options.fontSize : 11,
    pageSize: PAGE_SIZES[pageKey] ? pageKey : 'a4',
    landscape,
    toc: options.toc === true,
    footer: options.footer !== false,
    plainText: options.plainText === true,
    tocTitle: options.tocTitle || 'Contents',
    tocPlaceholder: options.tocPlaceholder || 'Right-click and choose "Update Field" to build the table of contents.',
    loadImage: typeof options.loadImage === 'function' ? options.loadImage : null,
    now: options.now instanceof Date ? options.now : new Date(),
    page,
    contentWidthTwips: (landscape ? page.height : page.width) - page.margin * 2,
  }
  return merged
}

/**
 * Build a .docx package.
 *
 * @param {string} markdown document body (Markdown, or plain text)
 * @param {object} [options] see `normalizeOptions`
 * @returns {Promise<{ buffer: Buffer, stats: object }>}
 */
export async function buildDocx(markdown, options = {}) {
  const normalized = normalizeOptions(options)
  const ctx = createContext(normalized)
  const parsed =
    normalized.plainText === true
      ? { blocks: plainTextBlocks(markdown) }
      : typeof markdown === 'string'
        ? parseMarkdown(markdown)
        : { blocks: markdown ?? [] }

  const stylesRel = addRelationship(
    ctx,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    'styles.xml',
  )
  const numberingRel = addRelationship(
    ctx,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
    'numbering.xml',
  )
  const settingsRel = addRelationship(
    ctx,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
    'settings.xml',
  )
  // Styles, numbering and settings are found through these typed relationships
  // rather than referenced from the body, so the ids only have to exist.
  void stylesRel
  void numberingRel
  void settingsRel
  let footerRel = null

  const pieces = []
  if (normalized.title || normalized.subtitle || normalized.author) pieces.push(titleXml(ctx))
  if (normalized.toc) pieces.push(tocXml(ctx))
  pieces.push(blocksToXml(ctx, parsed.blocks))

  if (normalized.footer) {
    footerRel = addRelationship(
      ctx,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
      'footer1.xml',
    )
    ctx.footerRelId = footerRel
  }

  const page = normalized.page
  const pageWidth = normalized.landscape ? page.height : page.width
  const pageHeight = normalized.landscape ? page.width : page.height
  const sectPr =
    '<w:sectPr>' +
    (footerRel ? `<w:footerReference w:type="default" r:id="${footerRel}"/>` : '') +
    `<w:pgSz w:w="${pageWidth}" w:h="${pageHeight}"${normalized.landscape ? ' w:orient="landscape"' : ''}/>` +
    `<w:pgMar w:top="${page.margin}" w:right="${page.margin}" w:bottom="${page.margin}" w:left="${page.margin}" ` +
    'w:header="851" w:footer="992" w:gutter="0"/>' +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="312"/>' +
    '</w:sectPr>'

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">` +
    `<w:body>${pieces.join('')}${sectPr}</w:body></w:document>`

  const documentRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    ctx.relationships
      .map(
        (rel) =>
          `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escAttr(rel.target)}"` +
          (rel.mode ? ` TargetMode="${rel.mode}"` : '') +
          '/>',
      )
      .join('') +
    '</Relationships>'

  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
    '</Relationships>'

  const entries = [
    { name: '[Content_Types].xml', data: contentTypesXml(ctx) },
    { name: '_rels/.rels', data: rootRels },
    { name: 'docProps/core.xml', data: coreXml(normalized, normalized.now) },
    { name: 'docProps/app.xml', data: appXml(normalized, ctx.stats) },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: documentRels },
    { name: 'word/styles.xml', data: stylesXml(normalized) },
    { name: 'word/numbering.xml', data: numberingXml() },
    { name: 'word/settings.xml', data: settingsXml() },
  ]
  if (footerRel) entries.push({ name: 'word/footer1.xml', data: footerXml(ctx) })
  for (const media of ctx.media) {
    entries.push({ name: media.name, data: media.data, compress: false })
  }

  return { buffer: createZip(entries, { now: normalized.now }), stats: ctx.stats }
}
