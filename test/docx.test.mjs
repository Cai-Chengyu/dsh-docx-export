import assert from 'node:assert/strict'
import test from 'node:test'
import { buildDocx, imageSize } from '../lib/docx.js'
import { assertWellFormedXml, makePng, readZip } from './helpers.mjs'

const SAMPLE = `# 季度报告

这是**第一**段，包含 *斜体*、\`代码\` 与 [链接](https://example.com/a?b=1&c=2)。

## 要点

- 第一项
- 第二项
  - 嵌套项
- [x] 已完成
- [ ] 待办

1. 步骤一
2. 步骤二

> 引用 **文字**
> 第二行

\`\`\`js
const a = 1 < 2 && 3 > 2
console.log("x & y")
\`\`\`

| 名称 | 数量 | 备注 |
|:-----|:----:|-----:|
| 甲   | 12   | 无   |
| 乙   | 3    | 待定 |

---

![示意图](pics/diagram.png)
`

async function build(overrides = {}, markdown = SAMPLE) {
  const png = makePng(640, 480)
  return buildDocx(markdown, {
    title: '季度报告',
    subtitle: '2025 Q1',
    author: 'DSH',
    toc: true,
    footer: true,
    now: new Date('2025-02-03T04:05:06Z'),
    loadImage: (src) => (src === 'pics/diagram.png' ? { data: png, ext: 'png' } : null),
    ...overrides,
  })
}

test('writes an OOXML package with every required part', async () => {
  const { buffer, stats } = await build()
  assert.ok(Buffer.isBuffer(buffer))
  assert.equal(buffer.subarray(0, 2).toString('ascii'), 'PK')

  const entries = readZip(buffer)
  for (const name of [
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/core.xml',
    'docProps/app.xml',
    'word/document.xml',
    'word/_rels/document.xml.rels',
    'word/styles.xml',
    'word/numbering.xml',
    'word/settings.xml',
    'word/footer1.xml',
    'word/media/image1.png',
  ]) {
    assert.ok(entries.has(name), `missing part ${name}`)
  }
  for (const [name, data] of entries) {
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      assertWellFormedXml(data.toString('utf8'), name)
    }
  }
  assert.equal(stats.tables, 1)
  assert.equal(stats.images, 1)
  assert.equal(stats.headings, 2)
  assert.ok(stats.words > 20)
})

test('maps Markdown structure onto native Word features', async () => {
  const { buffer } = await build()
  const entries = readZip(buffer)
  const document = entries.get('word/document.xml').toString('utf8')

  assert.match(document, /<w:pStyle w:val="Title"\/>/)
  assert.match(document, /<w:pStyle w:val="Heading1"\/>/)
  assert.match(document, /<w:pStyle w:val="Heading2"\/>/)
  assert.match(document, /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="1"\/><\/w:numPr>/)
  assert.match(document, /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="1"\/><\/w:numPr>/)
  assert.match(document, /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="2"\/><\/w:numPr>/)
  assert.match(document, /<w:b\/><w:bCs\/>/)
  assert.match(document, /<w:i\/><w:iCs\/>/)
  assert.match(document, /<w:pStyle w:val="CodeBlock"\/>/)
  assert.match(document, /<w:pStyle w:val="Quote"\/>/)
  assert.match(document, /<w:tbl>/)
  assert.match(document, /<w:hyperlink r:id="rId\d+" w:history="1">/)
  assert.match(document, /<w:drawing>/)
  assert.match(document, /<w:pgSz w:w="11906" w:h="16838"\/>/)
  assert.match(document, /TOC \\o "1-3"/)
  assert.match(document, /<w:bookmarkStart w:id="1" w:name="_Toc000001"\/>/)

  // Escaping: raw Markdown/HTML must never leak an unescaped angle bracket.
  assert.match(document, /const a = 1 &lt; 2 &amp;&amp; 3 &gt; 2/)
  assert.match(document, /console\.log\("x &amp; y"\)/)
  assert.doesNotMatch(document, /&lt;div&gt;/)

  const rels = entries.get('word/_rels/document.xml.rels').toString('utf8')
  assert.match(rels, /Target="https:\/\/example\.com\/a\?b=1&amp;c=2" TargetMode="External"/)
  assert.match(rels, /Target="media\/image1\.png"/)
  assert.match(rels, /Target="footer1\.xml"/)

  const core = entries.get('docProps/core.xml').toString('utf8')
  assert.match(core, /<dc:title>季度报告<\/dc:title>/)
  assert.match(core, /<dc:creator>DSH<\/dc:creator>/)
  assert.match(core, /<dcterms:created xsi:type="dcterms:W3CDTF">2025-02-03T04:05:06Z<\/dcterms:created>/)

  const contentTypes = entries.get('[Content_Types].xml').toString('utf8')
  assert.match(contentTypes, /<Default Extension="png" ContentType="image\/png"\/>/)
  assert.match(contentTypes, /wordprocessingml\.document\.main\+xml/)
})

test('the footer carries a PAGE field', async () => {
  const { buffer } = await build()
  const footer = readZip(buffer).get('word/footer1.xml').toString('utf8')
  assert.match(footer, /PAGE/)
  assert.match(footer, /<w:fldChar w:fldCharType="begin"\/>/)
  assert.match(footer, /季度报告/)
})

test('options change page geometry, fonts and optional parts', async () => {
  const { buffer } = await build({
    title: '',
    subtitle: '',
    author: '',
    toc: false,
    footer: false,
    pageSize: 'letter',
    landscape: true,
    fontFamily: 'Georgia',
    eastAsiaFont: '宋体',
    fontSize: 12,
  })
  const entries = readZip(buffer)
  const document = entries.get('word/document.xml').toString('utf8')
  assert.match(document, /<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>/)
  assert.doesNotMatch(document, /<w:pStyle w:val="Title"\/>/)
  assert.doesNotMatch(document, /TOC \\o/)
  assert.equal(entries.has('word/footer1.xml'), false)
  const styles = entries.get('word/styles.xml').toString('utf8')
  assert.match(styles, /w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="宋体"/)
  assert.match(styles, /<w:sz w:val="24"\/>/)
})

test('plain text without Markdown syntax stays one paragraph per line', async () => {
  const { buffer } = await build({ title: '', subtitle: '', author: '', toc: false }, '第一行\n第二行\n\n第三行\n')
  const document = readZip(buffer).get('word/document.xml').toString('utf8')
  assert.match(document, /第一行<\/w:t><w:br\/><w:t xml:space="preserve">第二行/)
  assert.equal((document.match(/<w:p>/g) ?? []).length, 2)
})

test('reads PNG, GIF and JPEG dimensions from file headers', () => {
  assert.deepEqual(imageSize(makePng(640, 480), 'png'), { width: 640, height: 480 })
  const gif = Buffer.alloc(16)
  gif.write('GIF89a', 0, 'ascii')
  gif.writeUInt16LE(320, 6)
  gif.writeUInt16LE(200, 8)
  assert.deepEqual(imageSize(gif, 'gif'), { width: 320, height: 200 })
  const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x01, 0x2c, 0x02, 0x58, 0x03])
  assert.deepEqual(imageSize(jpeg, 'jpg'), { width: 600, height: 300 })
  assert.equal(imageSize(Buffer.from('nope'), 'png'), null)
})

test('a missing image degrades to its alt text instead of failing the build', async () => {
  const { buffer, stats } = await build({ loadImage: null })
  const document = readZip(buffer).get('word/document.xml').toString('utf8')
  assert.equal(stats.images, 0)
  assert.doesNotMatch(document, /<w:drawing>/)
  assert.match(document, /示意图/)
})

test('an empty document still produces a valid package', async () => {
  const { buffer } = await build({ title: '', toc: false }, '')
  const entries = readZip(buffer)
  assertWellFormedXml(entries.get('word/document.xml').toString('utf8'))
  assert.match(entries.get('word/document.xml').toString('utf8'), /<w:sectPr>/)
})
