import assert from 'node:assert/strict'
import test from 'node:test'
import { parseInline, parseMarkdown } from '../lib/markdown.js'

const types = (blocks) => blocks.map((block) => block.type)

test('parses headings, paragraphs and rules', () => {
  const { blocks } = parseMarkdown('# 标题\n\n正文一段。\n\n---\n\n## Second\n')
  assert.deepEqual(types(blocks), ['heading', 'paragraph', 'hr', 'heading'])
  assert.equal(blocks[0].level, 1)
  assert.equal(blocks[0].text, '标题')
  assert.equal(blocks[3].level, 2)
})

test('parses setext headings', () => {
  const { blocks } = parseMarkdown('Title\n=====\n\nSub\n---\n')
  assert.deepEqual(types(blocks), ['heading', 'heading'])
  assert.deepEqual(
    blocks.map((block) => block.level),
    [1, 2],
  )
})

test('strips YAML frontmatter', () => {
  const { frontmatter, blocks } = parseMarkdown('---\ntitle: x\nauthor: y\n---\n\n# H\n')
  assert.match(frontmatter, /title: x/)
  assert.deepEqual(types(blocks), ['heading'])
})

test('parses bullet lists with nesting and tasks', () => {
  const { blocks } = parseMarkdown('- a\n- b\n  - b1\n  - b2\n- [x] done\n- [ ] todo\n')
  assert.deepEqual(types(blocks), ['list'])
  const list = blocks[0]
  assert.equal(list.ordered, false)
  assert.equal(list.items.length, 4)
  assert.equal(list.items[0].text, 'a')
  assert.deepEqual(types(list.items[1].blocks), ['list'])
  assert.equal(list.items[1].blocks[0].items.length, 2)
  assert.equal(list.items[2].checked, true)
  assert.equal(list.items[3].checked, false)
})

test('parses ordered lists with a start number', () => {
  const { blocks } = parseMarkdown('3. three\n4. four\n')
  assert.equal(blocks[0].ordered, true)
  assert.equal(blocks[0].start, 3)
  assert.equal(blocks[0].items.length, 2)
})

test('parses fenced code with a language and keeps blank lines', () => {
  const { blocks } = parseMarkdown('```js\nconst a = 1\n\nconst b = 2\n```\n')
  assert.equal(blocks[0].type, 'code')
  assert.equal(blocks[0].lang, 'js')
  assert.equal(blocks[0].text, 'const a = 1\n\nconst b = 2')
})

test('parses GFM tables with alignment', () => {
  const { blocks } = parseMarkdown('| a | b | c |\n|:--|:-:|--:|\n| 1 | 2 | 3 |\n| 4 | 5 |\n')
  const table = blocks[0]
  assert.equal(table.type, 'table')
  assert.deepEqual(table.header, ['a', 'b', 'c'])
  assert.deepEqual(table.align, ['left', 'center', 'right'])
  assert.deepEqual(table.rows[0], ['1', '2', '3'])
  assert.deepEqual(table.rows[1], ['4', '5', ''])
})

test('parses blockquotes recursively', () => {
  const { blocks } = parseMarkdown('> quoted **bold**\n> more\n')
  assert.equal(blocks[0].type, 'quote')
  assert.equal(blocks[0].blocks[0].type, 'paragraph')
  assert.match(blocks[0].blocks[0].text, /quoted/)
})

test('keeps a standalone image as an image block', () => {
  const { blocks } = parseMarkdown('![alt text](pics/a.png)\n')
  assert.deepEqual(blocks[0], { type: 'image', src: 'pics/a.png', alt: 'alt text' })
})

test('inline parser handles emphasis, code, links and escapes', () => {
  const runs = parseInline('a **b** *c* `d` [e](https://x.y) ~~f~~ \\*g\\*')
  assert.deepEqual(
    runs.map((run) => [run.text, run.bold ?? false, run.italic ?? false, run.code ?? false, run.strike ?? false]),
    [
      ['a ', false, false, false, false],
      ['b', true, false, false, false],
      [' ', false, false, false, false],
      ['c', false, true, false, false],
      [' ', false, false, false, false],
      ['d', false, false, true, false],
      [' ', false, false, false, false],
      ['e', false, false, false, false],
      [' ', false, false, false, false],
      ['f', false, false, false, true],
      [' *g*', false, false, false, false],
    ],
  )
  assert.equal(runs.find((run) => run.text === 'e').href, 'https://x.y')
})

test('inline parser does not italicize snake_case identifiers', () => {
  const runs = parseInline('call read_file_now() please')
  assert.equal(runs.length, 1)
  assert.equal(runs[0].text, 'call read_file_now() please')
})

test('inline parser decodes entities and drops raw HTML tags', () => {
  const { blocks } = parseMarkdown('<div>a &amp; b</div>\n')
  const runs = parseInline(blocks[0].text)
  assert.equal(
    runs.map((run) => run.text).join(''),
    'a & b',
  )
})
