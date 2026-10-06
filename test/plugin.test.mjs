import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { apply, fileStem, timestamp } from '../index.js'
import { readZip } from './helpers.mjs'

/** Register the plugin against a stub host and hand back the tool it mounted. */
function mount(config = {}, extraRoots = []) {
  const registered = []
  const ctx = { tools: { register: (definition) => registered.push(definition) } }
  const allowedRoots = [...extraRoots, ...(config.allowedRoots ?? [])]
  apply(ctx, allowedRoots.length > 0 ? { ...config, allowedRoots } : config)
  assert.equal(registered.length, 1)
  return registered[0]
}

const exec = { signal: new AbortController().signal }

/** Run `fn` with a scratch directory that is allowed as a write root. */
async function withTempDir(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'docx-export-'))
  try {
    return await fn(dir, (config = {}) => mount(config, [dir]))
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

test('registers a tool with a JSON-Schema parameter set', () => {
  const tool = mount()
  assert.equal(tool.name, 'export_docx')
  assert.match(tool.description, /Word document/)
  assert.equal(tool.parameters.type, 'object')
  assert.ok(tool.parameters.properties.content)
  assert.ok(tool.parameters.properties.outputPath)
  assert.equal(tool.parameters.required.length, 0)
  assert.equal(tool.isConcurrencySafe(), true)
  assert.equal(tool.presentCall({ outputPath: 'a.docx' }).card, 'generic')
})

test('config renames the tool and the defaults reach the document', () => {
  const tool = mount({ toolName: 'save_word', footer: false })
  assert.equal(tool.name, 'save_word')
})

test('writes a .docx at the requested path and reports it', async () => {
  await withTempDir(async (dir, mountHere) => {
    const tool = mountHere()
    const target = join(dir, 'report.docx')
    const result = await tool.execute({ content: '# 标题\n\n正文 **加粗**\n', title: '标题', outputPath: target }, exec)
    assert.match(result, /Wrote Word document:/)
    assert.match(result, /report\.docx/)
    assert.ok(existsSync(target))
    const entries = readZip(readFileSync(target))
    assert.match(entries.get('word/document.xml').toString('utf8'), /正文 /)
  })
})

test('defaults the output into the working directory with a title-derived name', async () => {
  await withTempDir(async (dir, mountHere) => {
    const tool = mountHere({ outputDir: dir })
    const result = await tool.execute({ content: '# 季度/报告\n\n正文\n' }, exec)
    const path = result.split('\n')[0].replace('Wrote Word document: ', '')
    assert.ok(path.startsWith(dir))
    assert.match(path, /季度 报告-\d{8}-\d{6}\.docx$/)
    assert.ok(existsSync(path))
  })
})

test('adds the .docx extension and creates missing directories', async () => {
  await withTempDir(async (dir, mountHere) => {
    const tool = mountHere()
    const target = join(dir, 'nested', 'deep', 'out')
    const result = await tool.execute({ content: 'text', outputPath: target }, exec)
    assert.match(result, /out\.docx/)
    assert.ok(existsSync(`${target}.docx`))
  })
})

test('overwrite: false keeps the existing file and picks a new name', async () => {
  await withTempDir(async (dir, mountHere) => {
    const tool = mountHere()
    const target = join(dir, 'once.docx')
    writeFileSync(target, 'placeholder')
    const result = await tool.execute({ content: 'text', outputPath: target, overwrite: false }, exec)
    assert.equal(readFileSync(target, 'utf8'), 'placeholder')
    assert.match(result, /once-1\.docx/)
    assert.ok(existsSync(join(dir, 'once-1.docx')))
  })
})

test('reads the body from sourcePath and resolves images beside it', async () => {
  await withTempDir(async (dir, mountHere) => {
    const tool = mountHere()
    writeFileSync(join(dir, 'body.md'), '# From file\n\n![pic](pic.png)\n')
    writeFileSync(
      join(dir, 'pic.png'),
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8AAAwAB/AGtE0zaAAAAAElFTkSuQmCC',
        'base64',
      ),
    )
    const target = join(dir, 'from-file.docx')
    const result = await tool.execute({ sourcePath: join(dir, 'body.md'), outputPath: target }, exec)
    assert.match(result, /1 images/)
    assert.ok(readZip(readFileSync(target)).has('word/media/image1.png'))
  })
})

test('rejects an empty body and a missing source; format text stays literal', async () => {
  await withTempDir(async (dir, mountHere) => {
    const tool = mountHere()
    await assert.rejects(() => tool.execute({}, exec), /Nothing to export/)
    await assert.rejects(() => tool.execute({ sourcePath: join(dir, 'nope.md') }, exec), /does not exist/)
    const target = join(dir, 'plain.docx')
    await tool.execute({ content: '**not bold**\n# not a heading', outputPath: target, format: 'text' }, exec)
    const xml = readZip(readFileSync(target)).get('word/document.xml').toString('utf8')
    assert.match(xml, /\*\*not bold\*\*/)
    assert.match(xml, /# not a heading/)
    assert.doesNotMatch(xml, /Heading1/)
  })
})

test('refuses to write outside the session workspace by default', async () => {
  const tool = mount()
  const outside = process.platform === 'win32' ? 'C:\\Windows\\Temp\\escape.docx' : '/tmp/escape.docx'
  await assert.rejects(() => tool.execute({ content: 'x', outputPath: outside }, exec), /outside the session workspace/)
})

test('allowedRoots extends the writable area', async () => {
  await withTempDir(async (dir) => {
    const tool = mount({ allowedRoots: [dir] })
    const result = await tool.execute({ content: 'x', outputPath: join(dir, 'ok.docx') }, exec)
    assert.match(result, /ok\.docx/)
    const any = mount({ allowOutsideWorkspace: true })
    const result2 = await any.execute({ content: 'x', outputPath: join(dir, 'ok2.docx') }, exec)
    assert.match(result2, /ok2\.docx/)
  })
})

test('fileStem and timestamp are filesystem-safe', () => {
  assert.equal(fileStem('季度/报告: 2025?'), '季度 报告 2025')
  assert.equal(fileStem('   '), 'document')
  assert.equal(fileStem('a'.repeat(120)).length, 60)
  assert.match(timestamp(new Date('2025-02-03T04:05:06')), /^20250203-040506$/)
})
