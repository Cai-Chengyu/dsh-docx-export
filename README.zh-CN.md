# Word 导出插件 · dsh-docx-export

[English](README.md) | **中文**

把会话里生成的文本导出成**真正的 Word 文档**（`.docx`）——不是 Markdown 改后缀，而是带原生样式的 OOXML 包：标题层级、多级列表、表格、代码块、引用、超链接、嵌入图片、目录域与页脚页码，用 Word / WPS / LibreOffice / Google Docs 打开都正常。

零依赖（只用 Node 内置的 `zlib`），不 import 任何 `@deepseek-ai/*` 包，因此对 DSH 版本不敏感；作为 bundle 安装，重启后自动挂载。

```
你：把上面那份分析整理成 Word 文档发我
DSH：（调用 export_docx，写入 G:\work\分析报告-20251006-1209.docx）
```

## 特性

| Markdown 写法 | Word 中的结果 |
|---|---|
| `# 标题` … `###### 标题` | Heading 1–6 样式（可套模板、可生成目录） |
| 段落、`**粗体**`、`*斜体*`、`~~删除线~~`、`` `行内代码` `` | 对应的字符格式（含中文 `w:eastAsia` 字体） |
| `- 项目` / `1. 步骤`（含缩进嵌套） | 真正的多级项目符号 / 编号列表（`numbering.xml`） |
| `- [x] 已完成` / `- [ ] 待办` | ☑ / ☐ 前缀的清单项 |
| ` ```lang ` 代码块 | 等宽字体 + 灰底边框的代码块，保留空行与缩进 |
| `> 引用` | 左侧竖线的引用样式 |
| GFM 表格（含 `:--:` 对齐） | 真正的 Word 表格：表头加粗加底纹、按列对齐、自动列宽 |
| `[文字](https://…)` | 可点击的超链接（外部关系） |
| `![说明](图.png)` | 嵌入图片，按正文宽度自动缩放（PNG/JPEG/GIF/BMP） |
| `---` | 分隔线 |
| 文档开头 `---\ntitle: …\n---` | 自动忽略 YAML frontmatter |

另外：`toc: true` 插入目录域（打开时自动更新），默认带页脚页码，文档属性里写入标题/作者/时间。

## 安装

三种方式任选其一，安装后**重启 DeepSeek Harness**（新 bundle 在下次启动时挂载；桌面版直接在应用内重启即可）。

```powershell
# 1) 从 npm（推荐）
dsh plugin --profile desktop add dsh-docx-export

# 2) 直接从 GitHub 安装
dsh plugin --profile desktop add github:Cai-Chengyu/dsh-docx-export

# 3) 本地目录 / 离线包
dsh plugin --profile desktop add G:\path\to\dsh-docx-export
dsh plugin --profile desktop add .\dsh-docx-export-1.0.0.tgz
```

`dsh` 不在 PATH 上时，用发行版自带 CLI：`"<DeepSeek Harness>\resources\runtime\cli\bin\dsh.cmd"`；也可以用仓库里的 `scripts/install.ps1`（Windows）/ `scripts/install.sh`（macOS、Linux），它会调用官方 CLI 并检查 bundle 是否登记成功。

手动安装（不想用 CLI）：在 profile 目录里 `pnpm add dsh-docx-export`，并把 `dsh-docx-export` 追加到该目录 `package.json` 的 `dsh.profile.bundles` 数组。

## 使用

在会话里直接说需求即可，模型会调用工具：

> 把上面的内容导出成 Word，标题「××分析报告」，要目录

也可以显式调用 `export_docx`：

```json
{
  "content": "# 标题\n\n正文 **加粗**\n\n| 列 | 值 |\n|:--|--:|\n| a | 1 |\n",
  "title": "××分析报告",
  "outputPath": "G:\\dsh\\报告.docx",
  "toc": true,
  "pageSize": "a4"
}
```

### 工具参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `content` | string | 文档正文，默认按 Markdown 解析 |
| `sourcePath` | string | 改为读取已有的 `.md` / `.txt` 文件 |
| `outputPath` | string | 输出 `.docx` 路径（绝对路径或相对会话工作目录）；缺省为 `<outputDir>/<标题>-<时间戳>.docx` |
| `title` / `subtitle` / `author` | string | 标题块与文档属性；`title` 缺省取正文第一个标题 |
| `format` | `markdown` \| `text` | `text` 时逐行原样输出，不解析 Markdown |
| `toc` | boolean | 插入目录域 |
| `pageSize` | `a4` \| `letter` \| `a5` | 纸张大小，默认 A4 |
| `landscape` | boolean | 横向 |
| `footer` | boolean | 页脚页码，默认 true |
| `overwrite` | boolean | 默认 true；false 时若文件已存在则自动加 `-1`、`-2` 后缀 |

工具返回一段文本：写入的绝对路径、段落/标题/表格/图片数量与文件大小。

### 命令行（不经过 DSH 也能用）

```bash
node bin/docx-export.mjs report.md -o report.docx --title "季度报告" --toc
cat notes.md | node bin/docx-export.mjs - -o notes.docx --text
```

`--page a4|letter|a5`、`--landscape`、`--no-footer`、`--font`、`--east-asia`、`--font-size` 均可指定。

## 配置（可选）

在 profile 的 `cordis.patch.yml` 里覆盖默认值：

```yaml
- id: docx-export
  name: dsh-docx-export
  config:
    toolName: export_docx          # 工具名（冲突时可改）
    outputDir: "D:/导出"           # 默认输出目录，空 = 会话工作目录
    fontFamily: Calibri            # 西文字体
    eastAsiaFont: 微软雅黑          # 中文字体
    monoFont: Consolas             # 代码块字体
    fontSize: 11                   # 正文字号（pt）
    pageSize: a4                   # 默认纸张
    toc: false                     # 默认是否插入目录
    footer: true                   # 默认是否带页脚页码
    allowOutsideWorkspace: false   # 是否允许写出会话工作目录
    allowedRoots: []               # 额外允许的目录
    maxContentChars: 400000        # 单次导出正文长度上限
```

### 写入范围（安全默认）

工具默认只允许读写在**会话工作目录**内（`allowOutsideWorkspace: false`），越界会直接拒绝并给出可操作的提示。需要写到桌面、文档目录等位置时，用 `allowedRoots` 精确授权，或把 `allowOutsideWorkspace` 设为 `true`。图片只有位于允许范围内才会被嵌入。

## 开发与测试

```bash
node test/markdown.test.mjs   # Markdown 解析（12 项）
node test/docx.test.mjs       # OOXML 生成与 ZIP 校验（8 项）
node test/plugin.test.mjs     # 工具注册、路径策略、真实落盘（11 项）
npm test                      # = node --test test/（CI 使用）
npm run demo                  # 用 examples/sample.md 生成 examples/sample.docx
```

测试自带 ZIP 中央目录解析 + CRC 校验、XML 良构检查和一个迷你 PNG 编码器，因此不需要任何开发依赖。

## 发布与分享

- 仓库加上 GitHub topic **`dsh-plugin`**，就能被 DSH 插件市场/发现流程收录。
- 提交到社区目录：向 [dshworks/awesome-dsh-plugins](https://github.com/dshworks/awesome-dsh-plugins) 的 `data/plugins.json` 提 PR（字段见该仓库 `data/schema.json`）。
- 发布 npm 包：`npm publish`；随后 `dsh plugin --profile <profile> add dsh-docx-export` 即可安装。
- 详细步骤见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。

## 兼容性

- Node ≥ 20（使用 `node:zlib`、ESM、`structuredClone` 无关特性）。
- DSH 0.1.7 及以上均可：插件不声明 `@deepseek-ai/*` peer 依赖，也不 import 宿主包，只用运行时的 `ctx.tools.register` 原始 JSON-Schema 通道。
- 已在 DSH 0.2.0-rc.2（Windows 桌面版）上实测：安装 → 挂载 → 生成文档 → 结构校验 → LibreOffice 渲染。

## 许可

[MIT](LICENSE)
