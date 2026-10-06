# 项目周报：DSH Word 导出插件

> 本示例同时充当转换器的回归样本：标题、强调、列表、表格、代码块、引用与图片都会被转成 Word 原生元素。

## 1. 本周进展

- 完成 `export_docx` 工具与 Markdown → OOXML 转换器
- 支持 **加粗**、*斜体*、`行内代码`、~~删除线~~ 与 [链接](https://example.com/docs)
- 列表支持嵌套：
  1. 有序列表
  2. 二级条目
     - 三级条目
- 任务清单：
  - [x] 生成 .docx 包结构
  - [ ] 补充更多模板样式

## 2. 关键数据

| 模块 | 代码行数 | 状态 | 说明 |
|:-----|--------:|:----:|------|
| 压缩写入器 | 120 | 完成 | 纯 Node，无依赖 |
| Markdown 解析 | 330 | 完成 | 支持 GFM 表格 |
| OOXML 生成 | 720 | 完成 | 样式/编号/表格/图片 |

## 3. 架构示意

![导出流程示意图](diagram.png)

## 4. 代码示例

```js
const { buildDocx } = await import('dsh-docx-export')
const { buffer } = await buildDocx('# 标题\n\n正文', { toc: true })
```

## 5. 引用与说明

> 导出的文档使用 Word 原生样式（Heading 1-6、List Paragraph、Quote、Code Block），
> 因此可以直接套用企业模板、自动生成目录与页码。

---

## 6. 下周计划

1. 增加 LaTeX 公式转 OMML
2. 支持页眉页脚自定义
3. 提供参考模板（公文 / 报告 / 论文）
