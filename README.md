# docx-format

`docx-format` 是用于中文 Word/DOCX 论文与课程作业交付物的默认格式化技能。它面向中文学术论文、课程论文、案例论文、理论框架、文献综述和研究报告等场景，提供稳定的 DOCX 生成、修订、排版与校验规则。

## 主要能力

- A4 页面设置，上下左右页边距均为 2.5 cm，页脚距底端 1.27 cm。
- 中文使用宋体，西文与数字使用 Times New Roman，正文颜色统一为黑色。
- 开头大标题设置为宋体小二、居中、加粗。
- 每页下方设置居中页码，页码为小四。
- 正文首行缩进 2 个中文字符，两端对齐，1.5 倍行距，段前段后为 0。
- 一级标题默认四号，正文默认小四。
- 表格使用三线表：顶线和底线 1.5 磅，表头下横线 0.75 磅。
- 表格整体居中，单元格文字水平和垂直居中，表格内文字小五，单倍行距。
- 图例、表例居中，五号字，单倍行距，按章节编号为 `图1-1`、`表1-1`。
- 缺少图例或表例时，会根据表头、相邻正文或当前章节标题拟定可用题名，不使用占位题名。
- 半角直双引号会转换为中文弯引号，并尽量避免中文引号落入西文字体。
- 参考文献支持普通 `[1]`、`[2]` 编号格式。
- 数学公式和变量解释要求以 MathML 为源表示，并尽量转换为 Word 可渲染公式。

## 使用方式

格式化已有 DOCX：

```powershell
python scripts\format_chinese_paper_docx.py input.docx --output output.docx --json
```

原地格式化用户指定的最终文档：

```powershell
python scripts\format_chinese_paper_docx.py input.docx --output input.docx --json
```

如果当前环境不适合调用 Word COM 引号全角化流程，可以跳过该步骤：

```powershell
python scripts\format_chinese_paper_docx.py input.docx --output output.docx --json --skip-word-com-fullwidth-quotes
```

## 模板

参考模板位于：

```text
assets/template.docx
```

模板包含 `一、`、`1.1`、`1.1.1` 等章节层级示例，并展示正文、标题、图表题注、三线表、参考文献和公式等默认格式。

## 仓库说明

本仓库只保留可公开复用的技能说明、脚本和模板。内部注册、索引或本地运行状态文件不应提交到远程仓库，`.json` 文件已通过 `.gitignore` 排除。
