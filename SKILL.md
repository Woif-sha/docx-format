---
name: chinese-paper-docx-format
description: Use this skill by default whenever the user asks to create, read, edit, revise, polish, format, or validate a Word/DOCX deliverable, especially Chinese academic papers, course papers, case papers, theory-framework papers, literature-review reports, thesis-like assignments, and Chinese research reports. This skill is the user's default DOCX route unless the current task provides a template or conflicting formatting requirements.
---

# Chinese Paper DOCX Format

## Core Rule

When producing or revising a Word/DOCX deliverable for the user, use this skill as the default DOCX route. For Chinese academic deliverables, apply these defaults unless the current task provides a template or a more specific conflicting requirement. Current user instructions override this skill.

Use `scripts/format_chinese_paper_docx.py` after creating or editing a DOCX when a deterministic formatting pass is appropriate. This script is now the default post-processing pass for generated DOCX files in this user's workspace.

```powershell
python scripts\format_chinese_paper_docx.py input.docx --output output.docx --json
```

If editing a user-named final document in place is expected, pass the same path to `--output`.

On Windows, the script also attempts a Word COM pass by default to mimic the user's referenced Word workflow: find Chinese curly quotes, select matches, apply full-width character width, and force `宋体`. Use `--skip-word-com-fullwidth-quotes` only when Word automation is unavailable or undesirable.

## Default Format

- Page setup: A4; top, bottom, left, and right margins all `2.5 cm`.
- Fonts: Chinese text uses `宋体`; western text and numbers use `Times New Roman`.
- Chinese punctuation: keep Chinese full-width punctuation in the East Asian font. Do not force Chinese punctuation into Times New Roman.
- Quotation marks: implement the Word workflow shown in the user's reference, conceptually: find Chinese curly quote marks, select all matches, and convert/format them as full-width Chinese punctuation. In automation, replace straight ASCII double quotes `"` with Chinese curly quotes `“”`; also process existing `“”`; put each Chinese quote mark in its own run and force `ascii`, `hAnsi`, and `eastAsia` fonts to `宋体`, not inherited `Times New Roman`. Reference: https://zhuanlan.zhihu.com/p/1973714955296596248.
- Body paragraphs: first-line indent `2` Chinese characters; justified alignment; `1.5` line spacing; space before and after `0`.
- Heading paragraphs: `1.5` line spacing; space before and after `0`; indentation follows the template or user request. Usually headings have no first-line indent.
- Figures, tables, and captions: images, tables, figure captions, and table captions are centered.
- Table text: horizontally centered and vertically centered.
- Three-line tables: center the table; keep only the top border, header-bottom border, and bottom border. Top and bottom borders are `1.5 pt`; the header-bottom border is `0.75 pt`. Do not shade the first row. Use `宋体` for Chinese and `Times New Roman` for western text and numbers.
- References: use ordinary `[1]`, `[2]` numbering when the user asks for plain numbered references. Reference paragraphs are justified, with space before and after `0`. If a template already has automatic numbering or cross-references, preserve that mechanism and avoid duplicate numbering.

## Conflict Handling

- A user-provided Word template wins over these defaults.
- A current-turn instruction wins over this skill.
- Do not rely on the old `docx` skill for this user. This skill is the default route for DOCX generation/editing; use ordinary DOCX libraries or scripts as needed, then run this skill's formatting/validation pass.
- Do not rewrite scholarly content merely to format it. Keep content changes separate from formatting unless the user asks for polishing.

## Template

Use `assets/template.docx` as the bundled reference template for this skill when the user asks for a blank Chinese paper DOCX template. The template demonstrates the default page setup, body paragraph style, heading levels such as `一、`, `1.1`, and `1.1.1`, centered captions, three-line table border weights, Chinese quote marks, and reference numbering.

## Validation Checklist

Before final delivery, read the output back and report the important checks:

- DOCX opens as a valid ZIP/package.
- Title, paragraph count, and table count can be read.
- Page size and all margins match the required defaults where applicable.
- Body and heading line spacing are `1.5`.
- Paragraph space before/after is `0`.
- Body paragraphs are justified and first-line indented by `2` Chinese characters.
- Tables are centered; table-cell paragraphs are centered; cells are vertically centered.
- Three-line tables have no first-row shading and only top, header-bottom, and bottom rules.
- Image paragraphs and figure/table captions are centered.
- Straight ASCII double quote count is `0` in Chinese prose unless intentionally preserved.
- Chinese quote runs use `宋体` for `ascii`, `hAnsi`, `cs`, and `eastAsia`; `hint` is `eastAsia`.
- On Windows with Word available, the Word COM full-width quote pass reports success or a clear fallback reason.
- References use `[1]`, `[2]` numbering when plain numbering is expected.
