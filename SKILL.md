---
name: chinese-paper-docx-format
description: Use this skill by default whenever the user asks to create, edit, revise, polish, format, or validate a Chinese academic Word/DOCX deliverable, including Chinese papers, course papers, case papers, theory-framework papers, literature-review reports, thesis-like assignments, and Chinese research reports. Also use it together with any DOCX/Word skill when generating or editing a Chinese academic DOCX, unless the user provides a template or conflicting formatting requirements.
---

# Chinese Paper DOCX Format

## Core Rule

When producing or revising a Chinese academic Word/DOCX deliverable for the user, apply these defaults unless the current task provides a template or a more specific conflicting requirement. Current user instructions override this skill.

Use `scripts/format_chinese_paper_docx.py` after creating or editing a DOCX when a deterministic formatting pass is appropriate.

```powershell
python scripts\format_chinese_paper_docx.py input.docx --output output.docx --json
```

If editing a user-named final document in place is expected, pass the same path to `--output`.

## Default Format

- Page setup: A4; top, bottom, left, and right margins all `2.5 cm`.
- Fonts: Chinese text uses `宋体`; western text and numbers use `Times New Roman`.
- Chinese punctuation: keep Chinese full-width punctuation in the East Asian font. Do not force Chinese punctuation into Times New Roman.
- Quotation marks: replace straight ASCII double quotes `"` in Chinese prose with Chinese curly quotes `“”`.
- Body paragraphs: first-line indent `2` Chinese characters; justified alignment; `1.5` line spacing; space before and after `0`.
- Heading paragraphs: `1.5` line spacing; space before and after `0`; indentation follows the template or user request. Usually headings have no first-line indent.
- Figures, tables, and captions: images, tables, figure captions, and table captions are centered.
- Table text: horizontally centered and vertically centered.
- Three-line tables: center the table; keep only the top border, header-bottom border, and bottom border. Do not shade the first row. Use `宋体` for Chinese and `Times New Roman` for western text and numbers.
- References: use ordinary `[1]`, `[2]` numbering when the user asks for plain numbered references. Reference paragraphs are justified, with space before and after `0`. If a template already has automatic numbering or cross-references, preserve that mechanism and avoid duplicate numbering.

## Conflict Handling

- A user-provided Word template wins over these defaults.
- A current-turn instruction wins over this skill.
- If another DOCX skill says different Chinese academic defaults, prefer this skill for Chinese papers and similar academic DOCX deliverables.
- Do not rewrite scholarly content merely to format it. Keep content changes separate from formatting unless the user asks for polishing.

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
- References use `[1]`, `[2]` numbering when plain numbering is expected.
