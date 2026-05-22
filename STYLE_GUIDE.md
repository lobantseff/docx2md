# Markdown Style Guide

This guide covers the Markdown conventions used in this repository. Following
these rules ensures that `md-docx` produces clean `.docx` output and that
roundtrip conversion (doc→md→doc) is reliable.

## Markdown Dialect

The converter uses Pandoc's `markdown` format with these extensions:

| Extension            | Status   | Meaning                            |
| -------------------- | -------- | ---------------------------------- | ------------- |
| `pipe_tables`        | enabled  | Tables use `                       | ` delimiters  |
| `fenced_code_blocks` | enabled  | Code blocks use ` ``` ` fences     |
| `line_blocks`        | enabled  | Lines starting with `              | ` are literal |
| `strikeout`          | enabled  | `~~text~~` for strikethrough       |
| `raw_html`           | enabled  | Inline HTML is passed through      |
| `simple_tables`      | disabled | —                                  |
| `multiline_tables`   | disabled | —                                  |
| `grid_tables`        | disabled | —                                  |
| `smart`              | disabled | Straight quotes only, no em-dashes |

## YAML Frontmatter

Every document starts with a YAML frontmatter block containing the title:

```yaml
---
title: Document Title Here
---
```

The `title` is used as the document title when converting to `.docx`. Additional
fields may be added in the future — only `title` is used by the converter today.

## Headings

- Use ATX-style headings (`#`, `##`, `###`, `####`).
- Maximum depth is `####` (h4). Deeper nesting usually signals a structural
  issue — consider restructuring instead.
- Leave a blank line before and after each heading.

## Tables

- **Pipe tables only.** Do not use grid tables, simple tables, or multiline
  tables — they are disabled in the converter.
- **No merged cells.** Pandoc pipe tables do not support cell spans. If the
  source `.docx` has merged cells, restructure into flat tables.
- Use `|---|` separator rows (alignment colons `:` are allowed).

Example:

```markdown
| Req ID  | Description            | Priority |
| ------- | ---------------------- | -------- |
| DRS.001 | Functional requirement | High     |
| DRS.002 | Interface requirement  | Medium   |
```

## Page Breaks

Page breaks are preserved as raw HTML divs:

```html
<div class="page-break"></div>
```

This div is converted to a Word page break (`w:br w:type="page"`) during
`md2doc` conversion. Do not modify the class name or add content inside the div.

## Images

- The converter stores images in `<document_stem>_images/` alongside the `.md`
  file.
- Reference images with relative paths:

```markdown
![Alt text](Document_Name_images/image1.png)
```

- Pandoc attributes (width/height) are preserved in curly braces:

```markdown
![](Document_Name_images/figure.png){width="5.0in" height="3.0in"}
```

## Text Formatting

- **Bold:** `**text**`
- _Italic:_ `*text*`
- ~~Strikethrough:~~ `~~text~~`
- `Code:` `` `inline code` ``
- Use straight quotes (`"`, `'`), not curly/smart quotes. The converter
  straightens smart quotes automatically during `doc2md`.

## Lists

- Use `-` for unordered lists.
- Use `1.` for ordered lists (Pandoc renumbers automatically).
- Leave a blank line before the first list item.
