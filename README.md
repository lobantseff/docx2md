# md-docx

Bidirectional converter between `.docx` and Markdown, built on Pandoc.

## Prerequisites

- [uv](https://docs.astral.sh/uv/) (Python package manager)

## Installation

Install as a global CLI tool:

```bash
uv tool install .
```

After installation, `md-docx` is available system-wide:

```bash
md-docx doc2md document.docx
md-docx md2doc document.md
```

## CLI Usage

Install and run via `uv`:

```bash
# Convert docx → markdown
uv run md-docx doc2md document.docx
uv run md-docx doc2md document.docx -o output.md

# Convert markdown → docx
uv run md-docx md2doc document.md
uv run md-docx md2doc document.md -o output.docx -s reference.docx

# Batch convert a directory of docx files
uv run md-docx doc2md -d docx/ -o md/
```

Or run the script directly:

```bash
uv run --script src/md_docx/main.py doc2md document.docx
```

## Markdown Conventions

See [STYLE_GUIDE.md](STYLE_GUIDE.md) for formatting rules that ensure clean
roundtrip conversion.

## Tests

```bash
uv run pytest
```

## VS Code Tips

## Recommended extensions for Visual Studio Code

- Install:
  - 1. Open Extensions: `Ctrl + Shift + X`
  - 2. Search for "@recommended" and install Workspace recommendations.

VScode hints:

- Open .md files in preview mode:
  - `Ctrl + Shift + P` -> `Markdown: Open Preview to the Side`
  - `Ctrl + Shift + P` -> `Markdown: Open Preview`

Add hotkeys for these two:

- `Ctrl + Shift + P` -> `Preferences: Open Keyboard Shortcuts`
- Type: "Markdown: Open Preview", assign Preferred Shortcut.
