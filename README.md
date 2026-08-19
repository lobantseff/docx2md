# doc2md

Bidirectional converter between `.docx` and Markdown, built on Pandoc.

## Prerequisites

- [uv](https://docs.astral.sh/uv/) (Python package manager)

## Installation

Install as a global CLI tool directly from GitHub:

```bash
uv tool install git+https://github.com/lobantseff/docx2md
```

To install a specific tag or branch:

```bash
uv tool install git+https://github.com/lobantseff/docx2md@v0.1.7
```

To upgrade later:

```bash
uv tool upgrade doc2md
```

Or install from a local clone:

```bash
uv tool install .
```

After installation, `doc2md` and `md2doc` are available system-wide:

```bash
doc2md document.docx
md2doc document.md
```

## CLI Usage

```bash
# Convert docx → markdown
doc2md document.docx
doc2md document.docx -o output.md

# Convert markdown → docx
md2doc document.md
md2doc document.md -o output.docx -s reference.docx

# Batch convert a directory of docx files
doc2md -d docx/ -o md/
```

Without installing, run from a clone via `uv`:

```bash
uv run doc2md document.docx
uv run md2doc document.md
```

Or run the module directly:

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
