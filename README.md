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

## CLI Usage

After installation, `doc2md` and `md2doc` are available system-wide:

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

## Markdown Conventions

See [STYLE_GUIDE.md](STYLE_GUIDE.md) for formatting rules that ensure clean
roundtrip conversion.
## License

[MIT](LICENSE)