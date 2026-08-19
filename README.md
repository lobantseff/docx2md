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

After installation, `doc2md` and `md2doc` are available system-wide.

### doc2md — docx → markdown

```bash
doc2md document.docx               # writes document.md next to the input
doc2md document.docx -o output.md  # explicit output file
doc2md document.docx -o out/       # output directory (created if missing)

# Batch mode: mirror a directory of .docx files into markdown
doc2md -d docx/ -o md/
doc2md                             # same as: doc2md -d docx/ -o md/
```

| Flag | Meaning |
| --- | --- |
| `input` | Input `.docx` file (single-file mode) |
| `-o`, `--output` | Output `.md` file, or output directory (with `-d`) |
| `-d`, `--input-dir` | Input directory for batch conversion (default: `docx`) |
| `-V`, `--version` | Print the version and exit |

### md2doc — markdown → docx

```bash
md2doc document.md                    # writes document.docx next to the input
md2doc document.md -o output.docx
md2doc document.md -s reference.docx  # borrow styling from an existing docx
```

| Flag | Meaning |
| --- | --- |
| `input` | Input `.md` file (required) |
| `-o`, `--output` | Output `.docx` file |
| `-s`, `--style-reference` | Reference `.docx` for fonts, heading styles, margins |
| `-V`, `--version` | Print the version and exit |

## Markdown Conventions

See [STYLE_GUIDE.md](STYLE_GUIDE.md) for formatting rules that ensure clean
roundtrip conversion.

## License

[MIT](LICENSE)
