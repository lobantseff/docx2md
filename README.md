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
uv tool install git+https://github.com/lobantseff/docx2md@v0.1.8
uv tool install git+https://github.com/lobantseff/docx2md@master
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

| Flag                | Meaning                                                |
| ------------------- | ------------------------------------------------------ |
| `input`             | Input `.docx` file (single-file mode)                  |
| `-o`, `--output`    | Output `.md` file, or output directory (with `-d`)     |
| `-d`, `--input-dir` | Input directory for batch conversion (default: `docx`) |
| `-V`, `--version`   | Print the version and exit                             |

### md2doc — markdown → docx

```bash
md2doc document.md                    # writes document.docx next to the input
md2doc document.md -o output.docx
md2doc document.md -s reference.docx  # borrow styling from an existing docx
```

| Flag                      | Meaning                                              |
| ------------------------- | ---------------------------------------------------- |
| `input`                   | Input `.md` file (required)                          |
| `-o`, `--output`          | Output `.docx` file                                  |
| `-s`, `--style-reference` | Reference `.docx` for fonts, heading styles, margins |
| `-V`, `--version`         | Print the version and exit                           |

## Markdown Conventions

See [STYLE_GUIDE.md](STYLE_GUIDE.md) for formatting rules that ensure clean
roundtrip conversion.

## Development: keeping environments in sync

There are two independent installations, and they drift apart silently:

| Environment          | Path                  | Refresh with                    |
| -------------------- | --------------------- | ------------------------------- |
| Dev venv (editable)  | `.venv/bin/doc2md`    | `uv sync`                       |
| Global CLI (uv tool) | `~/.local/bin/doc2md` | `uv tool install --reinstall .` |

When the dev venv is active your shell prompt shows `(doc2md)`, and
`.venv/bin` precedes `~/.local/bin` on `PATH` — so `uv tool install` appears to
do nothing, because you keep running the venv's copy. Always confirm which
binary you are testing:

```bash
type -a doc2md
```

To reinstall the global CLI and verify it, leave the venv first:

```bash
deactivate && uv tool install --reinstall . && doc2md --version
```

After a version bump, refresh both. `increment_version.py --install` covers
only the global tool:

```bash
./increment_version.py patch --install && uv sync
```

`--version` reads the installed distribution metadata, not `pyproject.toml`.
An editable install runs current source while still reporting the version
recorded at install time, so a stale `--version` means the venv needs `uv sync`
— not that your code is old.

### Recovering from a broken `pypandoc`

```
AttributeError: module 'pypandoc' has no attribute 'convert_file'
```

`pypandoc_binary` ships the `pypandoc` module *and* a bundled pandoc, so this
project depends on it alone. Environments created before that dependency was
deduplicated still contain plain `pypandoc` as well; when a sync prunes it, the
shared module files both distributions own are deleted, leaving a
`site-packages/pypandoc/` that holds only `files/`. Python imports that as a
namespace package without complaint, so the failure surfaces only at the call
site. Restore the module with:

```bash
uv sync --reinstall-package pypandoc-binary   # dev venv
uv tool install --reinstall .                 # global CLI
```

Prefer `uv sync --reinstall` over plain `uv sync` when recovering from any
unexplained import error: it reinstalls every package instead of incrementally
pruning, so it cannot leave a half-deleted package behind.

## License

[MIT](LICENSE)
