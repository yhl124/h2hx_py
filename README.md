# h2hx-py

Python package that converts `.hwp` files into `.hwpx`.

This project is a Python port of [`hwp2hwpx`](https://github.com/neolord0/hwp2hwpx), originally implemented in Java. The porting work was done with Codex (GPT-5.4).

It currently focuses on producing a valid HWPX package and preserving common document structures such as paragraphs, tables, equations, and fields.

## Status

- Python 3.11+
- CLI entry point: `h2hx`
- Library API: `h2hx_py.convert_file()`
- Test fixtures are bundled inside this repository

## Installation

```bash
uv sync
```

Or with pip:

```bash
pip install -e .
```

## Usage

```bash
h2hx input.hwp
```

Write to a specific output path:

```bash
h2hx input.hwp --output output.hwpx
```

From Python:

```python
from h2hx_py import convert_file

result = convert_file("input.hwp", "output.hwpx")
print(result.output_path)
print(result.warnings)
```

## Development

Run tests:

```bash
uv run python -m unittest discover -s tests
```

## Repository Notes

- `.hwp` fixtures in `tests/fixtures/` are small sample files used for regression tests.
- `uv.lock` is committed for reproducible local development.
- Virtual environments and OS/editor-generated files are excluded via `.gitignore`.
