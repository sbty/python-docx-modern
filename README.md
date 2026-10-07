# python-docx-modern

*python-docx-modern* is a fork of [python-docx](https://github.com/python-openxml/python-docx),
the Python library for reading, creating, and updating Microsoft Word 2007+ (.docx) files.

It starts from python-docx v1.2.0 and keeps the same API and import name (`docx`); behavior
differs only where bugs were fixed (see [HISTORY.rst](HISTORY.rst)). The fork modernizes the
build and tooling, types the whole package, and fixes a number of bugs, including several
open upstream issues.

## What's different from python-docx 1.2.0

- **Fully typed.** Every module under `src/docx` passes `mypy --strict` and pyright in strict
  mode. Both checks run as part of the test suite, so the package stays at zero type errors.
- **Modern build.** Packaging uses hatchling instead of setuptools (configured in
  `pyproject.toml`). Python 3.10+ is required; the lxml minimum is 4.9.0.
- **Bug fixes**, each covered by tests:
  - upstream issues #1475, #1494, #1539, #1541, #1559, #1600 and #1609;
  - table alignment and direction, comment styles, styles without `w:type`, image DPI and
    height, measures with fractional or invalid values, and inline shapes such as charts,
    SmartArt and pictures without image data.

See [HISTORY.rst](HISTORY.rst) (the "Unreleased" section) for the complete list.

## Installation

The fork is not published on PyPI. Install it from GitHub:

```
pip install "python-docx @ git+https://github.com/sbty/python-docx-modern.git@modernize"
```

The distribution is still named `python-docx`, so this replaces an installed upstream
python-docx in the same environment.

## Example

```python
>>> from docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

The API is the same as python-docx, so the
[python-docx documentation](https://python-docx.readthedocs.org/en/latest/) applies.

## Development

The project uses [uv](https://docs.astral.sh/uv/):

```
uv sync
uv run pytest                                  # unit tests, including the typing gate
uv run pytest tests/test_typing.py             # typing gate only (mypy + pyright, src/docx)
uv run behave --format progress --tags=-wip    # acceptance tests
uv run ruff check . && uv run ruff format --check .
```

Running `uv run pyright` directly also checks the tests in standard mode and reports known
errors there (mostly mocks passed to typed parameters), plus one in pyright's bundled typeshed
stubs. `src/docx` must stay at zero, which the typing gate enforces.

## Credits and license

python-docx is written by Steve Canny and contributors; this fork builds on their work.
Released under the MIT license; see [LICENSE](LICENSE).
