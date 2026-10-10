# rustpyxl

A Rust streaming XLSX core with an **openpyxl compatibility bridge** and a
separate fast streaming API. This is a working migration milestone;
**the complete openpyxl implementation has not been translated to Rust**.
The hot path reads worksheet XML from bounded ZIP member streams; writes spool
to temporary files instead of retaining a Python object for every cell.

## Drop-in API compatibility

Install the compatibility extra, which pins the reference engine tested here:

```sh
python -m pip install ".[compat]"
```

Keep existing openpyxl imports and calls by enabling the bridge at application
startup, before other modules import `load_workbook` by name:

```python
from rustpyxl.compat import install
install()

import openpyxl
from openpyxl.styles import Font

book = openpyxl.load_workbook("source.xlsx", read_only=True)
try:
    for row in book.active.iter_rows(values_only=True):
        print(row)
finally:
    book.close()

book = openpyxl.Workbook()
book.active["A1"] = "styled"
book.active["A1"].font = Font(bold=True)
book.save("edited.xlsx")
```

Alternatively, `import rustpyxl.compat as openpyxl` exposes the same top-level
Workbook and load_workbook API without installing the process-local patch.
Continue importing submodules such as styles from the original openpyxl package.
`rustpyxl.compat.uninstall()` restores the original loaders. Installation changes
only the current process; it does not replace package files or affect other apps.

The bridge returns actual openpyxl workbook, worksheet and cell objects. Normal
and write-only workbooks, cell objects, styles, dates, merges, names, drawings,
macros and mutable editing/saving continue through openpyxl. Supported read-only
value iteration uses Rust. Unsupported values encountered in a native batch
resume through the reference parser at the first row not delivered; shared
formulas, dates and large integers retain their reference types/semantics.

Shared strings use one Rust-backed sequence, including when the reference parser
handles a fallback. The reference loader does not allocate a duplicate Python
string table. File-like inputs, rich-text mode, dimensions reset to unknown, and
openpyxl versions other than 3.1.5 use reference paths. Only value iteration with
known width (from dimensions or explicit max_col) is accelerated in this bridge.
Mutable edits and compatible writes currently have the reference engine's speed.
The native writer below is faster but supports a narrower API.

See [benchmark report](docs/benchmarks.md) for separate measurements of the bridge
and standalone engine. Future port work replaces reference paths feature by
feature; this bridge preserves their behavior during that migration.

## Build and test

Requires Rust with edition 2024 support and CPython 3.10 or newer.

```sh
python -m venv .venv
source .venv/bin/activate
python -m pip install maturin pytest openpyxl==3.1.5
maturin develop --release
cargo test
cargo clippy --all-targets --features python -- -D warnings
python -m pytest -q
```

For a distributable wheel: `maturin build --release --out dist`.
The Python package has no openpyxl runtime dependency. Use the locked dependency
versions with `cargo build --locked` or `maturin build --release --locked`.
The native engine is also usable as a Rust library (`_native::read`,
`_native::write`, `_native::Cell`); the `python` feature enables bindings.

## Standalone streaming API

```python
from rustpyxl import Workbook, load_workbook

with Workbook(write_only=True) as book:
    sheet = book.create_sheet("Data")
    sheet.append(["id", "name", "amount"])
    sheet.append_rows(
        ((i, f"item-{i}", i * 0.25) for i in range(100_000)),
        batch_size=1024,
    )
    book.save("data.xlsx")

with load_workbook("data.xlsx", read_only=True, data_only=True) as book:
    for row in book["Data"].iter_rows(values_only=True, batch_size=1024):
        print(row)
```

`append()` supports the familiar one-row call. `append_rows()` consumes generators
in bounded batches to amortize Python/Rust calls. Bounds to `iter_rows()` are
inclusive and one-based. Access sheets through `.worksheets`, `.sheetnames`,
`.active` (first sheet in this milestone), or `book[title]`.

## Implemented compatibility

- Read-only value tuples and write-only append/save, multiple sheets, context
  managers, independent iterators, and row/column iteration bounds.
- None, booleans, signed 64-bit integers, finite floats, Unicode strings,
  shared strings, inline strings, sparse cells and missing rows.
- XML whitespace, entities, CDATA, plain text from rich strings, and Excel errors.
- Ordinary formulas as strings starting with `=`; `data_only=True` uses stored
  caches and returns None for absent caches. No formula calculation.
- Stored and DEFLATE ZIP compression. CRC and decoded sizes checked when a
  member is fully consumed. Stopping iteration early does not validate its tail.
- Atomic output publication: writes a sibling temporary archive, finalizes and
  synchronizes it, then replaces the destination. Failed saves can be retried.
  Temporary worksheet spools are discarded after save/close/destruction.

## Standalone engine limitations and differences

Mutable workbooks/cell objects, styles, numeric and ISO dates, durations,
tables, merges, charts, images, macros, hyperlinks, named ranges, pivot tables,
comments, preservation of workbook parts, external links, and arbitrary stream
inputs are not ported. Numeric date/custom-format cells and ISO date values
raise errors. Shared/array/data-table formulas raise errors unless requesting
cached values with `data_only=True`.

These limitations apply to the standalone API, not the compatibility bridge,
which delegates unported features. The standalone reader extracts supported
cell values; it does not preserve formatting or
unsupported workbook parts. It uses the standard `xl/workbook.xml` package
location. Active-sheet metadata is not interpreted. Only UTF-8 XML is supported.

Dimensions provide a minimum row width. Physical rows and columns beyond an
incorrect dimension are retained; trailing blank rows mentioned only in
dimensions are omitted. Empty gaps without a declared width yield tuples,
where openpyxl sometimes yields lists. Sheet names must be unique ignoring
case; duplicate titles raise instead of being automatically renamed. Strings
longer than Excel's 32,767-character limit raise instead of truncating. Integer
inputs outside signed 64-bit range raise. Empty strings follow openpyxl's
write-only normalization to empty cells. A failed bulk append can leave earlier
rows committed to the worksheet spool; each individual row validates before
being appended.

## Memory and speed

Reading uses approximately O(shared-string bytes + shared-string count + batch
cells + XML/DEFLATE buffers + package metadata) memory. Shared strings are held
once in compact Rust storage and referenced across iterators. Python value
objects are created only for the current batch. Holding yielded rows in a list
still uses memory proportional to those rows.

Writing uses O(batch cells + largest serialized row + sheet buffers + ZIP
metadata) RAM and O(uncompressed worksheet XML) temporary disk space. Inline
strings avoid a global deduplication map. More sheets require more open spool
files. Compression level defaults to 4 (2–5× faster than 6 with similar output
size on tabular XML) and can be set to 0–9 with `save(path, compression_level=6)`.

The engine reuses XML/row buffers, borrows text and frequently used attributes,
releases the Python GIL for parsing/writing/compression, and uses zlib-rs for
DEFLATE. A worker thread parses up to two batches ahead while Python consumes
the current one. Read/write batches default to 1,024 rows; decrease the batch size for
very wide or large-string rows.

See [benchmark report](docs/benchmarks.md) for measured results and limitations.
Run a fresh-process comparison with:

```sh
python benchmarks/benchmark.py --rows 100000 --repeats 3 --output docs/results.json
```

The benchmark compares openpyxl 3.1.5 read_only/write_only modes, includes
Python value conversion, row generation, disk spooling and ZIP save, and checks
written row digests outside timing. Numeric, mixed inline-string, and unique
shared-string fixtures are covered. RSS is whole-process peak memory, including
imports; elapsed times exclude imports. Avoid extrapolating to other hardware,
style-heavy files or full openpyxl workloads.

## Remaining port work

1. A full style and date/duration engine with 1900/1904 epochs, formula
   translation, correct dimension controls and broader package discovery.
2. Disk-backed shared strings for large cardinality, configurable resource
   limits, native column/row projection and Arrow-oriented batch exports.
3. Mutable workbook/cell storage, styles and lossless OOXML part preservation.
4. Feature-by-feature support for tables, names, merges, links, drawings,
   charts, comments, validations, macros and the rest of the openpyxl API.
5. Real-world file corpora and cross-platform wheels before a stable release.

Primary sources and compatibility research are in [research.md](docs/research.md).
