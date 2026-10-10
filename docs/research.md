# Streaming openpyxl port: primary-source research

Checked 10 October 2026. The local reference installation is openpyxl 3.1.5; its installed source was inspected directly. Public stable documentation currently identifies itself as 3.1.3. Crate documentation was checked for quick-xml 0.42.0, zip 9.0.0 and PyO3 0.29.3; search snippets can show older releases, so implementation must use the locked dependency versions.

## Initial compatibility contract

A useful first release is a streaming values reader and append-only writer with a familiar Python facade. Full openpyxl parity also includes mutable cells, styles, charts, images, macros, links and many other OOXML parts; do not imply full parity from a values-only implementation. [openpyxl tutorial](https://openpyxl.readthedocs.io/en/3.1/tutorial.html)

| Area | Required behavior / gotcha |
| --- | --- |
| Read-only lifecycle | Loading is lazy; explicitly close the workbook. Cells are ReadOnlyCell instances when values_only is false. |
| Dimensions | Worksheet dimensions can be wrong; expose reset_dimensions or document the difference. |
| Write-only lifecycle | No initial sheets; create_sheet is required. Only append is supported; save once, and reject subsequent append/save. |
| Writer metadata | Items preceding cell XML, such as freeze panes, must be finalized before the first append. |

These lifecycle requirements are documented in [openpyxl optimized modes](https://openpyxl.readthedocs.io/en/stable/optimized.html).

Iteration uses inclusive, one-based row/column bounds, normally starts at A1, and values_only yields tuples of values. [Worksheet implementation](https://openpyxl.readthedocs.io/en/latest/_modules/openpyxl/worksheet/worksheet.html)

Direct inspection of installed openpyxl 3.1.5 `worksheet/_read_only.py` and `worksheet/_reader.py` additionally establishes:

- Missing rows/cells are padded with None in values_only mode. Column width follows explicit max_col, dimension metadata, or the final cell of that row if dimensions were reset.
- Missing cell references infer successive columns; missing row references infer successive rows.
- A shared-string cell indexes the workbook string table. Boolean values use bool(int(text)); cached string cells remain strings.
- Normal formula values include the leading equals sign. Shared formulas translate the master formula to each follower coordinate; array and data-table formulas produce special objects.
- Numeric dates are recognized by style index and converted using the workbook epoch. Invalid date serials become a #VALUE! error with a warning.

The authoritative project source is [openpyxl on Heptapod](https://foss.heptapod.net/openpyxl/openpyxl). The local inspected files are under `/Library/Frameworks/Python.framework/Versions/3.14/lib/python3.14/site-packages/openpyxl/`.

data_only means return the cached value stored by Excel, instead of the formula. An absent cache remains None; this option does not promise formula calculation. [load_workbook source and documentation](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/reader/excel.html)

## Strings, dates and styles

openpyxl loads all shared strings into a list, concatenating rich text content in the ordinary reader. It also removes the OOXML `x005F_` escape marker. Streaming rows alone therefore do not imply constant total memory: the shared-string table can dominate. [String-table source](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/reader/strings.html)

Recommendation: use inline strings for streaming writes to avoid a workbook-wide deduplication map. For reads, start with a compact shared string table and clearly report its memory complexity; a disk-backed string index is a later optimization. Keep row batches bounded and never collect the worksheet XML or all rows.

Cell binding treats strings longer than one character starting with `=` as formulas, recognizes Excel error tokens, distinguishes booleans, truncates strings to 32,767 characters, and rejects forbidden XML control characters. A literal single `=` remains a string. [Cell source](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/cell/cell.html)

Dates can be ISO strings or serial numbers, are timezone-naive, and have millisecond precision. Workbooks use either the 1900 or 1904 epoch; the 1900 leap-year bug needs compatibility handling. Elapsed formats such as `[h]:mm:ss` map to timedelta. [Dates and times](https://openpyxl.readthedocs.io/en/stable/datetime.html)

Cell `s` is an index into cellXfs, not a number-format identifier. Map that XF's numFmtId to a built-in/custom format, then classify date/duration styles. [Stylesheet normalization](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/stylesheet.html)

Date-format recognition examines the first format section, strips quoted text and locale/color groups, and detects unescaped date/time symbols. A substring check for `m` or `d` would misclassify literal text. [Number-format source](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/numbers.html)

## Rust implementation choices

quick-xml offers a pull reader with reusable buffers. Avoid global trim_text(true): whitespace in cells is data, and trimming individual XML text events can also corrupt text split by comments/CDATA. [Parser configuration](https://docs.rs/quick-xml/0.42.0/quick_xml/reader/struct.Config.html)

Version 0.42 emits Text, CData and GeneralRef events separately. Resolve character references and the predefined XML entities and concatenate content; reading only Text events silently drops escaped characters. Borrowed event data cannot outlive reuse of the input buffer without ownership. [Event API](https://docs.rs/quick-xml/0.42.0/quick_xml/events/enum.Event.html)

The default reader is appropriate for UTF-8. Supporting non-ASCII-compatible encodings such as UTF-16 requires DecodingReader or an equivalent decoding layer; otherwise reject them explicitly. [quick-xml encoding documentation](https://docs.rs/quick-xml/0.42.0/quick_xml/)

zip's normal by_name/by_index return a ZipFile borrowing the archive. The stable API does not provide an owned member iterator constructor. Metadata can be shared across readers, but unsafe_new_with_metadata requires matching the exact underlying archive. [ZipArchive API](https://docs.rs/zip/latest/zip/read/struct.ZipArchive.html)

For an owned worksheet stream, one option is to read validated member metadata and reopen the XLSX, seek to data_start, limit input by compressed_size, and own either a stored reader or raw DEFLATE decoder. The API provides size, crc32, compression and encrypted status; data_start in zip 9 is Option<u64>. Reject encryption and unsupported compression. Verify actual decoded size and CRC at EOF. This architecture is an implementation recommendation, rather than a zip-provided reader. [ZipFile API](https://docs.rs/zip/latest/zip/read/struct.ZipFile.html)

Use independent File::open calls for simultaneously active streams: File::try_clone shares the file offset. [Rust File API](https://doc.rust-lang.org/std/fs/struct.File.html#method.try_clone)

Use DeflateDecoder rather than a zlib/gzip decoder for raw ZIP DEFLATE. Its bufread form exposes total_in/total_out and stops after one DEFLATE stream. Compare decoded size and checksum against ZIP metadata; compressed-input limits prevent accidental reads into the next member. [flate2 decoder](https://docs.rs/flate2/latest/flate2/bufread/struct.DeflateDecoder.html)

Ensure the XML reader reaches member EOF when naturally exhausted, so integrity checks run. Early stopping at max_row can intentionally avoid reading the rest; document that a partial iterator cannot validate the entire member.

ZipWriter streams member content through start_file and std::io::copy and needs a seekable sink. Call finish explicitly to observe finalization errors. Tempfiles per worksheet allow append-only multi-sheet writing without retaining cell rows in RAM. [ZipWriter API](https://docs.rs/zip/latest/zip/write/struct.ZipWriter.html)

PyO3 0.29 uses Bound/Py smart pointers, #[pyclass], #[pymethods] and #[pymodule]; use cdylib plus rlib for Python and Rust entry points. Maturin provides package builds, and abi3 can target the stable Python ABI. [PyO3 API and build example](https://docs.rs/pyo3/0.29.3/pyo3/)

Detach from Python during Rust-only decompression/parsing/writing, then attach for conversion to Python values. Bounded row batches amortize interpreter-boundary overhead. Keep Python objects out of detached Rust buffers. [PyO3 performance guide](https://github.com/PyO3/pyo3/blob/main/guide/src/performance.md)

## Attribution

openpyxl identifies its license as MIT/Expat. Its installed `openpyxl-3.1.5.dist-info/LICENCE.rst` names `Copyright (c) 2010 openpyxl` and requires preserving copyright and permission notices in copies or substantial portions. Include the original notice when translating code or substantial algorithms; retain Rust dependencies' license notices in release artifacts as appropriate. [Project license declaration](https://openpyxl.readthedocs.io/en/stable/)

## Validation priorities

Differential tests against the installed reference should cover sparse rows/columns, bounds, empty rows and cells, Unicode/whitespace/entities, shared and inline strings, boolean versus integer types, errors, cached and uncached formulas, date1904, millisecond rounding and duration formats. Include more than one simultaneously active sheet iterator and repeated iteration to expose shared-cursor/lifecycle bugs. Benchmark release builds in fresh processes using both elapsed time and peak RSS, with numeric, inline-string and high-cardinality shared-string inputs; stream a checksum instead of collecting rows.
