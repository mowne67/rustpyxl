# Measured performance

Measured 10 October 2026 on macOS arm64, CPython 3.14.5, openpyxl 3.1.5, rustpyxl 0.1.0, Rust 1.99.0 release build. Each fixture contains 100,000 rows × 10 columns (one million cells). Values below are medians of three fresh-process samples.

## Drop-in compatibility bridge

This path constructs actual openpyxl workbooks, keeps reference behavior for unported features and uses native shared-string storage/value parsing. Worksheets have accurate dimensions. Compatible mutable editing/writing is delegated unchanged and has no claimed speed gain.

| Input | openpyxl read | Bridge read | Speedup | openpyxl peak RSS | Bridge peak RSS | RSS reduction |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| numeric | 2.220 s | 0.340 s | 6.52× | 46.2 MiB | 39.3 MiB | 14.9% |
| mixed | 3.543 s | 0.429 s | 8.26× | 46.2 MiB | 39.6 MiB | 14.1% |
| shared | 2.599 s | 0.300 s | 8.67× | 61.6 MiB | 54.6 MiB | 11.3% |

## Standalone Rust streaming API

This stricter API rejects unimplemented types/features and has no openpyxl runtime dependency. Its fixtures follow write-only output, without declared dimensions, so these numbers should not be compared directly with the bridge table. Both engines use streaming modes. Writes include generated rows, spooling, compression (openpyxl's default level 6; rustpyxl's default level 4, whose archives range from 6% smaller to 6% larger than openpyxl's on these fixtures) and completed archive publication. All written row SHA256 digests match the reference.

| Input | Operation | openpyxl | Rust (batches) | Speedup | openpyxl peak RSS | Rust peak RSS | RSS reduction |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| numeric | read | 3.342 s | 0.327 s | 10.21× | 46.1 MiB | 32.7 MiB | 29.1% |
| numeric | write | 5.069 s | 0.315 s | 16.12× | 36.9 MiB | 32.7 MiB | 11.5% |
| mixed | read | 4.921 s | 0.489 s | 10.06× | 47.1 MiB | 32.8 MiB | 30.5% |
| mixed | write | 5.678 s | 0.356 s | 15.96× | 37.1 MiB | 32.5 MiB | 12.4% |
| shared | read | 3.748 s | 0.312 s | 12.01× | 62.2 MiB | 46.6 MiB | 25.2% |
| shared | write | 5.147 s | 0.344 s | 14.98× | 37.0 MiB | 32.9 MiB | 11.1% |

## Methodology and limits

- Each timed sample runs in a separate interpreter. Import time is excluded from latency, while whole-process peak RSS includes imports, native/Python buffers and allocator overhead.
- Complete read consumption, Python object conversion and the same small checksum are included. Row/cell counts and checksums match. Filesystem cache is warm and engine order alternates.
- Numeric input has ten integers per row. Mixed input includes integers, floats, booleans, text, blanks, Unicode and formulas. Shared input includes 100,000 unique strings of roughly 90–100 characters, with numeric cells in the remaining columns.
- Standalone writes use inline strings, including outputs for the shared-input scenario. This avoids a global string dictionary; output ZIP sizes can differ between compression backends.
- Peak RSS is not a measure of library allocations alone. The bridge retains openpyxl imports and objects, which reduces its total RSS advantage compared with the standalone core.
- Read memory still scales with shared-string bytes/count and chosen batch size. Write memory scales with batch size, largest row and sheet buffers; disk usage scales with uncompressed worksheet XML.
- These synthetic results do not establish full Rust-only openpyxl parity or performance on style-heavy, formula-heavy or unsupported files. A late fallback reparses the earlier XML before continuing, adding work for that iteration.
- Only read-only values with known dimensions/explicit width are accelerated through the bridge. Cell objects, unknown dimensions, rich text and file-like sources use reference parsing; mutable and compatible write-only workbooks use reference writing.

## Reproduce

```sh
python benchmarks/benchmark.py --rows 100000 --repeats 3 --output docs/results.json
python benchmarks/compatibility.py --rows 100000 --repeats 3 --output docs/compat-results.json
```

Raw samples and environment metadata: [standalone results](results.json), [compatibility results](compat-results.json).

## Validation

48 Python differential/integration tests and 2 Rust integration tests pass. Rust clippy passes with warnings denied. A separately installed macOS arm64 wheel was round-tripped without openpyxl installed for the standalone engine. Compatibility tests exercise actual openpyxl classes, styles, dates, named ranges, merges, mutable save/edit, shared formulas, big integers, native/reference cell access, concurrent iterators, file-like input, reset dimensions, process-local installation and late batch fallback.
