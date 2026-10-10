"""Fresh-process XLSX benchmarks with end-to-end latency and peak process RSS.

Run: python benchmarks/benchmark.py --rows 100000 --repeats 3 --output results.json
Fixture creation and correctness verification occur outside timed workers.
"""
import argparse
import gc
import hashlib
import json
import platform
import resource
import statistics
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


def rows(count, dataset):
    for i in range(count):
        if dataset == "numeric":
            yield tuple(i * 10 + j for j in range(10))
        elif dataset == "shared":
            yield (i, f"unique-{i}-" + "shared-string-payload-" * 4, *range(i + 2, i + 10))
        else:
            yield (i, i * 0.25, i % 2 == 0, f"item-{i}", " repeated & <text> ",
                   None, i + 7, "தமிழ்", -i, "=A1+1")


def worker(args):
    if args.engine == "rustpyxl":
        import rustpyxl as engine
    elif args.engine == "compat":
        import rustpyxl.compat as engine
    else:
        import openpyxl as engine
    gc.collect()
    started = time.perf_counter()
    if args.operation == "write":
        wb = engine.Workbook(write_only=True)
        ws = wb.create_sheet("Data")
        if args.engine == "rustpyxl" and args.mode == "batch":
            ws.append_rows(rows(args.rows, args.dataset), batch_size=args.batch_size)
        else:
            for row in rows(args.rows, args.dataset):
                ws.append(row)
        wb.save(args.file)
        count, cells, checksum = args.rows, args.rows * 10, None
    else:
        wb = engine.load_workbook(args.file, read_only=True, data_only=False)
        count = cells = checksum = 0
        kw = {"batch_size": args.batch_size} if args.engine == "rustpyxl" else {}
        for row in wb.active.iter_rows(values_only=True, **kw):
            count += 1
            cells += len(row)
            # Same small computation for both engines; includes Python conversion.
            checksum += sum(v if isinstance(v, int) else len(v) if isinstance(v, str) else 0 for v in row)
        wb.close()
    elapsed = time.perf_counter() - started
    rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    rss_mib = rss / (1024 ** 2 if sys.platform == "darwin" else 1024)
    print(json.dumps({"seconds": elapsed, "peak_rss_mib": rss_mib, "rows": count, "cells": cells, "checksum": checksum, "file_bytes": Path(args.file).stat().st_size}))


def verify(path):
    import openpyxl
    digest = hashlib.sha256()
    count = 0
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        for row in wb.active.iter_rows(values_only=True):
            digest.update(json.dumps(row, ensure_ascii=False, separators=(",", ":")).encode())
            count += 1
    finally:
        wb.close()
    return count, digest.hexdigest()


def shared_fixture(path, count):
    """Create a high-cardinality shared-string package without retaining rows."""
    replacement = path.with_suffix(".shared.xlsx")
    reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(replacement, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as dest:
        for name in source.namelist():
            if name == "xl/worksheets/sheet1.xml":
                continue
            data = source.read(name)
            if name == "xl/_rels/workbook.xml.rels":
                data = data.replace(b"</Relationships>", f'<Relationship Id="sst" Type="{reltype}" Target="sharedStrings.xml"/></Relationships>'.encode())
            if name == "[Content_Types].xml":
                data = data.replace(b"</Types>", b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
            dest.writestr(name, data)
        ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        with dest.open("xl/sharedStrings.xml", "w") as out:
            out.write(f'<sst {ns} count="{count}" uniqueCount="{count}">'.encode())
            for row in rows(count, "shared"):
                out.write(f'<si><t>{escape(row[1])}</t></si>'.encode())
            out.write(b"</sst>")
        with dest.open("xl/worksheets/sheet1.xml", "w") as out:
            out.write(f'<worksheet {ns}><sheetData>'.encode())
            for i, row in enumerate(rows(count, "shared"), 1):
                out.write(f'<row r="{i}">'.encode())
                for j, value in enumerate(row):
                    coord = f"{chr(65 + j)}{i}"
                    value_xml = f'<c r="{coord}" t="s"><v>{i - 1}</v></c>' if j == 1 else f'<c r="{coord}"><v>{value}</v></c>'
                    out.write(value_xml.encode())
                out.write(b"</row>")
            out.write(b"</sheetData></worksheet>")
    replacement.replace(path)


def run(args):
    import openpyxl
    import rustpyxl
    results = []
    with tempfile.TemporaryDirectory(prefix="rustpyxl-bench-") as directory:
        directory = Path(directory)
        for dataset in ["numeric", "mixed", "shared"]:
            source = directory / f"{dataset}.xlsx"
            wb = openpyxl.Workbook(write_only=True)
            ws = wb.create_sheet("Data")
            for row in rows(args.rows, dataset):
                ws.append(row)
            wb.save(source)
            if dataset == "shared":
                shared_fixture(source, args.rows)
            expected_digest = verify(source)
            configurations = [("read", "openpyxl", "row"), ("read", "rustpyxl", "batch"),
                              ("write", "openpyxl", "row"), ("write", "rustpyxl", "row"),
                              ("write", "rustpyxl", "batch")]
            raw = {config: [] for config in configurations}
            # Alternate order to reduce systematic cache/order advantages.
            for repeat in range(args.repeats):
                for operation, engine, mode in configurations[::1 if repeat % 2 == 0 else -1]:
                    path = source if operation == "read" else directory / f"{dataset}-{engine}-{mode}.xlsx"
                    command = [sys.executable, str(Path(__file__).resolve()), "--worker", "--engine", engine,
                               "--operation", operation, "--mode", mode, "--dataset", dataset,
                               "--rows", str(args.rows), "--file", str(path), "--batch-size", str(args.batch_size)]
                    result = json.loads(subprocess.check_output(command, text=True))
                    raw[(operation, engine, mode)].append(result)
                    print(f"{dataset} {operation} {engine}/{mode} #{repeat + 1}: {result['seconds']:.3f}s, {result['peak_rss_mib']:.1f} MiB", file=sys.stderr, flush=True)
            read_checksums = [sample["checksum"] for key, samples in raw.items() if key[0] == "read" for sample in samples]
            if len(set(read_checksums)) != 1:
                raise AssertionError("read checksums differ")
            for (operation, engine, mode), samples in raw.items():
                if operation == "write":
                    if verify(directory / f"{dataset}-{engine}-{mode}.xlsx") != expected_digest:
                        raise AssertionError(f"output correctness mismatch: {dataset}/{engine}/{mode}")
                if any(sample["rows"] != args.rows or sample["cells"] != args.rows * 10 for sample in samples):
                    raise AssertionError("row/cell counts differ")
                results.append({"dataset": dataset, "operation": operation, "engine": engine, "mode": mode,
                                "median_seconds": statistics.median(s["seconds"] for s in samples),
                                "median_peak_rss_mib": statistics.median(s["peak_rss_mib"] for s in samples), "samples": samples})
    report = {"environment": {"platform": platform.platform(), "machine": platform.machine(),
                              "python": sys.version, "openpyxl": openpyxl.__version__, "rustpyxl": rustpyxl.__version__},
              "rows": args.rows, "columns": 10, "repeats": args.repeats, "batch_size": args.batch_size,
              "methodology": "Fresh subprocess per sample. Imports excluded from elapsed time, included in peak RSS. Streaming baselines; openpyxl at its default compression level 6, rustpyxl at its default level 4; complete read conversion and checksum; write generation, disk spooling, compression, save included. Warm OS cache. Order alternated. Exact SHA256 of rows from written workbooks verified outside timing. Numeric and mixed fixtures use inline strings; shared input has one unique 90+ character shared string per row. All output writers use inline strings.",
              "results": results}
    Path(args.output).write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--worker", action="store_true")
    p.add_argument("--rows", type=int, default=100000)
    p.add_argument("--repeats", type=int, default=3)
    p.add_argument("--output", default="results.json")
    p.add_argument("--batch-size", type=int, default=1024)
    p.add_argument("--engine", choices=["openpyxl", "rustpyxl", "compat"])
    p.add_argument("--operation", choices=["read", "write"])
    p.add_argument("--mode", choices=["row", "batch"], default="batch")
    p.add_argument("--dataset", choices=["numeric", "mixed", "shared"])
    p.add_argument("--file")
    args = p.parse_args()
    if args.rows < 1 or args.repeats < 1:
        p.error("rows and repeats must be positive")
    worker(args) if args.worker else run(args)
