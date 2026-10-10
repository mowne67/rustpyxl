"""Benchmark the drop-in reader bridge, including reference object loading."""
import argparse
import json
import platform
import shutil
import statistics
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from benchmark import rows, shared_fixture


def dimensions(path, count):
    replacement = path.with_suffix(".dimension.xlsx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(replacement, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as dest:
        for name in source.namelist():
            if name != "xl/worksheets/sheet1.xml":
                dest.writestr(name, source.read(name))
                continue
            with source.open(name) as inp, dest.open(name, "w") as out:
                header = inp.read(65536)
                end = header.find(b"</sheetPr>")
                end = end + len(b"</sheetPr>") if end >= 0 else header.index(b">") + 1
                out.write(header[:end])
                out.write(f'<dimension ref="A1:J{count}"/>'.encode())
                out.write(header[end:])
                shutil.copyfileobj(inp, out, length=65536)
    replacement.replace(path)


def main(args):
    report = {"rows": args.rows, "columns": 10, "repeats": args.repeats,
              "python": sys.version, "platform": platform.platform(), "openpyxl": openpyxl.__version__,
              "methodology": "Fresh subprocess per read, imports excluded from time and included in process peak RSS. Full reference workbook construction plus native acceleration included. Explicit accurate worksheet dimensions, warm filesystem cache, alternating engine order, equal row/cell counts and checksum checked. Unique shared-string dataset holds one 90+ character string per row. Mutable/writing compatibility uses unmodified openpyxl and is not accelerated.",
              "results": []}
    with tempfile.TemporaryDirectory(prefix="rustpyxl-compat-") as directory:
        for dataset in ["numeric", "mixed", "shared"]:
            path = Path(directory) / f"{dataset}.xlsx"
            book = openpyxl.Workbook(write_only=True)
            sheet = book.create_sheet("Data")
            for row in rows(args.rows, dataset):
                sheet.append(row)
            book.save(path)
            if dataset == "shared":
                shared_fixture(path, args.rows)
            dimensions(path, args.rows)
            raw = {"openpyxl": [], "compat": []}
            for repeat in range(args.repeats):
                for engine in (["openpyxl", "compat"] if repeat % 2 == 0 else ["compat", "openpyxl"]):
                    command = [sys.executable, str(Path(__file__).with_name("benchmark.py")), "--worker", "--engine", engine,
                               "--operation", "read", "--dataset", dataset, "--rows", str(args.rows), "--file", str(path)]
                    sample = json.loads(subprocess.check_output(command, text=True))
                    assert sample["rows"] == args.rows and sample["cells"] == args.rows * 10
                    raw[engine].append(sample)
                    print(f"compat/{dataset}/{engine} #{repeat + 1}: {sample['seconds']:.3f}s, {sample['peak_rss_mib']:.1f} MiB", file=sys.stderr, flush=True)
            assert len({s["checksum"] for samples in raw.values() for s in samples}) == 1
            for engine, samples in raw.items():
                report["results"].append({"dataset": dataset, "engine": engine,
                                          "median_seconds": statistics.median(s["seconds"] for s in samples),
                                          "median_peak_rss_mib": statistics.median(s["peak_rss_mib"] for s in samples), "samples": samples})
    Path(args.output).write_text(json.dumps(report, indent=2) + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=100000)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--output", default="docs/compat-results.json")
    args = parser.parse_args()
    if args.rows < 1 or args.repeats < 1:
        parser.error("rows and repeats must be positive")
    main(args)
