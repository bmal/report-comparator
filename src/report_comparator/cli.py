from __future__ import annotations

import argparse
import json
from pathlib import Path

from .comparator import compare_runs
from .config import load_config


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="report-comparator")
    parser.add_argument("--old", required=True, help="Directory containing reference .pptx files")
    parser.add_argument("--new", required=True, help="Directory containing candidate .pptx files")
    parser.add_argument("--mode", choices=["strict", "lenient"], help="Severity handling mode")
    parser.add_argument("--config", help="YAML config path")
    parser.add_argument("--out", help="Write JSON report to this path instead of stdout")
    parser.add_argument("--quiet", action="store_true", default=None, help="Suppress accepted/minor report entries")
    parser.add_argument("--calibrate", action="store_true", default=None, help="Emit measured metrics for all compared elements")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = load_config(args.config, {"mode": args.mode, "quiet": args.quiet, "calibrate": args.calibrate})
    report = compare_runs(args.old, args.new, config)
    rendered = json.dumps(report, indent=2)
    if args.out:
        Path(args.out).write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
