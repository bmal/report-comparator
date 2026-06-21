from __future__ import annotations

import argparse
import json
from pathlib import Path

from . import __version__
from .comparator import compare_runs
from .config import load_config


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="report-comparator",
        description="Compare a directory of reference .pptx decks against candidate decks "
        "and emit a JSON report of only what differs.",
    )
    parser.add_argument("--old", help="Directory containing reference .pptx files")
    parser.add_argument("--new", help="Directory containing candidate .pptx files")
    parser.add_argument("--mode", choices=["strict", "lenient"], help="Severity handling mode (overrides config)")
    parser.add_argument("--config", help="YAML config path")
    parser.add_argument("--out", help="Write JSON report to this path instead of stdout")
    parser.add_argument("--quiet", action="store_true", default=None, help="Suppress accepted/minor report entries")
    parser.add_argument(
        "--calibrate", action="store_true", default=None, help="Emit measured metrics for all compared elements"
    )
    parser.add_argument(
        "--print-config",
        action="store_true",
        help="Print the effective config (defaults + overrides) as JSON and exit",
    )
    parser.add_argument("--version", action="version", version=f"report-comparator {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    config = load_config(args.config, {"mode": args.mode, "quiet": args.quiet, "calibrate": args.calibrate})

    if args.print_config:
        print(json.dumps(config, indent=2))
        return 0

    if not args.old or not args.new:
        parser.error("--old and --new are required")

    report = compare_runs(args.old, args.new, config)
    rendered = json.dumps(report, indent=2)
    if args.out:
        Path(args.out).write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
