# report-comparator

A Linux, rendering-free command-line tool that compares an `old/` directory of
reference PowerPoint decks against a `new/` directory of candidate decks
(generated from identical source data, different framework version) and produces
a single human-readable JSON report listing only what differs, where
(file → slide → element), and how severely.

Because both versions come from identical source data, every reported difference
is attributable to the framework change. The tool is deterministic and needs no
LibreOffice/PowerPoint — it reads `.pptx` files directly with `python-pptx`.

## Severity model

Each difference is classified into one of three buckets:

- **OK** — present, in place, data equal. Silent (emits no findings).
- **MINOR** — slight shift or formatting drift; nothing overlaps, runs off-slide, or is missing.
- **FAIL** — a structural or content break (missing/blank/colliding picture, off-slide overflow, changed/empty table, wording change, …).

Two switchable strategies differ **only** in how MINOR is handled:

- **strict (default)** — surface MINORs as warnings to eyeball.
- **lenient** — auto-accept MINORs, retaining them as an audit trail.

## Install

```bash
pip install -e .[test]   # editable install with the test extras
```

Dependencies are mainstream only: `python-pptx`, `Pillow`, `numpy`, `PyYAML`.

## Usage

```bash
report-comparator --old path/to/old --new path/to/new --out report.json
```

CLI flags:

| Flag | Meaning |
|------|---------|
| `--old DIR` | Directory of reference `.pptx` files (required) |
| `--new DIR` | Directory of candidate `.pptx` files (required) |
| `--mode {strict,lenient}` | Severity handling (overrides config) |
| `--config FILE` | YAML config path (see below) |
| `--out FILE` | Write JSON to this path instead of stdout |
| `--quiet` | Suppress accepted/minor entries |
| `--calibrate` | Emit measured metrics for every compared element |

CLI flags override values from `--config`.

## Configuration

All tunables live in a single documented YAML file. Start from
[`config.example.yaml`](config.example.yaml), which lists **every** option (mode,
filename pairing pattern, volatile-text patterns, geometry tolerances, picture
thresholds, overlap rules, ignore-list) with its default and an explanation.

`--calibrate` exists to support tuning those defaults against a handful of real
old/new deck pairs.

## Report shape

```json
{
  "summary": { "files": 0, "ok": 0, "warnings": 0, "failures": 0 },
  "files": [
    { "key": "...", "old": "...", "new": "...", "status": "ok|warning|fail",
      "errors": [],
      "slides": [
        { "index": 1, "status": "ok|warning|fail",
          "findings": [
            { "severity": "minor|fail|accepted", "type": "...", "element": "...", "message": "..." }
          ] } ] } ]
}
```

OK slides emit no findings, so the report shows only what needs attention.
