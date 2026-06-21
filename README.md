# report-comparator

A Linux, rendering-free command-line tool that compares an `old/` directory of
reference PowerPoint decks against a `new/` directory of candidate decks and
produces a **single JSON report listing only what differs** — where (file →
slide → element) and how severely.

## What it does

A generation framework auto-produces ~50 `.pptx` decks (wafer overlay/focus
maps and MATLAB charts as raster images, data tables, and short text
descriptions) from fixed source data. When the framework is updated you must
regenerate every deck and check that the change didn't silently break or alter
anything — because these decks go to customers. Done by hand that means opening
old and new side by side and eyeballing every slide of ~50 decks.

This tool does that mechanical pass for you. You regenerate everything with the
new framework version, point it at the `old/` and `new/` directories, and it
emits one JSON report. Because both runs come from **identical source data**,
every reported difference is attributable to the framework change. You open the
report, ignore the OK decks, and manually check only the flagged slides.

The tool is deterministic and needs no LibreOffice/PowerPoint — it reads `.pptx`
files directly with `python-pptx`.

## Severity model

Every difference is classified into one of three buckets:

- **OK** — present, in place, data equal. Silent (emits no findings).
- **MINOR** — same elements present, but a slight shift or formatting drift;
  nothing overlaps, runs off-slide, or is missing.
- **FAIL** — a structural or content break: missing/blank/colliding picture,
  off-slide overflow, changed/empty/blank table, row/column count change, text
  wording change, stray element.

Two switchable strategies differ **only** in how MINOR is handled:

- **strict (default)** — surface MINORs as warnings to eyeball. The default
  because decks are customer-facing and nothing should be auto-accepted.
- **lenient** — auto-accept MINORs (recorded as `accepted` audit entries, no
  status escalation). Use once you trust the tool.

## Requirements

- **Python 3.11+**
- **Linux** (rendering-free; no LibreOffice/PowerPoint required)
- Mainstream libraries only: `python-pptx`, `Pillow`, `numpy`, `PyYAML`

## Install

```bash
pip install -e .          # or: pip install -e .[test] for the test extras
```

## Quickstart

```bash
report-comparator --old old/ --new new/ --mode strict --config my-config.yaml --out report.json
```

`--mode`, `--config`, and `--out` are optional — `--mode` defaults to `strict`,
config falls back to built-in defaults, and without `--out` the JSON is printed
to stdout. The minimal form is just `report-comparator --old old/ --new new/`.

### Reading the report

Running against three decks — one identical, one with a picture nudged beyond
tolerance, one with a changed table cell — produces:

```json
{
  "summary": {
    "files": 3,
    "ok": 1,
    "warnings": 1,
    "failures": 1
  },
  "files": [
    {
      "key": "focus_map",
      "old": "old/focus_map_20260101_120000.pptx",
      "new": "new/focus_map_20260115_120000.pptx",
      "status": "ok",
      "errors": [],
      "slides": []
    },
    {
      "key": "wafer_overlay",
      "old": "old/wafer_overlay_20260101_120000.pptx",
      "new": "new/wafer_overlay_20260115_120000.pptx",
      "status": "warning",
      "errors": [],
      "slides": [
        {
          "index": 1,
          "status": "warning",
          "findings": [
            {
              "severity": "minor",
              "type": "placement_changed",
              "element": "picture:OverlayMap",
              "message": "element moved/resized beyond tolerance (shift 274320, resize 0)"
            }
          ]
        }
      ]
    },
    {
      "key": "yield_table",
      "old": "old/yield_table_20260101_120000.pptx",
      "new": "new/yield_table_20260115_120000.pptx",
      "status": "fail",
      "errors": [],
      "slides": [
        {
          "index": 1,
          "status": "fail",
          "findings": [
            {
              "severity": "fail",
              "type": "table_cell_changed",
              "element": "table:YieldTable",
              "message": "table cell R2C2 changed"
            }
          ]
        }
      ]
    }
  ]
}
```

How to read it:

- **`summary`** is your triage line: 3 decks compared, 1 clean, 1 to eyeball, 1
  broken. Start here.
- Each entry in **`files`** is one old↔new deck pair, matched by `key` (the
  filename with its trailing timestamp stripped). `status` rolls up the worst
  finding in the deck (`ok` / `warning` / `fail`).
- **`focus_map`** is identical, so it has an empty `slides` list — OK slides
  emit no findings. You can ignore it.
- **`wafer_overlay`** has one `minor` finding: a picture moved beyond tolerance
  but stayed on-slide with no new overlap. In strict mode that makes the deck a
  `warning` — worth a glance, not a break.
- **`yield_table`** has one `fail` finding pointing at the exact cell
  (`R2C2`) that changed. Go straight to that slide.
- Deck-level problems (unreadable file, slide-count mismatch) appear in the
  file's **`errors`** list instead of under a slide.

`--quiet` drops the `accepted`/`minor` entries so only failures remain;
`--calibrate` adds a `measured_metrics` finding to every compared element.

## Modes & flags

| Flag | Meaning |
|------|---------|
| `--old DIR` | Directory of reference `.pptx` files (required for a comparison) |
| `--new DIR` | Directory of candidate `.pptx` files (required for a comparison) |
| `--mode strict` | **Default.** Surface MINORs as `minor` warnings to eyeball |
| `--mode lenient` | Auto-accept MINORs as `accepted` audit entries (no status escalation) |
| `--config FILE` | YAML config path (see [Configuration](#configuration)); CLI flags override it |
| `--out FILE` | Write the JSON report to this path instead of stdout |
| `--quiet` | Suppress `accepted`/`minor` entries (failures always kept) |
| `--calibrate` | Emit measured metrics for every compared element, including OK ones |
| `--print-config` | Print the effective config (defaults + overrides) as JSON and exit |
| `--version` | Print the tool version and exit |

## Configuration

All tunables live in a single YAML file. Start from
[`config.example.yaml`](config.example.yaml) — copy it, edit it, and pass it
with `--config my-config.yaml`. Every key is optional; anything you omit falls
back to the default shown below. CLI flags override the file.

For per-knob detail and calibration guidance, see
[CONFIGURATION.md](CONFIGURATION.md). To inspect the config a run will actually
use (defaults + your file + CLI overrides), run `report-comparator --print-config`.

<!-- BEGIN GENERATED CONFIG TABLE -->
| Option | Default | Description |
|--------|---------|-------------|
| `mode` | `strict` | How MINOR findings are handled: strict surfaces them as warnings, lenient auto-accepts them. |
| `quiet` | `false` | Drop accepted/minor entries from the report entirely for terse output. |
| `calibrate` | `false` | Emit measured metrics for every compared element (including OK ones) to tune thresholds. |
| `filename_key_pattern` | `^(?P<key>.+)_\d{8}(_\d{6})?\.pptx$` | Regex with a named `key` group used to pair old/new decks after stripping the trailing timestamp. |
| `volatile_text_patterns` | `3 date/filename patterns` | Regexes stripped from free text (never table cells) before exact comparison. |
| `shift_tolerance` | `0` | Max element move (EMU) still treated as OK; beyond it (on-slide, no new overlap) is MINOR. |
| `resize_tolerance` | `0` | Max element size change (EMU) still treated as OK; beyond it is MINOR. |
| `picture_pixel_threshold` | `0.0` | Max percentage of changed pixels still treated as MINOR; above it is FAIL. |
| `picture_dimension_tolerance` | `0` | Allowed pixel difference in decoded width/height before a dimension FAIL. |
| `picture_normalize_resize` | `false` | If true, resize the candidate to the reference size and compare pixels instead of failing on a dimension mismatch. |
| `ignore_list` | `[] (empty)` | Rules that silence specific shapes by name, type, and/or region. |
<!-- END GENERATED CONFIG TABLE -->

> This table is generated from the config schema in
> [`config.py`](src/report_comparator/config.py); regenerate the docs with
> `python -m report_comparator.docgen`.

### Calibration

The threshold defaults above (`shift_tolerance`, `picture_pixel_threshold`, …)
and the initial `ignore_list` are meant to be **tuned against a handful of real
old/new deck pairs**, not guessed. Run with `--calibrate` to dump every measured
metric — including for OK elements — then read the actual shift/resize/changed-
pixel numbers off real decks and set tolerances just above the noise floor. The
architecture doesn't depend on the numbers, only the defaults do. See
[CONFIGURATION.md](CONFIGURATION.md) for the full workflow.

## Report shape

```json
{
  "summary": { "files": 0, "ok": 0, "warnings": 0, "failures": 0 },
  "files": [
    { "key": "...", "old": "...", "new": "...", "status": "ok|warning|fail",
      "errors": [ { "severity": "fail", "type": "...", "message": "..." } ],
      "slides": [
        { "index": 1, "status": "ok|warning|fail",
          "findings": [
            { "severity": "minor|fail|accepted", "type": "...", "element": "...", "message": "..." }
          ] } ] } ]
}
```

OK slides emit no findings, so the report shows only what needs attention. There
is no exit-code gate — the tool is purely interactive; you read the JSON.

## Limitations (out of scope)

- **In-cell text clipping** (text squished/clipped *inside* a table cell) is not
  detectable without rendering; only a table frame that overflows off-slide is
  caught.
- **Swapped pictures** are reported as content-differs at both slots, not
  labeled as a "swap" (still flagged).
- **No version history / trend tracking** — strictly 2-way (old vs. new).
- **JSON only** — no HTML report, visual-diff artifacts, or heatmaps.
- **No CI / exit-code gating** — purely interactive.
- **Native charts, SmartArt, OLE, and media** are not deep-compared — they're
  surfaced as "couldn't compare (type: X)" notes, with presence/geometry rules
  still applied.
- **The generation framework is not modified** (e.g. to emit stable shape
  names); that's a recommended but separate effort.
- **Non-Linux packaging** is not provided.
