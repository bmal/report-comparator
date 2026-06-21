# Configuration reference

Every tunable lives in a single YAML file. Copy
[`config.example.yaml`](config.example.yaml), edit it, and pass it with
`--config my-config.yaml`. Every key is optional — anything you omit falls back
to its default. CLI flags (`--mode`, `--quiet`, `--calibrate`) override the file.

The options below are the single source of truth: they're defined once in the
config schema in [`config.py`](src/report_comparator/config.py), and the example
file, the README table, and this reference are all generated from it (run
`python -m report_comparator.docgen` to regenerate). To see the effective config
a given invocation will use, run:

```bash
report-comparator --print-config --config my-config.yaml
```

## Options

<!-- BEGIN GENERATED CONFIG REFERENCE -->
### Mode & verbosity

#### `mode`

_Default: `strict`_

How MINOR findings are handled: strict surfaces them as warnings, lenient auto-accepts them.

The only difference between the two strategies; OK and FAIL behave identically in both. strict  — surface MINORs as warnings to eyeball (default; decks are customer-facing). lenient — auto-accept MINORs as 'accepted' audit entries with no status escalation (use once the tool is trusted).

#### `quiet`

_Default: `false`_

Drop accepted/minor entries from the report entirely for terse output.

Failures are always kept. Equivalent to the --quiet CLI flag.

#### `calibrate`

_Default: `false`_

Emit measured metrics for every compared element (including OK ones) to tune thresholds.

Adds a 'measured_metrics' finding per element carrying the raw shift, resize, picture sizes, and changed-pixel percentage. Equivalent to the --calibrate CLI flag. See the calibration guidance below.

### Filename pairing

#### `filename_key_pattern`

_Default: `^(?P<key>.+)_\d{8}(_\d{6})?\.pptx$`_

Regex with a named `key` group used to pair old/new decks after stripping the trailing timestamp.

A key must be unique within each directory and have exactly one counterpart in the other directory, or it becomes a file-level error (the tool never silently compares two unrelated decks).

### Free-text normalization

#### `volatile_text_patterns`

_Default: `3 date/filename patterns`_

Regexes stripped from free text (never table cells) before exact comparison.

Keeps generation dates / source filenames from showing up as false wording changes. Never applied to table cells, which stay strictly compared.

### Geometry tolerances (EMU; 914400 EMU = 1 inch)

#### `shift_tolerance`

_Default: `0`_

Max element move (EMU) still treated as OK; beyond it (on-slide, no new overlap) is MINOR.

Measured as the largest of the left/top deltas. Off-slide overflow is always FAIL regardless of this tolerance.

#### `resize_tolerance`

_Default: `0`_

Max element size change (EMU) still treated as OK; beyond it is MINOR.

Measured as the largest of the width/height deltas.

### Picture comparison

#### `picture_pixel_threshold`

_Default: `0.0`_

Max percentage of changed pixels still treated as MINOR; above it is FAIL.

Pixels are compared after decoding and stripping metadata, so harmless re-encoding noise does not register. 0.0 means any pixel change above zero is a FAIL.

#### `picture_dimension_tolerance`

_Default: `0`_

Allowed pixel difference in decoded width/height before a dimension FAIL.

0 means any decoded dimension mismatch is a FAIL (unless picture_normalize_resize is on).

#### `picture_normalize_resize`

_Default: `false`_

If true, resize the candidate to the reference size and compare pixels instead of failing on a dimension mismatch.

Use when charts legitimately wobble by a few pixels between framework versions.

### Ignore-list

#### `ignore_list`

_Default: `[] (empty)`_

Rules that silence specific shapes by name, type, and/or region.

Each rule may filter by `name`, `type` (picture/table/text/decoration/unknown), and/or `region` (left/top/right/bottom in EMU). A shape matching every field present in a rule is dropped before comparison. Grow this from --calibrate output when decoration proves noisy.
<!-- END GENERATED CONFIG REFERENCE -->

## A note on units

Geometry values (`shift_tolerance`, `resize_tolerance`, and the `region` bounds
in an `ignore_list` rule) are in **English Metric Units (EMU)**, the native
PowerPoint unit: **914400 EMU = 1 inch**, and 1 cm ≈ 360000 EMU. A tolerance of
`91440`, for example, allows roughly a tenth of an inch of movement.

## A fixed invariant: picture↔picture overlap

Overlap is judged **relative to the reference**: only an overlap that is new in
the candidate (and wasn't present in the old deck) fails, so intentional
template overlaps like a label over a chart don't spam the report. The **one
exception is picture↔picture overlap, which is always an absolute FAIL** —
colliding wafer maps are always caught, whether or not they already overlapped
in the reference. This rule is intentionally always on and is **not**
configurable. Off-slide overflow is likewise always a FAIL.

## Calibration workflow

The threshold defaults and the initial `ignore_list` are meant to be tuned
against a handful of **real** old/new deck pairs, not guessed. The architecture
does not depend on the numbers — only the defaults do — so calibration is purely
about choosing good values for your decks.

1. **Dump the measurements.** Run the tool with `--calibrate` against a few
   representative deck pairs:

   ```bash
   report-comparator --old old/ --new new/ --calibrate --out calibration.json
   ```

   In calibrate mode every compared element — including OK ones — gets a
   `measured_metrics` finding carrying the raw numbers:

   ```json
   {
     "severity": "ok",
     "type": "measured_metrics",
     "element": "picture:OverlayMap",
     "message": "calibration metrics",
     "metrics": {
       "kind": "picture",
       "shift_delta": 12700,
       "resize_delta": 0,
       "old_bounds": { "left": 914400, "top": 914400, "right": 1828800, "bottom": 1828800 },
       "new_bounds": { "left": 927100, "top": 914400, "right": 1841500, "bottom": 1828800 },
       "old_picture_size": [10, 10],
       "new_picture_size": [10, 10],
       "picture_changed_percent": 0.0
     }
   }
   ```

2. **Pick shift / resize tolerances.** Look at `shift_delta` and `resize_delta`
   across elements you consider visually unchanged. Set `shift_tolerance` and
   `resize_tolerance` **just above that noise floor** so imperceptible movement
   reads as OK while genuine repositioning still trips a MINOR. Remember the
   values are in EMU (914400 = 1 inch).

3. **Pick the picture pixel threshold.** Read `picture_changed_percent` for
   pictures you consider equivalent (re-encoding noise, anti-aliasing). Set
   `picture_pixel_threshold` a little above the largest such value so harmless
   noise is MINOR and real image regressions stay FAIL. If charts legitimately
   change size by a few pixels, either raise `picture_dimension_tolerance` or set
   `picture_normalize_resize: true` to resize-and-compare instead of failing on
   the dimension mismatch.

4. **Grow the ignore-list from noise.** If the same decoration (a logo, footer,
   or slide number) keeps producing findings you don't care about, add it to
   `ignore_list`. Use the `element` field (e.g. `decoration:CompanyLogo`) and the
   `*_bounds` from the calibration output to write a rule by `name`, `type`,
   and/or `region`:

   ```yaml
   ignore_list:
     - name: CompanyLogo
     - type: decoration
     - region: { left: 0, top: 0, right: 914400, bottom: 914400 }
   ```

   A shape matching **every** field present in a rule is dropped before
   comparison. Start narrow (by `name`) and widen only if needed.

5. **Re-run without `--calibrate`** and confirm the previously noisy elements are
   now OK and real regressions still FAIL. Iterate until the report shows only
   what genuinely needs attention.
