# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-06-21

### Added
- `--print-config` flag: prints the effective configuration (defaults + config
  file + CLI overrides) as JSON and exits.
- `--version` flag: prints the tool version (sourced from the package metadata).
- A single-source-of-truth config schema (`ConfigOption` / `SCHEMA` in
  `config.py`); `DEFAULT_CONFIG` is now derived from it.
- `report_comparator.docgen`: generates `config.example.yaml`, the README
  configuration table, and the `CONFIGURATION.md` reference from the schema.
- `CONFIGURATION.md` — per-knob reference plus a calibration workflow.
- This changelog.
- A drift-guard test (`tests/test_docs_drift.py`) asserting the generated docs
  cover exactly the schema's keys, with no missing or extra options.

### Changed
- README rewritten for operators: what the tool does, requirements, install,
  an annotated quickstart with real example output, modes & flags, and the
  out-of-scope limitations.
- `config.example.yaml` is now generated from the schema and no longer documents
  the `overlap.pic_pic_absolute` knob, which the comparator never read (picture↔
  picture overlap is an always-on invariant, documented in `CONFIGURATION.md`).

## [0.1.0] - 2026-06-21

### Added
- `compare_runs(old_dir, new_dir, config) -> Report`: compares a directory of
  reference `.pptx` decks against candidate decks and returns a JSON-serializable
  report of only what differs.
- Filename pairing by configurable timestamp-stripping regex, with hard errors on
  non-unique or unmatched keys.
- Positional slide alignment with a slide-count-mismatch presentation FAIL.
- Group-recursive shape enumeration and robust shape matching (stable name/alt-
  text when unique, otherwise position+content) with anti-cascade behavior.
- Picture content comparison by decoded pixels (metadata stripped) with
  configurable pixel threshold, dimension tolerance, and resize normalization.
- Geometry and overlap engine: shift/resize tolerances, slide-bounds overflow,
  reference-relative overlap, and absolute picture↔picture overlap.
- Strict exact table-cell comparison plus structure, empty-row, and blank-table
  checks; table and text formatting drift reported as MINOR.
- Free-text comparison with configurable volatile-token normalization.
- Per-file error isolation, configurable ignore-list, and unsupported-shape
  notes.
- `strict`/`lenient` modes, `--quiet`, and `--calibrate`.
- `report-comparator` CLI (`--old`, `--new`, `--mode`, `--config`, `--out`) and a
  pytest suite built on a `python-pptx` fixture builder.

[Unreleased]: https://github.com/bmal/report-comparator/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/bmal/report-comparator/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/bmal/report-comparator/releases/tag/v0.1.0
