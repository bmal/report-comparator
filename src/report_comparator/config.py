from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


@dataclass(frozen=True)
class ConfigOption:
    """One configuration knob — the single source of truth for its name,
    default, and documentation. The example config, the README options table,
    and CONFIGURATION.md are all generated from the schema below, and a
    drift-guard test keeps them in sync (see tests/test_docs_drift.py)."""

    name: str
    default: Any
    summary: str  # one-line description (README table + YAML header comment)
    section: str  # grouping for the example YAML and CONFIGURATION.md
    detail: tuple[str, ...] = ()  # longer, wrapped prose for CONFIGURATION.md / YAML
    table_default: str | None = None  # README-table default-cell override
    item_comments: tuple[str, ...] = ()  # trailing comments for list items, by index
    example_lines: tuple[str, ...] = ()  # commented-out example lines after the value


# ── The schema: every tunable the comparator actually reads, in one place ─────
SCHEMA: tuple[ConfigOption, ...] = (
    ConfigOption(
        name="mode",
        default="strict",
        section="Mode & verbosity",
        summary="How MINOR findings are handled: strict surfaces them as warnings, lenient auto-accepts them.",
        detail=(
            "The only difference between the two strategies; OK and FAIL behave",
            "identically in both.",
            "  strict  — surface MINORs as warnings to eyeball (default; decks are",
            "            customer-facing).",
            "  lenient — auto-accept MINORs as 'accepted' audit entries with no",
            "            status escalation (use once the tool is trusted).",
        ),
    ),
    ConfigOption(
        name="quiet",
        default=False,
        section="Mode & verbosity",
        summary="Drop accepted/minor entries from the report entirely for terse output.",
        detail=(
            "Failures are always kept. Equivalent to the --quiet CLI flag.",
        ),
    ),
    ConfigOption(
        name="calibrate",
        default=False,
        section="Mode & verbosity",
        summary="Emit measured metrics for every compared element (including OK ones) to tune thresholds.",
        detail=(
            "Adds a 'measured_metrics' finding per element carrying the raw shift,",
            "resize, picture sizes, and changed-pixel percentage. Equivalent to the",
            "--calibrate CLI flag. See the calibration guidance below.",
        ),
    ),
    ConfigOption(
        name="filename_key_pattern",
        default=r"^(?P<key>.+)_\d{8}(_\d{6})?\.pptx$",
        section="Filename pairing",
        summary="Regex with a named `key` group used to pair old/new decks after stripping the trailing timestamp.",
        detail=(
            "A key must be unique within each directory and have exactly one",
            "counterpart in the other directory, or it becomes a file-level error",
            "(the tool never silently compares two unrelated decks).",
        ),
    ),
    ConfigOption(
        name="volatile_text_patterns",
        default=[
            r"\b\d{4}-\d{2}-\d{2}\b",
            r"\b\d{8}(_\d{6})?\b",
            r"\b[\w.-]+_\d{8}(_\d{6})?\.pptx\b",
        ],
        section="Free-text normalization",
        summary="Regexes stripped from free text (never table cells) before exact comparison.",
        detail=(
            "Keeps generation dates / source filenames from showing up as false",
            "wording changes. Never applied to table cells, which stay strictly",
            "compared.",
        ),
        table_default="3 date/filename patterns",
        item_comments=(
            "ISO date, e.g. 2026-06-21",
            "compact timestamp, e.g. 20260621_120000",
            "timestamped source filename",
        ),
    ),
    ConfigOption(
        name="shift_tolerance",
        default=0,
        section="Geometry tolerances (EMU; 914400 EMU = 1 inch)",
        summary="Max element move (EMU) still treated as OK; beyond it (on-slide, no new overlap) is MINOR.",
        detail=(
            "Measured as the largest of the left/top deltas. Off-slide overflow is",
            "always FAIL regardless of this tolerance.",
        ),
    ),
    ConfigOption(
        name="resize_tolerance",
        default=0,
        section="Geometry tolerances (EMU; 914400 EMU = 1 inch)",
        summary="Max element size change (EMU) still treated as OK; beyond it is MINOR.",
        detail=(
            "Measured as the largest of the width/height deltas.",
        ),
    ),
    ConfigOption(
        name="picture_pixel_threshold",
        default=0.0,
        section="Picture comparison",
        summary="Max percentage of changed pixels still treated as MINOR; above it is FAIL.",
        detail=(
            "Pixels are compared after decoding and stripping metadata, so harmless",
            "re-encoding noise does not register. 0.0 means any pixel change above",
            "zero is a FAIL.",
        ),
    ),
    ConfigOption(
        name="picture_dimension_tolerance",
        default=0,
        section="Picture comparison",
        summary="Allowed pixel difference in decoded width/height before a dimension FAIL.",
        detail=(
            "0 means any decoded dimension mismatch is a FAIL (unless",
            "picture_normalize_resize is on).",
        ),
    ),
    ConfigOption(
        name="picture_normalize_resize",
        default=False,
        section="Picture comparison",
        summary="If true, resize the candidate to the reference size and compare pixels instead of failing on a dimension mismatch.",
        detail=(
            "Use when charts legitimately wobble by a few pixels between framework",
            "versions.",
        ),
    ),
    ConfigOption(
        name="ignore_list",
        default=[],
        section="Ignore-list",
        summary="Rules that silence specific shapes by name, type, and/or region.",
        detail=(
            "Each rule may filter by `name`, `type`",
            "(picture/table/text/decoration/unknown), and/or `region`",
            "(left/top/right/bottom in EMU). A shape matching every field present in",
            "a rule is dropped before comparison. Grow this from --calibrate output",
            "when decoration proves noisy.",
        ),
        table_default="[] (empty)",
        example_lines=(
            "- name: CompanyLogo",
            "- type: decoration",
            "- region: { left: 0, top: 0, right: 914400, bottom: 914400 }",
        ),
    ),
)


DEFAULT_CONFIG: dict[str, Any] = {option.name: deepcopy(option.default) for option in SCHEMA}


def load_config(path: str | Path | None = None, overrides: dict[str, Any] | None = None) -> dict[str, Any]:
    """Build the effective config: schema defaults, then YAML file, then CLI overrides."""
    config = deepcopy(DEFAULT_CONFIG)
    if path:
        with Path(path).open("r", encoding="utf-8") as handle:
            loaded = yaml.safe_load(handle) or {}
        if not isinstance(loaded, dict):
            raise ValueError("config file must contain a YAML mapping")
        config.update(loaded)
    for key, value in (overrides or {}).items():
        if value is not None:
            config[key] = value
    return config
