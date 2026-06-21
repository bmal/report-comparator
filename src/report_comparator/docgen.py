"""Generate user-facing docs from the config schema (the single source of truth).

Three surfaces are derived from ``config.SCHEMA``:

* ``config.example.yaml``      — the file users copy and edit
* the README "Configuration" options table
* the per-knob reference block in ``CONFIGURATION.md``

Run ``python -m report_comparator.docgen`` to regenerate all three in place.
``tests/test_docs_drift.py`` re-runs these renderers and fails if the committed
files have drifted, so the docs can never silently fall out of sync with the
code.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .config import SCHEMA, ConfigOption

README_BEGIN = "<!-- BEGIN GENERATED CONFIG TABLE -->"
README_END = "<!-- END GENERATED CONFIG TABLE -->"
CONFIG_REF_BEGIN = "<!-- BEGIN GENERATED CONFIG REFERENCE -->"
CONFIG_REF_END = "<!-- END GENERATED CONFIG REFERENCE -->"

_HEADER_WIDTH = 78


def _yaml_scalar(value: Any) -> str:
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float):
        return repr(value)
    if isinstance(value, int):
        return str(value)
    return "'" + str(value).replace("'", "''") + "'"


def _section_header(section: str) -> str:
    prefix = f"# ── {section} "
    return prefix + "─" * max(0, _HEADER_WIDTH - len(prefix))


def render_example_yaml() -> str:
    """Render the full text of config.example.yaml from the schema."""
    lines: list[str] = [
        "# report-comparator configuration — every tunable, documented in one place.",
        "#",
        "# GENERATED from report_comparator/config.py (the schema is the source of",
        "# truth). Regenerate with `python -m report_comparator.docgen`.",
        "#",
        "# Copy this file, edit it, and pass it with `--config my-config.yaml`.",
        "# Every key is optional: anything you omit falls back to the default shown",
        "# here. CLI flags (--mode, --quiet, --calibrate) override the file.",
    ]

    current_section: str | None = None
    for option in SCHEMA:
        if option.section != current_section:
            current_section = option.section
            lines.append("")
            lines.append(_section_header(option.section))
        lines.append(f"# {option.summary}")
        for detail_line in option.detail:
            lines.append(f"#{'' if not detail_line else ' '}{detail_line}")
        lines.extend(_yaml_value_lines(option))
        for example in option.example_lines:
            lines.append(f"#   {example}")

    return "\n".join(lines) + "\n"


def _yaml_value_lines(option: ConfigOption) -> list[str]:
    value = option.default
    if isinstance(value, list):
        if not value:
            return [f"{option.name}: []"]
        rendered = [f"{option.name}:"]
        for index, item in enumerate(value):
            line = f"  - {_yaml_scalar(item)}"
            if index < len(option.item_comments) and option.item_comments[index]:
                line = f"{line:<38} # {option.item_comments[index]}"
            rendered.append(line)
        return rendered
    return [f"{option.name}: {_yaml_scalar(value)}"]


def _doc_default(option: ConfigOption) -> str:
    """Markdown-friendly rendering of a default (no YAML quoting for strings)."""
    if option.table_default is not None:
        return option.table_default
    value = option.default
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float):
        return repr(value)
    return str(value)


def render_readme_table() -> str:
    """Render the README "Configuration" options table from the schema."""
    lines = [
        "| Option | Default | Description |",
        "|--------|---------|-------------|",
    ]
    for option in SCHEMA:
        default = f"`{_doc_default(option)}`"
        lines.append(f"| `{option.name}` | {default} | {option.summary} |")
    return "\n".join(lines) + "\n"


def render_config_reference() -> str:
    """Render the per-knob reference block embedded in CONFIGURATION.md."""
    blocks: list[str] = []
    current_section: str | None = None
    for option in SCHEMA:
        if option.section != current_section:
            current_section = option.section
            blocks.append(f"### {option.section}")
        body = [
            f"#### `{option.name}`",
            "",
            f"_Default: `{_doc_default(option)}`_",
            "",
            option.summary,
        ]
        if option.detail:
            body.append("")
            body.append(" ".join(line.strip() for line in option.detail if line.strip()))
        blocks.append("\n".join(body))
    return "\n\n".join(blocks) + "\n"


def _splice(text: str, begin: str, end: str, block: str) -> str:
    start = text.index(begin) + len(begin)
    stop = text.index(end)
    return text[:start] + "\n" + block + text[stop:]


def extract_block(text: str, begin: str, end: str) -> str:
    start = text.index(begin) + len(begin)
    stop = text.index(end)
    return text[start:stop].strip("\n") + "\n"


def main() -> None:
    root = Path(__file__).resolve().parents[2]

    (root / "config.example.yaml").write_text(render_example_yaml(), encoding="utf-8")

    readme = root / "README.md"
    readme.write_text(
        _splice(readme.read_text(encoding="utf-8"), README_BEGIN, README_END, render_readme_table()),
        encoding="utf-8",
    )

    configuration = root / "CONFIGURATION.md"
    configuration.write_text(
        _splice(
            configuration.read_text(encoding="utf-8"),
            CONFIG_REF_BEGIN,
            CONFIG_REF_END,
            render_config_reference(),
        ),
        encoding="utf-8",
    )
    print("Regenerated config.example.yaml, README table, CONFIGURATION.md reference.")


if __name__ == "__main__":
    main()
