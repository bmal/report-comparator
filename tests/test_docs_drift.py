"""Drift guard: the generated docs must always match the config schema.

If these fail, the schema in ``config.py`` changed but the docs weren't
regenerated. Fix with: ``python -m report_comparator.docgen``.
"""

from __future__ import annotations

import re
from pathlib import Path

from report_comparator.config import DEFAULT_CONFIG, SCHEMA, load_config
from report_comparator import docgen

ROOT = Path(__file__).resolve().parents[1]
SCHEMA_KEYS = {option.name for option in SCHEMA}


def _readme_table_keys() -> set[str]:
    block = docgen.extract_block(
        (ROOT / "README.md").read_text(encoding="utf-8"),
        docgen.README_BEGIN,
        docgen.README_END,
    )
    keys = set()
    for line in block.splitlines():
        match = re.match(r"\|\s*`([^`]+)`\s*\|", line)
        if match:
            keys.add(match.group(1))
    return keys


def _example_yaml_keys() -> set[str]:
    keys = set()
    for line in (ROOT / "config.example.yaml").read_text(encoding="utf-8").splitlines():
        match = re.match(r"^([A-Za-z_]\w*):", line)
        if match:
            keys.add(match.group(1))
    return keys


def test_schema_default_config_keys_match():
    assert SCHEMA_KEYS == set(DEFAULT_CONFIG)


def test_example_yaml_covers_exactly_the_schema_keys():
    # No missing, no extra — the example file documents every knob and only the knobs.
    assert _example_yaml_keys() == SCHEMA_KEYS


def test_readme_table_covers_exactly_the_schema_keys():
    assert _readme_table_keys() == SCHEMA_KEYS


def test_example_yaml_is_regenerated_from_the_schema():
    assert (ROOT / "config.example.yaml").read_text(encoding="utf-8") == docgen.render_example_yaml()


def test_readme_table_is_regenerated_from_the_schema():
    block = docgen.extract_block(
        (ROOT / "README.md").read_text(encoding="utf-8"),
        docgen.README_BEGIN,
        docgen.README_END,
    )
    assert block == docgen.render_readme_table()


def test_configuration_reference_is_regenerated_from_the_schema():
    block = docgen.extract_block(
        (ROOT / "CONFIGURATION.md").read_text(encoding="utf-8"),
        docgen.CONFIG_REF_BEGIN,
        docgen.CONFIG_REF_END,
    )
    assert block == docgen.render_config_reference()


def test_example_yaml_values_equal_the_defaults():
    # Loading the example (which sets every key to its default) must reproduce
    # DEFAULT_CONFIG exactly — guards both the keys and the rendered values.
    assert load_config(ROOT / "config.example.yaml") == DEFAULT_CONFIG
