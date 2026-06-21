from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import yaml


DEFAULT_CONFIG: dict[str, Any] = {
    "mode": "strict",
    "quiet": False,
    "calibrate": False,
    "filename_key_pattern": r"^(?P<key>.+)_\d{8}(_\d{6})?\.pptx$",
    "ignore_list": [],
    "picture_pixel_threshold": 0.0,
    "picture_dimension_tolerance": 0,
    "picture_normalize_resize": False,
    "shift_tolerance": 0,
    "resize_tolerance": 0,
    "volatile_text_patterns": [
        r"\b\d{4}-\d{2}-\d{2}\b",
        r"\b\d{8}(_\d{6})?\b",
        r"\b[\w.-]+_\d{8}(_\d{6})?\.pptx\b",
    ],
}


def load_config(path: str | Path | None = None, overrides: dict[str, Any] | None = None) -> dict[str, Any]:
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
