from __future__ import annotations

import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def compare_runs(old_dir: str | Path, new_dir: str | Path, config: dict[str, Any] | None = None) -> dict[str, Any]:
    config = config or {}
    old_dir = Path(old_dir)
    new_dir = Path(new_dir)
    files: list[dict[str, Any]] = []

    old_pairs, old_errors = _collect_pairs(old_dir, config)
    new_pairs, new_errors = _collect_pairs(new_dir, config)
    files.extend(old_errors)
    files.extend(new_errors)

    old_keys = set(old_pairs)
    new_keys = set(new_pairs)
    for key in sorted(old_keys - new_keys):
        files.append(_file_error(key, str(old_pairs[key]), None, f"unmatched old deck for key '{key}'"))
    for key in sorted(new_keys - old_keys):
        files.append(_file_error(key, None, str(new_pairs[key]), f"unmatched new deck for key '{key}'"))

    for key in sorted(old_keys & new_keys):
        files.append(_compare_decks(key, old_pairs[key], new_pairs[key], config))

    return _with_summary(files)


def _collect_pairs(directory: Path, config: dict[str, Any]) -> tuple[dict[str, Path], list[dict[str, Any]]]:
    pattern = re.compile(config.get("filename_key_pattern", r"^(?P<key>.+)_\d{8}(_\d{6})?\.pptx$"))
    grouped: dict[str, list[Path]] = defaultdict(list)
    errors: list[dict[str, Any]] = []

    for path in sorted(directory.glob("*.pptx")):
        match = pattern.match(path.name)
        key = match.group("key") if match else path.stem
        grouped[key].append(path)

    pairs: dict[str, Path] = {}
    for key, paths in grouped.items():
        if len(paths) == 1:
            pairs[key] = paths[0]
        else:
            errors.append(_file_error(key, None, None, f"non-unique deck key '{key}' in {directory}"))
    return pairs, errors


def _compare_decks(key: str, old_path: Path, new_path: Path, config: dict[str, Any]) -> dict[str, Any]:
    file_result = {
        "key": key,
        "old": str(old_path),
        "new": str(new_path),
        "status": "ok",
        "errors": [],
        "slides": [],
    }
    try:
        old_deck = Presentation(str(old_path))
        new_deck = Presentation(str(new_path))
    except Exception as exc:  # python-pptx exposes several package/read errors.
        file_result["status"] = "fail"
        file_result["errors"].append({"severity": "fail", "type": "file_error", "message": f"could not read deck: {exc}"})
        return file_result

    old_count = len(old_deck.slides)
    new_count = len(new_deck.slides)
    if old_count != new_count:
        file_result["status"] = "fail"
        file_result["errors"].append(
            {
                "severity": "fail",
                "type": "slide_count_mismatch",
                "message": f"slide count changed from {old_count} to {new_count}",
            }
        )

    for index in range(min(old_count, new_count)):
        slide = _compare_slides(index + 1, old_deck.slides[index], new_deck.slides[index], config)
        if slide["findings"]:
            file_result["slides"].append(slide)

    if any(f["severity"] == "fail" for slide in file_result["slides"] for f in slide["findings"]):
        file_result["status"] = "fail"
    elif any(f["severity"] == "minor" for slide in file_result["slides"] for f in slide["findings"]):
        file_result["status"] = "warning"
    return file_result


def _compare_slides(index: int, old_slide: Any, new_slide: Any, config: dict[str, Any]) -> dict[str, Any]:
    old_elements = _visible_elements(old_slide.shapes, config)
    new_elements = _visible_elements(new_slide.shapes, config)
    old_counts = Counter(element["identity"] for element in old_elements)
    new_counts = Counter(element["identity"] for element in new_elements)
    by_identity = {element["identity"]: element for element in [*old_elements, *new_elements]}
    findings: list[dict[str, str]] = []

    for identity in sorted(old_counts.keys() - new_counts.keys()):
        element = by_identity[identity]
        findings.append(_finding("fail", "missing_element", element["element"], f"missing {element['kind']} '{element['name']}'"))
    for identity in sorted(new_counts.keys() - old_counts.keys()):
        element = by_identity[identity]
        findings.append(_finding("fail", "stray_element", element["element"], f"stray {element['kind']} '{element['name']}'"))
    for identity in sorted(old_counts.keys() & new_counts.keys()):
        if old_counts[identity] != new_counts[identity]:
            element = by_identity[identity]
            findings.append(_finding("fail", "element_count_mismatch", element["element"], f"element count changed for {element['kind']} '{element['name']}'"))

    for element in old_elements:
        if element["kind"] == "unknown" and old_counts[element["identity"]] and new_counts[element["identity"]]:
            findings.append(
                _finding(
                    "minor",
                    "unsupported_shape",
                    element["element"],
                    f"couldn't compare (type: {element['shape_type']})",
                )
            )

    status = "fail" if any(f["severity"] == "fail" for f in findings) else "warning" if findings else "ok"
    return {"index": index, "status": status, "findings": findings}


def _visible_elements(shapes: Iterable[Any], config: dict[str, Any]) -> list[dict[str, Any]]:
    elements: list[dict[str, Any]] = []
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            elements.extend(_visible_elements(shape.shapes, config))
            continue
        element = _element(shape)
        if not _ignored(element, config):
            elements.append(element)
    return elements


def _element(shape: Any) -> dict[str, Any]:
    kind = _kind(shape)
    name = getattr(shape, "name", "") or "<unnamed>"
    shape_type = str(shape.shape_type)
    element_id = f"{kind}:{name}"
    return {
        "identity": element_id,
        "element": element_id,
        "kind": kind,
        "name": name,
        "shape_type": shape_type,
        "bounds": {
            "left": int(shape.left),
            "top": int(shape.top),
            "right": int(shape.left + shape.width),
            "bottom": int(shape.top + shape.height),
        },
    }


def _kind(shape: Any) -> str:
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "picture"
    if getattr(shape, "has_table", False):
        return "table"
    if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        return "text"
    if _looks_like_decoration(shape):
        return "decoration"
    return "unknown"


def _looks_like_decoration(shape: Any) -> bool:
    name = (getattr(shape, "name", "") or "").lower()
    return any(token in name for token in ("logo", "footer", "slide number"))


def _ignored(element: dict[str, Any], config: dict[str, Any]) -> bool:
    for rule in config.get("ignore_list", []):
        if not isinstance(rule, dict):
            continue
        if rule.get("name") and rule["name"] != element["name"]:
            continue
        if rule.get("type") and rule["type"] != element["kind"]:
            continue
        if rule.get("region") and not _inside_region(element["bounds"], rule["region"]):
            continue
        return True
    return False


def _inside_region(bounds: dict[str, int], region: dict[str, int]) -> bool:
    return (
        bounds["left"] >= region.get("left", bounds["left"])
        and bounds["top"] >= region.get("top", bounds["top"])
        and bounds["right"] <= region.get("right", bounds["right"])
        and bounds["bottom"] <= region.get("bottom", bounds["bottom"])
    )


def _finding(severity: str, finding_type: str, element: str, message: str) -> dict[str, str]:
    return {"severity": severity, "type": finding_type, "element": element, "message": message}


def _file_error(key: str, old: str | None, new: str | None, message: str) -> dict[str, Any]:
    return {
        "key": key,
        "old": old,
        "new": new,
        "status": "fail",
        "errors": [{"severity": "fail", "type": "file_error", "message": message}],
        "slides": [],
    }


def _with_summary(files: list[dict[str, Any]]) -> dict[str, Any]:
    summary = {"files": len(files), "ok": 0, "warnings": 0, "failures": 0}
    for file_result in files:
        if file_result["status"] == "fail":
            summary["failures"] += 1
        elif file_result["status"] == "warning":
            summary["warnings"] += 1
        else:
            summary["ok"] += 1
    return {"summary": summary, "files": files}
