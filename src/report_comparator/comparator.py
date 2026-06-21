from __future__ import annotations

import re
from collections import defaultdict
from itertools import permutations
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
    findings: list[dict[str, str]] = []
    matched, missing, stray = _match_elements(old_elements, new_elements)

    for element in missing:
        findings.append(_finding("fail", "missing_element", element["element"], f"missing {element['kind']} '{element['name']}'"))
    for element in stray:
        findings.append(_finding("fail", "stray_element", element["element"], f"stray {element['kind']} '{element['name']}'"))

    for old_element, _new_element in matched:
        if old_element["kind"] == "unknown":
            findings.append(
                _finding(
                    "minor",
                    "unsupported_shape",
                    old_element["element"],
                    f"couldn't compare (type: {old_element['shape_type']})",
                )
            )

    status = "fail" if any(f["severity"] == "fail" for f in findings) else "warning" if findings else "ok"
    return {"index": index, "status": status, "findings": findings}


def _match_elements(
    old_elements: list[dict[str, Any]], new_elements: list[dict[str, Any]]
) -> tuple[list[tuple[dict[str, Any], dict[str, Any]]], list[dict[str, Any]], list[dict[str, Any]]]:
    if _can_match_by_stable_id(old_elements) and _can_match_by_stable_id(new_elements):
        return _match_by_stable_id(old_elements, new_elements)
    return _match_by_kind_position_and_content(old_elements, new_elements)


def _can_match_by_stable_id(elements: list[dict[str, Any]]) -> bool:
    stable_ids = [element["stable_id"] for element in elements]
    return bool(stable_ids) and all(stable_ids) and len(stable_ids) == len(set(stable_ids))


def _match_by_stable_id(
    old_elements: list[dict[str, Any]], new_elements: list[dict[str, Any]]
) -> tuple[list[tuple[dict[str, Any], dict[str, Any]]], list[dict[str, Any]], list[dict[str, Any]]]:
    old_by_id = {element["stable_id"]: element for element in old_elements}
    new_by_id = {element["stable_id"]: element for element in new_elements}
    shared_ids = sorted(old_by_id.keys() & new_by_id.keys())
    matched = [(old_by_id[stable_id], new_by_id[stable_id]) for stable_id in shared_ids]
    missing = [old_by_id[stable_id] for stable_id in sorted(old_by_id.keys() - new_by_id.keys())]
    stray = [new_by_id[stable_id] for stable_id in sorted(new_by_id.keys() - old_by_id.keys())]
    return matched, missing, stray


def _match_by_kind_position_and_content(
    old_elements: list[dict[str, Any]], new_elements: list[dict[str, Any]]
) -> tuple[list[tuple[dict[str, Any], dict[str, Any]]], list[dict[str, Any]], list[dict[str, Any]]]:
    matched: list[tuple[dict[str, Any], dict[str, Any]]] = []
    missing: list[dict[str, Any]] = []
    stray: list[dict[str, Any]] = []
    kinds = sorted({element["kind"] for element in [*old_elements, *new_elements]})

    for kind in kinds:
        old_group = [element for element in old_elements if element["kind"] == kind]
        new_group = [element for element in new_elements if element["kind"] == kind]
        group_matches, group_missing, group_stray = _optimal_group_match(old_group, new_group)
        matched.extend(group_matches)
        missing.extend(group_missing)
        stray.extend(group_stray)

    return matched, missing, stray


def _optimal_group_match(
    old_group: list[dict[str, Any]], new_group: list[dict[str, Any]]
) -> tuple[list[tuple[dict[str, Any], dict[str, Any]]], list[dict[str, Any]], list[dict[str, Any]]]:
    if not old_group:
        return [], [], new_group
    if not new_group:
        return [], old_group, []

    match_count = min(len(old_group), len(new_group))
    best_pairs: list[tuple[int, int]] = []
    best_cost: int | None = None

    if len(old_group) <= len(new_group):
        old_indexes = tuple(range(len(old_group)))
        for new_indexes in permutations(range(len(new_group)), match_count):
            pairs = list(zip(old_indexes, new_indexes))
            cost = sum(_match_cost(old_group[old_index], new_group[new_index]) for old_index, new_index in pairs)
            if best_cost is None or cost < best_cost:
                best_cost = cost
                best_pairs = pairs
    else:
        new_indexes = tuple(range(len(new_group)))
        for old_indexes in permutations(range(len(old_group)), match_count):
            pairs = list(zip(old_indexes, new_indexes))
            cost = sum(_match_cost(old_group[old_index], new_group[new_index]) for old_index, new_index in pairs)
            if best_cost is None or cost < best_cost:
                best_cost = cost
                best_pairs = pairs

    matched_old = {old_index for old_index, _new_index in best_pairs}
    matched_new = {new_index for _old_index, new_index in best_pairs}
    matched = [(old_group[old_index], new_group[new_index]) for old_index, new_index in best_pairs]
    missing = [element for index, element in enumerate(old_group) if index not in matched_old]
    stray = [element for index, element in enumerate(new_group) if index not in matched_new]
    return matched, missing, stray


def _match_cost(old_element: dict[str, Any], new_element: dict[str, Any]) -> int:
    return _bounds_distance(old_element["bounds"], new_element["bounds"]) + _content_distance(old_element, new_element)


def _bounds_distance(old_bounds: dict[str, int], new_bounds: dict[str, int]) -> int:
    return sum(abs(old_bounds[key] - new_bounds[key]) for key in ("left", "top", "right", "bottom"))


def _content_distance(old_element: dict[str, Any], new_element: dict[str, Any]) -> int:
    return 0 if old_element["content_key"] == new_element["content_key"] else 10_000_000_000


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
    stable_id = _stable_identifier(shape, name)
    shape_type = str(shape.shape_type)
    element_id = f"{kind}:{name}"
    return {
        "identity": element_id,
        "element": element_id,
        "kind": kind,
        "name": name,
        "stable_id": stable_id,
        "shape_type": shape_type,
        "content_key": _content_key(shape, kind),
        "bounds": {
            "left": int(shape.left),
            "top": int(shape.top),
            "right": int(shape.left + shape.width),
            "bottom": int(shape.top + shape.height),
        },
    }


def _stable_identifier(shape: Any, name: str) -> str | None:
    alt_text = _alt_text(shape)
    if alt_text:
        return alt_text
    if name and name != "<unnamed>" and not _is_default_shape_name(name):
        return name
    return None


def _alt_text(shape: Any) -> str | None:
    try:
        c_nv_pr = shape._element.xpath(".//p:cNvPr")[0]
    except (AttributeError, IndexError):
        return None
    return c_nv_pr.get("descr") or c_nv_pr.get("title") or None


def _is_default_shape_name(name: str) -> bool:
    return bool(re.match(r"^(Picture|TextBox|Table|Rectangle|Freeform|Group|Chart|GraphicFrame) \d+$", name))


def _content_key(shape: Any, kind: str) -> str:
    if kind == "picture":
        return getattr(getattr(shape, "image", None), "sha1", "")
    if kind == "table":
        return "\n".join("\t".join(cell.text.strip() for cell in row.cells) for row in shape.table.rows)
    if kind == "text":
        return shape.text_frame.text.strip()
    return ""


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
