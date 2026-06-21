from __future__ import annotations

from report_comparator import compare_runs

from .fixture_builder import DeckBuilder, make_png
from .test_walking_skeleton import save_pair


def test_reencoded_picture_with_same_pixels_produces_no_findings(tmp_path):
    old_image = make_png(tmp_path / "old.png", metadata={"source": "old"})
    new_image = make_png(tmp_path / "new.png", metadata={"source": "new"})
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("WaferMap", old_image)
    new.add_picture("WaferMap", new_image)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []


def test_picture_pixel_difference_uses_configured_threshold(tmp_path):
    old_image = make_png(tmp_path / "old.png")
    new_image = make_png(tmp_path / "new.png", changed_pixels={(0, 0): (0, 0, 255)})
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("WaferMap", old_image)
    new.add_picture("WaferMap", new_image)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    minor_report = compare_runs(old_dir, new_dir, {"picture_pixel_threshold": 1.0})
    fail_report = compare_runs(old_dir, new_dir, {"picture_pixel_threshold": 0.5})

    minor_finding = minor_report["files"][0]["slides"][0]["findings"][0]
    fail_finding = fail_report["files"][0]["slides"][0]["findings"][0]
    assert minor_report["summary"] == {"files": 1, "ok": 0, "warnings": 1, "failures": 0}
    assert minor_finding["severity"] == "minor"
    assert minor_finding["type"] == "picture_pixels_changed"
    assert fail_report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert fail_finding["severity"] == "fail"


def test_picture_dimension_mismatch_fails_unless_resize_normalization_is_enabled(tmp_path):
    old_image = make_png(tmp_path / "old.png", size=(10, 10))
    new_image = make_png(tmp_path / "new.png", size=(12, 12))
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("WaferMap", old_image)
    new.add_picture("WaferMap", new_image)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    fail_report = compare_runs(old_dir, new_dir, {})
    normalized_report = compare_runs(old_dir, new_dir, {"picture_normalize_resize": True})

    finding = fail_report["files"][0]["slides"][0]["findings"][0]
    assert fail_report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["type"] == "picture_dimension_changed"
    assert finding["severity"] == "fail"
    assert normalized_report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert normalized_report["files"][0]["slides"] == []


def test_blank_picture_where_old_had_content_is_failure(tmp_path):
    old_image = make_png(tmp_path / "old.png", color=(255, 0, 0))
    new_image = make_png(tmp_path / "new.png", color=(255, 255, 255))
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("WaferMap", old_image)
    new.add_picture("WaferMap", new_image)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {"picture_pixel_threshold": 100.0})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["type"] == "picture_blank"
    assert finding["severity"] == "fail"
