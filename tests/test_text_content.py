from __future__ import annotations

from report_comparator import compare_runs

from .fixture_builder import DeckBuilder
from .test_walking_skeleton import save_pair


def test_footer_differing_only_by_generation_date_is_ok(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Footer", "Generated on 2026-06-20")
    new.add_text("Footer", "Generated on 2026-06-21")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []


def test_genuine_text_wording_difference_is_failure(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Description", "Wafer map is centered")
    new.add_text("Description", "Wafer map is shifted")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["severity"] == "fail"
    assert finding["type"] == "text_changed"


def test_text_differing_only_by_timestamped_filename_is_ok(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Description", "Source file: old_result_20260101.pptx")
    new.add_text("Description", "Source file: new_result_20260102.pptx")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []


def test_table_cell_text_is_not_compared_as_free_text(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Description", "same")
    new.add_text("Description", "same")
    old.add_table("Data", "Generated on 2026-06-20")
    new.add_table("Data", "Generated on 2026-06-21")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    findings = [finding for slide in report["files"][0]["slides"] for finding in slide["findings"]]
    assert all(finding["type"] != "text_changed" for finding in findings)


def test_configured_volatile_text_pattern_is_normalized(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Description", "Calibration run ABC123 complete")
    new.add_text("Description", "Calibration run XYZ999 complete")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {"volatile_text_patterns": [r"\b[A-Z]{3}\d{3}\b"]})

    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []
