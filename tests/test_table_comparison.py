from __future__ import annotations

from report_comparator import compare_runs

from .fixture_builder import DeckBuilder
from .test_walking_skeleton import save_pair


def test_reformatted_table_cell_is_failure(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_table("Data", values=[["1.23"]])
    new.add_table("Data", values=[["1.230"]])
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["severity"] == "fail"
    assert finding["type"] == "table_cell_changed"


def test_table_row_or_column_count_change_is_failure(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_table("Data", values=[["A", "B"]])
    new.add_table("Data", values=[["A"], ["B"]])
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["severity"] == "fail"
    assert finding["type"] == "table_structure_changed"


def test_empty_row_where_old_had_data_is_failure(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_table("Data", values=[["A"], ["B"]])
    new.add_table("Data", values=[["A"], [""]])
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["severity"] == "fail"
    assert finding["type"] == "table_empty_row"


def test_fully_blank_table_is_failure(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_table("Data", values=[["A", "B"]])
    new.add_table("Data", values=[["", ""]])
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["severity"] == "fail"
    assert finding["type"] == "table_blank"


def test_table_frame_overflowing_slide_is_failure(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_table("Data", values=[["A"]])
    new.add_table("Data", values=[["A"]], left=9.5, width=1)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["severity"] == "fail"
    assert finding["type"] == "off_slide"
