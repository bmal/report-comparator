from __future__ import annotations

import json
import subprocess
import sys

from report_comparator import compare_runs

from .fixture_builder import DeckBuilder, make_png


def save_pair(tmp_path, old_builder: DeckBuilder, new_builder: DeckBuilder, name: str = "deck"):
    old_dir = tmp_path / "old"
    new_dir = tmp_path / "new"
    old_dir.mkdir()
    new_dir.mkdir()
    old_builder.save(old_dir / f"{name}_20260101_120000.pptx")
    new_builder.save(new_dir / f"{name}_20260102_120000.pptx")
    return old_dir, new_dir


def test_identical_decks_emit_json_serializable_report_with_no_ok_slide_findings(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Title", "same")
    new.add_text("Title", "same")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    json.dumps(report)
    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []


def test_pairs_files_by_normalized_key_and_reports_unmatched_without_guessing(tmp_path):
    old_dir = tmp_path / "old"
    new_dir = tmp_path / "new"
    old_dir.mkdir()
    new_dir.mkdir()
    DeckBuilder().save(old_dir / "alpha_20260101.pptx")
    DeckBuilder().save(new_dir / "beta_20260101.pptx")

    report = compare_runs(old_dir, new_dir, {})

    messages = [error["message"] for file in report["files"] for error in file["errors"]]
    assert "unmatched old deck for key 'alpha'" in messages
    assert "unmatched new deck for key 'beta'" in messages
    assert report["summary"]["failures"] == 2


def test_reports_non_unique_normalized_keys(tmp_path):
    old_dir = tmp_path / "old"
    new_dir = tmp_path / "new"
    old_dir.mkdir()
    new_dir.mkdir()
    DeckBuilder().save(old_dir / "alpha_20260101.pptx")
    DeckBuilder().save(old_dir / "alpha_20260102.pptx")
    DeckBuilder().save(new_dir / "alpha_20260103.pptx")

    report = compare_runs(old_dir, new_dir, {})

    messages = [error["message"] for file in report["files"] for error in file["errors"]]
    assert any("non-unique deck key 'alpha'" in message for message in messages)


def test_slide_count_mismatch_is_file_level_fail_but_prefix_slides_still_compare(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("OnlySharedSlide", "same")
    new.add_text("OnlySharedSlide", "same")
    old.add_slide()
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    file = report["files"][0]
    assert file["status"] == "fail"
    assert file["errors"][0]["type"] == "slide_count_mismatch"
    assert file["slides"] == []


def test_missing_and_stray_elements_are_failures(tmp_path):
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("StableText")
    new.add_text("StableText")
    old.add_picture("MissingPicture", image_path)
    new.add_table("StrayTable")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    findings = compare_runs(old_dir, new_dir, {})["files"][0]["slides"][0]["findings"]

    assert {finding["type"] for finding in findings} == {"missing_element", "stray_element"}
    assert all(finding["severity"] == "fail" for finding in findings)


def test_missing_default_named_picture_in_grid_does_not_cascade_to_later_pictures(tmp_path):
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture(None, image_path, left=1, top=1)
    old.add_picture(None, image_path, left=2.5, top=1)
    old.add_picture(None, image_path, left=4, top=1)
    new.add_picture(None, image_path, left=1, top=1)
    new.add_picture(None, image_path, left=4, top=1)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    findings = compare_runs(old_dir, new_dir, {})["files"][0]["slides"][0]["findings"]

    assert findings == [
        {
            "severity": "fail",
            "type": "missing_element",
            "element": "picture:Picture 2",
            "message": "missing picture 'Picture 2'",
        }
    ]


def test_unique_non_default_names_match_even_when_positions_and_creation_order_change(tmp_path):
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("LeftMap", image_path, left=1, top=1)
    old.add_picture("RightMap", image_path, left=4, top=1)
    new.add_picture("RightMap", image_path, left=1, top=1)
    new.add_picture("LeftMap", image_path, left=4, top=1)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []


def test_unknown_shapes_are_reported_and_ignore_list_can_suppress_them(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_unknown("NativeChartPlaceholder")
    new.add_unknown("NativeChartPlaceholder")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})
    findings = report["files"][0]["slides"][0]["findings"]
    assert findings[0]["type"] == "unsupported_shape"
    assert "couldn't compare" in findings[0]["message"]

    ignored = compare_runs(old_dir, new_dir, {"ignore_list": [{"name": "NativeChartPlaceholder"}]})
    assert ignored["files"][0]["slides"] == []


def test_unreadable_deck_is_file_level_error_and_run_continues(tmp_path):
    old_dir = tmp_path / "old"
    new_dir = tmp_path / "new"
    old_dir.mkdir()
    new_dir.mkdir()
    (old_dir / "bad_20260101.pptx").write_text("not a pptx", encoding="utf-8")
    DeckBuilder().save(new_dir / "bad_20260101.pptx")
    DeckBuilder().save(old_dir / "good_20260101.pptx")
    DeckBuilder().save(new_dir / "good_20260101.pptx")

    report = compare_runs(old_dir, new_dir, {})

    assert report["summary"] == {"files": 2, "ok": 1, "warnings": 0, "failures": 1}
    assert any(file["errors"] and file["errors"][0]["type"] == "file_error" for file in report["files"])


def test_cli_loads_config_and_writes_report_with_cli_mode_override(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Title")
    new.add_text("Title")
    old_dir, new_dir = save_pair(tmp_path, old, new)
    config = tmp_path / "config.yaml"
    config.write_text("mode: lenient\n", encoding="utf-8")
    out = tmp_path / "report.json"

    subprocess.run(
        [
            sys.executable,
            "-m",
            "report_comparator.cli",
            "--old",
            str(old_dir),
            "--new",
            str(new_dir),
            "--config",
            str(config),
            "--mode",
            "strict",
            "--out",
            str(out),
        ],
        check=True,
    )

    assert json.loads(out.read_text(encoding="utf-8"))["summary"]["ok"] == 1
