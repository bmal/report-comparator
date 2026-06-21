from __future__ import annotations

from pathlib import Path

import yaml

from report_comparator import compare_runs
from report_comparator.config import DEFAULT_CONFIG, load_config

from .fixture_builder import DeckBuilder, make_png
from .test_walking_skeleton import save_pair


def test_both_old_and_new_pictures_blank_is_failure(tmp_path):
    # PRD severity matrix: a blank candidate picture is FAIL even if the reference
    # was already blank (a blank wafer map is never acceptable output).
    blank = make_png(tmp_path / "blank.png", color=(255, 255, 255))
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("WaferMap", blank)
    new.add_picture("WaferMap", blank)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {"picture_pixel_threshold": 100.0})

    finding = report["files"][0]["slides"][0]["findings"][0]
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 0, "failures": 1}
    assert finding["type"] == "picture_blank"
    assert finding["severity"] == "fail"


def test_vanished_element_replaced_by_unrelated_one_elsewhere_is_missing_plus_stray(tmp_path):
    # Fallback matcher: a removed picture and an unrelated new picture at a clearly
    # different position must be reported as missing + stray, not forced together
    # and mislabeled as a content change.
    old_image = make_png(tmp_path / "old.png", color=(255, 0, 0))
    new_image = make_png(tmp_path / "new.png", color=(0, 0, 255))
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture(None, old_image, left=1, top=1)
    new.add_picture(None, new_image, left=7, top=5)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    findings = compare_runs(old_dir, new_dir, {})["files"][0]["slides"][0]["findings"]

    assert {finding["type"] for finding in findings} == {"missing_element", "stray_element"}
    assert all(finding["severity"] == "fail" for finding in findings)


def test_many_same_kind_pictures_match_without_factorial_blowup(tmp_path):
    # 12 identically-content, default-named pictures would be ~12! permutations under
    # a factorial matcher; the bounded matcher resolves them as an exact 1:1 match.
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    positions = [(left, top) for top in (0.5, 2.0, 3.5) for left in (0.5, 2.0, 3.5, 5.0)]
    for left, top in positions:
        old.add_picture(None, image_path, left=left, top=top)
        new.add_picture(None, image_path, left=left, top=top)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    assert report["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    assert report["files"][0]["slides"] == []


def test_shapes_inside_a_group_are_recursed_into_and_compared(tmp_path):
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_group("GroupLabel", "Wafer map is centered", "GroupPicture", image_path)
    new.add_group("GroupLabel", "Wafer map is shifted", "GroupPicture", image_path)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    findings = compare_runs(old_dir, new_dir, {})["files"][0]["slides"][0]["findings"]

    types = {finding["type"] for finding in findings}
    # The grouped text change is caught (recursion works) and the grouped picture is
    # matched rather than reported as a missing/stray shape.
    assert "text_changed" in types
    assert "missing_element" not in types
    assert "stray_element" not in types


def test_pictures_match_by_alt_text_when_names_are_default(tmp_path):
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture(None, image_path, left=1, top=1, alt_text="left-wafer")
    old.add_picture(None, image_path, left=4, top=1, alt_text="right-wafer")
    new.add_picture(None, image_path, left=4, top=1, alt_text="left-wafer")
    new.add_picture(None, image_path, left=1, top=1, alt_text="right-wafer")
    old_dir, new_dir = save_pair(tmp_path, old, new)

    report = compare_runs(old_dir, new_dir, {})

    findings = report["files"][0]["slides"][0]["findings"]
    # Matched by alt-text despite default names and swapped positions: only placement
    # drift, no missing/stray/content cascade.
    assert report["summary"] == {"files": 1, "ok": 0, "warnings": 1, "failures": 0}
    assert {finding["type"] for finding in findings} == {"placement_changed"}


def test_custom_filename_regex_with_named_key_group_pairs_decks(tmp_path):
    old_dir = tmp_path / "old"
    new_dir = tmp_path / "new"
    old_dir.mkdir()
    new_dir.mkdir()
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("Title", "same")
    new.add_text("Title", "same")
    old.save(old_dir / "wafer_run-ALPHA-2026.pptx")
    new.save(new_dir / "wafer_run-BETA-2026.pptx")
    pattern = r"^(?P<key>.+?)_run-[A-Z]+-\d{4}\.pptx$"

    paired = compare_runs(old_dir, new_dir, {"filename_key_pattern": pattern})
    default = compare_runs(old_dir, new_dir, {})

    assert paired["summary"] == {"files": 1, "ok": 1, "warnings": 0, "failures": 0}
    # The default timestamp pattern cannot key these names, so they stay unmatched.
    assert default["summary"]["failures"] == 2


def test_inserted_colliding_picture_produces_overlap_finding_in_addition_to_stray(tmp_path):
    image_path = make_png(tmp_path / "pic.png")
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_picture("MapA", image_path, left=1, top=3)
    new.add_picture("MapA", image_path, left=1, top=3)
    new.add_picture("MapB", image_path, left=1, top=3)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    findings = compare_runs(old_dir, new_dir, {})["files"][0]["slides"][0]["findings"]

    types = {finding["type"] for finding in findings}
    assert "stray_element" in types
    assert "picture_overlap" in types
    assert all(finding["severity"] == "fail" for finding in findings if finding["type"] == "picture_overlap")


def test_inserted_overlapping_text_is_a_new_overlap(tmp_path):
    old = DeckBuilder()
    new = DeckBuilder()
    old.add_text("LabelA", "same", left=1, top=1, width=2)
    new.add_text("LabelA", "same", left=1, top=1, width=2)
    new.add_text("LabelB", "other", left=1, top=1, width=2)
    old_dir, new_dir = save_pair(tmp_path, old, new)

    findings = compare_runs(old_dir, new_dir, {})["files"][0]["slides"][0]["findings"]

    types = {finding["type"] for finding in findings}
    assert "stray_element" in types
    assert "new_overlap" in types


def test_example_config_documents_every_tunable():
    example_path = Path(__file__).resolve().parents[1] / "config.example.yaml"
    documented = yaml.safe_load(example_path.read_text(encoding="utf-8"))

    assert set(DEFAULT_CONFIG) <= set(documented)
    # And the documented example loads cleanly through the real config loader.
    loaded = load_config(example_path)
    assert loaded["mode"] == "strict"
    assert loaded["filename_key_pattern"] == DEFAULT_CONFIG["filename_key_pattern"]
