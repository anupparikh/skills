#!/usr/bin/env python3
"""Regression checks for the 2026-07-05 light-first + comparison-columns cycle.

Runnable, no framework: `python3 test_archetypes.py`. Guards the behaviors that
re-anchored the engine on the approved decks (see genexp/FINDINGS.md):
  1. evidence resolves to the LIGHT native archetypes (study_intro / study_findings
     light) — NOT the old dark navy stat-tile compose grid.
  2. statement-with-support resolves to problem_flow (the light "Sample prep is
     part of the biology." archetype).
  3. enumerate items carrying a strengths list (`points`) resolve to
     comparison_columns (the "How labs handle X today" archetype), which builds and
     lints clean (0 brand defects) when it is not slide 1.
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from resolve import resolve                     # noqa: E402
from deck import Deck                           # noqa: E402
import lint_brand as L                          # noqa: E402

_BRAND = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "clients", "pcs", "brand")
TEMPLATE = os.path.join(_BRAND, "precision-cell-systems-deck-template.pptx")
CONFIG = os.path.join(_BRAND, "brand_config.json")
MANIFEST = os.path.join(_BRAND, "deck-manifest.json")


def test_evidence_is_light():
    intro = {"intent": "evidence", "title": "T", "citation": {"ref": "R", "author": "A"},
             "method": ["x"], "challenge": {"points": ["p"]}}
    calls = resolve(intro, {})
    assert any(c.get("native") == "study_intro" for c in calls), calls
    assert not any(c.get("dark") for c in calls), "evidence intro must not be a dark slide"

    findings = {"intent": "evidence", "title": "T", "stat": {"value": "143,051"},
                "findings": ["f1", "f2"]}
    calls = resolve(findings, {})
    sf = [c for c in calls if c.get("native") == "study_findings"]
    assert sf and sf[0]["kwargs"].get("light") is True, "findings must render light"


def test_statement_support_is_problem_flow():
    s = {"intent": "statement", "text": "Headline.", "body": "Body.",
         "support": ["Short chip A", "Short chip B",
                     "A long answer sentence that is clearly beyond a chip length threshold here."]}
    calls = resolve(s, {})
    assert len(calls) == 1 and calls[0].get("native") == "problem_flow", calls
    kw = calls[0]["kwargs"]
    assert kw["answer"] and len(kw["cards"]) == 2, kw


def test_enumerate_points_is_comparison():
    s = {"intent": "enumerate", "arrangement": "parallel", "title": "How labs do X today.",
         "items": [
             {"label": "Manual", "desc": "d", "points": ["a", "b"], "warn": "caution one"},
             {"label": "Auto", "desc": "d", "points": ["c"], "warn": "caution two"},
         ]}
    calls = resolve(s, {})
    assert len(calls) == 1 and calls[0].get("native") == "comparison_columns", calls
    cols = calls[0]["kwargs"]["columns"]
    assert len(cols) == 2 and cols[0]["warn"] == "caution one", cols


def test_enumerate_descs_is_checklist():
    s = {"intent": "enumerate", "title": "What labs need:",
         "items": [{"desc": "a"}, {"desc": "b"}, {"desc": "c"}, {"desc": "d"}]}
    calls = resolve(s, {})
    assert len(calls) == 1 and calls[0].get("native") == "checklist", calls
    assert calls[0]["kwargs"]["items"] == ["a", "b", "c", "d"]


def test_two_values_parallel_is_dueling():
    s = {"intent": "enumerate", "arrangement": "parallel", "title": "1 vs 8",
         "items": [{"label": "A", "value": "1", "desc": "one"},
                   {"label": "B", "value": "8", "desc": "eight"}]}
    calls = resolve(s, {})
    assert len(calls) == 1 and calls[0].get("native") == "dueling_numbers", calls
    kw = calls[0]["kwargs"]
    assert kw["left"]["value"] == "1" and kw["right"]["value"] == "8"


def test_section_with_contact_is_closing():
    s = {"intent": "section", "title": "Thank you.",
         "contact": {"name": "Rep", "detail": "a · b", "website": "x.com"}}
    calls = resolve(s, {})
    assert len(calls) == 1 and calls[0].get("native") == "closing", calls
    # a plain section (no contact) stays a section, not a closing
    s2 = {"intent": "section", "title": "Part II", "kicker": "K"}
    assert resolve(s2, {})[0].get("native") != "closing"


def test_closing_builds():
    d = Deck(template=TEMPLATE, footer="© 2026 PCS", manifest=MANIFEST)
    d.closing("Thank you.", {"name": "Rep", "detail": "t · e", "website": "x.com"})
    fd, path = tempfile.mkstemp(suffix="_close.pptx")
    os.close(fd)
    try:
        d.save(path)            # airy by design (like the cover) — build, don't require fill
        assert len(d.prs.slides._sldIdLst) == 1
    finally:
        os.unlink(path)


def test_comparison_builds_and_lints_clean():
    d = Deck(template=TEMPLATE, footer="© 2026 Precision Cell Systems · For research use only",
             manifest=MANIFEST)
    d.cover("Cover", "sub", "tag")                       # slide 1 so comparison isn't slide 1
    d.comparison_columns("How labs handle X today.", [
        {"label": "Manual", "desc": "hand tools", "strengths": ["cheap", "familiar"],
         "warn": "operator-dependent, variable"},
        {"label": "Semi-Auto", "desc": "instruments", "strengths": ["faster"],
         "warn": "less flexible, higher input"},
        {"label": "FFPE", "desc": "solvent washes", "strengths": ["no capex"],
         "warn": "toxic, variable recovery"},
    ])
    fd, path = tempfile.mkstemp(suffix="_cmp.pptx")
    os.close(fd)
    try:
        d.save(path)
        _, defects = L.lint(path, TEMPLATE, config_path=CONFIG)
        assert not defects, [f"{x.slide}:{x.check}:{x.detail}" for x in defects]
    finally:
        os.unlink(path)


def test_group_grid_three_columns_one_slide():
    # 3 groups must render on ONE slide as 3 columns, with NO synthesized filler
    # stat. Regression for the 2026-07-07 capability-grid bug: the old 2-per-slide
    # packing split a 3-column grid across two slides AND stamped a junk
    # "{N} publications in {label}" stat + a hardwired "{N} studies" fill strip.
    s = {"intent": "enumerate", "title": "One prep across samples and assays",
         "groups": [
             {"label": "SAMPLE STATES", "items": [{"desc": "Fresh"}, {"desc": "FFPE"}]},
             {"label": "TISSUES", "items": [{"desc": "Brain"}, {"desc": "Colon"}]},
             {"label": "ASSAYS", "items": [{"desc": "snRNA-seq"}, {"desc": "scATAC"}]},
         ]}
    calls = resolve(s, {})
    assert len(calls) == 1, f"3 groups must be ONE slide, got {len(calls)}"
    bullets = [b for b in calls[0]["blocks"] if b.get("type") == "bullets"]
    assert len(bullets) == 3, f"expected 3 bullet columns, got {len(bullets)}"
    flat = repr(calls)
    assert "publications" not in flat, "no synthesized '{N} publications' filler stat"
    assert "studies" not in flat, "no hardwired '{N} studies' fill strip"
    # a `source` summary fills the bottom band as a real callout banner (the
    # approved deck's summary line), NOT a synthesized count stat.
    s2 = dict(s, source="Fifteen studies · one automated workflow")
    c2 = resolve(s2, {})
    assert any(b.get("type") == "callout" and b.get("text") == "Fifteen studies · one automated workflow"
               for b in c2[0]["blocks"]), "source must render as a bottom callout banner"


def test_labeled_items_are_not_checklist():
    # enumerate/items whose items carry BOTH label and desc are titled rows, not
    # single-line requirements -> must NOT become a green-check checklist (which
    # renders desc-only and drops the label). Regression for the slide-5-vs-slide-7
    # inconsistency: value/customer-value lists must resolve to the same card family.
    s = {"intent": "enumerate", "title": "Fewer wasted reads",
         "items": [
             {"label": "Kersey et al. · Cell Reports Methods 2026", "desc": "lowest contamination of three methods"},
             {"label": "Zhang et al. · Nature Communications 2023", "desc": "38 samples under 30 min each"},
             {"label": "Tanoue et al. · Nature 2026", "desc": "one protocol across 8 samples"},
         ]}
    calls = resolve(s, {})
    assert not any(c.get("native") == "checklist" for c in calls), \
        "label+desc items must not resolve to checklist"
    assert "Kersey et al." in repr(calls), "citation label must be preserved, not dropped"


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    for fn in fns:
        fn()
        print(f"  [PASS] {fn.__name__}")
    print("ALL PASS")
