#!/usr/bin/env python3
"""
example_singulator.py — the reference proof + worked example for deck.py.

Rebuilds the first 15 slides of the Singulator differentiation deck from plain
content (example_content.json + the figures in assets/), using ONLY the deck.py
API — semantic methods for the recurring page shapes, and compose() for the two
one-off pages (metric table, two-way comparison) that have no semantic method.

    python example_singulator.py [out.pptx]

Read it top-to-bottom to see how a content agent drives the engine: no colour,
font, or coordinate appears here — only words, image filenames, and layout intent.
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck import Deck

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")


def build(content, out_path):
    d = Deck(footer="© 2026 Precision Cell Systems · For research use only")
    for s in content["slides"]:
        t = s["type"]

        if t == "cover":
            d.cover(s["title"], s["subtitle"], tagline=s.get("kicker"))

        elif t == "problem":
            d.problem_flow(
                s["title"], s["body"],
                cards=[(c, "") for c in s["chips"]],
                answer=s.get("callout"), caption=s.get("footnote"), eyebrow=s.get("eyebrow"))

        elif t == "glance":
            # agenda() auto-numbers, so strip any leading "N. " the content carries
            strip = lambda c: re.sub(r"^\s*\d+\.\s*", "", c)
            groups = [(g["label"], [(strip(it["cite"]), it["desc"]) for it in g["items"]])
                      for g in s["groups"]]
            d.agenda(s["title"], groups)

        elif t == "study_intro":
            cit = s["citation"]
            d.study_intro(
                s["title"],
                citation=(cit["ref"], f'{cit["author"]} · {cit["doi"]}'),
                method=s.get("method"),
                challenge_title=s.get("challenge_title", "The challenge, and where the Singulator fit"),
                challenge_points=s.get("challenge"),
                eyebrow=s.get("eyebrow"))

        elif t == "study_detail":
            kw = {"title": s["title"], "findings": s.get("findings"), "why": s.get("why")}
            fig = s.get("figure", {})
            if fig.get("image"):
                kw["figure"] = os.path.join(ASSETS, fig["image"])
                kw["figure_caption"] = fig.get("caption")
            elif s.get("steps"):
                kw["steps"] = [(st, "") for st in s["steps"]]
            big = s.get("stat", {})
            if big:
                # a number-led value → big stat; a word-led value → statement
                if any(ch.isdigit() for ch in str(big["big"])):
                    kw["stat"], kw["stat_label"] = big["big"], big["body"]
                else:
                    kw["statement"], kw["statement_sub"] = big["big"], big["body"]
            q = s.get("quote")
            if q:
                kw["quote"] = q["text"]
            d.study_findings(**kw)

        elif t == "metric_table":
            d.compose(
                title=s["title"], dark=True, eyebrow=s.get("eyebrow"), why=s.get("footnote"),
                blocks=[
                    {"type": "table", "col": 0, "colspan": 12, "row": 0, "rowspan": 4,
                     "columns": s["columns"], "rows": s["rows"]},
                    {"type": "callout", "col": 0, "colspan": 12, "row": 4, "rowspan": 1, "text": s["quote"]},
                ])

        elif t == "compare":
            d.compose(
                title=s["title"], dark=True, eyebrow=s.get("eyebrow"), why=s.get("footnote"),
                blocks=[
                    {"type": "card", "col": 0, "colspan": 12, "row": 0, "rowspan": 1,
                     "variant": "navy", "accent": "tint", "body": s["intro"]},
                    {"type": "card", "col": 0, "colspan": 6, "row": 1, "rowspan": 2, "variant": "navy",
                     "accent": "accent_bright", "heading": s["left"]["label"], "body": s["left"]["body"]},
                    {"type": "card", "col": 6, "colspan": 6, "row": 1, "rowspan": 2, "variant": "navy",
                     "accent": "warn", "heading": s["right"]["label"], "body": s["right"]["body"]},
                    {"type": "callout", "col": 0, "colspan": 12, "row": 3, "rowspan": 1, "text": s["quote"]},
                ])

        else:
            raise ValueError(f"unknown slide type: {t!r}")

    d.save(out_path)
    return len(content["slides"])


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "singulator-15.pptx")
    content = json.load(open(os.path.join(HERE, "example_content.json")))
    n = build(content, out)
    print(f"built {n} slides -> {out}")
