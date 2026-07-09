#!/usr/bin/env python3
"""map_guided.py — the Guided arm's driver: spec -> resolve.py -> deck.py.

    python3 map_guided.py SPEC.json OUT.pptx --template T.pptx --manifest M.json
                           [--images-dir D]
"""
import argparse
import json
import os
import sys

DECK_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, DECK_SCRIPTS_DIR)
from deck import Deck  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from resolve import resolve  # noqa: E402


def build(spec, template, images_dir, manifest):
    meta = spec["meta"]
    d = Deck(template=template, footer=meta.get("footer"), manifest=manifest)
    d._img_dir = images_dir   # native archetypes (study_findings) resolve figures against this
    ctx = {"images_dir": images_dir}
    for i, s in enumerate(spec["slides"]):
        for call in resolve(s, ctx):
            if "native" in call:
                getattr(d, call["native"])(**call["kwargs"])
            else:
                kwargs = {k: v for k, v in call.items() if k in ("blocks", "title", "dark", "eyebrow", "why")}
                d.compose(images_dir=images_dir, **kwargs)
    return d


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("spec")
    ap.add_argument("out")
    ap.add_argument("--template", required=True)
    ap.add_argument("--manifest", default=None)
    ap.add_argument("--images-dir", default=None)
    args = ap.parse_args()

    spec = json.load(open(args.spec))
    spec_dir = os.path.dirname(os.path.abspath(args.spec))
    images_dir = args.images_dir
    if images_dir and not os.path.isabs(images_dir):
        images_dir = os.path.join(spec_dir, images_dir)
    if images_dir is None and spec.get("meta", {}).get("images_dir"):
        rel = spec["meta"]["images_dir"]
        images_dir = rel if os.path.isabs(rel) else os.path.join(spec_dir, rel)

    d = build(spec, args.template, images_dir, args.manifest)
    d.save(args.out)
    print(f"built {len(spec['slides'])} spec slides -> {len(d.prs.slides._sldIdLst)} pptx slides -> {args.out}")


if __name__ == "__main__":
    main()
