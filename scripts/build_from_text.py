#!/usr/bin/env python3
"""
Builds .key files from hand-supplied text in ../manual-texts/*.json instead of
from a .pptx source.

Needed because a handful of hymns can't be taken from their .pptx as-is:
song 36's source has no line breaks at all inside a verse (one run-on
paragraph, and there's no punctuation to split on safely), and several
carry decisions that only a human can make - repeat marks stripped, an
over-long verse split across two slides, a duplicated word removed. Those
decisions live in the JSON, not in the .pptx.

JSON shape:
    {"number": "36", "title": "...", "refrain": null | "...",
     "verses": ["line\\nline", ...],
     "verse_labels": ["1","1","2","2"]   # optional; defaults to 1..n
    }

Usage:
    python3 build_from_text.py            # every json in ../manual-texts
    python3 build_from_text.py 36 137     # only these
"""

import json
import re
import sys
from pathlib import Path

from make_song import (
    open_and_save_as, set_texts, delete_slide, duplicate_after,
    slide_count, save, close_no_save,
)

HERE = Path(__file__).resolve().parent
TEXTS = HERE.parent / "manual-texts"
SONGS = HERE.parent / "songs"


def build(number, title, verses, refrain, labels, output_path):
    """Same slide surgery as make_song.build_song, but with per-slide verse
    labels so a verse split across two slides can keep its real number on
    both halves rather than being renumbered."""
    open_and_save_as(str(output_path))

    set_texts(2, {1: number, 2: title, 3: title})
    n = len(verses)

    if not refrain:
        delete_slide(8)
        delete_slide(6)
        delete_slide(4)
        while (slide_count() - 1) - 3 + 1 < n:
            blank = slide_count()
            duplicate_after(blank - 1, blank - 1)
        while (slide_count() - 1) - 3 + 1 > n:
            delete_slide(slide_count() - 1)
        for i, text in enumerate(verses):
            set_texts(3 + i, {1: labels[i], 2: text, 3: number, 5: text})
    else:
        while True:
            blank = slide_count()
            if (blank - 3) // 2 >= n:
                break
            last_verse, last_refrain = blank - 2, blank - 1
            duplicate_after(last_verse, last_refrain)
            duplicate_after(last_refrain, last_refrain + 1)
        while True:
            blank = slide_count()
            if (blank - 3) // 2 <= n:
                break
            delete_slide(blank - 1)
            delete_slide(blank - 2)
        for i, text in enumerate(verses):
            vpos = 3 + i * 2
            set_texts(vpos, {1: labels[i], 2: text, 3: number, 5: text})
            set_texts(vpos + 1, {2: refrain, 3: number, 5: refrain})

    save()
    close_no_save()


def main(only=None):
    files = sorted(TEXTS.glob("*.json"))
    if only:
        wanted = {int(x) for x in only}
        files = [f for f in files if int(f.stem) in wanted]

    ok, failed = [], []
    for f in files:
        d = json.loads(f.read_text(encoding="utf-8"))
        num, title = d["number"], d["title"]
        verses, refrain = d["verses"], d.get("refrain")
        labels = d.get("verse_labels") or [str(i) for i in range(1, len(verses) + 1)]

        safe = re.sub(r"[/:]", "-", title)
        out = SONGS / f"{num} - {safe}.key"
        try:
            build(num, title, verses, refrain, labels, out)
            print(f"OK   {num:>4}  {title[:44]:<46} {len(verses)} slides"
                  f"{' + refrain' if refrain else ''}", flush=True)
            ok.append(num)
        except Exception as e:
            print(f"FAIL {num:>4}  {str(e)[:90]}", flush=True)
            failed.append(num)

    print(f"\n{len(ok)} built, {len(failed)} failed")
    if failed:
        print("failed:", " ".join(failed))
    return failed


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1:] or None) else 0)
