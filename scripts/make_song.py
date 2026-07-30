#!/usr/bin/env python3
"""
Builds a Zions-Liederbuch style Keynote song presentation from the "900 - Muster.key" template.

Usage: edit the SONG dict at the bottom of this file and run:
    python3 make_song.py

Template slide layout (900 - Muster.key), positions before any editing:
    1        blank (opening)
    2        cover: [number, title, title-shadow, "Lied Nr."]
    3,5,7    verse slides:   [verse-number, verse-text, song-number, "", verse-text-shadow]
    4,6,8    refrain slides: ["Ref: ", refrain-text, song-number, "", refrain-text-shadow]
    9        blank (closing)

SAFETY RULE (learned the hard way, twice): Keynote autosaves in place, and its
AppleScript "save ... in POSIX file X" does NOT reliably re-home that autosave
target - it can keep writing back to whatever file was originally opened, even
after a scripted "save as". So this script never lets Keynote open the template
at all: it copies the template to the output path on the filesystem FIRST (plain
cp, no Keynote involved), and only then tells Keynote to open the copy. Keynote's
own autosave can only ever land on the copy this way.
"""

import shutil
import subprocess
import sys

TEMPLATE_PATH = "/Users/ohnedan/Developer/zions-liederbuch/scripts/900 - Muster.key"
KEYNOTE_ID = "com.apple.Keynote"


def osa(script: str) -> str:
    result = subprocess.run(
        ["osascript", "-"], input=script, capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error:\n{script}\n---\n{result.stderr}")
    return result.stdout.strip()


def as_str(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", '" & linefeed & "') + '"'


def open_and_save_as(output_path: str):
    shutil.copyfile(TEMPLATE_PATH, output_path)
    osa(f'''
    tell application id "{KEYNOTE_ID}"
        open POSIX file {as_str(output_path)}
        delay 2
    end tell
    tell application "System Events" to set visible of application process "Keynote" to false
    ''')


def slide_count() -> int:
    return int(osa(f'tell application id "{KEYNOTE_ID}" to count of slides of front document'))


def delete_slide(index: int):
    osa(f'tell application id "{KEYNOTE_ID}" to delete (slide {index} of front document)')


def duplicate_after(src_index: int, after_index: int):
    osa(f'''
    tell application id "{KEYNOTE_ID}"
        duplicate (slide {src_index} of front document) to after (slide {after_index} of front document)
    end tell
    ''')


def set_texts(slide_index: int, items: dict):
    lines = [f'set theSlide to slide {slide_index} of front document',
             'set t to text items of theSlide']
    for item_index, text in items.items():
        lines.append(f'set object text of (item {item_index} of t) to {as_str(text)}')
    body = "\n        ".join(lines)
    osa(f'''
    tell application id "{KEYNOTE_ID}"
        {body}
    end tell
    ''')


def save():
    osa(f'tell application id "{KEYNOTE_ID}" to save front document')


def close_no_save():
    osa(f'tell application id "{KEYNOTE_ID}" to close front document saving no')


def build_song(number: str, title: str, verses: list, refrain: str, output_path: str):
    """refrain: pass "" or None if the song has no refrain."""
    open_and_save_as(output_path)

    # cover slide
    set_texts(2, {1: number, 2: title, 3: title})

    if not refrain:
        # drop refrain slides (8, 6, 4) - highest index first so positions below stay stable
        delete_slide(8)
        delete_slide(6)
        delete_slide(4)
        # now: 1 blank, 2 cover, 3/4/5 verses, 6 blank
        n = len(verses)
        while (slide_count() - 1) - 3 + 1 < n:   # fewer verse slots than verses -> duplicate last one
            blank = slide_count()
            last_verse = blank - 1
            duplicate_after(last_verse, last_verse)
        while (slide_count() - 1) - 3 + 1 > n:   # more slots than verses -> trim from the end
            blank = slide_count()
            delete_slide(blank - 1)

        for i, verse_text in enumerate(verses, start=1):
            set_texts(2 + i, {1: str(i), 2: verse_text, 3: number, 5: verse_text})
    else:
        n = len(verses)
        while True:
            blank = slide_count()
            pairs = (blank - 1 - 2) // 2  # content slides = blank-1-2 (cover+blank excluded), /2 per pair
            if pairs >= n:
                break
            last_verse = blank - 2
            last_refrain = blank - 1
            duplicate_after(last_verse, last_refrain)   # new verse slot right after last pair
            new_verse_pos = last_refrain + 1
            duplicate_after(last_refrain, new_verse_pos)  # new refrain slot right after new verse
        while True:
            blank = slide_count()
            pairs = (blank - 1 - 2) // 2
            if pairs <= n:
                break
            delete_slide(blank - 1)  # refrain
            delete_slide(blank - 2)  # verse

        for i, verse_text in enumerate(verses, start=1):
            verse_pos = 3 + (i - 1) * 2
            refrain_pos = verse_pos + 1
            set_texts(verse_pos, {1: str(i), 2: verse_text, 3: number, 5: verse_text})
            set_texts(refrain_pos, {2: refrain, 3: number, 5: refrain})

    save()

if __name__ == "__main__":
    print("Use convert_song.py <number> for pptx-based songs, "
          "or import build_song() directly for manually supplied text.")
