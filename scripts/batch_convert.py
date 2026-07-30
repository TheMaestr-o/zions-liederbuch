#!/usr/bin/env python3
"""Batch-run convert_song.convert() over a list of song numbers, logging one line per song."""

import os
import subprocess
import sys
import time

from convert_song import convert, osa, KEYNOTE_ID

with open(os.path.join(os.path.dirname(__file__), "remaining_numbers.txt")) as f:
    NUMBERS = [line.strip() for line in f if line.strip()]

LOCKFILE = "/tmp/batch_convert.lock"
RESTART_EVERY = 1  # every single song - confirmed a batch of 79 songs that ALL failed
                    # identically at RESTART_EVERY=4 each succeeded individually in total
                    # isolation (fresh Keynote, one song, quit). So it's not memory bloat
                    # over many songs, it's state left behind by the PREVIOUS song specifically


def restart_keynote():
    try:
        osa(f'tell application id "{KEYNOTE_ID}" to quit saving no')
    except Exception:
        pass
    subprocess.run(["pkill", "-9", "-f", "Keynote Creator Studio"], check=False)
    time.sleep(2)


if __name__ == "__main__":
    if os.path.exists(LOCKFILE):
        with open(LOCKFILE) as f:
            old_pid = f.read().strip()
        print(f"ABORT: lockfile exists (pid {old_pid}). Another batch may still be running. "
              f"Check with `ps -p {old_pid}`, then rm {LOCKFILE} if it's stale.")
        sys.exit(1)
    with open(LOCKFILE, "w") as f:
        f.write(str(os.getpid()))

    fails = []
    try:
        for i, num in enumerate(NUMBERS):
            if i > 0 and i % RESTART_EVERY == 0:
                print(f"--- restarting Keynote after {i} songs ---")
                restart_keynote()
            try:
                ok = convert(num, do_verify=True)
                if not ok:
                    fails.append(num)
            except Exception as e:
                print(f"FAIL {num} EXCEPTION: {e}")
                fails.append(num)
                restart_keynote()
            sys.stdout.flush()
    finally:
        os.remove(LOCKFILE)

    print("\n=== SUMMARY ===")
    print(f"total: {len(NUMBERS)}  failed: {len(fails)}")
    if fails:
        print("failed numbers:", " ".join(fails))
