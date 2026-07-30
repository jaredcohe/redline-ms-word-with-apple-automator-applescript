#!/usr/bin/env python3
"""Compute a collision-free output base name for a redline.

The Redline workflow names its output after the *revised* document, bumping the version
token so the revised file is never overwritten. Given the target directory and the revised
document's name (without extension), print the next version base name (no extension):

  - Bump the `vN`/`vN.M` token: minor+1 when a minor is present (v7.1 -> v7.2), otherwise
    append `.1` (v7 -> v7.1). The token is anchored on the `v`, NOT on a word boundary, so
    underscore-separated names (v7_260718_...) increment correctly — a trailing `\\b` fails
    there because `_` is a word character.
  - Then, while `<directory>/<candidate>.docx` already exists, keep bumping so an existing
    file is never overwritten.
  - If the name has no version token at all, fall back to appending " redline" (then
    " redline 2", ...) so it still never overwrites.

Usage: next_version_name.py <directory> <revNameNoExt>
"""

from __future__ import annotations

import os
import re
import sys

VERSION_RE = re.compile(r"([vV])(\d+)(?:\.(\d+))?")
EXT = ".docx"


def bump(name: str) -> str | None:
    """Return `name` with its version token bumped, or None if it has no vN token."""
    m = VERSION_RE.search(name)
    if not m:
        return None
    if m.group(3) is not None:
        new = f"{m.group(1)}{m.group(2)}.{int(m.group(3)) + 1}"
    else:
        new = f"{m.group(1)}{m.group(2)}.1"
    return name[: m.start()] + new + name[m.end():]


def next_free_name(directory: str, name_no_ext: str) -> str:
    """Return a base name (no extension) whose `.docx` does not exist in `directory`."""
    def exists(base: str) -> bool:
        return os.path.exists(os.path.join(directory, base + EXT))

    candidate = bump(name_no_ext)
    if candidate is None:
        # No version token — append " redline", then a counter, until free.
        base = f"{name_no_ext} redline"
        if not exists(base):
            return base
        n = 2
        while exists(f"{base} {n}"):
            n += 1
        return f"{base} {n}"

    while exists(candidate):
        nxt = bump(candidate)
        if nxt is None or nxt == candidate:  # safety — should never happen after a bump
            break
        candidate = nxt
    return candidate


def main() -> None:
    if len(sys.argv) != 3:
        sys.stderr.write("usage: next_version_name.py <directory> <revNameNoExt>\n")
        sys.exit(2)
    directory, name_no_ext = sys.argv[1], sys.argv[2]
    print(next_free_name(directory, name_no_ext))


if __name__ == "__main__":
    main()
