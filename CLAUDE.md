# Notes for future Claude sessions

## Deployment

The Automator workflow bundle has its own copies of scripts — edits to the project directory have no effect until deployed.

### Edited `applescript-v2`? → use `install.sh`

```bash
./install.sh
```

Runs all tests, then rebuilds the entire bundle from scratch. **Required for any AppleScript change.** Do NOT use `sync.sh` for this — it will silently leave the old AppleScript in place.

### Edited a Python helper only (`clean_redline.py`, `normalize_docx.py`, `next_version_name.py`)? → use `sync.sh`

```bash
./sync.sh
```

Copies only the Python helper scripts into the bundle. Faster, but does not touch the AppleScript. (A **new** helper file also needs adding to the tuple in `build_workflow.py` and, since that's an AppleScript-adjacent bundle change, an `install.sh`.)

### Checking what was cleaned

Every workflow run appends to `/tmp/redline_clean.log`:

```
[2026-06-10 14:23:01] myfile.redline.docx: cleaned 5 artifact(s)
```

---

## Output versioning (`next_version_name.py`)

The output file is named after the **revised** document with its version token bumped, so the revised file is never overwritten. `next_version_name.py <dir> <revNameNoExt>` bumps the `vN`/`vN.M` token (minor+1, or append `.1`), then **keeps bumping while the target `.docx` already exists** so repeated runs never overwrite (`v7` → `v7.1` → `v7.2` …).

Gotcha that caused a v7-overwrite bug: the old inline regex used `([vV])(\d+)\b`, but `\b` fails on underscore-separated names (`v7_260718…`) because `_` is a word character — so no token matched and the name came back unchanged, overwriting the source. The helper anchors on the `v` instead (`([vV])(\d+)(?:\.(\d+))?`), so `_`/space/end all work.

## Word Compare artifact patterns and clean_redline.py rules

Word Compare produces misalignment artifacts when paragraph counts differ significantly between document versions (e.g., NU adds several paragraphs to a section near a heading). `clean_redline.py` has rules for each pattern observed so far:

| Artifact | Rule |
|---|---|
| Heading text appears as `<w:ins>` in a body paragraph; heading itself is empty | `clean_misplaced_heading_insertions` |
| Last word of a deletion in para N == first insertion of para N+1 | `clean_paragraph_boundary_noop` |
| Heading paragraph contains ONLY `<w:ins>T`; nearby para del ends with T | `clean_heading_full_ins_noop` |
| Inline `<w:ins>T` before `<w:del>…T` in same paragraph (del may have bookmarks) | Rule D in `clean_parent` |
| Deletion-only paragraphs appear AFTER a clean heading instead of before it | `clean_misplaced_deletions_after_heading` |
| Empty list paragraphs (`<w:numPr>` but no text/delText, no tracked change) render as orphan bullets/letters, e.g. `(a) (b) (c)…` | `clean_empty_list_paragraphs` (skips bookmark-anchoring paras and a table cell's last paragraph) |
| Unchanged heading appears as a trailing `<w:ins>T` at the tail of a body paragraph AND as a full `<w:del>T` in a nearby heading (within 3 paras, skipping del-only paras) | `clean_trailing_heading_ins_del_noop` (drops the trailing ins run, converts the heading's del back to unchanged text) |
| Word Compare silently injects the revised doc's section breaks (`detect format changes false` adopts them untracked, so they can't be rejected) | `clean_cosmetic_section_breaks` (removes a mid-doc `<w:sectPr>` whose `pgSz` matches the final body section; keeps landscape/differently-sized sections, the body-level section, and never deletes a paragraph) |
| Logo/picture shown as an insert+delete pair of the image (same logo, different embed relationship) | `clean_image_revisions` (accepts image-only insertions to an unmarked run, drops image-only deletions; skips revisions with text or an OLE object). `rewrite_docx` also runs the cleaner on `header*/footer*` parts since logos often live in the page header. |

### Verifying a fix

```bash
cp "<desktop redline>.docx" /tmp/test.docx
python3 clean_redline.py /tmp/test.docx
# then inspect paragraphs around the suspect heading with Python
```

### Running tests

```bash
python3 -m unittest tests/test_clean_redline.py -v   # 24 tests as of June 2026
```
