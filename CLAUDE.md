# Notes for future Claude sessions

## Deployment

The Automator workflow bundle has its own copies of scripts — edits to the project directory have no effect until deployed.

### Edited `applescript-v2`? → use `install.sh`

```bash
./install.sh
```

Runs all tests, then rebuilds the entire bundle from scratch. **Required for any AppleScript change.** Do NOT use `sync.sh` for this — it will silently leave the old AppleScript in place.

### Edited `clean_redline.py` or `normalize_docx.py` only? → use `sync.sh`

```bash
./sync.sh
```

Copies only the two Python scripts into the bundle. Faster, but does not touch the AppleScript.

### Checking what was cleaned

Every workflow run appends to `/tmp/redline_clean.log`:

```
[2026-06-10 14:23:01] myfile.redline.docx: cleaned 5 artifact(s)
```

---

## Word Compare artifact patterns and clean_redline.py rules

Word Compare produces misalignment artifacts when paragraph counts differ significantly between document versions (e.g., NU adds several paragraphs to a section near a heading). `clean_redline.py` has rules for each pattern observed so far:

| Artifact | Rule |
|---|---|
| Heading text appears as `<w:ins>` in a body paragraph; heading itself is empty | `clean_misplaced_heading_insertions` |
| Last word of a deletion in para N == first insertion of para N+1 | `clean_paragraph_boundary_noop` |
| Heading paragraph contains ONLY `<w:ins>T`; nearby para del ends with T | `clean_heading_full_ins_noop` |
| Inline `<w:ins>T` before `<w:del>…T` in same paragraph (del may have bookmarks) | Rule D in `clean_parent` |
| Deletion-only paragraphs appear AFTER a clean heading instead of before it | `clean_misplaced_deletions_after_heading` |

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
