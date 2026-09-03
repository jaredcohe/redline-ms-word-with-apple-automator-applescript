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

## Cloud-only Google Docs (`.gdoc`)

A `.gdoc` is a small JSON shortcut (`{"doc_id": …}`). When Google Drive keeps it "online only" (not opened recently), reading it raises `OSError: [Errno 11] Resource deadlock avoided` — Drive can't serve the dehydrated file synchronously. `gdocExportToDocx` first `open`s the file (which makes Drive download the shortcut and shows the doc in the browser), then **polls the read** for up to 60s until the JSON materializes. The upfront "downloading from cloud storage" step still skips `.gdoc` (it force-downloads Word docs by opening them in Word).

## Body formatting normalization (`normalize_body_formatting`)

Every run, `clean_redline.py` forces **body paragraphs to 10pt (`sz`/`szCs val="20"`) and 1.15 line spacing (`spacing line="276" lineRule="auto"`)** as silent (untracked) direct formatting. Applies to normal body + table paragraphs; **skips** headings (`Heading1`–`6`) and `Title`/`Subtitle`, and only the main document (no `<w:body>` in header/footer parts). New `spacing`/`rPr`/`sz` elements are inserted in OOXML schema order via `PPR_ORDER`/`RPR_ORDER` (`_get_or_insert_ordered`) so Word does not flag the file for repair. The pass is idempotent and runs once after the artifact-cleaning loop.

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
| Word Compare silently injects the revised doc's section breaks (`detect format changes false` adopts them untracked, so they can't be rejected) | `clean_cosmetic_section_breaks` (removes a mid-doc `<w:sectPr>` whose `pgSz` matches the final body section; keeps landscape/differently-sized sections and the body-level section). Also deletes the break's **host paragraph** when it is a pure spacer, else a stray blank line lands above the next heading — bookmarks move forward first; skipped for a cell's last paragraph, the body's last paragraph, and a paragraph sitting between two `<w:tbl>` (removing it would splice them) |
| Logo/picture shown as an insert+delete pair of the image (same logo, different embed relationship) | `clean_image_revisions` (accepts image-only insertions to an unmarked run, drops image-only deletions; skips revisions with text or an OLE object). `rewrite_docx` also runs the cleaner on `header*/footer*` parts since logos often live in the page header. |
| An unchanged hyperlink (e.g. a signature block e-mail) shown as ins+del differing only in whitespace, because the two files build the link differently (`HYPERLINK` field vs `<w:hyperlink>`) | `is_textlike_revision` — like `is_text_only_revision` but tolerates `HYPERLINK_TAGS`; used in `clean_parent`'s adjacent-pair branch, which resolves by *accepting* the insertion so the field is deep-copied verbatim. Drawings, OLE, tables, SDTs, bookmarks and comment anchors still block |
| Both documents' title spacers survive, so blank lines stack up above the document title | `clean_leading_empty_paragraphs` (removes every pure spacer between the last content paragraph and the first paragraph with text, folding the gap onto that content paragraph as `spacing/@w:after` — the largest explicit `after` removed, or `TITLE_GAP_AFTER`). The logo paragraph holds a `<w:drawing>`, so it bounds the scan instead of being deleted |
| A split table makes Compare delete a row's text from the original table and re-insert it into the new one, so content looks moved | `clean_relocated_table_noops` — matches a DEL-only against an INS-only paragraph within `RELOCATION_WINDOW` paragraphs of flattened order, in **different** parents, at least one inside a table. Unlike `clean_cross_paragraph_noops` it keeps the **deletion's** location (the original position is the true one), then drops the emptied paragraph and, if the table it lived in is now completely empty, the whole `<w:tbl>` |
| A converted revision splits a sentence where the page broke, so one paragraph stops mid-sentence and the next starts lower-case | `clean_conversion_paragraph_splits` (merges them, inserting one space). Requires: same `<w:body>` parent (so never inside a table), first ends in a letter, second starts with a **lower-case letter**, neither is a heading/`Title`/numbered/tracked-paragraph-mark/section-break paragraph, no pictures/fields/tables, at most `MERGE_MAX_SPACERS` blanks between (removed too) |

### Cross-part cleaning

Some artifacts span parts, so `rewrite_docx` no longer cleans each part in isolation:

- `clean_root(root)` runs every rule over one part to a fixed point.
- `clean_docx_parts(parts)` parses all cleanable parts, `clean_root`s each, runs the cross-part pass, then `clean_root`s again to a fixed point and normalizes formatting.
- `clean_document_xml(bytes)` still cleans one part alone (used by the tests' `clean_fragment`).

The one cross-part rule so far: **`clean_header_footer_relocations`**. A PDF-derived revision has no header/footer parts at all — the page furniture is baked into the body — so Compare deletes the original's footer text *and* inserts the same string mid-body, which is how a footer prints in the middle of a paragraph. It matches a header/footer DEL-only paragraph against a body INS-only paragraph with the same text (whitespace-insensitive), restores the footer del in place, and deletes the body insertion. Returns a per-part change count so the caller knows which parts to rewrite.

### `<w:ins>`/`<w:del>` as markers, not wrappers

A revision inside a property container (`REVISION_MARKER_PARENTS`: `rPr`, `trPr`, `tcPr`, `tblPr`, `tblPrEx`) marks the **paragraph mark**, row or cell itself and is empty by design. `clean_parent` skips these before its empty-wrapper rules.

This was a real bug: without the guard, every `pPr/rPr/w:ins|w:del` was pruned as an "empty revision wrapper", so a genuine "this paragraph break was deleted" edit vanished silently and could not be rejected. A fresh Atlas comparison carries **20** tracked paragraph marks (12 ins, 8 del); the shipped `v2.1`, built by the old cleaner, has **0**.

### Inserted vs deleted paragraph marks — which spacers may be removed

Fixing the above broke every rule that guarded on "paragraph mark not tracked", because real Word Compare output marks a converted revision's layout noise as **inserted** paragraph marks. The predicates are therefore split three ways:

| Predicate | Means | Used by |
|---|---|---|
| `_is_contentless_paragraph` | no text, no *content* revision, nothing but bookmarks — **ignores the mark** | the two below |
| `is_spacer_paragraph` | contentless **and** mark untracked | anywhere tracked state must not change |
| `is_rejectable_spacer_paragraph` | contentless **and** mark untracked *or inserted* | `clean_cosmetic_section_breaks`, `clean_leading_empty_paragraphs`, `_drop_emptied_container`, `_is_empty_table`, the merge rule's spacer scan |

The principle: **deleting a paragraph whose mark is an insertion is exactly what Word's Reject does to it** — nothing the author wrote is lost. A **deleted** mark is left alone; that one is the reviewer's to accept, and removing it would pre-empt them. `paragraph_mark_revision(p)` returns `INS`/`DEL`/`None`; `paragraph_mark_has_revision` is kept as a thin wrapper.

`_is_empty_table` follows the same rule — a stray table whose cells carry only inserted marks may go, one marked deleted stays.

### `clean_ignorable_namespaces` runs LAST

Removing a revision orphans the prefixes it used (`w16du:dateUtc` on every `w:ins`/`w:del` Word 2024 writes), so the `mc:Ignorable` prune must run *after* the cleaning loop. It used to run before, which left a stale prefix for the workflow's **second** cleaner run to notice — the file was rewritten on pass 2 and the run reported `Cleaned 1` instead of `Cleaned 0`.

### Idempotence is a hard requirement

[applescript-v2](applescript-v2) runs `clean_redline.py` **twice** per redline — once on the Desktop temp and again after the file is moved to its final location. A non-idempotent rule double-applies on every real run. Verify with a second invocation: it must print `Cleaned 0`.

### Verifying a fix

```bash
cp "<desktop redline>.docx" /tmp/test.docx
python3 clean_redline.py /tmp/test.docx
# then inspect paragraphs around the suspect heading with Python
```

The three-file Atlas set in `examples/atlas/` is the reference case for conversion artifacts
(`v1` = NU original, `v2` = Atlas's PDF-derived return, `v2.1` = generated redline).

**Do not verify against `v2.1`.** It is already post-clean *and* was produced by an older cleaner,
so its `<w:sectPr>`s and all 20 tracked paragraph marks are already gone — rules guarding on either
will silently appear to work. Generate a fresh comparison instead; the same applies to
`examples/cribl/v4.1`. `examples/cribl/` has no matched original/revised pair (`v3` and `v4` are
different engagements off one template), but it is the right **normal-document** control: the
conversion-specific rules should fire zero times on it.

Two scratch tools worth rebuilding if you need them: a per-rule fire-count reporter (calls each
rule on a pristine parse of every part, since `clean_root` only reports an aggregate and hides
over-firing), and an AppleScript that mirrors [applescript-v2](applescript-v2)'s pipeline but stops
before cleaning, so you can inspect the raw pre-clean state.

### Running tests

```bash
python3 -m unittest tests/test_clean_redline.py -v   # 98 tests as of September 2026
```
