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
| Logo/picture shown as an insert+delete pair of the image (same logo, different embed relationship) | `clean_image_revisions` — see **Image revisions are always false** below. `rewrite_docx` also runs the cleaner on `header*/footer*` parts since logos often live in the page header. |
| An unchanged hyperlink (e.g. a signature block e-mail) shown as ins+del differing only in whitespace, because the two files build the link differently (`HYPERLINK` field vs `<w:hyperlink>`) | `is_textlike_revision` — like `is_text_only_revision` but tolerates `HYPERLINK_TAGS`; used in `clean_parent`'s adjacent-pair branch, which resolves by *accepting* the insertion so the field is deep-copied verbatim. Drawings, OLE, tables, SDTs, bookmarks and comment anchors still block |
| Both documents' title spacers survive, so blank lines stack up above the document title | `clean_leading_empty_paragraphs` — **de-duplicates**: removes every pure spacer between the last content paragraph and the first paragraph with text **except the last**, which keeps the intended gap. A single spacer is the document's own gap, not a duplicate, so the rule does not fire. It used to remove all of them and fold a hardcoded `TITLE_GAP_AFTER` onto the logo instead, which only ever matched one letterhead — on Applied Systems' it pulled the title up ~0.14″. The logo paragraph holds a `<w:drawing>`, so it bounds the scan instead of being deleted |
| A split table makes Compare delete a row's text from the original table and re-insert it into the new one, so content looks moved | `clean_relocated_table_noops` — matches a DEL-only against an INS-only paragraph within `RELOCATION_WINDOW` paragraphs of flattened order, in **different** parents, at least one inside a table. Unlike `clean_cross_paragraph_noops` it keeps the **deletion's** location (the original position is the true one), then drops the emptied paragraph and, if the table it lived in is now completely empty, the whole `<w:tbl>` |
| A converted revision splits a sentence where the page broke, so one paragraph stops mid-sentence and the next starts lower-case | `clean_conversion_paragraph_splits` (merges them, inserting one space). Requires: same `<w:body>` parent (so never inside a table), first ends in a letter, second starts with a **lower-case letter**, neither is a heading/`Title`/numbered/tracked-paragraph-mark/section-break paragraph, no pictures/fields/tables, at most `MERGE_MAX_SPACERS` blanks between (removed too) |
| One document auto-numbers headings with `numPr`, the other has the number as literal text, so Compare emits a phantom paragraph holding an inserted bare number (`1.`, `10.`) whose own mark is deleted so it re-merges into the heading | `clean_phantom_list_number_paragraphs` (removes the whole phantom paragraph). Requires the content to be a bare number plus tabs, the mark to be **deleted**, and the *following* paragraph to carry `numPr` — see **Removing an invented paragraph** below |
| The reverse: the *original* auto-numbered its headings with `numPr` while the revision has the number as typed text, so Compare reports `1. ` as **inserted** on every heading (13 of 14 in the Coast letter) | `clean_inserted_list_number_prefixes` — **accepts** the insertion (`accepted_children`), keeping the number as plain text. Dropping it would leave the heading unnumbered, because the paragraph has lost its `numPr` too. Guards: the `<w:ins>` is the paragraph's first content child, its text is a bare number ending in a period, the paragraph has no `numPr` and **no `<w:del>` anywhere**, and unchanged text still follows |
| The revised file lost its numbering entirely, so the original's literal number has nothing to pair with and every heading reads `~~1.~~ Services` | `clean_orphaned_list_number_deletions` — turns the deletion back into unchanged text (`del_to_unchanged_runs`), keeping the number as literal text. The tell that it is accidental: each heading keeps the hanging indent that reserved space for a number. See **Overriding a deletion** below |
| A converted revision spaces paragraphs with `pPr/spacing` where the template uses empty paragraphs, so Compare inserts every blank line the other side has (56 of 157 paragraphs in the Nooks letter) | `clean_inserted_spacer_paragraphs` — removes a contentless paragraph whose mark is specifically **`INS`**, folding the gap onto the preceding paragraph as `spacing/@w:after` (keeping whichever value is larger). Untracked and `DEL`-marked blanks are left alone |
| The letterhead logo's paragraph mark arrives as `<w:ins>` while the empty paragraph below it is marked deleted — one break added, one removed, but a change bar across the top of page 1 | `clean_logo_spacer_mark_swap` (Accept semantics: untrack the logo's mark, drop the empty paragraph). Runs after `clean_image_revisions` has untracked the image and before `clean_inserted_spacer_paragraphs`, which would otherwise eat the inserted half |
| A blank header/footer paragraph's mark is marked inserted or deleted because the revised file lost its header parts in a conversion | `clean_empty_hf_paragraph_marks` — strips the mark revision, **keeps** the paragraph (a header part needs one). Called from `clean_docx_parts` for `header*`/`footer*` only: in the body a tracked mark on a blank line can be a real edit |
| The *original* held the surplus spacers, so Compare marks blank lines and the logo paragraph with a **deleted** paragraph mark — change bars on empty lines and across the top of page 1 | `clean_blank_paragraph_mark_deletions` — strips the `<w:del>` from the mark, **keeps** the paragraph, so layout is untouched. Only for paragraphs with no text either way (a `<w:drawing>` is allowed so the logo qualifies; a comment anchor, field, SDT or any content revision disqualifies). Skips a blank whose *preceding* sibling also carries a deleted mark — see **Merge chains** below |
| A converter sets the bottom page margin inside its own footer distance (`bottom=280`, `footer=720`), so body text prints under the footer — and it propagates into every redline built on that chain | `normalize_section_margins` — sets `bottom = top` when bottom is **both** smaller than top **and** smaller than the footer distance. Two conditions so a deliberately asymmetric setup survives. The document's own top margin supplies the value; no template is hardcoded (ours disagree: 1152 and 1440). Every counterparty return in `examples/` shows the identical `280` |
| A counterparty's stylesheet has a different default font, so every paragraph that names no font of its own silently reflows — SentiLink's `docDefaults` is Times New Roman where ours is Arial | `normalize_default_font` in `rewrite_docx` — repoints `docDefaults/rPrDefault/rPr/rFonts` at `BODY_FONT`, rewriting only the `ascii`/`hAnsi`/`cs` attributes that are **present and disagree**. Never creates an absent attribute (no `rFonts` means the theme supplies the font). Leaves `eastAsia` alone. A style or run naming its own font still wins. Applied as a **scoped text edit, not an ElementTree round-trip** — see the warning below |
| The *original* had the page furniture inline, so the header/footer copy arrives **inserted** while the body copy is **deleted** (reverse of the usual direction) | `clean_header_footer_relocations` handles both directions: match a header/footer `DEL` **or** `INS` against a body paragraph carrying the opposite. Either way the text becomes unchanged in the header/footer, where it belongs, and its body revision is dropped — the words survive in the header/footer part |
| A resolved logo revision leaves its relationship and media file behind, orphaned | `prune_orphaned_media` in `rewrite_docx` — drops `.../image` relationships no part still references and their media when nothing else points at them. Relationship ids are **per-part** (`fontTable.xml` uses `rId8` for a font while `document.xml.rels` uses it for an image), so the reference scan is keyed by owning part. Fails closed on any parse error |

### Converted documents differ structurally, not textually

A file returned via Google Docs or a PDF→Word round-trip says the same words a different way, and
Compare can only express the difference as edits. Two signatures account for most of the noise:

- **Auto-numbering flattened to literal text.** `numPr` on the heading becomes a typed `1.` + tab.
  Compare cannot say "auto-number became text", so it invents a carrier paragraph per heading.
- **Spacer paragraphs traded for `spacing` properties.** One side separates paragraphs with empty
  paragraphs, the other with `pPr/spacing`, so every blank line shows as an insertion.
- **The whole stylesheet is the revised document's.** Compare takes `styles.xml` and `theme1.xml`
  from the revised file wholesale, so a style *name* that survives can resolve to a different font
  than it did in the original. In the SentiLink letter not one run in any of the three documents
  carries an explicit `rFonts`, so the two paragraphs that came through with our `Normal` style
  switched from Arial to Times New Roman — same label, a different stylesheet behind it. The rest of
  the document escaped only because SentiLink's `Body` style happens to declare Arial itself.

Tells that a file has been through a converter: no embedded fonts, no `header*` parts, `ListParagraph`
styles appearing on lists that had none, and `w:sectPr` reduced to a single default section.

Worth knowing when reading such a redline: if the revised file lost its numbering entirely (no inline
`numPr`, no style defining one), the bare numbers show as **deletions** and accepting the redline
yields an unnumbered agreement. That is a defect in the source document, not an artifact — the cleaner
deliberately leaves it visible.

### Never round-trip `styles.xml` through ElementTree

`styles.xml` declares a dozen namespace prefixes and names most of them in `mc:Ignorable`, but ET
emits only the declarations some element actually *uses*. Re-serializing silently drops the rest
while `mc:Ignorable` still references them — and an `mc:Ignorable` naming an undeclared prefix is
exactly what makes Word offer to repair the file.

`normalize_default_font` therefore edits text in place: it bounds the search to the
`<w:rPrDefault>` … `</w:rPrDefault>` block, rewrites the one `<w:rFonts>` tag's attribute values, and
leaves every other byte identical (one changed line in a 34 KB stylesheet). It parses the result and
returns the original on `ParseError`, so a broken stylesheet is never written. A regression test
feeds it an unused `xmlns:w15` with `mc:Ignorable="w15"` and asserts both survive; swapping the
implementation for an ET round-trip fails that test.

`document.xml` is safe to round-trip only because `clean_ignorable_namespaces` runs last and prunes
`mc:Ignorable` to the prefixes still in use.

### Direction: which file holds the surplus decides the tag

Every spacer and furniture artifact arrives in one of two directions, and a rule written for one
is blind to the other. Whichever document holds the surplus blank lines decides it: when the
**revised** file does they come as inserted paragraph marks, and when the **original** does they come
as deleted marks. The Nooks letter was the first case, Applied Systems' the second, and the first
pass of these rules only handled `INS` — so Applied Systems' logo and five of its blank lines still
drew change bars.

So the spacer rules come in pairs: `clean_inserted_spacer_paragraphs` /
`clean_blank_paragraph_mark_deletions`, and both directions of
`clean_header_footer_relocations`. When adding a rule for a structural artifact, write the mirror
at the same time or note explicitly why it cannot occur.

The numbering artifact has **three** shapes, and each was found separately by a user noticing a
redline they could not read — which is the cost of not writing the mirror up front:

| Shape | Rule |
|---|---|
| Inserted number in its own carrier paragraph, mark deleted, next paragraph has `numPr` | `clean_phantom_list_number_paragraphs` |
| Number **deleted** inline, paragraph keeps the hanging indent but has no `numPr` | `clean_orphaned_list_number_deletions` |
| Number **inserted** inline at the head of a paragraph that has no `numPr` | `clean_inserted_list_number_prefixes` |

All three come from the same source: one document generates the number from `numPr`, the other has it
as literal text. Which side is which decides the shape.

**Product rule from the user: reformatting of bullets or paragraph numbers is never worth showing.**
A number or bullet that changes only in how it is produced is noise, so these rules resolve it
silently rather than surfacing it. A number that changes *value* is a real edit and always survives —
that is what the no-deletion and trailing-period guards protect.

### Merge chains

Consecutive deleted paragraph marks are a **merge chain**: each one says "this break goes", so
accepting the run rejoins everything in it into a single paragraph. Applied Systems' letter has a
sentence split by an inline footer — `…do not apply to` / blank / `ID:0` / blank / `any other role…`
— and the five deleted marks are precisely what put it back together.

`clean_blank_paragraph_mark_deletions` therefore skips a blank whose preceding sibling also carries
a deleted mark. Untracking one in the middle breaks the chain and leaves the sentence split across
paragraphs permanently. This does not show up in a word-level diff, because every word is still
present — only the paragraph boundaries are wrong. Check the merged Accept view, not just the text.

### Overriding a deletion

`clean_orphaned_list_number_deletions` is the only rule that *reverses* a deletion rather than
tidying a no-op, so it carries the most guards of anything here. It fires only when the `<w:del>` is
the paragraph's first content child, holds a `<w:tab/>`, reads as a bare number **ending in a
period**, sits in a paragraph with `pPr/ind/@w:hanging` but no `numPr`, with no `<w:ins>` anywhere in
the paragraph, and with unchanged text still following.

Each guard excludes a real edit. The trailing period and the first-child test keep `$35,000` →
`$40,000` tracked (the Nooks letter contains exactly that change two paragraphs away). The
no-insertion test is the important one: a genuine renumber swaps one number for another in place, so
its deletion is always paired with an insertion and can never reach this rule.

### Removing an invented paragraph

`is_rejectable_spacer_paragraph` holds that a **deleted** paragraph mark is the reviewer's to accept,
so a rule must not delete the paragraph out from under them. `clean_phantom_list_number_paragraphs`
and `clean_cosmetic_section_breaks` both break that rule deliberately, and the licence is the same in
each case: the paragraph exists in **neither input document**. Compare invented it as a carrier — for
a literal number, or for a section break — and rejecting it in Word leaves a stray blank line where
the carrier was. The guard that makes this safe is narrowness: the content must be *only* the thing
Compare was carrying.

### Image revisions are always false

**Product rule from the user: a redline must never show a logo/image replacement.** It is always a
false positive — the picture is identical, Word only marks it because the two files reference it
through different relationships.

So `clean_image_revisions` resolves images unconditionally: an image insertion is accepted (unwrapped
to a normal unmarked run), an image deletion is dropped, and a delete+insert pair collapses to one
unmarked image.

A revision that mixes an image with real text is **split, not skipped** — `_resolve_image_runs`
resolves the image runs and leaves the text tracked. This matters because Word sometimes sweeps a
stray character into the same `<w:del>` as the picture (a fresh Atlas comparison put a deleted `e`
next to the old logo), and the old "skip anything with text" guard let that block the whole rule.
An image sharing one `<w:r>` with text is lifted into its own unmarked run, carrying the `rPr` over.

Deliberate consequence, accepted with the product rule: a genuine *"figure and its caption deleted"*
edit shows the caption struck through while the figure simply disappears. Embedded OLE objects
(`OLE_TAGS`) are never touched — they are not logos.

Termination trap: the split branch is guarded on `_image_runs(child)` being non-empty. An image
nested below a direct `<w:r>` (inside a `<w:hyperlink>`, say) is left alone and **must** report zero
changes — reporting a change without making one spins `clean_root`'s fixed-point loop forever.

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

Three scratch tools worth rebuilding if you need them:

1. **A per-rule fire-count reporter** — calls each rule on a pristine parse of every part, since
   `clean_root` only reports an aggregate and hides over-firing.
2. **A text-preservation gate**, the most valuable of the three. Extract two word multisets from a
   docx across `document.xml` and every `header*`/`footer*` part: the **ACCEPT** view is every text
   node not inside a `<w:del>`, the **REJECT** view every text node not inside a `<w:ins>`. A cleaner
   run must lose nothing from either. Decide membership by **ancestry, not by the node's own tag** —
   Word nests revisions, and `<w:ins><w:del>8</w:del></w:ins>` is text inserted and then deleted, so
   it belongs to neither final state. Reading the tag alone reports a false loss on exactly that
   shape (cribl `v4` contains one). Multisets mean re-flowing does not register, so any loss is real.
   It does **not** catch wrong paragraph boundaries — see **Merge chains**.
3. **An AppleScript mirroring [applescript-v2](applescript-v2)'s pipeline but stopping before
   cleaning**, so you can inspect the raw pre-clean state.

Isolate a run by copying the *pre-clean* file and cleaning the copy. Comparing against a Desktop
file is unsound: these documents get edited in Word between runs, and such an edit will look like
the cleaner losing text.

### Running tests

```bash
python3 -m unittest tests/test_clean_redline.py -v   # 167 tests as of November 2026
```
