---
name: mikra
description: Extracts fixture data (manufacturer, voltage/dimming, finish colour, product photo, installation category) from a manufacturer spec-sheet PDF for an architectural lighting fixture, and inserts it as a new row into the studio's lighting legend Excel file (מקרא תאורה). Use this whenever the user uploads one or more lighting fixture PDFs from any manufacturer and wants the data added to a מקרא תאורה / lighting legend spreadsheet. Never invent or guess a value that is not clearly supported by the PDF text — always flag it, or ask, instead.
---

# Lighting Fixture PDF → Legend Row

## What this skill does

Takes one manufacturer PDF per fixture, reads its spec text, and appends one row to the
studio's lighting legend (`אומדן` sheet), in the correct section: ceiling feed
(מנורות הזנה מהתקרה), wall feed (מנורות הזנה מהקיר), or floor feed (מנורות הזנה מהרצפה).

Two stages, kept separate on purpose:

1. **Reading the PDF — Claude does this directly**, because manufacturer catalog layouts
   differ too much for one fixed regex parser. Claude reads the extracted text/images itself
   and decides field values with judgment.
2. **Writing the Excel row — the bundled script does this**, because the layout (colour bars,
   merged category headers, formulas, row heights, image embedding) must be handled precisely
   and consistently every time.

## Hard rule: no hallucination

If a field is not clearly stated in the PDF, or the wording is ambiguous, do **not** guess.
Pass its canonical field name (see Step 3) in `uncertain_fields` so the script highlights the
whole cell orange with a comment for the human to check. For a short single-value field (גוון,
מתח_עמעום, יצרן, סימון, תמונה) that's the right amount of highlighting, since the whole cell IS
the uncertain value. For **תיאור specifically**, don't flag the whole multi-line cell just
because one phrase inside it is shaky (e.g. an uncertain beam-angle line inside an otherwise
solid description) — pass that exact phrase via `uncertain_snippets` instead (see Step 3), which
highlights only that phrase (bold, orange text) and leaves the rest of the cell looking like any
confirmed entry. This applies especially to:
- the installation category (ceiling/wall/floor) when the PDF gives no clear installation type
- the finish colour, when several options are listed with no clear single choice
- voltage/dimming, when the PDF lists more than one option ("No Dim - Other DIM, please consult")
- **any field where the PDF offers more than one real option** (e.g. a beam angle of "38° / 72°",
  several colour-temperature SKUs) — this is common enough that it gets its own rule below.

### Fields with multiple real options — ask, don't pick one, batched

Some spec sheets offer more than one real configuration for the same field. Two different
situations come up in practice, and both are resolved by asking — never by guessing or
averaging — but the framing differs:

**A. The PDF lists concrete alternative values**, e.g. `Light beam angle: 38° / 72°`, or a
photometric table with a separate output row per beam angle. Present the *exact* choices as they
appear in *this* PDF and ask the person to pick one (e.g. "זווית הפיזור בגיליון הזה היא 38° או
72° — איזו מהן?"). Once they answer, use that specific option's values for every field the choice
affects — including any dependent value the PDF ties to it (if the lumen output differs by beam
angle in the PDF's own table, use the figure for the chosen angle, not a generic headline figure).

**B. The PDF gives one primary value plus vague "consult us" language**, e.g.
`Colour temperature: 2700K - Other K, please consult` or `Dimming: No Dim - Other DIM, please
consult`. This isn't a real menu of options in the PDF itself, but it's still not necessarily
settled — ask anyway, since the person may already know exactly what this specific fixture needs
even though the PDF doesn't spell it out (e.g. "ה-PDF נותן רק 2700K כברירת מחדל, אבל אתה יודע אם
זה מה שאתה רוצה לגוף הזה?"). If they give a value, use it. If they don't know yet, fall back to
the normal uncertain-field rule instead of blocking: leave it and flag the cell orange for later
review.

**Batching**: when processing several PDFs in one pass, don't stop after each one — read all of
them first, collect every open question (from both A and B) across all fixtures, and ask them
all together in one batch before inserting any rows. Use the AskUserQuestion tool (it supports up
to 4 questions per call — chain further calls if a batch has more than 4 open questions).

## Step 0: project header, every time a new legend workbook is started

Before creating a *new* legend workbook from the skeleton (not when adding another fixture to
one already in progress), there are **two required questions** to ask the person if they weren't
already given earlier in the conversation. Do not guess or reuse either from a previous project:

1. **שם הפרויקט** (project name)
2. **גוון תאורה בפרוייקט** (project-wide light colour temperature, e.g. `2700K` / `3000K`) —
   this is different from **גוון** the per-fixture field in step 2, which is the fixture's own
   body/finish colour, not the light's colour temperature.

- If they give both, write them into the header with `scripts/set_project_header.py --name ... --color ...`.
- If either is missing at this point, still build the workbook, but leave that header area
  highlighted orange with a comment (the script does this automatically), and say so explicitly
  in the chat reply.
- The date cell is never asked about — always set automatically to the date the workbook is being
  built, via the same script.

## Step-by-step

### 1. Extract the PDF content
Run `scripts/read_pdf.py <pdf_path>` to get per-page text and to dump embedded/rendered images to
a temp folder. If the file is a normal PDF, this uses `pdftotext -layout` and `pdftoppm`/`pdfimages`.
Read the printed text yourself in-context; view the dumped page images if a product photo needs
visual confirmation (e.g. to pick the correct photo when a page has several product variants).

**If `read_pdf.py` exits with an error, stop for that PDF and say so** — never fill a row from a
file it couldn't read:

| exit code | meaning | what to do |
|---|---|---|
| 2 | file not found | ask for the right file |
| 3 | not a PDF (e.g. a saved web page renamed to `.pdf`) | ask for the real spec sheet |
| 4 | damaged / incomplete PDF | ask for a fresh copy; insert nothing from it |
| 5 | no text layer (scanned) | page images were still written: you may read them, but pass **every** field you read from an image in `uncertain_fields` — a value read off a picture is never "confirmed" |

### 2. Decide the field values
From the text, work out:
- **יצרן** (manufacturer) — company name, usually in the header/footer or a "Category/Brand" line
- **סימון** — leave blank unless the user gives the project fixture code; this is a studio-assigned
  tag, not something the PDF knows
- **תיאור** — write in the studio's existing style, per `references/description_style.md`
  (short newline-separated phrase fragments, not full sentences; line order: fixture type +
  model name, LED spec line, beam angle if relevant, other trailing notes if relevant, then
  the fixed driver line). Keep English model/reference codes exactly as the manufacturer writes
  them, never translated or reworded. Two things the description must NOT contain: the finish
  colour (goes in the גוון column only) and any dimension/site-coordination note (that's filled
  in by studio staff by hand, never by the skill). The driver line
  "דרייברים למיקום באזור פריק ונגיש" is fixed wording and always included, in every fixture,
  regardless of type.
- **מתח_עמעום** — voltage (e.g. `48V`, `24V`, `350mA`) plus dimming protocol (e.g. `DIM DALI`), two lines
- **גוון** — the fixture body/finish colour (not the light colour temperature — that is the
  project-wide setting from Step 0, not per fixture)
- **תמונה** — pick the clearest product photo page (usually a rendering/hero shot page rather
  than a technical/polar-diagram page)
- **קטגוריה** (ceiling/wall/floor, chosen for the `--category` argument, not written as a data
  field itself) — guess from the installation type keywords:
  - ceiling: recessed, suspension, pendant, track, plaster-in, ceiling-mounted, chandelier
  - wall: wall-mounted, sconce, wall washer
  - floor: floor lamp, portable, standing
  If the PDF gives no installation-type hint at all, mark the category itself as uncertain and
  ask which section to place it in rather than guessing.

The skill only ever writes these six fields (five data fields above, plus the image). **Every
other column in the sheet — quantities, prices, totals, or anything project-specific like a
per-floor split (e.g. קרקע/מרתף seen in real project files) — is left completely untouched**,
regardless of what it's called or where it sits in the file.

### 3. Insert the row
Call `scripts/insert_fixture.py` with the target workbook path, the category, and a JSON dict of
**canonical field name → value** (not column letters — column position differs per project file,
since real files customize the base template). Recognized keys: `סימון`, `תיאור`, `יצרן`,
`מתח_עמעום`, `גוון`. The image goes through `--image`, not the data dict.

```
python3 scripts/insert_fixture.py <xlsx_path> --category ceiling \
  --data '{"סימון":"SH-2","תיאור":"...","יצרן":"VELTRANO","מתח_עמעום":"48V\nDIM DALI","גוון":"הכל לבן"}' \
  --image /path/to/photo.jpg --uncertain גוון
```

To flag only a specific phrase inside `תיאור` (rather than the whole field), use
`--uncertain-snippets` instead of `--uncertain` for that field, with the exact substring to mark:

```
--uncertain-snippets '{"תיאור":"זוית פיזור 38 מעלות"}'
```

The substring must match the text in `--data` verbatim (it's a plain `str.find`, not a regex or
fuzzy match). If it isn't found, the script falls back to flagging the whole `תיאור` cell instead
of silently skipping the flag.

The script:
- reads the workbook's own header row (row 8) to find which column is which, on that specific
  file — it does not assume a fixed layout, and raises a clear error if it can't find a column
  for something it was asked to write
- finds the first genuinely free row in that category block (skipping rows that only *look* empty
  because they're a leftover non-anchor cell of an older multi-row-merged entry), or inserts a new
  row before the next category header if the block is full — correctly shifting merged category
  headers below it and keeping the summary `SUM` formula's range valid either way
- writes a **single row per fixture, never merged across multiple rows** (older files sometimes
  have 3- or 5-row merged entries from manual formatting; this was a per-person choice, not a
  format rule — new rows are always one row, sized automatically to fit the description text)
- writes the values, right-aligned RTL
- embeds the photo into whichever column is recognized as `תמונה`, scaled to the row height
- fills any uncertain field orange (`FFC000`) and attaches a cell comment saying what to check —
  or, for a snippet-flagged `תיאור`, bolds/colours just that phrase instead of the whole cell
- auto-numbers `סעיף`
- **on every save**, automatically strips any leftover skeleton/example row (identified by a
  `תיאור` cell whose text still literally starts with `[` — the bracketed instructional
  placeholder baked into the base template, e.g. row 10 of `legend_template.xlsx`) from every category
  block, and attaches the standing review-request note (see Step 4) as a comment on `A2`. A
  delivered legend should never show a template/example row — only real fixtures that were
  actually added.

### 3.5. Check the new row against the fixture bank

Right after `insert_fixture.py` prints `Inserted row N in block '...'` (N is the row's final
position, after any template row above it was removed), check that same
fixture against the studio's fixture bank (`bank.xlsx` — in this repository, `examples/out/bank.xlsx`
after running `demo.py`). If the
bank file can't be found there, ask the person for its current path once rather than
guessing; if it truly isn't available, skip this whole step and say so in the Step 5
summary — never let a missing bank file block inserting the fixture into the project's own
legend, which is this skill's real job.

Run `../bank/match_fixture.py <bank_path> --manufacturer "<יצרן>" --description
"<תיאור>"` (the exact same values just written into the row). It prints one JSON verdict —
comparison is approximate on purpose (real PDFs and past entries rarely word a תיאור
identically twice), so treat the three verdicts as three different actions, not as three
degrees of the same warning:

- **`"strong"`** — this fixture is already in the bank. Do nothing to the row. Just note it
  in the Step 5 summary (e.g. "כבר קיים בבנק") — no comment, no highlight, since nothing
  needs the person's attention.
- **`"maybe"`** — a similar-looking bank row exists, but it's not confident enough to treat
  as the same fixture. Run `../bank/mark_bank_status.py <mikra_path> --row N --status
  maybe --bank-description "<the verdict's bank_description>"` — this attaches a review
  comment to the row (never touching the תיאור/גוון/etc. cells insert_fixture.py already
  wrote or flagged), asking a human to check and merge if it's really a duplicate.
- **`"new"`** — no match found. Run `../bank/mark_bank_status.py <mikra_path> --row N
  --status new`. This attaches a comment saying the fixture will be added to the bank
  automatically once the corrected file is returned (see Step 4), and that the person should
  delete the comment before sending the file back if this particular fixture is a
  one-off that doesn't belong in the shared bank (e.g. a client-supplied or discontinued
  item) — that's their opt-out, no separate chat exchange is needed for it.

Do this once per fixture inserted in the current pass, right after that fixture's own
`insert_fixture.py` call — don't batch it separately from the insertion loop.

### 4. Export the workbook — always a real file, never left "in the drive"
Once fixtures are inserted, the finished **`.xlsx` file itself** — not only a PDF/screenshot
preview — must be handed to the person doing the work. Convert the workbook to PDF or take a
screenshot region for a quick visual sanity check first (this is for Claude's own QA, to catch
layout problems before handing anything over), but the actual deliverable is the Excel file,
downloadable/attachable directly (e.g. via `SendUserFile` in Cowork) — not just edited in place in
a shared/synced location.

The workbook itself always carries a standing review note (a comment on cell `A2`, added
automatically by `insert_fixture.py` — see Step 3) asking the person reviewing to check every
orange-flagged cell and fix it if needed. **Repeat the same request in the chat message that
delivers the file** — both channels, since the file may get forwarded on its own (email, WhatsApp)
without this conversation attached. Ask them to send the corrected file back once they're done
reviewing/fixing it — that returned copy is what feeds later phases of this project (the fixture
bank), not something to act on immediately at this stage.

If any row got a "🆕 גוף חדש" comment in Step 3.5, mention in that same chat message that those
fixtures will be added to the shared fixture bank automatically from the returned file, and that
they can delete a "🆕" comment on any row that shouldn't go into the bank (a one-off, client-
supplied, or discontinued item) before sending the file back — that deletion is their entire
opt-out, nothing else is needed on their end.

**When the corrected file comes back**, before doing anything else with it, run
`../bank/process_new_fixtures.py <returned_mikra_path> <bank_path>` once. It finds every
row still carrying a "🆕" comment, adds each one to the bank (this is the confirmation step — the
person returning the file already had the chance to remove a wrongly-flagged row, so no further
chat approval is needed) — **except a row that still has an orange cell or an orange phrase**:
nobody has confirmed that value yet, and the bank is copied into every future project, so the row
is skipped with a reason and is picked up when the file comes back with the orange removed. It flips each added row's comment to "✓ נוסף לבנק" so a repeat run
never double-adds. It prints a short report of what it added — mention that in the chat reply
handling the returned file, the same way Step 5 mentions everything else this skill does.

### 5. Chat summary, after the workbook is ready
Write a short summary of what was done: which fixtures were added, to which category, which cells
were left flagged for human review and why, and the bank-check outcome from Step 3.5 for each
fixture (already in the bank / flagged for a manual bank check / new — will be added once the file
comes back). Because plain chat text mixes Hebrew and English codes/terms and can render in the
wrong direction, build this summary with the `rtl-hebrew-docs` skill — HTML `dir="rtl"` rendered to
PDF via `scripts/html_to_pdf.py` from that skill — rather than typing it directly as a chat message.
Present the PDF, and keep the chat message itself to one or two lines pointing at it, including the
review-and-send-back request from Step 4.

## Notes
- The colour bars next to each category header already exist in the template; do not recreate them.
- Any writable field that arrives empty (or not at all) is flagged orange by the script itself with
  a "חסר" comment — leaving a field out is never a way to skip it silently. `סימון` is the only
  exception (it's a project-specific code, blank by design).
- Column position is never assumed — the script always re-reads row 8 of the actual file it's
  given. This also means a corrupted or oddly-worded header cell (seen in the studio's own
  base template — the `סעיף` header cell had stray leftover text appended) is tolerated via
  prefix/substring matching rather than exact-match, as long as it's unambiguous.

