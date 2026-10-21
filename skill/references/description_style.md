# Style reference for the "תיאור" field

Built from the studio's own past legend descriptions plus the official structure of its base
template, and confirmed with Guy field by field — not invented. (In this public repository the
examples below are fictional: same structure, invented model names.)

## Overall tone

Short newline-separated phrase fragments, not full grammatical sentences. No periods at the end
of lines. Telegraphic / spec-sheet style, matching how the studio's own past legends read.

## Line order

1. **[סוג הגוף] מדגם [שם הדגם]** — fixture type in Hebrew, then `מדגם` (model), then the model
   name/number in English exactly as the manufacturer writes it (never translated, never
   reworded). A beam angle or diameter that's part of the model's own name can sit on this line
   too (e.g. `ספוט מדגם Kito 24°`, `שקוע תקרה ללא שוליים מדגם Arca Mini Ø55`).
2. *(optional)* a second context line — the system/track it belongs to, or the English product
   category repeated (`PENDANT`), when the PDF makes a point of it.
3. **LED spec line**: `LED [הספק]W [שטף אור]LM [גוון]K CRI[ערך] IP[דירוג]` — combine only the
   values the PDF actually states; drop any part it doesn't give (many real rows are missing IP
   or CRI). For linear/profile fixtures priced per meter, use the `\M` suffix on wattage and
   lumens (`LED 20W\M 2000LM\M`), matching how the studio already writes per-meter figures.
4. *(optional)* **זוית פיזור [X] מעלות** — only when the PDF gives one clear beam angle and it
   isn't already folded into the model line above. If the PDF offers more than one beam-angle
   option, this is a case for the "ask, don't guess" rule in SKILL.md — never write two options
   into one line.
5. *(optional)* trailing notes, one per line, only when the PDF/context supports them — seen in
   real rows: `אורך [בס"מ] עפ"י תוכנית` (length is project-specific, for profiles), `כולל מתאם
   לבית חכם` / `כולל מתאם עבור בית חכם` (smart-home adapter), `כולל רוזטות שקועות משפוכטלות`,
   `כולל תותב גבס מקורי של היצרן`, `כולל נורת [דגם נורה]`, `גודל סופי לבחירה בהמשך`. These can
   land before or after the driver line (line 6) depending on what reads naturally — real past
   rows do it both ways.
6. **דרייברים למיקום באזור פריק ונגיש** — fixed wording, always included in every fixture
   regardless of type, even ones where past legend files sometimes forgot it (confirmed with Guy:
   that was a past omission, not a rule to keep following).

## Explicitly NOT part of the תיאור text

- **Finish colour** — goes only in the separate **גוון** column, never repeated inside the
  description text (even though the official skeleton's placeholder text shows a `גוון הגוף -
  [צבע גימור]` line — Guy confirmed to drop that and keep colour column-only).
- **Dimension/site-coordination note** (the skeleton's `[הערת מידות/מיקום מיוחדת אם יש, למשל
  גובה הגוף לתיאום מול השטח]` placeholder) — this is filled in by the studio's own staff by hand
  later, never generated or guessed by the skill. Leave it out entirely; don't leave a bracketed
  placeholder for it either.

## Fixtures with no manufacturer chosen yet (TBD)

When the studio is still deciding on a fixture (no PDF yet), the whole entry collapses to one
short line instead of a full spec, e.g. `גוף תאורה תלוי דקורטיבי לבחירה בהמשך עפ"י אדריכל` or
`גוף תאורה לבחירה בהמשך`. This skill's job is always to write *from a PDF*, so this pattern is
here mainly so Claude recognizes an existing TBD row and doesn't try to "complete" it with
invented specs.

## Examples (fictional — same structure as the studio's real rows)

```
ספוט מדגם Kito 24°
מותאם למערכת Linea 48
LED 10W 980LM 2700K CRI95 IP20
דרייברים למיקום באזור פריק ונגיש
```

```
פרופיל לד שקוע ללא שוליים מדגם Strada 20
LED 18W\M 1850LM\M 3000K CRI90 IP20
דרייברים למיקום באזור פריק ונגיש
כולל תותב גבס מקורי של היצרן
```

(Note: some historical rows used "יבש ונגיש" for the driver area — per Guy, the wording is now
always fixed to "פריק ונגיש" regardless of what the PDF implies about the installation area.)

```
גוף תאורה תלוי מדגם Orbe
PENDANT
LED 2X4.5W 760LM 2700K
כולל מתאם עבור בית חכם
דרייברים למיקום באזור פריק ונגיש
כולל רוזטות שקועות משפוכטלות
```
