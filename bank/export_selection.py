#!/usr/bin/env python3
"""
Export selected fixtures from בנק גופי תאורה (fixture bank) into a project
מקרא workbook — Stage 3 v1 ("ממשק בחירה"), extended in Stage 6 to also
support appending into a project that's already in progress (not only
building a brand-new one from the empty skeleton).

Selection happens in one of two ways:
  - marking the "בחירה לפרויקט" column directly in the bank file — any
    non-empty value in that column counts as marked (e.g. "X"). This script
    does NOT do the marking; it reads marks the person already made.
  - --rows "10,13,47" — an explicit list of bank row numbers, which is what
    the web selection page (Artifact) produces. In this mode the בחירה
    column is not read at all and is left untouched.

Two modes:
  - Default (new project): copies the empty skeleton (legend_template.xlsx) and
    sets the project header (name / light colour / date) via mikra's own
    set_project_header() — same function stage 1 uses.
  - --append-to <existing מקרא.xlsx> (Stage 6): skips the skeleton and the
    header entirely — the project already has both — and inserts the newly
    marked fixtures straight into that existing file. --name/--color are
    ignored in this mode (there is nothing to set a header for).

Either way, every fixture is placed into the correct category block using
mikra's own insert_fixture() — the exact same tested row-insertion /
merged-header-shifting / image-anchoring logic used when a fixture is added
fresh from a PDF, so an exported row is structurally identical to one mikra
would produce (aside from image resolution — see note below). This is also
why the reverse direction is safe: a project file built or extended by this
script is just a regular מקרא file as far as mikra/insert_fixture.py is
concerned, so PDF-driven fixtures keep landing correctly in it afterwards
(verified — see CHANGELOG.md, Stage 6).

On success, clears the בחירה marks in the bank for the rows that were just
exported (so the column is ready for the next project) and saves the bank
back. Pass --keep-marks to skip this.

Image quality (decided 23.08.2026 — see CHANGELOG.md): bank images are
deliberately small (<=130x110px, JPEG q78) to keep the bank file light. This
script reuses them as-is; there is no separate full-resolution archive. If
this proves too low-quality in practice for a real exported legend, revisit.

סימון (fixture tag) is intentionally left BLANK in every exported row, even
when the bank row has one — per mikra/SKILL.md, סימון is a studio-assigned
code for ONE specific project, not a fact about the fixture itself. Carrying
a tag over from whichever old project happened to be the bank's source would
be actively misleading in the new project, not just stale.

Usage (new project):
    python3 export_selection.py <bank_xlsx> \
        --name "שם הפרויקט" [--color "3000K"] \
        [--skeleton /path/to/legend_template.xlsx] \
        [--out /path/to/מקרא\ -\ שם\ הפרויקט.xlsx] \
        [--bank-out /path/to/updated_bank.xlsx] \
        [--keep-marks]

Usage (append to an existing, already-in-progress project):
    python3 export_selection.py <bank_xlsx> \
        --append-to /path/to/מקרא\ -\ פרויקט\ קיים.xlsx \
        [--out /path/to/output.xlsx] \
        [--bank-out /path/to/updated_bank.xlsx] \
        [--keep-marks]

If --skeleton is omitted, it defaults to "legend_template.xlsx" next to the bank
file. If --out is omitted: in new-project mode it defaults to
"מקרא - <name>.xlsx" next to the bank file (requires --name); in --append-to
mode it defaults to updating the existing project file in place. If
--bank-out is omitted, the bank file is updated in place (same path it was
read from).
"""
import argparse
from bank_guard import bank_guard
import re
import datetime
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "skill" / "scripts"))
from xlsx_io import WorkbookError, open_workbook, replace_file, save_workbook  # noqa: E402
from openpyxl.styles import Alignment

sys.path.insert(0, str(Path(__file__).resolve().parent))
import migrate_mikra_to_bank as M

MIKRA_SCRIPTS = Path(__file__).resolve().parent.parent / "skill" / "scripts"
sys.path.insert(0, str(MIKRA_SCRIPTS))
import insert_fixture as IF
import set_project_header as SPH

# עברית בפלט גם בווינדוס, כשהפלט נקרא דרך pipe (ברירת המחדל שם היא cp1252)
for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")


SELECTION_HEADER = "בחירה לפרויקט"
CATEGORY_HE_TO_EN = {"תקרה": "ceiling", "קיר": "wall", "רצפה": "floor"}


def ensure_selection_column(ws, header_row):
    """Find the בחירה לפרויקט column by header text (on the bank's actual
    column-header row, wherever that is); add it as a new last column
    (styled like the other headers) if it isn't there yet. Returns the
    1-indexed column number."""
    for col in range(1, ws.max_column + 1):
        if M.normalize(ws.cell(row=header_row, column=col).value) == SELECTION_HEADER:
            return col

    col = ws.max_column + 1
    cell = ws.cell(row=header_row, column=col, value=SELECTION_HEADER)
    from openpyxl.styles import Font, PatternFill, Border, Side
    cell.font = Font(name="Arial", size=10)
    cell.border = Border(*(Side(style="thin"),) * 4)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    from openpyxl.utils import get_column_letter
    ws.column_dimensions[get_column_letter(col)].width = 16
    last_row = max(ws.max_row, header_row)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(col)}{last_row}"
    return col


def read_marked_fixtures(bank_path, only_rows=None):
    """Returns a plain list of fixture dicts (all values already extracted as
    python primitives / raw image bytes). Deliberately does NOT keep the
    workbook object alive past this function; clearing marks later
    (clear_marks_in_bank) re-opens the bank fresh instead of reusing this
    object."""
    wb = open_workbook(bank_path, sheet=M.BANK_SHEET_NAME)
    ws = wb[M.BANK_SHEET_NAME]
    header_row = M.find_bank_header_row(ws)
    if header_row is None:
        raise ValueError("Could not find the bank's column-header row (expected קטגוריה/יצרן/תיאור).")
    selection_col = ensure_selection_column(ws, header_row)
    img_map = M.images_by_anchor_row(ws)  # 0-indexed anchor row -> Image

    fixtures = []
    for r in range(header_row + 1, ws.max_row + 1):
        if only_rows is not None:
            # explicit row list (e.g. coming from the web selection page) —
            # the בחירה column is ignored entirely in this mode.
            if r not in only_rows:
                continue
        else:
            mark = M.normalize(ws.cell(row=r, column=selection_col).value)
            if not mark:
                continue
        vals = {c: (M.normalize_text if c in M.MULTILINE_FIELDS else M.normalize)(
                    ws.cell(row=r, column=M.col_index(c)).value) for c, _ in M.BANK_COLUMNS}
        if not vals.get("קטגוריה") or not vals.get("תיאור") or not vals.get("יצרן"):
            continue  # shouldn't happen for a bank row, but don't crash the export over it
        img = img_map.get(r - 1)
        fixtures.append({
            "row": r,
            "category": vals["קטגוריה"],
            "יצרן": vals["יצרן"],
            "תיאור": vals["תיאור"],
            "גוון": vals["גוון"],
            "מתח_עמעום": vals["מתח_עמעום"],
            "image_bytes": img._data() if img else None,
        })
    # if a fresh header was just added but nothing ended up marked, still
    # persist the new column so it's there for next time
    if not fixtures:
        save_workbook(wb, bank_path)
    # openpyxl workbooks with embedded images form reference cycles (cell <->
    # worksheet <-> image <-> anchor), so the file handle this load_workbook()
    # call opened is NOT guaranteed to be released just because `wb` goes out
    # of scope here -- CPython's cyclic GC may not run before the caller opens
    # the SAME path again a moment later (clear_marks_in_bank), and that
    # second open+save was observed to silently no-op while this one was
    # still alive. Close explicitly rather than relying on GC timing.
    wb.close()
    return fixtures


def clear_marks_in_bank(bank_path, bank_out_path, rows):
    """Re-opens the bank fresh (see read_marked_fixtures docstring for why),
    blanks the בחירה לפרויקט cell for the given row numbers, and saves."""
    wb = open_workbook(bank_path, sheet=M.BANK_SHEET_NAME)
    ws = wb[M.BANK_SHEET_NAME]
    header_row = M.find_bank_header_row(ws)
    selection_col = ensure_selection_column(ws, header_row)
    for r in rows:
        # ws.cell(..., value=None) is a silent no-op in openpyxl -- None is
        # its sentinel for "no value given", not "clear this cell". Must set
        # .value on the Cell object directly to actually blank it.
        ws.cell(row=r, column=selection_col).value = None
    save_workbook(wb, bank_out_path)


DATE_TAIL_RE = re.compile(r"(?P<sep>[-–—]?\s*)(?P<date>\d{1,2}\.\d{1,2}\.\d{2,4})\s*$")


def dated_version_name(existing_path, today=None):
    """שם הגרסה החדשה של מקרא קיים: אותו שם, עם התאריך של היום במקום התאריך הישן.

    "ניסוי 20.08.26.xlsx"      -> "ניסוי 28.08.26.xlsx"      (התאריך הישן יורד, לא מצטבר)
    "מקרא - בית לדוגמה - 12.3.25" -> "מקרא - בית לדוגמה - 28.08.26" (המפריד שהיה לפני התאריך נשמר)
    "מקרא - ניסוי.xlsx"        -> "מקרא - ניסוי 28.08.26.xlsx" (אין תאריך קודם — מתווסף)
    """
    p = Path(existing_path)
    stem = p.stem
    today = today or datetime.date.today().strftime("%d.%m.%y")
    m = DATE_TAIL_RE.search(stem)
    if m:
        stem = stem[:m.start()] + m.group("sep") + today
    else:
        stem = stem.rstrip() + " " + today
    return p.with_name(stem + p.suffix)


def build_export(fixtures, out_path, tmpdir, name=None, color=None, set_header=True):
    """Places every fixture into out_path via insert_fixture(). Assumes
    out_path already exists (either a freshly-copied skeleton, or a copy of
    an existing in-progress project — main() is responsible for getting the
    right starting file to out_path before calling this). Only sets the
    project header when set_header is True (new-project mode) — an
    already-in-progress project keeps its existing header untouched."""
    if set_header:
        SPH.set_project_header(str(out_path), project_name=name, light_color=color)

    report = {"placed": [], "skipped_no_image": [], "skipped_unknown_category": []}
    for fx in fixtures:
        category_en = CATEGORY_HE_TO_EN.get(fx["category"])
        if not category_en:
            report["skipped_unknown_category"].append(fx)
            continue
        if not fx["image_bytes"]:
            report["skipped_no_image"].append(fx)
            # still place the row — a missing photo shouldn't drop a whole
            # fixture from the export — just without an image embedded.
        data = {
            "תיאור": fx["תיאור"],
            "יצרן": fx["יצרן"],
            "מתח_עמעום": fx["מתח_עמעום"],
            "גוון": fx["גוון"],
        }
        image_path = None
        if fx["image_bytes"]:
            image_path = str(Path(tmpdir) / f"row{fx['row']}.jpg")
            Path(image_path).write_bytes(fx["image_bytes"])
        IF.insert_fixture(str(out_path), category_en, data, image_path=image_path, out_path=str(out_path))
        report["placed"].append(fx)
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("bank_xlsx")
    ap.add_argument("--name", default=None, help="שם הפרויקט")
    ap.add_argument("--color", default=None, help='גוון תאורה בפרויקט, e.g. "3000K"')
    ap.add_argument("--skeleton", default=None)
    ap.add_argument("--append-to", default=None,
                     help="קובץ מקרא קיים (פרויקט שכבר בעבודה) לצרף אליו את הבחירה, "
                          "במקום לבנות מקרא חדש מהשלד הריק")
    ap.add_argument("--out", default=None)
    ap.add_argument("--bank-out", default=None)
    ap.add_argument("--keep-marks", action="store_true")
    ap.add_argument("--in-place", action="store_true",
                    help="במצב --append-to: לעדכן את הקובץ הקיים עצמו במקום ליצור גרסה חדשה מתוארכת")
    ap.add_argument("--rows", default=None,
                    help="רשימת שורות בנק מפורשת לייצוא, מופרדת בפסיקים (למשל \"10,13,47\"). "
                         "במצב הזה עמודת 'בחירה לפרויקט' לא נקראת ולא נוקה.")
    args = ap.parse_args()

    bank_path = Path(args.bank_xlsx)

    if args.append_to:
        append_path = Path(args.append_to)
        if not append_path.exists():
            sys.exit(f"קובץ המקרא לצירוף לא נמצא: {append_path}")
        if args.name or args.color:
            print("שים לב: --name/--color מתעלמים עם --append-to — כותרת הפרויקט כבר קיימת "
                  "בקובץ ולא משתנה.")
        if args.skeleton:
            print("שים לב: --skeleton מתעלם עם --append-to.")
        if args.out:
            out_path = Path(args.out)
        elif args.in_place:
            out_path = append_path
        else:
            # ברירת המחדל: גרסה חדשה של אותו מקרא, עם התאריך של היום במקום התאריך הישן
            out_path = dated_version_name(append_path)
            if out_path.exists():
                print(f"שים לב: כבר קיימת גרסה בשם {out_path.name} — היא תתעדכן.")
        copy_from = append_path if out_path != append_path else None
        set_header = False
    else:
        skeleton_path = Path(args.skeleton) if args.skeleton else bank_path.parent / "legend_template.xlsx"
        if not skeleton_path.exists():
            sys.exit(f"Skeleton not found: {skeleton_path}")

        if args.out:
            out_path = Path(args.out)
        else:
            if not args.name:
                sys.exit("--out omitted and --name not given — can't build a default output filename.")
            out_path = bank_path.parent / f"מקרא - {args.name}.xlsx"

        copy_from = skeleton_path
        set_header = True

    bank_out_path = Path(args.bank_out) if args.bank_out else bank_path

    only_rows = None
    if args.rows:
        try:
            only_rows = {int(x) for x in args.rows.replace(" ", "").split(",") if x}
        except ValueError:
            sys.exit(f"--rows לא תקין: {args.rows!r} — צפוי מספרי שורות מופרדים בפסיקים, למשל 10,13,47. לא נוצר כלום.")
        if not only_rows:
            sys.exit("--rows ריק — לא נוצר כלום.")

    fixtures = read_marked_fixtures(bank_path, only_rows=only_rows)
    if not fixtures:
        if only_rows is not None:
            sys.exit(f"אף אחת מהשורות שהתבקשו לא נמצאה בבנק ({sorted(only_rows)}). לא נוצר כלום.")
        sys.exit("לא נמצאו שורות מסומנות בעמודת 'בחירה לפרויקט' בבנק. לא נוצר כלום.")

    if only_rows is not None:
        missing = sorted(only_rows - {fx["row"] for fx in fixtures})
        if missing:
            print(f"אזהרה: שורות שהתבקשו ולא נמצאו/לא תקינות בבנק ודולגו: {missing}")

    # Build in a scratch copy and swap it in only at the end: if fixture 3 of 5
    # fails, there must be no half-built legend sitting where the real one goes.
    with tempfile.TemporaryDirectory() as tmpdir:
        work_path = Path(tmpdir) / "work.xlsx"
        shutil.copy(copy_from if copy_from is not None else out_path, work_path)
        report = build_export(fixtures, work_path, tmpdir, name=args.name, color=args.color,
                               set_header=set_header)
        replace_file(work_path, out_path)
    if copy_from is not None and set_header is False:
        print(f"נוצרה גרסה חדשה: {out_path.name}  (מתוך {copy_from.name})")

    if only_rows is None and not args.keep_marks:
        exported_rows = [fx["row"] for fx in report["placed"] + report["skipped_no_image"]]
        with bank_guard(bank_path, "ניקוי סימוני בחירה בבנק"):
            clear_marks_in_bank(bank_path, bank_out_path, exported_rows)

    by_cat = {}
    for fx in report["placed"]:
        by_cat.setdefault(fx["category"], 0)
        by_cat[fx["category"]] += 1

    if args.append_to:
        print(f"נוספו גופים לקובץ מקרא קיים: {out_path}")
    else:
        print(f"נוצר מקרא חדש: {out_path}")
        print(f"שם פרויקט: {args.name or '[לא סופק — מסומן כתום בקובץ]'} | גוון תאורה: {args.color or '[לא סופק — מסומן כתום בקובץ]'}")
    print(f"הוכנסו {len(report['placed'])} גופים:")
    for cat, n in sorted(by_cat.items()):
        print(f"  {cat}: {n}")
    if report["skipped_no_image"]:
        print(f"הוכנסו בלי תמונה (לא נמצאה תמונה בבנק) — {len(report['skipped_no_image'])}:")
        for fx in report["skipped_no_image"]:
            print(f"  שורת בנק {fx['row']}: {fx['יצרן']} — {fx['תיאור'][:40].replace(chr(10), ' · ')}")
    if report["skipped_unknown_category"]:
        print(f"דולגו — קטגוריה לא מזוהה — {len(report['skipped_unknown_category'])}:")
        for fx in report["skipped_unknown_category"]:
            print(f"  שורת בנק {fx['row']}: קטגוריה='{fx['category']}' {fx['יצרן']} — {fx['תיאור'][:40].replace(chr(10), ' · ')}")
    print("\nלתשומת לב: שדה סימון הושאר ריק בכל השורות שיוצאו (זה קוד ייעודי לפרויקט הקודם, לא רלוונטי כאן) — יש למלא ידנית.")
    if only_rows is not None:
        print("הייצוא נעשה לפי רשימת שורות מפורשת (--rows) — עמודת הבחירה בבנק לא נקראה ולא שונתה.")
    elif not args.keep_marks:
        print(f"סימוני הבחירה בבנק נוקו ({bank_out_path}) — מוכן לפרויקט הבא.")
    else:
        print("סימוני הבחירה בבנק נשארו כפי שהיו (--keep-marks).")


if __name__ == "__main__":
    try:
        main()
    except (M.WorkbookError, ValueError) as e:
        raise SystemExit(f"שגיאה: {e}")
