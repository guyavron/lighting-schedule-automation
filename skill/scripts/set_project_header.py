#!/usr/bin/env python3
"""
Set the project-header cells of the lighting legend workbook: project name,
build date, and the project-wide light colour temperature (e.g. 2700K/3000K).

The date is always today's date. Project name and light colour are two
required questions the skill (Claude) must ask the person before running
this script, if they weren't already given — this script never guesses them.
If either is still missing when this runs, its cell is left highlighted
orange with a comment, exactly like the studio's existing "flag it, don't
invent it" convention for uncertain fixture data.

Usage:
    python3 set_project_header.py <xlsx_path> [--name "שם הפרויקט"] \
        [--color "2700K"] [--out <path>]
"""
import argparse
import datetime
import sys

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

from xlsx_io import WorkbookError, open_workbook, save_workbook

# Hebrew output on Windows too, when stdout is a pipe (the default there is cp1252)
for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")

SHEET_NAME = "אומדן"
UNCERTAIN_FILL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")


def _set_flagged(ws, coord, label_prefix, value, missing_comment):
    cell = ws[coord]
    if value:
        cell.value = f"{label_prefix}{value}"
        cell.fill = PatternFill(fill_type=None)
        cell.comment = None
    else:
        # leave existing text (or blank it if this is a fresh skeleton) and flag it
        if not cell.value or label_prefix.strip(": ") in str(cell.value):
            cell.value = label_prefix
        cell.fill = UNCERTAIN_FILL
        cell.comment = Comment(missing_comment, "mikra skill")


def set_project_header(xlsx_path, project_name=None, light_color=None, out_path=None, build_date=None):
    wb = open_workbook(xlsx_path, sheet=SHEET_NAME)
    ws = wb[SHEET_NAME]

    build_date = build_date or datetime.date.today()
    ws["C3"] = f"תאריך: {build_date.strftime('%d.%m.%y')}"

    _set_flagged(ws, "C2", "שם הפרויקט: ", project_name,
                 "שם הפרויקט לא הוזן, נא להשלים ידנית.")
    _set_flagged(ws, "C5", "גוון תאורה בפרוייקט: ", light_color,
                 "גוון תאורה בפרויקט לא הוזן, נא להשלים ידנית.")

    out_path = out_path or xlsx_path
    save_workbook(wb, out_path)
    print(f"Header set. Date: {build_date.strftime('%d.%m.%y')} | "
          f"Project name: {project_name or '[NOT PROVIDED - flagged]'} | "
          f"Light colour: {light_color or '[NOT PROVIDED - flagged]'} | Saved to {out_path}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx_path")
    ap.add_argument("--name", default=None)
    ap.add_argument("--color", default=None, help='e.g. "2700K" or "3000K"')
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    try:
        set_project_header(args.xlsx_path, args.name, args.color, args.out)
    except WorkbookError as e:
        raise SystemExit(f"שגיאה: {e}")


if __name__ == "__main__":
    main()
