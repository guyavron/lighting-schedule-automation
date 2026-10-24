#!/usr/bin/env python3
"""
Compare one candidate fixture (as it's about to be / was just inserted into a
project's מקרא) against every row already in the fixture bank, using
approximate matching — real project files and PDFs rarely word a תיאור
identically twice (extra spaces, "מק"ט" vs "דגם", slightly different line
order), so an exact-string or exact-model-key comparison (like the one
migrate_mikra_to_bank.py uses for its own bulk-duplicate merge) is too brittle
for a live per-fixture check. This script is intentionally a separate module
rather than a change to that exact-key logic, so the tested bulk-migration
path is not touched.

Verdict bands (tuned, not derived from a formula — adjust the two constants
below if real usage shows too many/few "maybe"s):
  strong  >= STRONG_THRESHOLD  -> already in the bank, no action
  maybe   >= MAYBE_THRESHOLD   -> flag for a human to check manually
  new     <  MAYBE_THRESHOLD   -> not found, candidate for the bank

Usage (single candidate vs the whole bank):
    python3 match_fixture.py <bank.xlsx> --manufacturer "VELTRANO" --description "..."

Prints one JSON object to stdout:
    {"verdict": "strong"|"maybe"|"new", "score": 0.91,
     "bank_row": 42, "bank_manufacturer": "...", "bank_description": "..."}
"""
import argparse
import json
import re
import sys
from difflib import SequenceMatcher

import openpyxl

sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
import migrate_mikra_to_bank as M  # noqa: E402  (reuse find_bank_header_row, BANK_COLUMNS, col_index)

# עברית בפלט גם בווינדוס, כשהפלט נקרא דרך pipe (ברירת המחדל שם היא cp1252)
for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")


STRONG_THRESHOLD = 0.80
# Raised from 0.45 to 0.60 after a check on the real bank: every genuinely new
# fixture from a manufacturer that was already in the bank scored 0.51-0.55 (the
# shared brand name alone is worth ~0.25), while truly similar pairs scored
# 0.71-0.79. At 0.45 every new fixture landed in "maybe" and none was added.
# 0.60 sits in the gap between the two groups.
MAYBE_THRESHOLD = 0.60

_NIQQUD = re.compile(r"[֑-ׇ]")  # Hebrew diacritics, if any slip in
_PUNCT = re.compile(r"[\"'׳״״׳.,;:()\[\]{}\-–—/\\]")
_WS = re.compile(r"\s+")

# Every fixture description in this studio's style always includes this exact
# fixed driver line (see mikra SKILL.md step 2) — left in, two unrelated
# fixtures would always share a long common substring and inflate a plain
# SequenceMatcher score regardless of the actual product, so it's stripped
# before the full-text similarity is computed (the model-key comparison below
# is the real discriminator; full-text is only a secondary signal).
_BOILERPLATE = [
    "דרייברים למיקום באזור פריק ונגיש",
]


def fold(text):
    """Normalize for comparison: strip niqqud/punctuation, collapse
    whitespace, lowercase (affects only the Latin-script portions, e.g. model
    codes — Hebrew has no case)."""
    t = _NIQQUD.sub("", str(text or ""))
    t = _PUNCT.sub(" ", t)
    t = _WS.sub(" ", t).strip().lower()
    return t


def strip_boilerplate(text):
    t = str(text or "")
    for phrase in _BOILERPLATE:
        t = t.replace(phrase, " ")
    return t


def similarity(a, b):
    a, b = fold(a), fold(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


# Where the model NAME ends. After "מדגם <name>" every description continues with
# the same kind of spec line ("LED 10W 980LM 3000K CRI90 ..."), so a key that runs
# past the name compares spec boilerplate, not products: a floodlight "Beam 30"
# scored 0.82 against an in-ground "Terra 60" of the same brand ("strong" = "already
# in the bank"), and the new fixture would never have been added. Bank rows may
# be stored on one line, so a line break alone can't be the only stop sign.
_MODEL_STOP = re.compile(r"\n|\s(?:LED|PENDANT|מותאם|אורך|זוית|כולל|דרייברים)\b|\s\d+(?:[.,]\d+)?W\b", re.I)
_NUMBER = re.compile(r"\d+(?:[.,]\d+)?")


def extract_model_name(description):
    text = str(description or "")
    m = re.search(r"מדגם\s+", text)
    if not m:
        return M.extract_model_key(text)
    rest = text[m.end():]
    stop = _MODEL_STOP.search(rest)
    return M.normalize(rest[:stop.start()] if stop else rest)[:40].lower()


def model_numbers_differ(desc_a, desc_b):
    """True when both model names carry numbers and the number lists differ
    ("Kestra 2642" vs "Kestra 2632"). Letter-by-letter similarity rates such
    names as near-identical (~0.95), but in a catalogue a different number is
    usually a different product. Such a pair may never be "strong"."""
    na = _NUMBER.findall(extract_model_name(desc_a))
    nb = _NUMBER.findall(extract_model_name(desc_b))
    return bool(na) and bool(nb) and na != nb


def model_numbers_conflict(desc_a, desc_b):
    """Like model_numbers_differ, but only when neither number list is contained
    in the other.

      "Nexa 35" vs "Nexa 50"                         -> conflict (two products)
      "Orbo Floor 2" vs "Orbo Floor 2 F7741020"    -> no conflict: the same
          product with an extra part number, so it stays "maybe" for a person.
    """
    na = _NUMBER.findall(extract_model_name(desc_a))
    nb = _NUMBER.findall(extract_model_name(desc_b))
    if not na or not nb:
        return False
    sa, sb = set(na), set(nb)
    return not (sa <= sb or sb <= sa)


def score_pair(cand_manufacturer, cand_description, bank_manufacturer, bank_description):
    """Manufacturer must be a reasonable match on its own (a totally
    different brand should never score high just because the description
    text happens to overlap) — combined as a weighted blend. The model
    identifier (the token after "מדגם", same extraction migrate_mikra_to_bank
    uses for its own dedup) is the real discriminator between two fixtures of
    the same type/brand-family, so it carries the most weight; full-text
    similarity (boilerplate stripped) is only a secondary signal."""
    mfr_sim = similarity(cand_manufacturer, bank_manufacturer)
    if mfr_sim < 0.5:
        # different-enough manufacturer caps the score hard, regardless of
        # how much spec boilerplate the descriptions happen to share
        desc_sim = similarity(strip_boilerplate(cand_description), strip_boilerplate(bank_description))
        return desc_sim * 0.25

    model_sim = similarity(extract_model_name(cand_description), extract_model_name(bank_description))
    desc_sim = similarity(strip_boilerplate(cand_description), strip_boilerplate(bank_description))
    score = mfr_sim * 0.25 + model_sim * 0.55 + desc_sim * 0.20
    if model_numbers_conflict(cand_description, bank_description):
        return min(score, MAYBE_THRESHOLD - 0.01)      # different model number = a different product: "new"
    if score >= STRONG_THRESHOLD and model_numbers_differ(cand_description, bank_description):
        score = STRONG_THRESHOLD - 0.01                # at most "maybe": a person checks
    return score


def find_match(manufacturer, description, bank_path, sheet_name=M.BANK_SHEET_NAME):
    wb = M.open_workbook(bank_path, sheet=sheet_name, data_only=True)
    ws = wb[sheet_name]
    header_row = M.find_bank_header_row(ws)
    if header_row is None:
        raise ValueError("Could not find the bank's column-header row (expected קטגוריה/יצרן/תיאור).")

    best = None
    for r in range(header_row + 1, ws.max_row + 1):
        bank_mfr = ws.cell(row=r, column=M.col_index("יצרן")).value
        bank_desc = ws.cell(row=r, column=M.col_index("תיאור")).value
        if not bank_mfr and not bank_desc:
            continue
        score = score_pair(manufacturer, description, bank_mfr, bank_desc)
        if best is None or score > best["score"]:
            best = {
                "score": score, "bank_row": r,
                "bank_manufacturer": M.normalize(bank_mfr),
                "bank_description": M.normalize(bank_desc),
            }

    if best is None:
        return {"verdict": "new", "score": 0.0, "bank_row": None,
                "bank_manufacturer": "", "bank_description": ""}

    if best["score"] >= STRONG_THRESHOLD:
        verdict = "strong"
    elif best["score"] >= MAYBE_THRESHOLD:
        verdict = "maybe"
    else:
        verdict = "new"

    return {"verdict": verdict, "score": round(best["score"], 3),
            "bank_row": best["bank_row"],
            "bank_manufacturer": best["bank_manufacturer"],
            "bank_description": best["bank_description"]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("bank_xlsx")
    ap.add_argument("--manufacturer", required=True)
    ap.add_argument("--description", required=True)
    args = ap.parse_args()
    result = find_match(args.manufacturer, args.description, args.bank_xlsx)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    try:
        main()
    except (M.WorkbookError, ValueError) as e:
        raise SystemExit(f"שגיאה: {e}")
