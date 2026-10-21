"""
Failure-mode tests: what happens when the input is wrong, not only when it's right.

The rule they enforce: a visible error (or an orange cell) is acceptable; a silent
wrong result -- an invented value, an empty cell that looks finished, a half-written
file -- is not.

    python -m unittest discover -s tests -v
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skill" / "scripts"))
sys.path.insert(0, str(ROOT / "bank"))
sys.path.insert(0, str(ROOT / "examples"))

import openpyxl                              # noqa: E402
from openpyxl.cell.rich_text import CellRichText  # noqa: E402

import insert_fixture as IF                  # noqa: E402
import xlsx_io                               # noqa: E402
import make_template                         # noqa: E402
import draw_products                         # noqa: E402
import migrate_mikra_to_bank as M            # noqa: E402
import match_fixture as MF                   # noqa: E402
import build_bank_template as BBT            # noqa: E402
from bank_guard import bank_guard            # noqa: E402

SPECS = ROOT / "examples" / "spec-sheets"
PY = sys.executable
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}

FX = {"תיאור": "ספוט שקוע מדגם Testa 10\nLED 8W 700LM 3000K CRI90\nדרייברים למיקום באזור פריק ונגיש",
      "יצרן": "VELTRANO", "מתח_עמעום": "350mA\nDIM DALI", "גוון": "לבן"}


def setUpModule():
    if not (SPECS / "broken" / "truncated.pdf").exists():
        subprocess.run([PY, str(ROOT / "examples" / "make_sample_data.py")], check=True,
                       capture_output=True, env=ENV)


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def run(*args):
    r = subprocess.run([PY, *map(str, args)], cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=ENV)
    return r.returncode, r.stdout + r.stderr


def orange(cell):
    return cell.fill.fill_type == "solid" and str(cell.fill.fgColor.rgb).upper().endswith("FFC000")


class Tmp(unittest.TestCase):
    def setUp(self):
        self.d = Path(tempfile.mkdtemp())
        self.legend = make_template.build(self.d / "legend.xlsx")
        self.photo = draw_products.draw("downlight", "לבן", self.d / "p.jpg")

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def sheet(self, path=None):
        return openpyxl.load_workbook(path or self.legend, rich_text=True)[IF.SHEET_NAME]

    def cell(self, ws, row, field):
        return ws.cell(row=row, column=IF.build_header_map(ws)[field])


# ------------------------------------------------------------------ reading PDFs
class TestReadPdf(Tmp):
    def test_good_pdf_returns_text(self):
        code, out = run("skill/scripts/read_pdf.py", SPECS / "corvella_track-one-15.pdf", "--out", self.d / "x")
        self.assertEqual(code, 0)
        self.assertIn("Luminous flux", out)

    def test_scanned_pdf_is_a_loud_error_not_empty_success(self):
        code, out = run("skill/scripts/read_pdf.py", SPECS / "broken" / "scanned_no-text-layer.pdf", "--out", self.d / "x")
        self.assertEqual(code, 5)
        self.assertIn("no text layer", out)
        self.assertTrue(list((self.d / "x").glob("page*.jpg")), "page images should still be rendered for a human")

    def test_truncated_pdf(self):
        code, out = run("skill/scripts/read_pdf.py", SPECS / "broken" / "truncated.pdf", "--out", self.d / "x")
        self.assertEqual(code, 4)
        self.assertIn("damaged", out)
        self.assertNotIn("Traceback", out)

    def test_not_a_pdf(self):
        code, out = run("skill/scripts/read_pdf.py", SPECS / "broken" / "not-really-a.pdf", "--out", self.d / "x")
        self.assertEqual(code, 3)
        self.assertNotIn("Traceback", out)

    def test_missing_file(self):
        code, _ = run("skill/scripts/read_pdf.py", self.d / "nope.pdf")
        self.assertEqual(code, 2)


# ------------------------------------------------------------- writing a legend row
class TestInsertFixture(Tmp):
    def test_happy_path_fields_and_image(self):
        row = IF.insert_fixture(str(self.legend), "ceiling", FX, image_path=str(self.photo))
        ws = self.sheet()
        self.assertEqual(self.cell(ws, row, "יצרן").value, "VELTRANO")
        self.assertEqual(len(ws._images), 1)
        self.assertFalse(any(orange(self.cell(ws, row, f)) for f in ("יצרן", "גוון", "תיאור")))

    def test_empty_field_is_flagged_not_left_looking_finished(self):
        row = IF.insert_fixture(str(self.legend), "ceiling", {**FX, "גוון": ""}, image_path=str(self.photo))
        c = self.cell(self.sheet(), row, "גוון")
        self.assertTrue(orange(c))
        self.assertIn("חסר", c.comment.text)

    def test_field_not_given_at_all_is_flagged(self):
        data = {k: v for k, v in FX.items() if k != "מתח_עמעום"}
        row = IF.insert_fixture(str(self.legend), "wall", data, image_path=str(self.photo))
        self.assertTrue(orange(self.cell(self.sheet(), row, "מתח_עמעום")))

    def test_missing_image_is_flagged(self):
        row = IF.insert_fixture(str(self.legend), "floor", FX)
        self.assertTrue(orange(self.cell(self.sheet(), row, "תמונה")))

    def test_snippet_not_found_flags_whole_cell(self):
        row = IF.insert_fixture(str(self.legend), "ceiling", FX, image_path=str(self.photo),
                                uncertain_snippets={"תיאור": "זוית פיזור 72 מעלות"})
        self.assertTrue(orange(self.cell(self.sheet(), row, "תיאור")))

    def test_unknown_category_changes_nothing(self):
        before = sha(self.legend)
        with self.assertRaises(ValueError):
            IF.insert_fixture(str(self.legend), "roof", FX)
        self.assertEqual(sha(self.legend), before)

    def test_unsupported_field_changes_nothing(self):
        before = sha(self.legend)
        with self.assertRaises(ValueError):
            IF.insert_fixture(str(self.legend), "ceiling", {**FX, "מחיר": "100"})
        self.assertEqual(sha(self.legend), before)

    def test_corrupt_xlsx_is_a_clear_error(self):
        bad = self.d / "bad.xlsx"
        bad.write_bytes(self.legend.read_bytes()[:500])
        before = sha(bad)
        code, out = run("skill/scripts/insert_fixture.py", bad, "--category", "ceiling", "--data", json.dumps(FX))
        self.assertNotEqual(code, 0)
        self.assertIn("פגום", out)
        self.assertNotIn("Traceback", out)
        self.assertEqual(sha(bad), before)

    def test_wrong_workbook_is_a_clear_error(self):
        other = self.d / "other.xlsx"
        openpyxl.Workbook().save(other)
        with self.assertRaises(xlsx_io.WorkbookError):
            IF.insert_fixture(str(other), "ceiling", FX)

    def test_file_open_in_excel_leaves_original_intact(self):
        """Excel on Windows locks an open file; replacing it raises PermissionError."""
        before = sha(self.legend)
        with mock.patch("xlsx_io.os.replace", side_effect=PermissionError("locked")):
            with self.assertRaises(xlsx_io.WorkbookError) as e:
                IF.insert_fixture(str(self.legend), "ceiling", FX, image_path=str(self.photo))
        self.assertIn("פתוח", str(e.exception))
        self.assertEqual(sha(self.legend), before)
        self.assertEqual(list(self.d.glob(".~saving-*")), [], "temp file left behind")

    def test_crash_during_save_leaves_original_intact(self):
        before = sha(self.legend)
        with mock.patch("openpyxl.workbook.workbook.Workbook.save", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                IF.insert_fixture(str(self.legend), "ceiling", FX)
        self.assertEqual(sha(self.legend), before)

    def test_new_row_does_not_inherit_orange_from_row_above(self):
        """Found in testing: a full block copied the style of the flagged row above."""
        rows = [IF.insert_fixture(str(self.legend), "ceiling", {**FX, "גוון": g}, image_path=str(self.photo))
                for g in ("", "שחור", "לבן")]
        ws = self.sheet()
        flags = [orange(self.cell(ws, r, "גוון")) for r in rows]
        self.assertEqual(flags, [True, False, False])

    def test_orange_phrase_survives_later_scripts(self):
        """Found in testing: mark_bank_status re-saved the file without rich text."""
        row = IF.insert_fixture(str(self.legend), "ceiling", FX, image_path=str(self.photo),
                                uncertain_snippets={"תיאור": "3000K"})
        code, _ = run("bank/mark_bank_status.py", self.legend, "--row", row, "--status", "new")
        self.assertEqual(code, 0)
        self.assertIsInstance(self.cell(self.sheet(), row, "תיאור").value, CellRichText)


# --------------------------------------------------------------------- the bank
class BankBase(Tmp):
    def make_bank(self, rows):
        """rows: list of (category_en, data) -> builds an old legend and migrates it into a new bank."""
        old = self.d / "old.xlsx"
        shutil.copy(self.legend, old)
        for cat, data in rows:
            IF.insert_fixture(str(old), cat, data, image_path=str(self.photo))
        bank = self.d / "bank.xlsx"
        BBT.build(bank)
        M.migrate(str(old), str(bank), str(bank), {})
        return bank


class TestBank(BankBase):
    def test_variants_with_special_characters_are_not_merged(self):
        a = {**FX, "תיאור": "שקוע תקרה מדגם Arca Mini Ø55\nLED 7W 640LM"}
        b = {**FX, "תיאור": "שקוע תקרה מדגם Arca Mini Ø55 IP65\nLED 7W 610LM"}
        bank = self.make_bank([("ceiling", a), ("ceiling", b)])
        ws = openpyxl.load_workbook(bank)[M.BANK_SHEET_NAME]
        hr = M.find_bank_header_row(ws)
        n = sum(1 for r in range(hr + 1, ws.max_row + 1) if ws.cell(row=r, column=2).value)
        self.assertEqual(n, 2)

    def test_line_breaks_survive_the_bank(self):
        bank = self.make_bank([("ceiling", FX)])
        ws = openpyxl.load_workbook(bank)[M.BANK_SHEET_NAME]
        hr = M.find_bank_header_row(ws)
        self.assertEqual(ws.cell(row=hr + 1, column=M.col_index("תיאור")).value, FX["תיאור"])
        self.assertEqual(ws.cell(row=hr + 1, column=M.col_index("מתח_עמעום")).value, FX["מתח_עמעום"])

    def test_match_different_product_same_brand_is_not_strong(self):
        """Found in testing: 'Beam 30' scored 0.82 ('already in the bank') against 'Terra 60'."""
        terra = {**FX, "יצרן": "HALVARD LUMEN", "תיאור": "שקוע רצפה מדגם Terra 60\nLED 5W 410LM 3000K CRI90 IP67\nזוית פיזור 40 מעלות\nדרייברים למיקום באזור פריק ונגיש"}
        bank = self.make_bank([("floor", terra)])
        beam = "פנס הצפה לחזית מדגם Beam 30\nLED 20W 2100LM 3000K CRI80 IP66\nזוית פיזור 30 מעלות\nדרייברים למיקום באזור פריק ונגיש"
        self.assertNotEqual(MF.find_match("HALVARD LUMEN", beam, bank)["verdict"], "strong")
        self.assertEqual(MF.find_match("HALVARD LUMEN", terra["תיאור"], bank)["verdict"], "strong")

    def test_match_unknown_brand_is_new(self):
        bank = self.make_bank([("ceiling", FX)])
        self.assertEqual(MF.find_match("TOBEK", "גוף תאורה לינארי לקיר מדגם Riva 120\nLED 14W", bank)["verdict"], "new")

    def test_concurrent_run_is_blocked_and_bank_untouched(self):
        bank = self.make_bank([("ceiling", FX)])
        before = sha(bank)
        with bank_guard(bank, "first run"):
            code, out = run("bank/process_new_fixtures.py", self.legend, bank, "--no-rebuild")
        self.assertNotEqual(code, 0)
        self.assertIn("תפוס", out)
        self.assertEqual(sha(bank), before)

    def test_stale_lock_is_ignored(self):
        bank = self.make_bank([("ceiling", FX)])
        state = bank.parent / ".bank_state.json"
        state.write_text(json.dumps({"lock": {"who": "x", "action": "crashed", "ts": time.time() - 3600}}), encoding="utf8")
        with bank_guard(bank, "after a crash"):
            pass
        self.assertIsNone(json.loads(state.read_text(encoding="utf8"))["lock"])

    def test_export_unknown_rows_creates_nothing(self):
        bank = self.make_bank([("ceiling", FX)])
        out = self.d / "new.xlsx"
        code, msg = run("bank/export_selection.py", bank, "--rows", "500,501", "--name", "x", "--color", "3000K",
                        "--skeleton", self.legend, "--out", out)
        self.assertNotEqual(code, 0)
        self.assertFalse(out.exists(), "an empty legend was left behind")

    def test_export_bad_rows_argument(self):
        bank = self.make_bank([("ceiling", FX)])
        code, msg = run("bank/export_selection.py", bank, "--rows", "9,abc", "--skeleton", self.legend, "--out", self.d / "n.xlsx")
        self.assertNotEqual(code, 0)

    def test_export_partly_missing_rows_warns(self):
        bank = self.make_bank([("ceiling", FX)])
        out = self.d / "new.xlsx"
        code, msg = run("bank/export_selection.py", bank, "--rows", "9,500", "--name", "x", "--color", "3000K",
                        "--skeleton", self.legend, "--out", out)
        self.assertEqual(code, 0)
        self.assertIn("500", msg)
        self.assertTrue(out.exists())


class TestReturnedFile(BankBase):
    def setUp(self):
        super().setUp()
        self.bank = self.make_bank([("ceiling", FX)])

    def add_new(self, data, image=True):
        row = IF.insert_fixture(str(self.legend), "wall", data, image_path=str(self.photo) if image else None)
        code, _ = run("bank/mark_bank_status.py", self.legend, "--row", row, "--status", "new")
        self.assertEqual(code, 0)
        return row

    def bank_count(self):
        ws = openpyxl.load_workbook(self.bank)[M.BANK_SHEET_NAME]
        hr = M.find_bank_header_row(ws)
        return sum(1 for r in range(hr + 1, ws.max_row + 1) if ws.cell(row=r, column=2).value)

    def test_unverified_orange_row_does_not_enter_bank(self):
        self.add_new({**FX, "תיאור": "מנורת קיר מדגם Nova\nLED 6W", "גוון": ""})
        code, out = run("bank/process_new_fixtures.py", self.legend, self.bank, "--no-rebuild")
        self.assertEqual(code, 0)
        self.assertIn("כתום", out)
        self.assertEqual(self.bank_count(), 1)

    def test_row_without_image_is_skipped_with_reason(self):
        self.add_new({**FX, "תיאור": "מנורת קיר מדגם Nova\nLED 6W"}, image=False)
        code, out = run("bank/process_new_fixtures.py", self.legend, self.bank, "--no-rebuild")
        self.assertEqual(self.bank_count(), 1)
        self.assertIn("דולגו", out)

    def test_clean_row_is_added_once(self):
        self.add_new({**FX, "תיאור": "מנורת קיר מדגם Nova\nLED 6W"})
        run("bank/process_new_fixtures.py", self.legend, self.bank, "--no-rebuild")
        run("bank/process_new_fixtures.py", self.legend, self.bank, "--no-rebuild")
        self.assertEqual(self.bank_count(), 2)


# ------------------------------------------------------------------ the web page
class TestPage(BankBase):
    def test_description_cannot_break_out_of_the_script_tag(self):
        bank = self.make_bank([("ceiling", {**FX, "תיאור": "ספוט מדגם X</script><b>\nLED 1W"})])
        code, out = run("pages/build_pages.py", "--bank", bank, "--outdir", self.d)
        self.assertEqual(code, 0)
        html = (self.d / "picker.html").read_text(encoding="utf8")
        self.assertNotIn("X</script>", html)

    def test_not_a_bank_file(self):
        code, out = run("pages/build_pages.py", "--bank", self.legend, "--outdir", self.d)
        self.assertNotEqual(code, 0)
        self.assertNotIn("Traceback", out)


if __name__ == "__main__":
    unittest.main()
