#!/usr/bin/env python3
"""
Rebuild the selection page after the bank changes.

The page holds a snapshot of the bank (data + thumbnails embedded in the HTML),
it does not read the bank live -- so every change to the bank needs a rebuild,
or the page shows yesterday's bank. rebuild() is called automatically at the end
of every script that modifies the bank, so nobody has to remember.

A failure here is never fatal: the bank change has already been saved, so it
only prints a warning.
"""
import os, subprocess, sys
from pathlib import Path

# עברית בפלט גם בווינדוס, כשהפלט נקרא דרך pipe (ברירת המחדל שם היא cp1252)
for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")


PROJECT = Path(__file__).resolve().parents[1]


def rebuild(bank_path, quiet=False):
    """Rebuild the selection page from THIS bank file. The bank path is passed
    explicitly: rebuilding from a default path would silently show a different
    bank than the one that was just changed (e.g. after --bank-out)."""
    bank_path = Path(bank_path).resolve()
    script = PROJECT / "pages" / "build_pages.py"
    if not script.exists():
        print(f"אזהרה: {script.name} לא נמצא — הדף לא נבנה מחדש.")
        return False
    if not quiet:
        print("\nבונה מחדש את דף הבחירה...")
    try:
        # Hebrew output: Windows defaults to cp1252 for pipes, which crashes on Hebrew -> force UTF-8
        r = subprocess.run([sys.executable, str(script), "--bank", str(bank_path),
                            "--outdir", str(bank_path.parent)],
                           cwd=str(script.parent), capture_output=True, text=True, timeout=600,
                           encoding="utf-8", errors="replace",
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    except Exception as e:
        print(f"אזהרה: בניית הדף נכשלה ({e}). השינוי בבנק נשמר. אפשר להריץ ידנית: pages/build_pages.py")
        return False
    if r.returncode != 0:
        print("אזהרה: בניית הדף נכשלה. השינוי בבנק נשמר.")
        print((r.stderr or "").strip()[-500:])
        return False
    if not quiet:
        for line in (r.stdout or "").splitlines():
            if line.strip():
                print("  " + line)
    print("דף הבחירה עודכן.")
    return True


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: python bank/rebuild_pages.py <bank.xlsx>")
    sys.exit(0 if rebuild(sys.argv[1]) else 1)
