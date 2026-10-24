#!/usr/bin/env python3
"""
Manually release the bank's lock, for when a run died (crash, shutdown) and left it behind.

You rarely need this: a lock without a heartbeat expires by itself after ~2 minutes.

Usage:
    python unlock_bank.py <path to the bank .xlsx>
If the lock is still fresh (someone may really be working), the script refuses
unless --force is added.
"""
import argparse
import sys
import time
from pathlib import Path

import bank_guard as G

for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("bank_xlsx")
    ap.add_argument("--force", action="store_true", help="release even a fresh lock")
    a = ap.parse_args(argv)

    sp = G._state_path(Path(a.bank_xlsx))
    st = G._read_state(sp)
    lk = st.get("lock")
    if not lk:
        print("הבנק לא נעול. אין מה לשחרר.")
        return 0
    age = int(time.time() - float(lk.get("ts", 0)))
    print(f"נעילה של {lk.get('who','?')} (\"{lk.get('action','')}\"), עודכנה לפני {age} שניות.")
    if age < G.STALE_SECONDS and not a.force:
        sys.exit("הנעילה עדיין טרייה — ייתכן שמישהו עובד עכשיו. אם בטוחים שלא, להריץ שוב עם --force.")
    st["lock"] = None
    G._write_state(sp, st)
    print("הנעילה שוחררה.")
    return 0


if __name__ == "__main__":
    main()
