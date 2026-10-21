#!/usr/bin/env python3
"""
Protects the fixture bank from two people writing to it at the same time, and
backs it up automatically before every change.

Why: the bank lives in a shared, cloud-synced folder. Sync clients don't merge
two simultaneous edits to the same .xlsx -- they create a "conflict copy" and one
of the two updates is silently lost.

What it does:
  1. Soft lock -- before writing, a "busy" record (who, what, when) is written to
     `.bank_state.json` next to the bank. A second run that sees a fresh lock
     stops with a clear message instead of overwriting. A lock older than 10
     minutes is treated as stale (e.g. a run that crashed) and ignored.
  2. Rolling backup -- a copy of the bank is saved to `backups/` before every
     change, rotating through 5 slots, so you can always go back.

Both only ever create or overwrite files, never delete -- so they keep working
in folders where the user has no delete permission.

Usage:
    from bank_guard import bank_guard
    with bank_guard(bank_path, "הוספת גופים חדשים"):
        ...  # all reads and writes to the bank
"""
import getpass, json, os, shutil, socket, time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

STALE_SECONDS = 10 * 60
BACKUP_SLOTS = 5


def _state_path(bank_path):
    return Path(bank_path).parent / ".bank_state.json"


def _read_state(p):
    try:
        return json.loads(p.read_text(encoding="utf8"))
    except Exception:
        return {}


def _write_state(p, data):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf8")


def _who():
    try:
        return f"{getpass.getuser()}@{socket.gethostname()}"
    except Exception:
        return "unknown"


def backup(bank_path, state=None, state_path=None):
    """שומר עותק של הבנק באחד מ-5 מקומות הגיבוי המתחלפים. מחזיר את הנתיב שנכתב."""
    bank_path = Path(bank_path)
    d = bank_path.parent / "backups"
    d.mkdir(parents=True, exist_ok=True)
    state = state if state is not None else {}
    slot = (int(state.get("backup_slot", 0)) % BACKUP_SLOTS) + 1
    dest = d / f"bank-backup-{slot}.xlsx"
    shutil.copy(bank_path, dest)
    state["backup_slot"] = slot
    state.setdefault("backups", {})[str(slot)] = datetime.now(timezone.utc).astimezone().strftime("%d.%m.%Y %H:%M")
    if state_path is not None:
        _write_state(state_path, state)
    return dest


@contextmanager
def bank_guard(bank_path, action="עדכון הבנק", skip_backup=False):
    bank_path = Path(bank_path)
    sp = _state_path(bank_path)
    state = _read_state(sp)
    lock = state.get("lock")
    if lock:
        age = time.time() - float(lock.get("ts", 0))
        if age < STALE_SECONDS:
            mins = int(age // 60)
            raise SystemExit(
                f"קובץ הבנק תפוס כרגע ע\"י {lock.get('who','מישהו אחר')} "
                f"(\"{lock.get('action','')}\", לפני {mins} דקות).\n"
                f"כדי לא לדרוס את העבודה שלו — לא בוצע שום שינוי. כדאי לנסות שוב בעוד כמה דקות.")
        print(f"שים לב: נמצאה נעילה ישנה מ-{lock.get('who','?')} (יותר מ-10 דקות) — כנראה הרצה שנפלה. ממשיכים.")

    state["lock"] = {"who": _who(), "action": action, "ts": time.time(),
                     "when": datetime.now(timezone.utc).astimezone().strftime("%d.%m.%Y %H:%M")}
    _write_state(sp, state)
    try:
        if not skip_backup and bank_path.exists():
            dest = backup(bank_path, state, sp)
            print(f"גיבוי נשמר: {dest.name}")
        yield
    finally:
        state = _read_state(sp)
        state["lock"] = None
        state["last_change"] = {"who": _who(), "action": action,
                                "when": datetime.now(timezone.utc).astimezone().strftime("%d.%m.%Y %H:%M")}
        _write_state(sp, state)
