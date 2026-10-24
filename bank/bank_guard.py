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
     stops with a clear message instead of overwriting. While a run is alive a
     small background "heartbeat" refreshes the lock every 15 seconds; a lock that
     has not been refreshed for 120 seconds belongs to a run that crashed, and is
     ignored. (A fixed 10-minute expiry meant a crash blocked everyone for 10
     minutes.) `unlock_bank.py` releases it by hand.
  2. Rolling backup -- a copy of the bank is saved to `backups/` before every
     change, rotating through 5 slots, so you can always go back. A backup is
     only made if the bank opens as a valid .xlsx: otherwise five failed runs
     would overwrite all five good backups with copies of the broken file.

Both only ever create or overwrite files, never delete -- so they keep working
in folders where the user has no delete permission.

Usage:
    from bank_guard import bank_guard
    with bank_guard(bank_path, "הוספת גופים חדשים"):
        ...  # all reads and writes to the bank
"""
import getpass, json, os, shutil, socket, threading, time, uuid, zipfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

STALE_SECONDS = 120        # no heartbeat for this long = the run died
HEARTBEAT_SECONDS = 15
_STATE_LOCK = threading.Lock()
BACKUP_SLOTS = 5


def _state_path(bank_path):
    return Path(bank_path).parent / ".bank_state.json"


def _read_state(p):
    try:
        return json.loads(p.read_text(encoding="utf8"))
    except Exception:
        return {}


def _write_state(p, data):
    """Atomic write: to a temp file first, then swapped in, so a crash in the
    middle cannot leave a half-written state file (which would read as "no lock")."""
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(".~" + p.name)
    with _STATE_LOCK:        # the heartbeat thread and the main thread both write
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf8")
        os.replace(tmp, p)


def bank_is_healthy(path):
    """True if the file opens as a complete .xlsx (an intact zip containing a workbook).
    Deliberately without openpyxl: fast, and it cannot itself fail on a damaged file."""
    try:
        with zipfile.ZipFile(path) as z:
            return z.testzip() is None and "xl/workbook.xml" in z.namelist()
    except Exception:
        return False


def _who():
    try:
        return f"{getpass.getuser()}@{socket.gethostname()}"
    except Exception:
        return "unknown"


def backup(bank_path, state=None, state_path=None):
    """Saves a copy of the bank into one of the 5 rotating backup slots and returns the path
    written -- or None if the bank is damaged (then no backup is touched)."""
    bank_path = Path(bank_path)
    d = bank_path.parent / "backups"
    d.mkdir(parents=True, exist_ok=True)
    state = state if state is not None else {}
    slot = (int(state.get("backup_slot", 0)) % BACKUP_SLOTS) + 1
    dest = d / f"bank-backup-{slot}.xlsx"
    if not bank_is_healthy(bank_path):
        # a damaged bank must never overwrite a good backup
        print("אזהרה: קובץ הבנק לא נפתח כראוי — לא נשמר גיבוי חדש, כדי לא לדרוס גיבוי תקין. "
              "כדאי לבדוק את הבנק או לשחזר מתיקיית הגיבויים.")
        return None
    shutil.copy(bank_path, dest)
    state["backup_slot"] = slot
    state.setdefault("backups", {})[str(slot)] = datetime.now(timezone.utc).astimezone().strftime("%d.%m.%Y %H:%M")
    if state_path is not None:
        _write_state(state_path, state)
    return dest


def _heartbeat(sp, token, stop):
    """Refreshes the lock's timestamp every few seconds while the run is alive."""
    while not stop.wait(HEARTBEAT_SECONDS):
        try:
            st = _read_state(sp)
            lk = st.get("lock")
            if lk and lk.get("id") == token:     # only our own lock
                lk["ts"] = time.time()
                _write_state(sp, st)
        except Exception:
            pass


@contextmanager
def bank_guard(bank_path, action="עדכון הבנק", skip_backup=False):
    bank_path = Path(bank_path)
    sp = _state_path(bank_path)
    state = _read_state(sp)
    lock = state.get("lock")
    if lock:
        age = time.time() - float(lock.get("ts", 0))
        if age < STALE_SECONDS:
            secs = int(age)
            ago = f"לפני {secs} שניות" if secs < 90 else f"לפני {secs // 60} דקות"
            raise SystemExit(
                f"קובץ הבנק תפוס כרגע ע\"י {lock.get('who','מישהו אחר')} "
                f"(\"{lock.get('action','')}\", {ago}).\n"
                f"כדי לא לדרוס את העבודה שלו — לא בוצע שום שינוי.\n"
                f"אם זו הייתה הרצה שלך שנפלה (קריסה/כיבוי): המתינו כשתי דקות ונסו שוב, "
                f"או הריצו unlock_bank.py. אם מישהו אחר עובד עכשיו — נסו שוב בעוד כמה דקות.")
        print(f"שים לב: נמצאה נעילה ישנה מ-{lock.get('who','?')} — כנראה הרצה שנפלה. ממשיכים.")

    token = uuid.uuid4().hex       # identifies THIS run's lock, so the heartbeat never refreshes someone else's
    state["lock"] = {"who": _who(), "action": action, "ts": time.time(), "id": token,
                     "when": datetime.now(timezone.utc).astimezone().strftime("%d.%m.%Y %H:%M")}
    _write_state(sp, state)
    stop = threading.Event()
    hb = threading.Thread(target=_heartbeat, args=(sp, token, stop), daemon=True)
    hb.start()
    try:
        if not skip_backup and bank_path.exists():
            dest = backup(bank_path, state, sp)
            if dest is not None:
                print(f"גיבוי נשמר: {dest.name}")
        yield
    finally:
        stop.set()
        hb.join(timeout=5)
        state = _read_state(sp)
        state["lock"] = None
        state["last_change"] = {"who": _who(), "action": action,
                                "when": datetime.now(timezone.utc).astimezone().strftime("%d.%m.%Y %H:%M")}
        _write_state(sp, state)
