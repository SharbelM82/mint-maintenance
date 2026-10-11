"""SQLite storage. Only input fields are stored; formulas are computed on read."""
import json
import sqlite3
from datetime import datetime

from flask import current_app, g

from . import calc
from .fields import BY_KEY, INPUT_FIELDS, INPUT_KEYS

_SQL_TYPES = {"money": "REAL", "percent": "REAL", "int": "INTEGER"}


def schema():
    cols = []
    for f in INPUT_FIELDS:
        if f.key == "wo_number":
            cols.append("wo_number TEXT NOT NULL UNIQUE COLLATE NOCASE")
        else:
            cols.append(f"{f.key} {_SQL_TYPES.get(f.kind, 'TEXT')}")
    return f"""
    CREATE TABLE IF NOT EXISTS work_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        {", ".join(cols)},
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS activity (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        work_order_id INTEGER,
        wo_number TEXT,
        action TEXT NOT NULL,
        details TEXT,
        at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS activity_wo ON activity(work_order_id);
    """


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(current_app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(_exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db(path):
    con = sqlite3.connect(path)
    con.executescript(schema())
    # Add any input column introduced after the database was created.
    have = {r[1] for r in con.execute("PRAGMA table_info(work_orders)")}
    for f in INPUT_FIELDS:
        if f.key not in have:
            con.execute(f"ALTER TABLE work_orders ADD COLUMN {f.key} "
                        f"{_SQL_TYPES.get(f.kind, 'TEXT')}")
    con.commit()
    con.close()


def now():
    return datetime.now().isoformat(timespec="seconds")


def enrich(row, today=None):
    """Row from the database -> dict with inputs plus every computed column."""
    r = dict(row)
    r.update(calc.compute(r, today))
    return r


def all_work_orders(today=None):
    rows = get_db().execute(
        "SELECT * FROM work_orders ORDER BY date_in IS NULL, date_in DESC, id DESC")
    return [enrich(r, today) for r in rows]


def get_work_order(wo_id):
    row = get_db().execute("SELECT * FROM work_orders WHERE id = ?", (wo_id,)).fetchone()
    return enrich(row) if row else None


def find_by_number(wo_number):
    return get_db().execute("SELECT * FROM work_orders WHERE wo_number = ?",
                            (wo_number,)).fetchone()


def _changes(before, after):
    diff = {}
    for k in INPUT_KEYS:
        old, new = (before or {}).get(k), after.get(k)
        if old != new and not (old in (None, "") and new in (None, "")):
            diff[BY_KEY[k].label] = [old, new]
    return diff


def log(db, wo_id, wo_number, action, details=None):
    db.execute("INSERT INTO activity (work_order_id, wo_number, action, details, at) "
               "VALUES (?, ?, ?, ?, ?)",
               (wo_id, wo_number, action,
                json.dumps(details, default=str) if details else None, now()))


def insert(data, action="Created"):
    db = get_db()
    ts = now()
    keys = INPUT_KEYS + ["created_at", "updated_at"]
    vals = [data.get(k) for k in INPUT_KEYS] + [ts, ts]
    cur = db.execute(f"INSERT INTO work_orders ({', '.join(keys)}) "
                     f"VALUES ({', '.join('?' for _ in keys)})", vals)
    log(db, cur.lastrowid, data["wo_number"], action, _changes(None, data))
    db.commit()
    return cur.lastrowid


def update(wo_id, data, action="Updated", only=None):
    """Update a work order. `only` limits which input keys are written."""
    db = get_db()
    before = dict(db.execute("SELECT * FROM work_orders WHERE id = ?",
                             (wo_id,)).fetchone())
    keys = only or INPUT_KEYS
    merged = {**before, **{k: data.get(k) for k in keys}}
    diff = _changes(before, merged)
    if not diff:
        return False
    db.execute(f"UPDATE work_orders SET {', '.join(f'{k} = ?' for k in keys)}, "
               "updated_at = ? WHERE id = ?",
               [data.get(k) for k in keys] + [now(), wo_id])
    log(db, wo_id, merged["wo_number"], action, diff)
    db.commit()
    return True


def delete(wo_id):
    db = get_db()
    row = db.execute("SELECT wo_number FROM work_orders WHERE id = ?", (wo_id,)).fetchone()
    if row:
        db.execute("DELETE FROM work_orders WHERE id = ?", (wo_id,))
        log(db, wo_id, row["wo_number"], "Deleted")
        db.commit()


def activity(wo_id=None, limit=50):
    db = get_db()
    if wo_id is None:
        rows = db.execute("SELECT * FROM activity ORDER BY id DESC LIMIT ?", (limit,))
    else:
        rows = db.execute("SELECT * FROM activity WHERE work_order_id = ? "
                          "ORDER BY id DESC LIMIT ?", (wo_id, limit))
    out = []
    for r in rows:
        d = dict(r)
        d["details"] = json.loads(d["details"]) if d["details"] else {}
        out.append(d)
    return out


def distinct_values(key):
    if key not in INPUT_KEYS:
        return []
    rows = get_db().execute(f"SELECT DISTINCT {key} FROM work_orders "
                            f"WHERE {key} IS NOT NULL AND {key} != '' ORDER BY {key}")
    return [r[0] for r in rows]
