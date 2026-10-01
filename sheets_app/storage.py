"""SQLite asosidagi workbook (kitob) saqlash tizimi."""
import sqlite3
import string
from contextlib import contextmanager

SCHEMA = """
CREATE TABLE IF NOT EXISTS sheets (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    pos  INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS cells (
    sheet INTEGER NOT NULL,
    r     INTEGER NOT NULL,
    c     INTEGER NOT NULL,
    raw   TEXT,
    PRIMARY KEY (sheet, r, c)
);
CREATE INDEX IF NOT EXISTS idx_cells_sheet ON cells(sheet);
"""


def col_letter(c: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = ""
    c += 1
    while c:
        c, m = divmod(c - 1, 26)
        s = string.ascii_uppercase[m] + s
    return s


def letter_col(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


class Workbook:
    """Bitta hujjat = SQLite fayl."""

    def __init__(self, path=None):
        self.path = path
        if path is None:
            self.conn = sqlite3.connect(":memory:")
        else:
            self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)
        if self.conn.execute("SELECT COUNT(*) FROM sheets").fetchone()[0] == 0:
            self.add_sheet("Varaq1")

    # ---------- sheets ----------
    def list_sheets(self):
        return self.conn.execute(
            "SELECT id, name FROM sheets ORDER BY pos, id").fetchall()

    def sheet_id(self, name):
        row = self.conn.execute(
            "SELECT id FROM sheets WHERE name = ?", (name,)).fetchone()
        return row["id"] if row else None

    def add_sheet(self, name="Varaq"):
        base, i = name, 1
        while self.sheet_id(name) is not None:
            i += 1
            name = f"{base}{i}"
        pos = (self.conn.execute(
            "SELECT COALESCE(MAX(pos),0)+1 FROM sheets").fetchone()[0])
        cur = self.conn.execute(
            "INSERT INTO sheets (name, pos) VALUES (?, ?)", (name, pos))
        self.conn.commit()
        return cur.lastrowid

    def rename_sheet(self, sid, name):
        self.conn.execute("UPDATE sheets SET name=? WHERE id=?", (name, sid))
        self.conn.commit()

    def delete_sheet(self, sid):
        if len(self.list_sheets()) <= 1:
            raise ValueError("Oxirgi varaqni o'chirib bo'lmaydi")
        self.conn.execute("DELETE FROM cells WHERE sheet=?", (sid,))
        self.conn.execute("DELETE FROM sheets WHERE id=?", (sid,))
        self.conn.commit()

    def duplicate_sheet(self, sid):
        src = self.conn.execute(
            "SELECT name FROM sheets WHERE id=?", (sid,)).fetchone()["name"]
        new_id = self.add_sheet(src + " (nusxa)")
        for row in self.conn.execute(
                "SELECT r, c, raw FROM cells WHERE sheet=?", (sid,)):
            self.conn.execute(
                "INSERT OR REPLACE INTO cells (sheet, r, c, raw) VALUES (?,?,?,?)",
                (new_id, row["r"], row["c"], row["raw"]))
        self.conn.commit()
        return new_id

    def clear_sheet(self, sid):
        self.conn.execute("DELETE FROM cells WHERE sheet=?", (sid,))
        self.conn.commit()

    # ---------- cells ----------
    def get_raw(self, sid, r, c):
        row = self.conn.execute(
            "SELECT raw FROM cells WHERE sheet=? AND r=? AND c=?",
            (sid, r, c)).fetchone()
        return row["raw"] if row else None

    def set_raw(self, sid, r, c, raw):
        if raw is None or raw == "":
            self.conn.execute("DELETE FROM cells WHERE sheet=? AND r=? AND c=?",
                              (sid, r, c))
        else:
            self.conn.execute(
                "INSERT OR REPLACE INTO cells (sheet, r, c, raw) VALUES (?,?,?,?)",
                (sid, r, c, raw))
        self.conn.commit()

    def iter_cells(self, sid):
        for row in self.conn.execute(
                "SELECT r, c, raw FROM cells WHERE sheet=?", (sid,)):
            yield row["r"], row["c"], row["raw"]

    def used_range(self, sid):
        row = self.conn.execute(
            "SELECT MIN(r) a, MAX(r) b, MIN(c) c, MAX(c) d FROM cells WHERE sheet=?",
            (sid,)).fetchone()
        if row["a"] is None:
            return 0, 0, 0, 0
        return row["a"], row["b"], row["c"], row["d"]

    def non_empty_rows(self, sid):
        return [r for (r,) in self.conn.execute(
            "SELECT DISTINCT r FROM cells WHERE sheet=? ORDER BY r", (sid,))]

    # ---------- struktura o'zgartirish ----------
    def insert_rows(self, sid, at, n=1):
        rows = self.conn.execute(
            "SELECT r, c FROM cells WHERE sheet=? AND r>=? ORDER BY r DESC",
            (sid, at)).fetchall()
        with self.tx():
            for row in rows:
                self.conn.execute(
                    "UPDATE cells SET r=? WHERE sheet=? AND r=? AND c=?",
                    (row["r"] + n, sid, row["r"], row["c"]))

    def delete_rows(self, sid, at, n=1):
        self.conn.execute(
            "DELETE FROM cells WHERE sheet=? AND r>=? AND r<?", (sid, at, at + n))
        self.conn.execute(
            "UPDATE cells SET r=r-? WHERE sheet=? AND r>?", (n, sid, at + n - 1))
        self.conn.commit()

    def insert_cols(self, sid, at, n=1):
        rows = self.conn.execute(
            "SELECT r, c FROM cells WHERE sheet=? AND c>=? ORDER BY c DESC",
            (sid, at)).fetchall()
        with self.tx():
            for row in rows:
                self.conn.execute(
                    "UPDATE cells SET c=? WHERE sheet=? AND r=? AND c=?",
                    (row["c"] + n, sid, row["r"], row["c"]))

    def delete_cols(self, sid, at, n=1):
        self.conn.execute(
            "DELETE FROM cells WHERE sheet=? AND c>=? AND c<?", (sid, at, at + n))
        self.conn.execute(
            "UPDATE cells SET c=c-? WHERE sheet=? AND c>?", (n, sid, at + n - 1))
        self.conn.commit()

    # ---------- umumiy ----------
    @contextmanager
    def tx(self):
        try:
            yield self.conn
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def close(self):
        self.conn.close()
