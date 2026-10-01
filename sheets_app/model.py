"""Qt jadval modeli (Google Sheets uslubidagi katak)."""
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QColor, QFont

from .storage import col_letter, letter_col
from .formula import Evaluator, fmt

MAX_COLS = 60
MAX_ROWS = 999
FORMULA_BG = QColor("#e8f0fe")


class SheetModel(QAbstractTableModel):
    def __init__(self, store, sheet_id, parent=None):
        super().__init__(parent)
        self.store = store
        self.sid = sheet_id
        self.ev = Evaluator(store, sheet_id)

    # ---------- o'lcham ----------
    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else MAX_ROWS

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else MAX_COLS

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return col_letter(section)
        return str(section + 1)

    # ---------- qiymatlar ----------
    def raw_at(self, r, c):
        return self.store.get_raw(self.sid, r, c)

    def value_at(self, r, c):
        return self.ev.raw_value(self.sid, r, c)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        raw = self.raw_at(r, c)
        if role in (Qt.DisplayRole, Qt.EditRole):
            if role == Qt.EditRole:
                return raw if raw is not None else ""
            if raw is None or raw == "":
                return ""
            if isinstance(raw, str) and raw.startswith("="):
                v = self.ev.raw_value(self.sid, r, c)
                return fmt(v) if not isinstance(v, str) else v
            return raw
        if role == Qt.TextAlignmentRole:
            raw0 = self.raw_at(r, c)
            v = self.ev.raw_value(self.sid, r, c) if raw0 else None
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        if role == Qt.BackgroundRole and raw and str(raw).startswith("="):
            return FORMULA_BG
        if role == Qt.ForegroundRole:
            raw0 = self.raw_at(r, c)
            if raw0:
                v = self.ev.raw_value(self.sid, r, c)
                if isinstance(v, str) and v in ("#CYCLE!", "#REF!", "#NAME?",
                                                "#VALUE!", "#DIV/0!", "#PARSE!"):
                    return QColor("#d93025")
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or role != Qt.EditRole:
            return False
        text = value if isinstance(value, str) else ("" if value is None else str(value))
        text = text.strip() if not (text.startswith("=") and len(text) > 1) else text
        self.store.set_raw(self.sid, index.row(), index.column(), text)
        self.ev.invalidate()
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole,
                                             Qt.TextAlignmentRole, Qt.BackgroundRole])
        return True

    def clear_cells(self, indexes):
        for i in indexes:
            self.store.set_raw(self.sid, i.row(), i.column(), None)
        if indexes:
            self.ev.invalidate()
            self.dataChanged.emit(indexes[0], indexes[-1])

    # ---------- struktura ----------
    def insert_rows(self, at, n=1):
        self.beginInsertRows(QModelIndex(), at, at + n - 1)
        self.store.insert_rows(self.sid, at, n)
        self.ev.invalidate()
        self.endInsertRows()

    def remove_rows(self, at, n=1):
        self.beginRemoveRows(QModelIndex(), at, at + n - 1)
        self.store.delete_rows(self.sid, at, n)
        self.ev.invalidate()
        self.endRemoveRows()

    def insert_cols(self, at, n=1):
        self.beginInsertColumns(QModelIndex(), at, at + n - 1)
        self.store.insert_cols(self.sid, at, n)
        self.ev.invalidate()
        self.endInsertColumns()

    def remove_cols(self, at, n=1):
        self.beginRemoveColumns(QModelIndex(), at, at + n - 1)
        self.store.delete_cols(self.sid, at, n)
        self.ev.invalidate()
        self.endRemoveColumns()

    def sort_by(self, col, descending=False, start_row=1):
        rows = self.store.non_empty_rows(self.sid)
        rows = [r for r in rows if r >= start_row]
        if not rows:
            return
        def key(r):
            v = self.ev.raw_value(self.sid, r, col)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return (0, float(v), "")
            return (1, 0.0, fmt(v).lower())
        rows.sort(key=key, reverse=descending)
        # faqat satrlarni qayta joylash (rawlarni vaqtinchalik saqlash)
        data = []
        _, rmax, _, cmax = self.store.used_range(self.sid)
        for r in rows:
            line = {}
            for c in range(0, cmax + 1):
                line[c] = self.store.get_raw(self.sid, r, c)
            data.append(line)
        with self.store.tx():
            for r, line in zip(rows, data):
                for c, raw in line.items():
                    if raw is None:
                        self.store.conn.execute(
                            "DELETE FROM cells WHERE sheet=? AND r=? AND c=?",
                            (self.sid, r, c))
                    else:
                        self.store.conn.execute(
                            "INSERT OR REPLACE INTO cells (sheet,r,c,raw) VALUES (?,?,?,?)",
                            (self.sid, r, c, raw))
        self.ev.invalidate()
        self.layoutChanged.emit()

    def refresh(self):
        self.ev.invalidate()
        self.layoutChanged.emit()
