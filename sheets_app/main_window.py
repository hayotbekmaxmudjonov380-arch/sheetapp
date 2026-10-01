"""Asosiy oyna — Google Sheets uslubidagi interfeys."""
import os
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QKeySequence, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableView, QHeaderView,
    QTabBar, QPushButton, QFileDialog, QMessageBox, QLineEdit, QLabel,
    QDialog, QFormLayout, QComboBox, QDialogButtonBox,
    QAbstractItemView, QMenu, QInputDialog, QTableWidget, QTableWidgetItem,
    QToolBar, QStatusBar, QGroupBox, QCheckBox, QStyledItemDelegate)

from .storage import Workbook, col_letter, letter_col
from .model import SheetModel, MAX_ROWS, MAX_COLS
from . import xls_io, reports

FILTER_BG = QColor("#fff8e1")


class SheetTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setSortingEnabled(False)
        self.verticalHeader().setDefaultSectionSize(24)
        self.horizontalHeader().setDefaultSectionSize(110)
        self.horizontalHeader().setMinimumSectionSize(40)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.NoDragDrop)
        self.setWordWrap(False)
        self.setEditTriggers(QAbstractItemView.DoubleClicked |
                             QAbstractItemView.EditKeyPressed |
                             QAbstractItemView.AnyKeyPressed)

    def keyPressEvent(self, event):
        k = event.key()
        editing = self.state() == QAbstractItemView.EditingState
        if not editing and k in (Qt.Key_Delete, Qt.Key_Backspace):
            m = self.model()
            idxs = self.selectionModel().selectedIndexes()
            if m is not None and hasattr(m, "clear_cells") and idxs:
                m.clear_cells(idxs)
                self.viewport().update()
                return
        if not editing and k in (Qt.Key_Return, Qt.Key_Enter):
            idx = self.currentIndex()
            if idx.isValid():
                self.edit(idx)
                return
        super().keyPressEvent(event)


class CellDelegate(QStyledItemDelegate):
    """Enter bosilganda saqlab, katakni pastga tushiradi."""

    def __init__(self, view):
        super().__init__(view)
        self.view = view
        self._idx = None

    def createEditor(self, parent, option, index):
        self._idx = index
        ed = QLineEdit(parent)
        ed.setFrame(False)
        ed.returnPressed.connect(lambda: self._commit_move(ed))
        return ed

    def setEditorData(self, editor, index):
        editor.setText(index.data(Qt.EditRole) or "")

    def _commit_move(self, ed):
        idx = self._idx
        self.commitData(ed)
        self.closeEditor(ed, QStyledItemDelegate.NoEditReason)
        m = self.view.model()
        if m is None or idx is None:
            return
        r = min(idx.row() + 1, m.rowCount() - 1)
        nxt = m.index(r, idx.column())
        self.view.setCurrentIndex(nxt)
        self.view.setFocus()


# ---------------- Qidiruv ----------------
class FindDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Qidiruv")
        self.setModal(False)
        lay = QHBoxLayout(self)
        self.q = QLineEdit()
        self.q.setPlaceholderText("Matn yoki formula...")
        lay.addWidget(self.q)
        self.case = QCheckBox("Register hisobga olinadi")
        lay.addWidget(self.case)
        b = QPushButton("Keyingi →")
        b.clicked.connect(self.parent().find_next)
        lay.addWidget(b)
        self.q.returnPressed.connect(self.parent().find_next)
        self.resize(420, 60)


# ---------------- Filtr ----------------
class FilterDialog(QDialog):
    def __init__(self, parent, col_names):
        super().__init__(parent)
        self.setWindowTitle("Filtr")
        f = QFormLayout(self)
        self.col = QComboBox()
        self.col.addItems(col_names)
        self.cond = QComboBox()
        self.cond.addItems(["Ichida", "Teng", "Teng emas", "Katta", "Kichik",
                            "Boshlaydi", "Bo'sh", "Bo'sh emas"])
        self.val = QLineEdit()
        f.addRow("Ustun:", self.col)
        f.addRow("Shart:", self.cond)
        f.addRow("Qiymat:", self.val)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Qo'llash")
        bb.button(QDialogButtonBox.Cancel).setText("Bekor")
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        f.addRow(bb)


# ---------------- Hisobot oynasi ----------------
class ReportDialog(QDialog):
    def __init__(self, parent, title, rows, issues=None, extra_cols=0):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(720, 520)
        v = QVBoxLayout(self)
        self.table = QTableWidget(len(rows), 3 + extra_cols)
        self.table.setHorizontalHeaderLabels(
            ["Kategoriya", "Summa (so'm)", "Izoh"] + [""] * extra_cols)
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                if isinstance(val, float):
                    it = QTableWidgetItem(f"{val:,.0f}")
                    it.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
                    it.setData(Qt.UserRole, val)
                else:
                    it = QTableWidgetItem(str(val))
                if row[0] in ("SOF FOYDA", "Yalpi foyda", "Jami xarajat") or \
                        (extra_cols and row[0] == "SOF FOYDA"):
                    it.setBackground(QColor("#d3e3fd"))
                self.table.setItem(i, j, it)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        v.addWidget(self.table)
        if issues:
            g = QGroupBox("Ma'lumot sifati — topilgan xatolar")
            gl = QVBoxLayout(g)
            for kind, text in issues:
                lb = QLabel(("⚠ " if kind == "warn" else "ℹ ") + text)
                lb.setWordWrap(True)
                lb.setStyleSheet("color:#b06000" if kind == "warn" else "color:#1a73e8")
                gl.addWidget(lb)
            v.addWidget(g)
        bb = QDialogButtonBox()
        b_close = bb.addButton("Yopish", QDialogButtonBox.RejectRole)
        b_write = bb.addButton("Natijani 'Hisobot' varaqga yozish", QDialogButtonBox.AcceptRole)
        b_close.clicked.connect(self.reject)
        b_write.clicked.connect(lambda: self.done(1))
        v.addWidget(bb)
        self.result_rows = rows


# ---------------- Asosiy oyna ----------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SheetApp — Google Sheets uslubidagi hisob-kitob")
        self.resize(1200, 720)
        self.store = Workbook(None)
        self.path = None
        self.models = {}
        self._find_pos = 0
        self._build_ui()
        self._load_sheets()
        self._update_title()

    # ---------- UI ----------
    def _build_ui(self):
        tb = QToolBar("Asosiy")
        tb.setMovable(False)
        self.addToolBar(tb)

        def act(text, slot, shortcut=None, tip=None):
            a = QAction(text, self)
            a.triggered.connect(slot)
            if shortcut:
                a.setShortcut(shortcut)
            if tip:
                a.setStatusTip(tip)
            tb.addAction(a)
            return a

        act("📄 Yangi", self.new_file, QKeySequence.New)
        act("📂 Ochish", self.open_file, QKeySequence.Open)
        act("💾 Saqlash", self.save_file, QKeySequence.Save)
        tb.addSeparator()
        act("⬇ Excel'dan import", self.import_excel)
        act("⬆ Excel'ga export", self.export_excel)
        tb.addSeparator()
        act("🔍 Qidiruv", self.show_find, QKeySequence.Find)
        act("🧮 Saralash", self.sort_dialog)
        act("🔽 Filtr", self.show_filter)
        tb.addSeparator()
        act("📊 Balans hisoboti", self.show_balance)
        act("📅 DDS hisoboti", self.show_dds)

        # formula satri
        fb = QWidget()
        fl = QHBoxLayout(fb)
        fl.setContentsMargins(4, 2, 4, 2)
        fl.setSpacing(4)
        self.cell_label = QLabel("A1")
        self.cell_label.setFixedWidth(56)
        self.cell_label.setStyleSheet(
            "background:#e8eaed;padding:4px;border-radius:4px;font-weight:bold;")
        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText(
            "Katakka yozing: matn, raqam yoki formula (masalan: =SUM(A1:A10))")
        self.formula_edit.returnPressed.connect(self._apply_formula_bar)
        fl.addWidget(self.cell_label)
        fl.addWidget(self.formula_edit)

        # jadval
        self.view = SheetTableView()
        self.view.setItemDelegate(CellDelegate(self.view))
        hh = self.view.horizontalHeader()
        hh.setContextMenuPolicy(Qt.CustomContextMenu)
        hh.customContextMenuRequested.connect(self._col_menu_pos)
        vh = self.view.verticalHeader()
        vh.setContextMenuPolicy(Qt.CustomContextMenu)
        vh.customContextMenuRequested.connect(self._row_menu_pos)

        # pastki: varaq tablari (Google Sheets uslubidagi)
        bottom = QWidget()
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(4, 0, 4, 2)
        self.tabs = QTabBar()
        self.tabs.setExpanding(False)
        self.tabs.setMovable(True)
        self.tabs.setDocumentMode(True)
        self.tabs.currentChanged.connect(self._tab_changed)
        self.tabs.tabMoved.connect(self._tab_moved)
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.customContextMenuRequested.connect(self._tab_menu)
        self.tabs.tabBarDoubleClicked.connect(self._rename_tab_dialog)
        btn_add = QPushButton("＋")
        btn_add.setFixedSize(28, 26)
        btn_add.setToolTip("Yangi varaq")
        btn_add.clicked.connect(self.add_sheet)
        bl.addWidget(self.tabs, 1)
        bl.addWidget(btn_add)

        central = QWidget()
        v = QVBoxLayout(central)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        v.addWidget(fb)
        v.addWidget(self.view, 1)
        v.addWidget(bottom)
        self.setCentralWidget(central)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Tayyor. Katakka yozing — Enter (pastga), Tab (o'ngga).")

    # ---------- varaq boshqaruvi ----------
    def _load_sheets(self):
        self.tabs.blockSignals(True)
        while self.tabs.count():
            self.tabs.removeTab(0)
        for i, sh in enumerate(self.store.list_sheets()):
            self.tabs.addTab(sh["name"])
        idx = max(0, self.tabs.currentIndex())
        self.tabs.blockSignals(False)
        if self.tabs.count():
            self.tabs.setCurrentIndex(idx if idx < self.tabs.count() else 0)
            self._tab_changed(self.tabs.currentIndex())

    def _sheet_id_by_index(self, i):
        sheets = self.store.list_sheets()
        return sheets[i]["id"] if 0 <= i < len(sheets) else None

    def _tab_changed(self, i):
        sid = self._sheet_id_by_index(i)
        if sid is None:
            return
        if sid not in self.models:
            self.models[sid] = SheetModel(self.store, sid)
            self.models[sid].dataChanged.connect(lambda *a: self._autosave())
        m = self.models[sid]
        old = self.view.model()
        if old is not m:
            self.view.setModel(m)
            self.view.selectionModel().selectionChanged.connect(self._on_selection)
        self._update_cell_label()

    def _tab_moved(self, frm, to):
        sheets = self.store.list_sheets()
        order = [sh["id"] for sh in sheets]
        moved = order.pop(frm)
        order.insert(to, moved)
        with self.store.tx():
            for pos, sid in enumerate(order):
                self.store.conn.execute("UPDATE sheets SET pos=? WHERE id=?",
                                        (pos, sid))

    def add_sheet(self):
        self.store.add_sheet("Varaq")
        self.models.clear()
        self._load_sheets()
        self.tabs.setCurrentIndex(self.tabs.count() - 1)

    def _tab_menu(self, pos):
        idx = self.tabs.tabAt(pos)
        if idx < 0:
            return
        sid = self._sheet_id_by_index(idx)
        m = QMenu(self)
        a1 = m.addAction("✏ Nomini o'zgartirish")
        a2 = m.addAction("📋 Nusxalash")
        a3 = m.addAction("🗑 O'chirish")
        act = m.exec(self.tabs.mapToGlobal(pos))
        if act is a1:
            self._rename_tab_dialog(idx)
        elif act is a2:
            self.store.duplicate_sheet(sid)
            self.models.clear()
            self._load_sheets()
        elif act is a3:
            try:
                self.store.delete_sheet(sid)
            except ValueError as e:
                QMessageBox.warning(self, "Xato", str(e))
                return
            self.models.pop(sid, None)
            self._load_sheets()

    def _rename_tab_dialog(self, idx):
        if not isinstance(idx, int) or idx < 0:
            idx = self.tabs.currentIndex()
        sid = self._sheet_id_by_index(idx)
        old = self.tabs.tabText(idx)
        name, ok = QInputDialog.getText(self, "Varaq nomi", "Yangi nom:", text=old)
        if ok and name.strip():
            try:
                self.store.rename_sheet(sid, name.strip())
                self.tabs.setTabText(idx, name.strip())
                self.models.pop(sid, None)
                self._tab_changed(idx)
            except Exception as e:
                QMessageBox.warning(self, "Xato", str(e))

    # ---------- katak ----------
    def current_cell(self):
        idx = self.view.currentIndex()
        if not idx.isValid():
            return 0, 0
        return idx.row(), idx.column()

    def _on_selection(self, *args):
        self._update_cell_label()
        self._selection_stats()

    def _update_cell_label(self, r=None, c=None):
        if r is None:
            r, c = self.current_cell()
        self.cell_label.setText(f"{col_letter(c)}{r + 1}")
        m = self.view.model()
        if m:
            raw = m.raw_at(r, c) or ""
            if self.formula_edit.hasFocus():
                return
            self.formula_edit.setText(raw)

    def _apply_formula_bar(self):
        r, c = self.current_cell()
        m = self.view.model()
        if not m:
            return
        m.setData(m.index(r, c), self.formula_edit.text())
        self.view.viewport().update()
        self._autosave()

    def _selection_stats(self):
        m = self.view.model()
        if not m:
            return
        sel = self.view.selectionModel()
        idxs = sel.selectedIndexes()
        n, s, nums = 0, 0.0, 0
        for i in idxs:
            v = m.value_at(i.row(), i.column())
            n += 1
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                s += float(v)
                nums += 1
        if n:
            avg = s / nums if nums else 0
            self.status.showMessage(
                f"Tanlandi: {n} katak | Yig'indi: {s:,.0f} | "
                f"O'rtacha: {avg:,.0f} | Raqamlar: {nums}")
        else:
            self.status.showMessage("Tayyor.")

    def keyPressEvent(self, event):
        super().keyPressEvent(event)

    # ---------- ustun / qator menyusi ----------
    def _col_menu_pos(self, pos):
        header = self.view.horizontalHeader()
        logical = header.logicalIndexAt(pos)
        if logical >= 0:
            self._col_menu(logical, header.mapToGlobal(pos))

    def _row_menu_pos(self, pos):
        header = self.view.verticalHeader()
        logical = header.logicalIndexAt(pos)
        if logical >= 0:
            self._row_menu(logical, header.mapToGlobal(pos))

    def _col_menu(self, logical, global_pos=None):
        m = QMenu(self)
        m.addAction("Ustunni o'ngga qo'shish", lambda: self._add_col(logical + 1))
        m.addAction("Ustunni chapga qo'shish", lambda: self._add_col(logical))
        m.addAction("Ustunni o'chirish", lambda: self._del_col(logical))
        m.addSeparator()
        m.addAction("⬇ Saralash (kattadan)", lambda: self._sort(logical, True))
        m.addAction("⬆ Saralash (kichikdan)", lambda: self._sort(logical, False))
        if global_pos is None:
            header = self.view.horizontalHeader()
            global_pos = header.mapToGlobal(
                header.rect().center())
        m.exec(global_pos)

    def _row_menu(self, logical, global_pos=None):
        m = QMenu(self)
        m.addAction("Yuqoriga qator qo'shish", lambda: self._add_row(logical))
        m.addAction("Pastga qator qo'shish", lambda: self._add_row(logical + 1))
        m.addAction("Qatorni o'chirish", lambda: self._del_row(logical))
        if global_pos is None:
            header = self.view.verticalHeader()
            global_pos = header.mapToGlobal(header.rect().center())
        m.exec(global_pos)

    def _add_col(self, at):
        self.view.model().insert_cols(at)
        self._autosave()

    def _del_col(self, at):
        self.view.model().remove_cols(at)
        self._autosave()

    def _add_row(self, at):
        self.view.model().insert_rows(at)
        self._autosave()

    def _del_row(self, at):
        self.view.model().remove_rows(at)
        self._autosave()

    def _sort(self, col, desc):
        self.view.model().sort_by(col, desc, start_row=1)
        self.status.showMessage(f"{col_letter(col)} ustuni bo'yicha saralandi", 3000)
        self._autosave()

    def sort_dialog(self):
        names = [f"{col_letter(c)}" for c in range(0, 30)]
        d = FilterDialog(self, names)
        d.setWindowTitle("Saralash")
        d.cond.hide()
        d.val.hide()
        if d.exec() == QDialog.Accepted:
            self._sort(d.col.currentIndex(), False)

    # ---------- qidiruv / filtr ----------
    def show_find(self):
        if not hasattr(self, "find_dlg") or self.find_dlg is None:
            self.find_dlg = FindDialog(self)
        self.find_dlg.show()
        self.find_dlg.q.setFocus()
        self.find_dlg.q.selectAll()

    def find_next(self):
        m = self.view.model()
        if not m:
            return
        text = self.find_dlg.q.text()
        if not text:
            return
        case = self.find_dlg.case.isChecked()
        total = MAX_ROWS * MAX_COLS
        start = self._find_pos
        for k in range(total):
            i = (start + k) % total
            r, c = divmod(i, MAX_COLS)
            raw = m.raw_at(r, c) or ""
            hay = str(raw) if case else str(raw).lower()
            needle = text if case else text.lower()
            if needle in hay:
                self._find_pos = i + 1
                idx = m.index(r, c)
                self.view.setCurrentIndex(idx)
                self.view.scrollTo(idx)
                self.status.showMessage(f"Topildi: {col_letter(c)}{r + 1}", 3000)
                return
        QMessageBox.information(self, "Qidiruv", "Topilmadi.")

    def show_filter(self):
        m = self.view.model()
        if not m:
            return
        names = [f"{col_letter(c)} ustun" for c in range(MAX_COLS)]
        d = FilterDialog(self, names)
        if d.exec() != QDialog.Accepted:
            return
        col = d.col.currentIndex()
        cond = d.cond.currentText()
        val = d.val.text()
        hidden = 0
        for r in range(MAX_ROWS):
            v = m.value_at(r, col)
            sv = str(v) if v is not None else ""
            try:
                nv = float(val) if val else None
            except ValueError:
                nv = None
            ok = True
            if cond == "Ichida":
                ok = val.lower() in sv.lower()
            elif cond == "Teng":
                ok = sv.lower() == val.lower()
            elif cond == "Teng emas":
                ok = sv.lower() != val.lower()
            elif cond == "Katta":
                ok = nv is not None and isinstance(v, (int, float)) and v > nv
            elif cond == "Kichik":
                ok = nv is not None and isinstance(v, (int, float)) and v < nv
            elif cond == "Boshlaydi":
                ok = sv.lower().startswith(val.lower())
            elif cond == "Bo'sh":
                ok = sv.strip() == ""
            elif cond == "Bo'sh emas":
                ok = sv.strip() != ""
            self.view.setRowHidden(r, not ok)
            if not ok:
                hidden += 1
        self.status.showMessage(
            f"Filtr qo'llandi: {hidden} ta yashirildi. "
            f"Tozalash → 'Filtr' → Bekor", 6000)

    # ---------- fayl ----------
    def _update_title(self):
        name = os.path.basename(self.path) if self.path else "Yangi hujjat"
        self.setWindowTitle(f"{name} — SheetApp")

    def _autosave(self):
        if self.path:
            QTimer.singleShot(600, self._do_save)

    def _do_save(self):
        try:
            if self.path and os.path.exists(self.path):
                xls_io.export_excel(self.store, self.path)
        except Exception as e:
            self.status.showMessage(f"Saqlashda xato: {e}", 5000)

    def new_file(self):
        if QMessageBox.question(
                self, "Yangi hujjat",
                "Joriy hujjat yopiladi. Davom etasizmi?",
                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        self.store.close()
        self.store = Workbook(None)
        self.models = {}
        self.path = None
        self._load_sheets()
        self._update_title()

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Ochish", "", "Excel (*.xlsx);;Barchasi (*)")
        if not path:
            return
        try:
            if path.lower().endswith(".xlsx"):
                new = Workbook(None)
                xls_io.import_excel(new, path)
                self.store.close()
                self.store = new
                self.path = None
            else:
                QMessageBox.warning(self, "Xato", "Faqat .xlsx fayllar qo'llab-quvvatlanadi")
                return
        except Exception as e:
            QMessageBox.critical(self, "Import xatosi", str(e))
            return
        self.models = {}
        self._load_sheets()
        self._update_title()
        self.status.showMessage(f"Ochildi: {path}", 5000)

    def save_file(self):
        if self.path:
            self._do_save()
            self.status.showMessage(f"Saqlandi: {self.path}", 5000)
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Saqlash", "hisob.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            xls_io.export_excel(self.store, path)
            self.path = path
            self._update_title()
            self.status.showMessage(f"Saqlandi: {path}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Saqlash xatosi", str(e))

    def import_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Excel'dan import", "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            n = xls_io.import_excel(self.store, path)
        except Exception as e:
            QMessageBox.critical(self, "Import xatosi", str(e))
            return
        self.models = {}
        self._load_sheets()
        self.status.showMessage(f"{n} ta varaq import qilindi: {path}", 6000)

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Excel'ga export", "hisob.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            xls_io.export_excel(self.store, path)
            self.status.showMessage(f"Export qilindi: {path}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Export xatosi", str(e))

    # ---------- hisobotlar ----------
    def _write_report_sheet(self, title, header, rows):
        sid = self.store.sheet_id(title)
        if sid is None:
            sid = self.store.add_sheet(title)
        for r in range(0, 200):
            for c in range(0, 10):
                self.store.set_raw(sid, r, c, None)
        for c, h in enumerate(header):
            self.store.set_raw(sid, 0, c, h)
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                if isinstance(val, float):
                    txt = str(int(val)) if float(val).is_integer() else str(val)
                else:
                    txt = str(val)
                self.store.set_raw(sid, i + 1, j, txt)
        self.models.pop(sid, None)
        self._load_sheets()
        for i in range(self.tabs.count()):
            if self._sheet_id_by_index(i) == sid:
                self.tabs.setCurrentIndex(i)
                break

    def show_balance(self):
        rep = reports.balance_report(self.store)
        d = ReportDialog(self, "📊 Balans va sof foyda (avtomatik)",
                         rep["rows"], rep["issues"])
        if d.exec() == 1:
            self._write_report_sheet(
                "Hisobot", ["Kategoriya", "Summa", "Izoh"], rep["rows"])

    def show_dds_report(self):
        rows = reports.dds_report(self.store)
        if not rows:
            QMessageBox.information(self, "DDS", "Kunlik ma'lumot topilmadi.")
            return
        disp = [(d, k - c - i, f"kirim={k:,.0f} | chiqim={c:,.0f} | "
                               f"ish haqi={i:,.0f} | balans={b:,.0f}")
                for d, k, c, i, b in rows]
        d = ReportDialog(self, "📅 DDS — kunlik kassa hisoboti", disp)
        if d.exec() == 1:
            self._write_report_sheet(
                "DDS", ["Sana", "Kunlik sof o'zgarish", "Tafsilot"],
                disp)

    def show_dds(self):
        self.show_dds_report()

    def closeEvent(self, event):
        if self.path:
            self._do_save()
        self.store.close()
        super().closeEvent(event)
