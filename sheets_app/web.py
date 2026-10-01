"""SheetApp WEB — brauzerda ishlaydigan server (localhost)."""
import io
import os

from flask import (Flask, request, jsonify, render_template, send_file,
                   Response)

from .storage import Workbook, col_letter
from .formula import Evaluator
from . import xls_io, reports

BASE = os.path.dirname(os.path.abspath(__file__))
if os.environ.get("SHEETAPP_DB"):
    DB_PATH = os.environ["SHEETAPP_DB"]
elif os.environ.get("VERCEL"):
    DB_PATH = "/tmp/web_workbook.db"      # Vercel'da faqat /tmp yoziladi
else:
    DB_PATH = os.path.join(os.path.dirname(BASE), "web_workbook.db")
UPLOAD_DIR = "/tmp" if os.environ.get("VERCEL") else os.path.dirname(BASE)

app = Flask(__name__, template_folder=os.path.join(BASE, "templates"))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

store = Workbook(DB_PATH)


# ---------- yordamchilar ----------
def fmt_val(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v != v:
            return "#NUM!"
        if v.is_integer():
            return str(int(v))
        return f"{v:.10g}"
    return str(v)


def sheets_payload():
    return [{"id": sh["id"], "name": sh["name"]} for sh in store.list_sheets()]


def cells_payload(sid):
    ev = Evaluator(store, sid)
    out = {}
    for r, c, raw in store.iter_cells(sid):
        v = ev.raw_value(sid, r, c)
        out[f"{r},{c}"] = {"raw": raw, "val": fmt_val(v)}
    return out


def write_report_sheet(title, header, rows):
    sid = store.sheet_id(title)
    if sid is None:
        sid = store.add_sheet(title)
    for r in range(0, 300):
        for c in range(0, 12):
            store.set_raw(sid, r, c, None)
    for c, h in enumerate(header):
        store.set_raw(sid, 0, c, h)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            if isinstance(val, float):
                txt = str(int(val)) if float(val).is_integer() else f"{val:.10g}"
            else:
                txt = str(val)
            store.set_raw(sid, i + 1, j, txt)
    return sid


# ---------- sahifalar ----------
@app.route("/")
def index():
    return render_template("index.html")


# ---------- API ----------
@app.get("/api/workbook")
def api_workbook():
    return jsonify({"sheets": sheets_payload(), "db": os.path.basename(DB_PATH)})


@app.get("/api/sheet/<int:sid>")
def api_sheet(sid):
    name = None
    for sh in store.list_sheets():
        if sh["id"] == sid:
            name = sh["name"]
    if name is None:
        return jsonify({"error": "varaq topilmadi"}), 404
    r0, r1, c1, c2 = store.used_range(sid)
    return jsonify({"id": sid, "name": name, "cells": cells_payload(sid),
                    "used": [r1, c2], "sheets": sheets_payload()})


@app.post("/api/cell")
def api_cell():
    d = request.get_json(force=True)
    sid, r, c = int(d["sid"]), int(d["r"]), int(d["c"])
    store.set_raw(sid, r, c, d.get("raw") or None)
    return jsonify({"cells": cells_payload(sid)})


@app.post("/api/cells")
def api_cells():
    """Ko'p katakni bir vaqtda yozish (nusxalash/kirish uchun)."""
    d = request.get_json(force=True)
    sid = int(d["sid"])
    for item in d["data"]:
        r, c, raw = int(item[0]), int(item[1]), item[2]
        store.set_raw(sid, r, c, raw if raw not in ("", None) else None)
    return jsonify({"cells": cells_payload(sid)})


@app.post("/api/clear")
def api_clear():
    d = request.get_json(force=True)
    sid = int(d["sid"])
    for r, c in d["cells"]:
        store.set_raw(sid, int(r), int(c), None)
    return jsonify({"cells": cells_payload(sid)})


@app.post("/api/struct")
def api_struct():
    d = request.get_json(force=True)
    sid, op, at = int(d["sid"]), d["op"], int(d["at"])
    n = int(d.get("n", 1))
    if op == "ins_row":
        store.insert_rows(sid, at, n)
    elif op == "del_row":
        store.delete_rows(sid, at, n)
    elif op == "ins_col":
        store.insert_cols(sid, at, n)
    elif op == "del_col":
        store.delete_cols(sid, at, n)
    else:
        return jsonify({"error": "noma'lum amal"}), 400
    return jsonify({"cells": cells_payload(sid)})


@app.post("/api/sort")
def api_sort():
    d = request.get_json(force=True)
    sid, col = int(d["sid"]), int(d["col"])
    ev = Evaluator(store, sid)
    rows = [r for r in store.non_empty_rows(sid) if r >= int(d.get("start", 1))]

    def key(r):
        v = ev.raw_value(sid, r, col)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return (0, float(v), "")
        return (1, 0.0, str(v).lower())
    rows.sort(key=key, reverse=bool(d.get("desc")))
    _, rmax, cmax, _ = store.used_range(sid)
    data = [{c: store.get_raw(sid, r, c) for c in range(cmax + 1)} for r in rows]
    with store.tx():
        for r, line in zip(rows, data):
            for c, raw in line.items():
                store.set_raw(sid, r, c, raw)
    return jsonify({"cells": cells_payload(sid)})


@app.post("/api/sheets")
def api_sheets():
    d = request.get_json(force=True)
    op = d["op"]
    try:
        if op == "add":
            store.add_sheet(d.get("name") or "Varaq")
        elif op == "rename":
            store.rename_sheet(int(d["sid"]), d["name"])
        elif op == "delete":
            store.delete_sheet(int(d["sid"]))
        elif op == "duplicate":
            store.duplicate_sheet(int(d["sid"]))
        else:
            return jsonify({"error": "noma'lum amal"}), 400
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"sheets": sheets_payload()})


@app.post("/api/import")
def api_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "fayl yo'q"}), 400
    tmp = os.path.join(UPLOAD_DIR, "_import_tmp.xlsx")
    f.save(tmp)
    try:
        xls_io.import_excel(store, tmp)
    except Exception as e:
        return jsonify({"error": f"Import xatosi: {e}"}), 400
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return jsonify({"sheets": sheets_payload()})


@app.get("/api/export")
def api_export():
    tmp = os.path.join(UPLOAD_DIR, "_export_tmp.xlsx")
    xls_io.export_excel(store, tmp)
    with open(tmp, "rb") as fh:
        data = fh.read()
    os.remove(tmp)
    return send_file(io.BytesIO(data), as_attachment=True,
                     download_name="SheetApp.xlsx",
                     mimetype="application/vnd.openxmlformats-"
                             "officedocument.spreadsheetml.sheet")


@app.get("/api/report/balance")
def api_balance():
    rep = reports.balance_report(store)
    return jsonify({
        "rows": [[n, float(v), note] for n, v, note in rep["rows"]],
        "issues": [[k, t] for k, t in rep["issues"]],
        "kirim": rep["kirim_total"], "chiqim": rep["chiqim_total"]})


@app.get("/api/report/dds")
def api_dds():
    rows = reports.dds_report(store)
    return jsonify({"rows": rows})


@app.post("/api/report/write")
def api_report_write():
    d = request.get_json(force=True)
    if d.get("kind") == "balance":
        rep = reports.balance_report(store)
        write_report_sheet("Hisobot", ["Kategoriya", "Summa", "Izoh"], rep["rows"])
        title = "Hisobot"
    else:
        rows = reports.dds_report(store)
        disp = [[d_, k - c - i,
                 f"kirim={k:,.0f} | chiqim={c:,.0f} | ish haqi={i:,.0f} | "
                 f"balans={b:,.0f}"] for d_, k, c, i, b in rows]
        write_report_sheet("DDS", ["Sana", "Kunlik sof o'zgarish", "Tafsilot"], disp)
        title = "DDS"
    return jsonify({"sheets": sheets_payload(), "title": title})


def main():
    port = int(os.environ.get("SHEETWEB_PORT", 5001))
    print("=" * 50)
    print(f"  SheetApp WEB ishga tushdi!")
    print(f"  Brauzerda oching:  http://127.0.0.1:{port}")
    print("=" * 50)
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False,
            threaded=False)


if __name__ == "__main__":
    main()
