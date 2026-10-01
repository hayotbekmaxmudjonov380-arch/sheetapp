"""SheetApp WEB — brauzerda ishlaydigan server (localhost)."""
import io
import os
from datetime import date

from flask import (Flask, request, jsonify, render_template, send_file,
                   Response)
from openpyxl import Workbook

from .storage import Workbook as _Wb, col_letter
from .formula import Evaluator
from . import xls_io, reports, db

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

store = _Wb(DB_PATH)


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


# ---------- ZKTeco ADMS (Face ID) push ----------
def _text(lines):
    return Response("\n".join(lines) + "\n", mimetype="text/plain")


@app.get("/iclock/cdata")
def zk_cdata_get():
    """Qurilma qo'shilganda sozlamalarni oladi (handshake)."""
    sn = request.args.get("SN", "")
    allow = os.environ.get("ZK_SN", "").strip()
    if allow and sn and sn not in [s.strip() for s in allow.split(",")]:
        return _text(["ERROR: unknown SN"])
    return _text([
        f"GET OPTION FROM: {sn}",
        "ATTLOGStamp=None",
        "OpStamp=0",
        "ErrorDelay=30",
        "Delay=10",
        "TransTimes=00:00;14:05",
        "Transinterval=1",
        "TransFlag=TransData AttLog OpLog AttPhoto EnrollUser ChgUser "
        "EnrollFP ChgFP UserPic",
        "TimeZone=5",
        "Realtime=1",
        "Encrypt=None",
    ])


def _parse_attlog(body, sn):
    rows = []
    for line in body.replace("\r\n", "\n").split("\n"):
        if not line.strip():
            continue
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        pin, ts = parts[0].strip(), parts[1].strip()
        if not pin or not ts:
            continue
        try:
            date(int(ts[0:4]), int(ts[5:7]), int(ts[8:10]))
        except (ValueError, IndexError):
            continue

        def _i(v):
            try:
                return int(str(v).strip())
            except (TypeError, ValueError):
                return None
        rows.append({
            "user_id": pin,
            "punch_time": ts,
            "status": _i(parts[2]) if len(parts) > 2 else 0,
            "verify_mode": _i(parts[3]) if len(parts) > 3 else None,
            "work_code": (parts[4].strip() or None) if len(parts) > 4 else None,
            "device_sn": sn,
            "raw": line,
        })
    return rows


@app.post("/iclock/cdata")
def zk_cdata_post():
    """Qurilma davomat yozuvlarini yuboradi (ATTLOG)."""
    sn = request.args.get("SN", "")
    table = (request.args.get("table") or "").upper()
    body = request.get_data(as_text=True, cache=False) or ""
    if table == "ATTLOG" and body.strip() and db.available():
        try:
            db.insert_punches(_parse_attlog(body, sn))
        except Exception as e:
            print(f"[ZK] push saqlashda xato: {e}")
    return _text(["OK"])


@app.route("/iclock/getrequest")
def zk_getrequest():
    return _text(["OK"])


@app.route("/iclock/devicecmd", methods=["GET", "POST"])
def zk_devicecmd():
    return _text(["OK"])


@app.route("/iclock/registry", methods=["GET", "POST"])
def zk_registry():
    return _text(["OK"])


# ---------- Davomat (Face ID hisoboti) ----------
def _attendance_rows(dfrom, dto):
    punches = db.list_attendance(dfrom, dto)
    workers = {w["user_id"]: w for w in db.list_workers()}
    start = os.environ.get("SHEET_START_TIME", "09:00")
    groups = {}
    for p in punches:
        pt = str(p.get("punch_time") or "").replace("T", " ")
        if " " not in pt:
            continue
        day, tstr = pt.split(" ", 1)
        tstr = tstr[:8]
        g = groups.setdefault((day, p["user_id"]),
                              {"first": tstr, "last": tstr, "count": 0})
        g["count"] += 1
        if tstr < g["first"]:
            g["first"] = tstr
        if tstr > g["last"]:
            g["last"] = tstr
    rows = []
    for (day, uid), g in sorted(groups.items(), reverse=True):
        w = workers.get(uid) or {}
        late = 0
        if start and g["first"] > start:
            try:
                fm = int(g["first"][3:5]) + int(g["first"][0:2]) * 60
                sm = int(start[3:5]) + int(start[0:2]) * 60
                late = max(0, fm - sm)
            except (ValueError, IndexError):
                late = 0
        rows.append({"date": day, "user_id": uid,
                     "name": w.get("full_name") or uid,
                     "first": g["first"], "last": g["last"],
                     "count": g["count"], "late": late})
    return rows, start


@app.get("/api/attendance")
def api_attendance():
    if not db.available():
        return jsonify({"error": "Supabase sozlanmagan"}), 503
    dfrom = request.args.get("from") or date.today().isoformat()
    dto = request.args.get("to") or dfrom
    try:
        rows, start = _attendance_rows(dfrom, dto)
    except Exception as e:
        return jsonify({"error": f"Davomat o'qishda xato: {e}"}), 502
    return jsonify({"from": dfrom, "to": dto, "start_time": start,
                    "rows": rows})


@app.get("/api/attendance/export")
def api_attendance_export():
    if not db.available():
        return jsonify({"error": "Supabase sozlanmagan"}), 503
    dfrom = request.args.get("from") or date.today().isoformat()
    dto = request.args.get("to") or dfrom
    rows, _ = _attendance_rows(dfrom, dto)
    wb = Workbook()
    ws = wb.active
    ws.title = "Davomat"
    ws.append(["Sana", "ID", "Ism", "Kelgan", "Chiqqan",
               "Yozuvlar", "Kechikish (daq)"])
    for r in rows:
        ws.append([r["date"], r["user_id"], r["name"], r["first"],
                   r["last"], r["count"], r["late"] or ""])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"Davomat_{dfrom}_{dto}.xlsx",
                     mimetype="application/vnd.openxmlformats-"
                             "officedocument.spreadsheetml.sheet")


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
