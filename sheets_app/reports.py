"""Avtomatik hisobotlar: Balans, DDS va ma'lumot sifati tekshiruvi.

Bu modul Excel'dagi qo'lda yozilgan xatolarni avtomatik tuzatadi:
  1) ish haqi -- barcha kun va barcha qatorlar yig'indisi (C4 formulasi emas)
  2) savdo -- Кирим varag'idan olinadi (qo'lda 100 mln emas)
  3) xarajatlar -- Чиким kategoriyalaridan, bo'sh varaq emas
  4) tungi smenalar -- chiqish <= kirish bo'lsa +24 soat
  5) nomlar farqi -- xodimlar ro'yxati solishtiriladi
  6) ma'lumot kam kunlarda -- statistika chiqariladi
  7) bo'sh varaqalar -- ogohlantirish beriladi
"""
import re
from datetime import datetime, date, time

from .storage import Workbook

DATE_FMT = ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y", "%d/%m/%Y")


def _parse_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if not isinstance(v, str):
        return None
    s = v.strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s
    for f in DATE_FMT:
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def _parse_time(v):
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, str):
        m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", v.strip())
        if m:
            return int(m.group(1)) * 60 + int(m.group(2))
    return None


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(" ", "").replace(",", ""))
        except ValueError:
            return None
    return None


def _norm_name(s):
    return re.sub(r"\s+", " ", str(s).strip().lower())


def find_sheet(store, *keys):
    for sh in store.list_sheets():
        n = _norm_name(sh["name"])
        for k in keys:
            if k in n:
                return sh["id"], sh["name"]
    return None, None


# ---------------- Kirim / Chiqim ----------------
def _detect_tx(store, sid, kind):
    """Ustunlarni sarlavhadan aniqlaydi. kind='kirim'|'chiqim'."""
    rows = {}
    for r, c, raw in store.iter_cells(sid):
        rows.setdefault(r, {})[c] = raw
    if not rows:
        return None
    max_r = max(rows)
    header_row = None
    amount_col = cat_col = date_col = kontr_col = None
    aliases_amt = {"сумма", "summa", "sum", "сум"}
    aliases_cat = {"статья", "статя", "modda", "category", "kategoriya",
                   "мақсат", "харачат"}
    aliases_date = {"sana", "дата", "date", "санаси"}
    for r in range(0, min(max_r + 1, 6)):
        for c, raw in rows.get(r, {}).items():
            t = _norm_name(raw) if isinstance(raw, str) else ""
            if t in aliases_amt:
                amount_col, header_row = c, r
        if amount_col is not None:
            break
    if amount_col is None:
        # qo'lda odatiy joylashuv
        if kind == "kirim":
            return {"header": 0, "cat": 3, "amt": 4, "date": 5, "kontr": None}
        return {"header": 0, "kontr": 3, "cat": 4, "amt": 5, "date": 6}
    start = (header_row or 0) + 1
    # sana ustunini aniqlash
    maxc = max(max(cols) for cols in rows.values())
    best, best_n = None, 0
    for c in range(0, maxc + 1):
        if c == amount_col:
            continue
        n = sum(1 for r in rows if r >= start and _parse_date(rows[r].get(c)))
        if n > best_n:
            best, best_n = c, n
    date_col = best if best_n else None
    # kategoriya / kontragent
    maxc = max(max(cols) for cols in rows.values())
    others = [c for c in range(0, maxc + 1)
              if c not in (amount_col, date_col)]
    for r in range(0, start):
        for c, raw in rows.get(r, {}).items():
            t = _norm_name(raw) if isinstance(raw, str) else ""
            if t in aliases_cat and c not in (amount_col, date_col):
                cat_col = c
            if t in ("kontragent", "контрагент") and c != amount_col:
                kontr_col = c
    if cat_col is None and others:
        cat_col = others[0]
    if kind == "chiqim":
        kontr_col = next((c for c in others if c != cat_col), None)
    return {"header": header_row or 0, "cat": cat_col, "amt": amount_col,
            "date": date_col, "kontr": kontr_col}


def transactions(store, kind):
    sid, name = find_sheet(store, "кирим" if kind == "kirim" else "чиким",
                           "kirim" if kind == "kirim" else "chiqim")
    if sid is None:
        return [], name
    cols = _detect_tx(store, sid, kind)
    out = []
    rows = {}
    for r, c, raw in store.iter_cells(sid):
        rows.setdefault(r, {})[c] = raw
    for r in sorted(rows):
        if r <= cols["header"]:
            continue
        row = rows[r]
        amt = _num(row.get(cols["amt"])) if cols["amt"] is not None else None
        if not amt:
            continue
        d = _parse_date(row.get(cols["date"])) if cols["date"] is not None else None
        cat = row.get(cols["cat"]) if cols["cat"] is not None else ""
        kontr = row.get(cols["kontr"]) if cols.get("kontr") is not None else ""
        out.append({"row": r, "date": d, "category": str(cat or "").strip(),
                    "kontragent": str(kontr or "").strip(), "amount": amt})
    return out, name


# ---------------- Ish haqi ----------------
def salary_records(store):
    """'ish haqi' varag'idan barcha kunlik yozuvlarni o'qiydi."""
    sid, name = find_sheet(store, "ish haqi", "иш haqi", "зп", "salary")
    if sid is None:
        return [], None, []
    rows = {}
    for r, c, raw in store.iter_cells(sid):
        rows.setdefault(r, {})[c] = raw
    if not rows:
        return [], name, []
    # sana ustunlari (odat 2-qator)
    date_cols = {}   # col -> sana
    date_row = None
    for r in sorted(rows):
        found = {}
        for c, raw in rows[r].items():
            d = _parse_date(raw)
            if d:
                found[c] = d
        if len(found) >= 2:
            date_cols, date_row = found, r
            break
        if found:
            date_cols, date_row = found, r
    if not date_cols:
        return [], name, []
    recs, issues = [], []
    for r in sorted(rows):
        if r <= date_row:
            continue
        nm = rows[r].get(1)
        if not nm or _norm_name(nm) == "ism sharif":
            continue
        for c, d in date_cols.items():
            ci = rows[r].get(c)
            co = rows[r].get(c + 1)
            amt = _num(rows[r].get(c + 2))
            ci_m, co_m = _parse_time(ci), _parse_time(co)
            if ci_m is None and co_m is None and not amt:
                continue
            hours = None
            night = False
            if ci_m is not None and co_m is not None:
                if co_m <= ci_m:
                    night = True
                    hours = round((co_m + 1440 - ci_m) / 60, 2)
                else:
                    hours = round((co_m - ci_m) / 60, 2)
                if hours > 16:
                    night = True
            recs.append({"row": r, "name": str(nm).strip(), "date": d,
                         "check_in": ci or "", "check_out": co or "",
                         "amount": amt or 0.0, "hours": hours, "night": night})
    return recs, name, issues


def salary_totals(store):
    recs, name, _ = salary_records(store)
    by_day, by_emp = {}, {}
    for x in recs:
        by_day[x["date"]] = by_day.get(x["date"], 0) + x["amount"]
        by_emp[x["name"]] = by_emp.get(x["name"], 0) + x["amount"]
    return recs, name, by_day, by_emp


# ---------------- Balans ----------------
def balance_report(store):
    kirim, kirim_sheet = transactions(store, "kirim")
    chiqim, chiqim_sheet = transactions(store, "chiqim")
    recs, ish_sheet, by_day, by_emp = salary_totals(store)

    savdo = sum(t["amount"] for t in kirim
                if "savdo" in t["category"].lower() or "савдо" in t["category"].lower())
    if savdo == 0:
        savdo = sum(t["amount"] for t in kirim)
    tovar = sum(t["amount"] for t in chiqim
                if "tovar" in t["category"].lower() or "товар" in t["category"].lower())
    ish_haqi = sum(x["amount"] for x in recs)

    def by_cat(word_ru, word_uz=""):
        return sum(t["amount"] for t in chiqim
                   if word_ru in t["category"].lower()
                   or (word_uz and word_uz in t["category"].lower()))
    kommunal = by_cat("kommunal", "коммунальн")
    arenda = by_cat("arenda", "аренда")
    marketing = by_cat("marketing", "маркетинг")
    remont = by_cat("remont", "ремонт")
    hisoblangan = tovar + kommunal + arenda + marketing + remont
    boshqa_chiqim = sum(t["amount"] for t in chiqim) - hisoblangan
    nalog = round(savdo * 0.02, 2)
    yalpi = savdo - tovar
    jami_xarajat = ish_haqi + kommunal + arenda + marketing + remont + nalog
    sof = yalpi - jami_xarajat

    rows = [
        ("Savdo (Кирим)", savdo, "avtomatik: " + (kirim_sheet or "topilmadi")),
        ("Sotib olingan tovar (Чиким)", -tovar, "avtomatik: " + (chiqim_sheet or "topilmadi")),
        ("Yalpi foyda", yalpi, "Savdo - Tovar"),
        ("Ish haqi", -ish_haqi, f"avtomatik: {len(recs)} ta yozuv, {len(by_emp)} ta xodim"),
        ("Kommunal", -kommunal, "Чиким kategoriyasi"),
        ("Arenda", -arenda, "Чиким kategoriyasi"),
        ("Marketing", -marketing, "Чиким kategoriyasi"),
        ("Remont", -remont, "Чиким kategoriyasi"),
        ("Boshqa chiqimlar", -boshqa_chiqim, "hisoblanmagan kategoriya"),
        ("Nalog (Savdoning 2%)", -nalog, "= Savdo * 0.02"),
        ("Jami xarajat", -jami_xarajat, "yig'indi"),
        ("SOF FOYDA", sof, "Yalpi foyda - Jami xarajat"),
    ]
    issues = quality_issues(store, kirim, chiqim, recs)
    return {"rows": rows, "issues": issues,
            "kirim_total": sum(t["amount"] for t in kirim),
            "chiqim_total": sum(t["amount"] for t in chiqim)}


# ---------------- DDS ----------------
def dds_report(store):
    kirim, _ = transactions(store, "kirim")
    chiqim, _ = transactions(store, "chiqim")
    recs, _, by_day, _ = salary_totals(store)
    dates = sorted({t["date"] for t in kirim if t["date"]}
                   | {t["date"] for t in chiqim if t["date"]}
                   | set(by_day))
    rows, bal = [], 0.0
    for d in dates:
        k = sum(t["amount"] for t in kirim if t["date"] == d)
        c = sum(t["amount"] for t in chiqim if t["date"] == d)
        i = by_day.get(d, 0.0)
        bal = bal + k - c - i
        rows.append((d, k, c, i, bal))
    return rows


# ---------------- Ma'lumot sifati (7 ta xato) ----------------
def quality_issues(store, kirim=None, chiqim=None, recs=None):
    if kirim is None:
        kirim, _ = transactions(store, "kirim")
    if chiqim is None:
        chiqim, _ = transactions(store, "chiqim")
    if recs is None:
        recs, _, _, _ = salary_totals(store)
    issues = []

    # 1) ish haqi jami (eski C4 xatosi bilan solishtirish)
    if recs:
        issues.append(("info",
                       f"Ish haqi jami: {sum(x['amount'] for x in recs):,.0f} so'm "
                       f"({len(recs)} ta yozuv) — barcha kunlar va qatorlar yig'indisi."))

    # 4) tungi smenalar
    nights = [x for x in recs if x["night"]]
    if nights:
        det = "; ".join(f"{x['name'].split()[0]} {x['date']} "
                        f"({x['check_in']}→{x['check_out']})" for x in nights[:6])
        issues.append(("warn",
                       f"Tungi smena: {len(nights)} ta yozuvda chiqish kirishdan oldin. "
                       f"Soat to'g'ri hisoblandi (+24h). {det}"))

    # 5) nomlar farqi
    sid, _ = find_sheet(store, "kontragent")
    if sid and recs:
        kontr = set()
        for r, c, raw in store.iter_cells(sid):
            if c == 1 and r >= 2 and raw:
                kontr.add(_norm_name(raw))
        ish_names = {_norm_name(x["name"]) for x in recs}
        only_ish = [n for n in ish_names if not any(
            k == n or k.startswith(n) or n.startswith(k) for k in kontr)]
        if only_ish:
            issues.append(("warn",
                           "Nomlar mos kelmaydi (importda ikki xodim bo'lib ketadi): "
                           + ", ".join(sorted(only_ish)[:5])))

    # 6) kam kunlarda ma'lumot
    if recs:
        days = {x["date"] for x in recs}
        dmin, dmax = min(days), max(days)
        span = (datetime.strptime(dmax, "%Y-%m-%d")
                - datetime.strptime(dmin, "%Y-%m-%d")).days + 1
        if len(days) < span:
            issues.append(("warn",
                           f"{span} kundan faqat {len(days)} tasida ish haqi ma'lumoti bor. "
                           f"DDS balansi faqat mavjud kunlar bo'yicha hisoblanadi."))

    # 7) bo'sh varaqalar
    for key, title in (("kommunal", "Kommunal"), ("marketing", "Marketing"),
                       ("remont", "Remont"), ("arenda", "Arenda")):
        sid, nm = find_sheet(store, key)
        if sid is not None:
            filled = sum(1 for _, _, raw in store.iter_cells(sid)
                         if raw not in (None, "", ","))
            if filled <= 1:
                issues.append(("warn", f"'{nm}' varaqasi bo'sh — xarajatlar "
                                       f"Чиким kategoriyalaridan olinadi."))

    # 2/3) eski Balans varag'i bilan solishtirish
    sid, _ = find_sheet(store, "balans")
    if sid:
        vals = {}
        for r, c, raw in store.iter_cells(sid):
            if c == 0 and raw:
                vals[_norm_name(raw)] = None
            if c == 1 and _norm_name(store.get_raw(sid, r, 0) or "") in vals:
                v = _num(raw)
                if v is not None:
                    vals[_norm_name(store.get_raw(sid, r, 0))] = v
        old = vals.get("savdo")
        if old is not None and kirim:
            real = sum(t["amount"] for t in kirim)
            if abs(old - real) > 1:
                issues.append(("warn",
                               f"Eski 'Balans' varag'idagi Savdo {old:,.0f} so'm, "
                               f"lekin Кирим varag'idan {real:,.0f} so'm topildi. "
                               f"Hisobot haqiqiy ma'lumotdan hisoblandi."))
    if not kirim:
        issues.append(("warn", "Кирим varaqasi topilmadi yoki bo'sh."))
    if not chiqim:
        issues.append(("warn", "Чиким varaqasi topilmadi yoki bo'sh."))
    return issues
