"""Supabase (Postgres) REST ulanish — faqat stdlib (urllib) bilan."""
import json
import os
import urllib.error
import urllib.parse
import urllib.request

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SECRET_FILE = os.path.join(BASE_DIR, "supabase_secret.json")


def _cfg():
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE")
    if (not url or not key) and os.path.exists(SECRET_FILE):
        try:
            with open(SECRET_FILE, encoding="utf-8") as f:
                data = json.load(f)
            url = url or data.get("url")
            key = key or data.get("service_role")
        except Exception:
            pass
    return (url or "").rstrip("/"), key or ""


def available():
    url, key = _cfg()
    return bool(url and key)


def _req(method, path, payload=None, headers=None, params=None):
    url, key = _cfg()
    if not url or not key:
        raise RuntimeError("Supabase sozlanmagan (SUPABASE_URL / SERVICE_ROLE)")
    query = ""
    if params:
        query = "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url + "/rest/v1/" + path + query,
                                 method=method)
    req.add_header("apikey", key)
    req.add_header("Authorization", "Bearer " + key)
    req.add_header("Content-Type", "application/json")
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    with urllib.request.urlopen(req, data=data, timeout=25) as resp:
        body = resp.read().decode("utf-8")
    try:
        return json.loads(body) if body else None
    except ValueError:
        return None


def insert_punches(rows):
    """Takroriy yozuvlar skip qilinadi (unique constraint + ignore-duplicates)."""
    if not rows:
        return 0
    _req("POST", "attendance", payload=rows,
         headers={"Prefer": "resolution=ignore-duplicates, return=minimal"},
         params=[("on_conflict", "user_id,punch_time,status")])
    return len(rows)


def list_attendance(date_from, date_to):
    return _req("GET", "attendance", params=[
        ("select", "user_id,punch_time,status,verify_mode,device_sn"),
        ("punch_time", f"gte.{date_from} 00:00:00"),
        ("punch_time", f"lte.{date_to} 23:59:59"),
        ("order", "punch_time.asc"),
        ("limit", "5000"),
    ]) or []


def list_workers():
    return _req("GET", "workers",
                params=[("select", "user_id,full_name,position"),
                        ("order", "user_id.asc")]) or []


def upsert_worker(user_id, full_name, position=""):
    _req("POST", "workers",
         payload=[{"user_id": user_id, "full_name": full_name,
                   "position": position}],
         headers={"Prefer": "resolution=merge-duplicates, return=minimal"})
