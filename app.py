from flask import Flask, render_template, request, jsonify, send_file, session, g
import json
import os
import io
import re
import hmac
import copy
import hashlib
from datetime import datetime, date as date_cls, time as time_cls, timedelta, timezone
from collections import OrderedDict
from contextlib import contextmanager

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import psycopg2
from werkzeug.security import generate_password_hash, check_password_hash
from psycopg2.extras import RealDictCursor, Json

app = Flask(__name__)
app.json.ensure_ascii = False
app.json.sort_keys = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# ადმინ პანელის პაროლი — დააყენეთ Vercel-ის Environment Variables-ში
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
ADMIN_ACTOR = "ადმინისტრატორი"

# Vercel აყენებს VERCEL=1 — მაშინ IP-ს ვიღებთ Vercel-ის მიერ გადაწერილი ჰედერებიდან (თაღლითობისგან დაცული)
ON_VERCEL = bool(os.environ.get("VERCEL"))
TRUST_PROXY = ON_VERCEL or os.environ.get("TRUST_PROXY") == "1"

# სესიის გასაღები: SECRET_KEY, ან (თუ არ არის) ADMIN_PASSWORD-იდან და ბაზის მისამართიდან მიღებული
app.secret_key = os.environ.get("SECRET_KEY") or hashlib.sha256(
    ("brigade-ops|" + ADMIN_PASSWORD + "|" + os.environ.get("POSTGRES_URL", "")).encode("utf-8")).hexdigest()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=ON_VERCEL,
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)

CONFIG_VERSION = 4
TBILISI = "თბილისი"
LEADER_POSITIONS = {"ბრიგადირი", "სარემონტო ბრიგადის უფროსი"}
DEFAULT_POSITIONS = ["მუშა", "ბრიგადირი", "სარემონტო ბრიგადის უფროსი"]

ABSENT_LABEL = "არ გამოცხადდა"
PLACEHOLDER = "@"

# ძველ ჩანაწერებში შენახული მნიშვნელობები → Excel-ის (აღრიცხვის) ფორმატი
COMMENT_MAP = {
    "დღიური": "დღიური ანაზღაურება",
    "ბიულეტენი": "ბიულეტინი",
    "სავარაუდო ბიულეტენი": "სავარაუდო ბიულეტინი",
}


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------
def get_db_connection():
    url = os.environ.get("POSTGRES_URL")
    if not url:
        raise RuntimeError("POSTGRES_URL გარემოს ცვლადი არ არის მითითებული")
    return psycopg2.connect(url, cursor_factory=RealDictCursor)


@contextmanager
def db_cursor(commit=False):
    conn = get_db_connection()
    try:
        cur = conn.cursor()
        yield cur
        if commit:
            conn.commit()
        cur.close()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with db_cursor(commit=True) as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS records (
                db_id VARCHAR(100) PRIMARY KEY,
                id VARCHAR(100) NOT NULL,
                data JSONB NOT NULL,
                created_at DATE NOT NULL
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key VARCHAR(50) PRIMARY KEY,
                value JSONB NOT NULL,
                updated_at TIMESTAMP NOT NULL DEFAULT NOW()
            );
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS activity_log (
                id BIGSERIAL PRIMARY KEY,
                ts TIMESTAMPTZ NOT NULL,
                username VARCHAR(100) NOT NULL DEFAULT '',
                ip VARCHAR(64) NOT NULL DEFAULT '',
                action VARCHAR(40) NOT NULL,
                record_id VARCHAR(100) NOT NULL DEFAULT '',
                db_id VARCHAR(100) NOT NULL DEFAULT '',
                details TEXT NOT NULL DEFAULT '',
                user_agent VARCHAR(300) NOT NULL DEFAULT ''
            );
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS activity_log_ts ON activity_log (ts);")
        cur.execute("CREATE INDEX IF NOT EXISTS records_created_at ON records (created_at);")
    # აქტივობის ჟურნალი — მხოლოდ დამატება: ცვლილება/წაშლა ბაზის დონეზე იბლოკება
    try:
        with db_cursor(commit=True) as cur:
            cur.execute("""
                CREATE OR REPLACE FUNCTION activity_log_append_only() RETURNS trigger AS $$
                BEGIN RAISE EXCEPTION 'activity_log is append-only'; END; $$ LANGUAGE plpgsql;
            """)
            cur.execute("DROP TRIGGER IF EXISTS activity_log_no_change ON activity_log;")
            cur.execute("""
                CREATE TRIGGER activity_log_no_change BEFORE UPDATE OR DELETE ON activity_log
                FOR EACH ROW EXECUTE FUNCTION activity_log_append_only();
            """)
    except Exception as e:
        print("activity_log trigger skipped:", e)


try:
    init_db()
except Exception as e:
    print("Startup DB init skipped/failed:", e)


# ---------------------------------------------------------------------------
# Config
# ბაზაში ინახება (Vercel-ზე ფაილში ჩაწერა შეუძლებელია).
# config.json — საწყისი მონაცემები და ახალი ვერსიის ნაგულისხმევი მნიშვნელობები.
#
# {
#   "version": 4,
#   "brigades":   {"1": {"city_defaults": {"თბილისი": {"car": "", "engineer": ""}}, "members": [...]}},
#   "work_types": [{"name": "...", "worker": 0.74, "brigade": 0.3, "engineer": 0.01}],
#   "cars":       ["CJ-752-JC", ...],
#   "engineers":  ["ნაცვლიშვილი გიორგი", ...]
# }
# ---------------------------------------------------------------------------
def wt_key(name):
    """სამუშაოს სახელის შედარების გასაღები (ჰარების გარეშე) — „<20 მმ“ == „<20მმ“."""
    return re.sub(r"\s+", "", str(name or "")).lower()


def load_file_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {"version": CONFIG_VERSION, "brigades": {}, "work_types": [], "cars": [], "engineers": []}


def to_number(v, default=0.0):
    try:
        if v is None or str(v).strip() == "":
            return default
        n = float(str(v).replace(",", "."))
        return n if n == n else default  # NaN guard
    except (TypeError, ValueError):
        return default


def clean_str_list(items):
    out, seen = [], set()
    for x in items or []:
        s = str(x).strip()
        if s and s.lower() not in seen:
            seen.add(s.lower())
            out.append(s)
    return out


def _migrate_v1_to_v2(cfg, defaults):
    # სამუშაოები: ახალი სია ტარიფებით + ძველი სიიდან ის, რაც ახალში არ არის
    new_types = [dict(w) for w in defaults.get("work_types", []) if isinstance(w, dict)]
    known = {wt_key(w["name"]) for w in new_types}
    for w in cfg.get("work_types", []) or []:
        name = w.get("name") if isinstance(w, dict) else w
        if name and wt_key(name) not in known:
            known.add(wt_key(name))
            new_types.append({"name": str(name).strip(), "worker": 0, "brigade": 0, "engineer": 0})
    cfg["work_types"] = new_types
    cfg["cars"] = cfg.get("cars") or defaults.get("cars", [])
    cfg["engineers"] = cfg.get("engineers") or defaults.get("engineers", [])
    cfg["version"] = 2
    return cfg


def _migrate_v2_to_v3(cfg, defaults):
    """ნაგულისხმევი ავტომობილი/ინჟინერი ქალაქების მიხედვით + თბილისის ბრიგადების დამატება."""
    brigades = {}
    for k, v in (cfg.get("brigades") or {}).items():
        v = v or {}
        members = v.get("members", []) or []
        cd = dict(v.get("city_defaults") or {})
        car, eng = v.get("car", ""), v.get("engineer", "")
        if car or eng:   # v2: ბრიგადის დონის მნიშვნელობა → ბრიგადის ყველა ქალაქზე
            for city in {m.get("city") for m in members if m.get("city")}:
                cd.setdefault(city, {"car": car, "engineer": eng})
        brigades[str(k)] = {"city_defaults": cd, "members": members}

    # თბილისის ბრიგადები config.json-იდან — ემატება მხოლოდ ის, ვინც ბაზაში ჯერ არ არის
    existing = {str(m.get("personal_id") or m.get("name")).strip()
                for b in brigades.values() for m in b["members"]}
    for k, v in (defaults.get("brigades") or {}).items():
        tb_members = [m for m in v.get("members", []) if m.get("city") == TBILISI]
        new = [m for m in tb_members if str(m.get("personal_id") or m.get("name")).strip() not in existing]
        if not tb_members:
            continue
        b = brigades.setdefault(str(k), {"city_defaults": {}, "members": []})
        b["members"].extend(dict(m) for m in new)
        tb_def = (v.get("city_defaults") or {}).get(TBILISI)
        if tb_def:
            b["city_defaults"].setdefault(TBILISI, dict(tb_def))
    cfg["brigades"] = brigades
    cfg["version"] = 3
    return cfg


def _migrate_v3_to_v4(cfg, defaults):
    """თბილისის შემადგენლობის სრული ჩანაცვლება config.json-ის სიით. რეგიონებს არ ეხება."""
    brigades = cfg.get("brigades") or {}
    had_tbilisi = set()
    for k, b in brigades.items():
        before = len(b.get("members", []))
        b["members"] = [m for m in b.get("members", []) if m.get("city") != TBILISI]
        if len(b["members"]) != before:
            had_tbilisi.add(k)
    for k, v in (defaults.get("brigades") or {}).items():
        tb = [dict(m) for m in v.get("members", []) if m.get("city") == TBILISI]
        if not tb:
            continue
        b = brigades.setdefault(str(k), {"city_defaults": {}, "members": []})
        b["members"].extend(tb)
        tb_def = (v.get("city_defaults") or {}).get(TBILISI)
        b.setdefault("city_defaults", {})
        if tb_def and not b["city_defaults"].get(TBILISI):
            b["city_defaults"][TBILISI] = dict(tb_def)
    # ბრიგადა, რომელიც მხოლოდ თბილისის იყო და სიაში აღარ არის — იშლება
    for k in [k for k in had_tbilisi if not brigades[k]["members"]]:
        del brigades[k]
    for b in brigades.values():
        cities = {m.get("city") for m in b["members"]}
        b["city_defaults"] = {c: d for c, d in (b.get("city_defaults") or {}).items() if c in cities}
    cfg["brigades"] = brigades
    cfg["version"] = 4
    return cfg


def migrate_config(cfg):
    """ძველი კონფიგურაციის ავტომატური განახლება უახლეს ვერსიაზე, არსებული მონაცემების შენარჩუნებით."""
    if isinstance(cfg, dict) and cfg.get("version", 1) >= CONFIG_VERSION:
        return cfg, False
    cfg = copy.deepcopy(cfg) if isinstance(cfg, dict) else {}
    defaults = load_file_config()
    if cfg.get("version", 1) < 2:
        cfg = _migrate_v1_to_v2(cfg, defaults)
    if cfg.get("version", 2) < 3:
        cfg = _migrate_v2_to_v3(cfg, defaults)
    if cfg.get("version", 3) < 4:
        cfg = _migrate_v3_to_v4(cfg, defaults)
    return cfg, True


def load_config():
    try:
        with db_cursor() as cur:
            cur.execute("SELECT value FROM settings WHERE key = 'config';")
            row = cur.fetchone()
        cfg = row["value"] if row and isinstance(row["value"], dict) else None
        if cfg is None:
            cfg = load_file_config()   # პირველი გაშვება
            changed = True
        else:
            changed = False
        cfg, migrated = migrate_config(cfg)
        if changed or migrated:
            try:
                save_config(normalize_config(cfg))
            except Exception as e:
                print("Config seed/migrate failed:", e)
        return cfg
    except Exception as e:
        print("Config DB read failed, using config.json:", e)
        return migrate_config(load_file_config())[0]


def save_config(cfg):
    with db_cursor(commit=True) as cur:
        cur.execute("""
            INSERT INTO settings (key, value, updated_at)
            VALUES ('config', %s, NOW())
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW();
        """, (Json(cfg),))


def brigade_sort_key(x):
    s = str(x)
    return (0, int(s), "") if s.isdigit() else (1, 0, s)


def normalize_config(raw):
    """ამოწმებს და ასუფთავებს კონფიგურაციას (ადმინ პანელიდან ან მიგრაციიდან)."""
    if not isinstance(raw, dict):
        raise ValueError("არასწორი ფორმატი")

    brigades_in = raw.get("brigades")
    if not isinstance(brigades_in, dict):
        raise ValueError("ბრიგადების სია არასწორია")

    cars = clean_str_list(raw.get("cars"))
    engineers = clean_str_list(raw.get("engineers"))

    brigades = {}
    seen_ids = {}
    for key, data in brigades_in.items():
        bkey = str(key).strip()
        if not bkey:
            raise ValueError("ბრიგადის ნომერი ცარიელია")
        if len(bkey) > 20:
            raise ValueError(f"ბრიგადის ნომერი ძალიან გრძელია: {bkey}")
        data = data or {}
        members_in = data.get("members", [])
        if not isinstance(members_in, list):
            raise ValueError(f"ბრიგადა {bkey}: წევრების სია არასწორია")

        members = []
        for m in members_in:
            if not isinstance(m, dict):
                continue
            name = str(m.get("name", "")).strip()
            position = str(m.get("position", "")).strip()
            city = str(m.get("city", "")).strip()
            pid = str(m.get("personal_id", "")).strip()
            if not name:
                raise ValueError(f"ბრიგადა {bkey}: წევრს სახელი აკლია")
            if not city:
                raise ValueError(f"ბრიგადა {bkey}: „{name}“-ს ქალაქი აკლია")
            if not position:
                raise ValueError(f"ბრიგადა {bkey}: „{name}“-ს თანამდებობა აკლია")
            if pid:
                if pid in seen_ids:
                    raise ValueError(
                        f"პირადი # {pid} მეორდება (ბრიგადა {seen_ids[pid]} და ბრიგადა {bkey})")
                seen_ids[pid] = bkey
            members.append({"name": name, "position": position, "city": city, "personal_id": pid})

        member_cities = {m["city"] for m in members}
        city_defaults = {}
        for city, d in (data.get("city_defaults") or {}).items():
            city = str(city).strip()
            d = d or {}
            car = str(d.get("car", "") or "").strip()
            engineer = str(d.get("engineer", "") or "").strip()
            if not city or city not in member_cities or not (car or engineer):
                continue   # ქალაქი, რომელშიც ბრიგადას წევრი აღარ ჰყავს, ან ცარიელი მნიშვნელობა
            if car and car.lower() not in {c.lower() for c in cars}:
                cars.append(car)
            if engineer and engineer.lower() not in {e.lower() for e in engineers}:
                engineers.append(engineer)
            city_defaults[city] = {"car": car, "engineer": engineer}
        brigades[bkey] = {"city_defaults": city_defaults, "members": members}

    work_types, seen_wt = [], set()
    for w in raw.get("work_types", []) or []:
        if isinstance(w, str):
            w = {"name": w}
        if not isinstance(w, dict):
            continue
        name = str(w.get("name", "")).strip()
        if not name or wt_key(name) in seen_wt:
            continue
        seen_wt.add(wt_key(name))
        entry = {"name": name}
        for f in ("worker", "brigade", "engineer"):
            n = to_number(w.get(f), 0.0)
            if n < 0:
                raise ValueError(f"„{name}“: ტარიფი არ შეიძლება იყოს უარყოფითი")
            entry[f] = round(n, 6)
        work_types.append(entry)

    ordered = {k: brigades[k] for k in sorted(brigades.keys(), key=brigade_sort_key)}
    return {
        "version": CONFIG_VERSION,
        "brigades": ordered,
        "work_types": work_types,
        "cars": cars,
        "engineers": engineers,
    }


def all_cities(cfg):
    cities = set()
    for data in cfg.get("brigades", {}).values():
        for m in data.get("members", []):
            if m.get("city"):
                cities.add(m["city"])
    return sorted(cities)


def tariff_map(cfg):
    return {
        wt_key(w["name"]): {
            "worker": to_number(w.get("worker")),
            "brigade": to_number(w.get("brigade")),
            "engineer": to_number(w.get("engineer")),
        }
        for w in cfg.get("work_types", []) if isinstance(w, dict) and w.get("name")
    }


def check_admin():
    """აბრუნებს შეცდომის პასუხს, ან None-ს თუ პაროლი სწორია."""
    if not ADMIN_PASSWORD:
        return jsonify({"error": "ადმინ პანელი გამორთულია: სერვერზე ADMIN_PASSWORD ცვლადი არ არის დაყენებული"}), 503
    given = request.headers.get("X-Admin-Password", "")
    if not hmac.compare_digest(given.encode("utf-8"), ADMIN_PASSWORD.encode("utf-8")):
        return jsonify({"error": "პაროლი არასწორია"}), 401
    return None


# ---------------------------------------------------------------------------
# IP, ოპერატორები, სესია, აქტივობის ჟურნალი
# ---------------------------------------------------------------------------
def client_ip():
    """IP სერვერის მხარეს. Vercel-ზე X-Forwarded-For-ს თავად Vercel წერს — კლიენტი ვერ გააყალბებს."""
    if TRUST_PROXY:
        for h in ("X-Vercel-Forwarded-For", "X-Real-IP", "X-Forwarded-For"):
            v = request.headers.get(h, "")
            if v:
                return v.split(",")[0].strip()[:64]
    return (request.remote_addr or "")[:64]


def load_users():
    if "users" in g:
        return g.users
    users = []
    try:
        with db_cursor() as cur:
            cur.execute("SELECT value FROM settings WHERE key = 'users';")
            row = cur.fetchone()
        if row and isinstance(row["value"], list):
            users = row["value"]
    except Exception as e:
        print("users read failed:", e)
    g.users = users
    return users


def save_users(users):
    with db_cursor(commit=True) as cur:
        cur.execute("""
            INSERT INTO settings (key, value, updated_at) VALUES ('users', %s, NOW())
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW();
        """, (Json(users),))
    g.users = users


def find_user(username, users=None):
    u = str(username or "").strip().lower()
    return next((x for x in (users if users is not None else load_users())
                 if str(x.get("username", "")).lower() == u), None)


def login_required():
    """შესვლა სავალდებულოა, როცა ადმინ პანელში ერთი აქტიური ოპერატორი მაინც არსებობს."""
    return any(u.get("active", True) for u in load_users())


def _session_version(user):
    return hashlib.sha256(str(user.get("password_hash", "")).encode()).hexdigest()[:16]


def current_user():
    if "cu" in g:
        return g.cu
    user = None
    s = session.get("user")
    if isinstance(s, dict):
        u = find_user(s.get("username"))
        # პაროლის შეცვლა ან დეაქტივაცია ძველ სესიებს აუქმებს
        if u and u.get("active", True) and s.get("v") == _session_version(u):
            user = u
    g.cu = user
    return user


def actor_name():
    u = current_user()
    return u["username"] if u else ""


def log_activity(action, record_id="", db_id="", details="", username=None):
    """აქტივობის ჩაწერა (ცალკე ტრანზაქციით, ძირითადი მოქმედების შემდეგ)."""
    row = (
        datetime.now(timezone.utc),
        str(actor_name() if username is None else username)[:100],
        client_ip(),
        str(action)[:40],
        str(record_id or "")[:100],
        str(db_id or "")[:100],
        str(details or "")[:2000],
        (request.headers.get("User-Agent") or "")[:300],
    )
    try:
        with db_cursor(commit=True) as cur:
            cur.execute("""
                INSERT INTO activity_log (ts, username, ip, action, record_id, db_id, details, user_agent)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s);
            """, row)
    except Exception as e:
        print("activity log failed:", e)


def too_many_failures(action, limit=8, minutes=15):
    since = datetime.now(timezone.utc) - timedelta(minutes=minutes)
    try:
        with db_cursor() as cur:
            cur.execute("SELECT COUNT(*) AS n FROM activity_log WHERE ip = %s AND action = %s AND ts > %s;",
                        (client_ip(), action, since))
            return int((cur.fetchone() or {}).get("n") or 0) >= limit
    except Exception:
        return False


PUBLIC_PATHS = {"/", "/admin", "/logo.png", "/favicon.ico", "/api/me", "/api/login", "/api/logout"}


@app.before_request
def require_operator_login():
    path = request.path
    if path in PUBLIC_PATHS or path.startswith("/api/admin/") or path.startswith("/static/"):
        return None
    if not login_required() or current_user():
        return None
    return jsonify({"error": "საჭიროა სისტემაში შესვლა", "login_required": True}), 401


# ---------- ცვლილებების მოკლე აღწერა ჟურნალისთვის
def _short_list(items, n=5):
    items = [str(x) for x in items]
    return ", ".join(items[:n]) + (f" და კიდევ {len(items) - n}" if len(items) > n else "")


def _member_key(m):
    pid = str(m.get("personal_id") or "").strip()
    return ("id", pid) if pid else ("n", str(m.get("name") or "").strip())


def _member_status(m):
    if m.get("absent") or str(m.get("note") or "").strip():
        return str(m.get("note") or "").strip() or "არ გამოცხადდა"
    return "იმუშავა"


def describe_record(rec):
    members = rec.get("members") or []
    absent = sum(1 for m in members if _member_status(m) != "იმუშავა")
    return (f"ბრ. {rec.get('brigade', '')} · {rec.get('city', '')} · {rec.get('date', '')} · "
            f"{len(members)} წევრი" + (f" ({absent} გაცდ.)" if absent else "") +
            f" · {len(rec.get('works') or [])} სამუშაო · {rec.get('address', '')}")


RECORD_FIELDS = [("id", "ID"), ("date", "თარიღი"), ("brigade", "ბრიგადა"), ("city", "ქალაქი"),
                 ("address", "მისამართი"), ("object_name", "ობიექტი"), ("coefficient", "კოეფიციენტი"),
                 ("car", "ავტომობილი"), ("ganwesi", "განწესი"), ("engineer", "ინჟინერი"),
                 ("overall_comment", "კომენტარი"), ("note", "შენიშვნა")]


def describe_record_changes(old, new):
    parts = []
    for f, label in RECORD_FIELDS:
        a, b = str(old.get(f) or "").strip(), str(new.get(f) or "").strip()
        if a != b:
            parts.append(f"{label}: {a or '—'} → {b or '—'}")
    om = {_member_key(m): m for m in old.get("members") or [] if isinstance(m, dict)}
    nm = {_member_key(m): m for m in new.get("members") or [] if isinstance(m, dict)}
    added = [nm[k].get("name") for k in nm if k not in om]
    removed = [om[k].get("name") for k in om if k not in nm]
    status = [f"{nm[k].get('name')}: {_member_status(om[k])} → {_member_status(nm[k])}"
              for k in nm if k in om and _member_status(om[k]) != _member_status(nm[k])]
    if added:
        parts.append("დაემატა: " + _short_list(added))
    if removed:
        parts.append("ამოიშალა: " + _short_list(removed))
    if status:
        parts.append(_short_list(status, 4))
    sig = lambda w: (w.get("work_type"), w.get("start"), w.get("end"), str(w.get("quantity")), w.get("comment"))
    ow = [sig(w) for w in old.get("works") or [] if isinstance(w, dict)]
    nw = [sig(w) for w in new.get("works") or [] if isinstance(w, dict)]
    if ow != nw:
        parts.append(f"სამუშაოები შეიცვალა ({len(ow)} → {len(nw)})")
    return "; ".join(parts) or "ცვლილება არ არის"


def summarize_config_change(old, new):
    parts = []
    ob, nb = old.get("brigades") or {}, new.get("brigades") or {}
    if [k for k in nb if k not in ob]:
        parts.append("დაემატა ბრიგადა: " + _short_list([k for k in nb if k not in ob]))
    if [k for k in ob if k not in nb]:
        parts.append("წაიშალა ბრიგადა: " + _short_list([k for k in ob if k not in nb]))

    def members(cfg):
        return {_member_key(m): (k, m) for k, b in (cfg.get("brigades") or {}).items() for m in b.get("members", [])}
    om, nm = members(old), members(new)
    add = [nm[k][1].get("name") for k in nm if k not in om]
    rem = [om[k][1].get("name") for k in om if k not in nm]
    moved = [f"{nm[k][1].get('name')} ({om[k][0]}→{nm[k][0]})" for k in nm if k in om and om[k][0] != nm[k][0]]
    edited = [nm[k][1].get("name") for k in nm if k in om and om[k][0] == nm[k][0] and om[k][1] != nm[k][1]]
    if add:
        parts.append("დაემატა თანამშრომელი: " + _short_list(add))
    if rem:
        parts.append("წაიშალა თანამშრომელი: " + _short_list(rem))
    if moved:
        parts.append("გადაყვანა: " + _short_list(moved))
    if edited:
        parts.append("რედაქტირდა: " + _short_list(edited))

    ow = {wt_key(w["name"]): w for w in old.get("work_types") or [] if isinstance(w, dict)}
    nw = {wt_key(w["name"]): w for w in new.get("work_types") or [] if isinstance(w, dict)}
    if [nw[k]["name"] for k in nw if k not in ow]:
        parts.append("ახალი სამუშაო: " + _short_list([nw[k]["name"] for k in nw if k not in ow]))
    if [ow[k]["name"] for k in ow if k not in nw]:
        parts.append("წაიშალა სამუშაო: " + _short_list([ow[k]["name"] for k in ow if k not in nw]))
    tar = [f"{nw[k]['name']} ({ow[k].get('worker')}/{ow[k].get('brigade')}/{ow[k].get('engineer')} → "
           f"{nw[k].get('worker')}/{nw[k].get('brigade')}/{nw[k].get('engineer')})"
           for k in nw if k in ow and tuple(to_number(ow[k].get(f)) for f in ("worker", "brigade", "engineer"))
           != tuple(to_number(nw[k].get(f)) for f in ("worker", "brigade", "engineer"))]
    if tar:
        parts.append("ტარიფი: " + _short_list(tar, 3))
    for key, label in (("cars", "ავტომობილი"), ("engineers", "ინჟინერი")):
        a, b = set(old.get(key) or []), set(new.get(key) or [])
        if b - a:
            parts.append(f"დაემატა {label}: " + _short_list(sorted(b - a)))
        if a - b:
            parts.append(f"წაიშალა {label}: " + _short_list(sorted(a - b)))
    dchg = [k for k in nb if k in ob and (ob[k].get("city_defaults") or {}) != (nb[k].get("city_defaults") or {})]
    if dchg:
        parts.append("ნაგულისხმევი ავტომობილი/ინჟინერი: ბრ. " + _short_list(dchg))
    return "; ".join(parts) or "ცვლილება არ არის"


# ---------------------------------------------------------------------------
# სამუშაო დროის თანხვედრის კონტროლი
# ---------------------------------------------------------------------------
def _abs_intervals(rec):
    """სამუშაოების დრო აბსოლუტურ წუთებში (ღამის ცვლა შუაღამის გადაკვეთით)."""
    try:
        base = datetime.strptime(str(rec.get("date")), "%Y-%m-%d").date().toordinal() * 1440
    except (TypeError, ValueError):
        return []
    out = []
    for w in rec.get("works") or []:
        if not isinstance(w, dict):
            continue
        st, en = _parse_time(w.get("start")), _parse_time(w.get("end"))
        if not st or not en:
            continue
        a, b = st.hour * 60 + st.minute, en.hour * 60 + en.minute
        if a == b:
            continue
        if b < a:
            b += 1440
        out.append((base + a, base + b, str(w.get("work_type") or "")))
    return out


def _working_members(rec):
    return [m for m in rec.get("members") or []
            if isinstance(m, dict) and str(m.get("name") or "").strip() and _member_status(m) == "იმუშავა"]


def _same_person(a, b):
    pa, pb = str(a.get("personal_id") or "").strip(), str(b.get("personal_id") or "").strip()
    if pa and pb:
        return pa == pb
    return str(a.get("name") or "").strip() == str(b.get("name") or "").strip()


def _fmt_min(m, ref_ordinal=None):
    day, mins = divmod(int(m), 1440)
    hhmm = f"{mins // 60:02d}:{mins % 60:02d}"
    if ref_ordinal is not None and day != ref_ordinal:
        return f"{date_cls.fromordinal(day).strftime('%d.%m')} {hhmm}"
    return hhmm


def _range_str(a, b, ref_ordinal):
    start = _fmt_min(a, ref_ordinal)
    # დასრულება მომდევნო დღეზე — ღამის ცვლა: ვაჩვენებთ ჩვეულებრივ HH:MM-ს
    end = f"{(b % 1440) // 60:02d}:{(b % 1440) % 60:02d}"
    return f"{start}–{end}"


def _dur_str(minutes):
    return f"{minutes // 60}სთ {minutes % 60:02d}წ"


def find_overlaps(new_rec, exclude_db_id=None):
    new_iv = _abs_intervals(new_rec)
    people = _working_members(new_rec)
    if not new_iv or not people:
        return []
    d = datetime.strptime(str(new_rec.get("date")), "%Y-%m-%d").date()
    ref = d.toordinal()
    with db_cursor() as cur:
        cur.execute("SELECT db_id, data FROM records WHERE created_at BETWEEN %s AND %s;",
                    (d - timedelta(days=1), d + timedelta(days=1)))
        rows = cur.fetchall()

    conflicts = []
    for row in rows:
        if exclude_db_id and str(row["db_id"]) == str(exclude_db_id):
            continue
        other = row["data"] if isinstance(row["data"], dict) else {}
        o_iv = _abs_intervals(other)
        if not o_iv:
            continue
        o_people = _working_members(other)
        try:
            o_ref = datetime.strptime(str(other.get("date")), "%Y-%m-%d").date().toordinal()
        except ValueError:
            o_ref = ref
        for m in people:
            om = next((x for x in o_people if _same_person(m, x)), None)
            if not om:
                continue
            pairs = [(a, b) for a in new_iv for b in o_iv if max(a[0], b[0]) < min(a[1], b[1])]
            if not pairs:
                continue
            existing = sorted({b for _, b in pairs})
            mine = sorted({a for a, _ in pairs})
            overlaps = [(max(a[0], b[0]), min(a[1], b[1])) for a, b in pairs]
            conflicts.append({
                "name": m.get("name"),
                "personal_id": str(m.get("personal_id") or ""),
                "other_db_id": str(row["db_id"]),
                "other_id": str(other.get("id") or ""),
                "other_brigade": str(other.get("brigade") or ""),
                "other_city": str(other.get("city") or ""),
                "other_date": str(other.get("date") or ""),
                "existing_times": [_range_str(b[0], b[1], o_ref) for b in existing],
                "existing_works": sorted({b[2] for b in existing if b[2]}),
                "new_times": [_range_str(a[0], a[1], ref) for a in mine],
                "overlaps": [f"{_range_str(a, b, ref)} ({_dur_str(b - a)})" for a, b in overlaps],
                "overlap_minutes": max(b - a for a, b in overlaps),
            })
    conflicts.sort(key=lambda c: (c["name"], c["other_date"], c["other_id"]))
    return conflicts


# ---------------------------------------------------------------------------
# Excel — „აღრიცხვა“ ფორმატი (Book2.xlsx-ის სვეტების მიხედვით)
# ---------------------------------------------------------------------------
# (სათაური, სიგანე, ფორმატი, ჯგუფი)
COLUMNS = [
    ("გვარი სახელი",               22, "@",          "main"),     # A
    ("პირადი #",                   22, "@",          "main"),     # B
    ("თარიღი",                     11, "dd.mm.yyyy", "main"),     # C
    ("ბრიგადა",                     8, "0",          "main"),     # D
    ("ID:",                         9, "0",          "main"),     # E
    ("ობიექტის\nდასახელება",        18, "@",          "main"),     # F
    ("მისამართი",                   34, "@",          "main"),     # G
    ("შესრულებული\nსამუშაო",        30, "@",          "main"),     # H
    ("ბრიგ.\nწევრ.\nრაოდ.",          7, "0",          "main"),     # I
    ("დაწ.",                        7, "h:mm",       "time"),     # J
    ("დას.",                        7, "h:mm",       "time"),     # K
    ("რაოდენობა\nჯამი",             12, "0.00",       "qty"),      # L
    ("რაოდ.\nკაცი",                  9, "0.00",       "qty"),      # M
    ("რაოდენობა\nსაათი",            12, "[h]:mm",     "time"),     # N
    ("ავტომობილი",                 12, "@",          "main"),     # O
    ("განწესი",                     9, "@",          "main"),     # P
    ("კომენტარი",                  20, "@",          "main"),     # Q
    ("შენიშვნა",                    36, "@",          "main"),     # R
    ("ტარიფი\nმუშა",                 8, "0.000",      "worker"),   # S
    ("თანხა\nმუშა",                 11, "#,##0.00",   "worker"),   # T
    ("ტარიფი\nბრიგ.",                8, "0.000",      "brigade"),  # U
    ("თანხა\nბრიგ.",                11, "#,##0.00",   "brigade"),  # V
    ("ტარიფი\nინჟინერი",            11, "0.0000",     "engineer"), # W
    ("თანხა\nინჟინერი",             12, "#,##0.00",   "engineer"), # X
    ("ინჟინერი საბონუსე\nსისტემაში", 24, "@",          "main"),     # Y
    ("ქალაქი",                     12, "@",          "extra"),    # Z
    ("ქალაქის\nკოეფიციენტი",       14, "General",    "extra"),    # AA
    ("სხვა\nბრიგადიდან",            13, "@",          "extra"),    # AB
]
NCOLS = len(COLUMNS)
IDX = {name: i for i, name in enumerate([
    "name", "pid", "date", "brigade", "id", "object", "address", "work", "count",
    "start", "end", "qty", "per", "hours", "car", "ganwesi", "comment", "note",
    "t_worker", "a_worker", "t_brigade", "a_brigade", "t_engineer", "a_engineer",
    "engineer", "city", "coef", "extra",
])}

FONT_NAME = "Sylfaen"   # ქართული შრიფტი Windows-ის Excel-ში
C_NAVY = "1F3864"
HEADER_FILL = {
    "main": "1F3864", "time": "1F3864", "qty": "1F3864",
    "worker": "385723", "brigade": "1F4E79", "engineer": "5B3A7A", "extra": "595959",
}
TINT = {"worker": "EEF5E9", "brigade": "E8F0F8", "engineer": "F1ECF6"}
BAND = "F5F8FC"
ABSENT_FILL = "FCE9E7"
ABSENT_FONT = "B42318"
GRID = "D6DCE4"


def _fill(hex_):
    return PatternFill(start_color=hex_, end_color=hex_, fill_type="solid")


def _num_or_str(v):
    s = str(v or "").strip()
    return int(s) if s.isdigit() and len(s) < 16 and not (len(s) > 1 and s.startswith("0")) else s


def _parse_date(s):
    try:
        return datetime.strptime(str(s), "%Y-%m-%d")
    except (TypeError, ValueError):
        return s


def _parse_time(s):
    try:
        h, m = str(s).split(":")
        return time_cls(int(h), int(m))
    except (TypeError, ValueError):
        return None


def _duration_minutes(start, end):
    s, e = _parse_time(start), _parse_time(end)
    if not s or not e:
        return None
    return ((e.hour * 60 + e.minute) - (s.hour * 60 + s.minute)) % (24 * 60)


def _clean_qty(v):
    n = to_number(v, 0.0)
    return int(n) if n == int(n) else n


def _map_comment(v):
    v = str(v or "").strip()
    return COMMENT_MAP.get(v, v)


def build_rows(records, cfg):
    """აბრუნებს Excel-ის სტრიქონებს Book2.xlsx-ის ლოგიკით."""
    current_tariffs = tariff_map(cfg)
    zero = {"worker": 0, "brigade": 0, "engineer": 0}

    records = sorted(records, key=lambda r: (
        str(r.get("date", "")), brigade_sort_key(r.get("brigade", "")), str(r.get("id", "")), str(r.get("db_id", ""))
    ))

    out = []           # [(values, kind, record_index)]
    absent_seen = set()

    for ri, record in enumerate(records):
        members = record.get("members", []) or []
        works = record.get("works", []) or []
        rec_date = _parse_date(record.get("date", ""))
        brigade = _num_or_str(record.get("brigade", ""))
        rid = _num_or_str(record.get("id", ""))
        car = str(record.get("car", "") or "").strip()
        ganwesi = str(record.get("ganwesi", "") or "").strip()
        engineer = str(record.get("engineer", "") or "").strip()
        comment = _map_comment(record.get("overall_comment", ""))
        rec_note = str(record.get("note", "") or "").strip()
        city = record.get("city", "")
        # კოეფიციენტი — თავისუფალი ტექსტი; რიცხვი Excel-ში რიცხვად ჩაიწერება
        coef_raw = str(record.get("coefficient", "") or "").strip()
        coef = to_number(coef_raw, None) if re.fullmatch(r"-?\d+([.,]\d+)?", coef_raw) else (coef_raw or None)

        active, absent = [], []
        for m in members:
            note = str(m.get("note") or "").strip()
            entry = {
                "name": m.get("name", ""),
                "position": m.get("position", ""),
                "pid": str(m.get("personal_id", "") or ""),
                "extra": m.get("home_brigade") if m.get("extra") else "",
                "note": _map_comment(note),
            }
            (absent if (m.get("absent") or note) else active).append(entry)

        leader_first = lambda x: 0 if x["position"] in LEADER_POSITIONS else 1
        active.sort(key=leader_first)
        absent.sort(key=leader_first)
        worker_count = sum(1 for m in active if m["position"] not in LEADER_POSITIONS)

        def base(m):
            row = [None] * NCOLS
            row[IDX["name"]] = m["name"]
            row[IDX["pid"]] = f"{m['position']} {m['pid']}".strip()
            row[IDX["date"]] = rec_date
            row[IDX["brigade"]] = brigade
            row[IDX["engineer"]] = engineer or None
            row[IDX["city"]] = city
            row[IDX["coef"]] = coef
            row[IDX["extra"]] = str(m["extra"]) if m["extra"] else None
            return row

        for w in works:
            qty = _clean_qty(w.get("quantity"))
            tariff = w.get("tariff") or current_tariffs.get(wt_key(w.get("work_type")), zero)
            mins = _duration_minutes(w.get("start"), w.get("end"))
            work_note = str(w.get("comment", "") or "").strip()
            note = "; ".join(x for x in (work_note, rec_note) if x) or None

            for m in active:
                count = 1 if m["position"] in LEADER_POSITIONS else worker_count
                per = qty / count if count else 0
                row = base(m)
                row[IDX["id"]] = rid
                row[IDX["object"]] = record.get("object_name") or None
                row[IDX["address"]] = record.get("address", "")
                row[IDX["work"]] = w.get("work_type", "")
                row[IDX["count"]] = count
                row[IDX["start"]] = _parse_time(w.get("start"))
                row[IDX["end"]] = _parse_time(w.get("end"))
                row[IDX["qty"]] = qty
                row[IDX["per"]] = per
                row[IDX["hours"]] = (mins / 1440) if mins is not None else None
                row[IDX["car"]] = car or None
                row[IDX["ganwesi"]] = ganwesi or None
                row[IDX["comment"]] = comment or None
                row[IDX["note"]] = note
                for f in ("worker", "brigade", "engineer"):
                    t = to_number(tariff.get(f))
                    row[IDX["t_" + f]] = t
                    row[IDX["a_" + f]] = t * per
                out.append((row, "work", ri))

        for m in absent:
            key = (m["pid"] or m["name"], str(record.get("date", "")))
            if key in absent_seen:
                continue           # ერთ დღეში ერთი სტრიქონი გაცდენაზე
            absent_seen.add(key)
            row = base(m)
            row[IDX["id"]] = PLACEHOLDER
            row[IDX["object"]] = PLACEHOLDER
            row[IDX["address"]] = PLACEHOLDER
            row[IDX["work"]] = ABSENT_LABEL
            row[IDX["count"]] = 1
            row[IDX["start"]] = time_cls(9, 0)
            row[IDX["end"]] = time_cls(18, 0)
            row[IDX["qty"]] = 1
            row[IDX["per"]] = 1
            row[IDX["hours"]] = 9 / 24
            row[IDX["car"]] = PLACEHOLDER
            row[IDX["ganwesi"]] = PLACEHOLDER
            row[IDX["comment"]] = m["note"] or None
            for f in ("worker", "brigade", "engineer"):
                row[IDX["t_" + f]] = 0
                row[IDX["a_" + f]] = 0
            out.append((row, "absent", ri))

    return out, records


def _style_header(ws, row_idx, height=42):
    thin_white = Side(style="thin", color="FFFFFF")
    for i, (title, width, _, group) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=row_idx, column=i, value=title)
        c.font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
        c.fill = _fill(HEADER_FILL[group])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin_white, right=thin_white)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row_idx].height = height


def build_main_sheet(ws, rows):
    ws.title = "აღრიცხვა"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    _style_header(ws, 1)

    grid = Side(style="thin", color=GRID)
    border = Border(left=grid, right=grid, top=grid, bottom=grid)
    body_font = Font(name=FONT_NAME, size=9)
    bold_font = Font(name=FONT_NAME, size=9, bold=True)
    absent_font = Font(name=FONT_NAME, size=9, bold=True, color=ABSENT_FONT)
    left_cols = {IDX["name"], IDX["pid"], IDX["object"], IDX["address"], IDX["work"], IDX["note"], IDX["engineer"]}

    band_on, last_ri = False, None
    for r_i, (values, kind, ri) in enumerate(rows, start=2):
        if ri != last_ri:
            band_on, last_ri = (not band_on), ri
        is_leader = str(values[IDX["pid"]] or "").split(" ")[0] in LEADER_POSITIONS
        for c_i, v in enumerate(values):
            cell = ws.cell(row=r_i, column=c_i + 1, value=v)
            _, _, fmt, group = COLUMNS[c_i]
            cell.number_format = fmt
            cell.border = border
            cell.font = body_font
            cell.alignment = Alignment(
                horizontal="left" if c_i in left_cols else "center",
                vertical="center",
                wrap_text=(c_i in (IDX["address"], IDX["note"])),
            )
            if kind == "absent":
                cell.fill = _fill(ABSENT_FILL)
                if c_i in (IDX["work"], IDX["comment"]):
                    cell.font = absent_font
            elif group in TINT:
                cell.fill = _fill(TINT[group])
            elif band_on:
                cell.fill = _fill(BAND)
        if is_leader:
            ws.cell(row=r_i, column=IDX["name"] + 1).font = bold_font

    last = len(rows) + 1
    # ჯამები (SUBTOTAL — ფილტრის გათვალისწინებით)
    total_row = last + 2
    top = Side(style="medium", color=C_NAVY)
    ws.cell(row=total_row, column=1, value="ჯამი (გაფილტრული)").font = Font(name=FONT_NAME, size=10, bold=True, color=C_NAVY)
    for key in ("hours", "a_worker", "a_brigade", "a_engineer"):
        col = IDX[key] + 1
        L = get_column_letter(col)
        c = ws.cell(row=total_row, column=col, value=f"=SUBTOTAL(9,{L}2:{L}{last})")
        c.number_format = COLUMNS[IDX[key]][2]
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=C_NAVY)
        c.alignment = Alignment(horizontal="center")
    for col in range(1, NCOLS + 1):
        ws.cell(row=total_row, column=col).border = Border(top=top)
        ws.cell(row=total_row, column=col).fill = _fill("EEF2F8")

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(NCOLS)}{max(last, 2)}"
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.3


def build_summary_sheet(wb, rows, records):
    ws = wb.create_sheet("შეჯამება")
    ws.sheet_view.showGridLines = False
    dates = [r.get("date") for r in records if r.get("date")]
    period = ""
    if dates:
        d1, d2 = min(dates), max(dates)
        fmt = lambda d: datetime.strptime(d, "%Y-%m-%d").strftime("%d.%m.%Y")
        period = fmt(d1) if d1 == d2 else f"{fmt(d1)} – {fmt(d2)}"

    ws["A1"] = "შესრულებული სამუშაოების შეჯამება"
    ws["A1"].font = Font(name=FONT_NAME, size=15, bold=True, color=C_NAVY)
    ws["A2"] = f"პერიოდი: {period}   ·   ჩანაწერები: {len(records)}   ·   შექმნილია: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws["A2"].font = Font(name=FONT_NAME, size=9, color="595959")

    headers = [("გვარი სახელი", 24), ("პირადი #", 22), ("ბრიგადა", 10), ("სამუშაო\nდღეები", 10),
               ("გაცდენა\n(დღე)", 10), ("საათები", 10), ("თანხა\nმუშა", 13), ("თანხა\nბრიგ.", 13), ("თანხა\nინჟინერი", 13)]
    groups = ["main"] * 6 + ["worker", "brigade", "engineer"]
    hr = 4
    for i, ((t, w), g) in enumerate(zip(headers, groups), start=1):
        c = ws.cell(row=hr, column=i, value=t)
        c.font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
        c.fill = _fill(HEADER_FILL[g])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[hr].height = 34

    people = OrderedDict()
    for values, kind, _ in rows:
        key = values[IDX["pid"]] or values[IDX["name"]]
        p = people.setdefault(key, {
            "name": values[IDX["name"]], "pid": values[IDX["pid"]], "brigades": set(),
            "work_days": set(), "absent_days": set(), "hours": 0.0, "aw": 0.0, "ab": 0.0, "ae": 0.0,
        })
        p["brigades"].add(str(values[IDX["brigade"]]))
        d = values[IDX["date"]]
        if kind == "absent":
            p["absent_days"].add(d)
        else:
            p["work_days"].add(d)
            p["hours"] += values[IDX["hours"]] or 0
            p["aw"] += values[IDX["a_worker"]] or 0
            p["ab"] += values[IDX["a_brigade"]] or 0
            p["ae"] += values[IDX["a_engineer"]] or 0

    grid = Side(style="thin", color=GRID)
    border = Border(left=grid, right=grid, top=grid, bottom=grid)
    fmts = ["@", "@", "@", "0", "0", "[h]:mm", "#,##0.00", "#,##0.00", "#,##0.00"]
    ordered = sorted(people.values(), key=lambda p: (sorted(p["brigades"], key=brigade_sort_key)[0], p["name"]))
    r = hr + 1
    for i, p in enumerate(ordered):
        vals = [p["name"], p["pid"], ", ".join(sorted(p["brigades"], key=brigade_sort_key)),
                len(p["work_days"]), len(p["absent_days"] - p["work_days"]), p["hours"], p["aw"], p["ab"], p["ae"]]
        for c_i, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=c_i, value=v)
            c.number_format = fmts[c_i - 1]
            c.border = border
            c.font = Font(name=FONT_NAME, size=9)
            c.alignment = Alignment(horizontal="left" if c_i <= 2 else "center", vertical="center")
            if c_i >= 7:
                c.fill = _fill(TINT[groups[c_i - 1]])
            elif i % 2:
                c.fill = _fill(BAND)
        r += 1

    top = Side(style="medium", color=C_NAVY)
    ws.cell(row=r, column=1, value="ჯამი")
    for c_i in range(1, len(headers) + 1):
        c = ws.cell(row=r, column=c_i)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=C_NAVY)
        c.border = Border(top=top)
        c.fill = _fill("EEF2F8")
        c.alignment = Alignment(horizontal="left" if c_i == 1 else "center")
        if c_i >= 4 and r > hr + 1:
            L = get_column_letter(c_i)
            c.value = f"=SUM({L}{hr + 1}:{L}{r - 1})"
            c.number_format = fmts[c_i - 1]
    ws.freeze_panes = f"A{hr + 1}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build_tariff_sheet(wb, cfg):
    ws = wb.create_sheet("ტარიფები")
    ws.sheet_view.showGridLines = False
    heads = [("შესრულებული სამუშაო", 54, "main", "@"), ("ტარიფი\nმუშა", 11, "worker", "0.000"),
             ("ტარიფი\nბრიგ.", 11, "brigade", "0.000"), ("ტარიფი\nინჟინერი", 11, "engineer", "0.0000")]
    for i, (t, w, g, _) in enumerate(heads, start=1):
        c = ws.cell(row=1, column=i, value=t)
        c.font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
        c.fill = _fill(HEADER_FILL[g])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 32
    grid = Side(style="thin", color=GRID)
    border = Border(left=grid, right=grid, top=grid, bottom=grid)
    for r, w in enumerate(cfg.get("work_types", []), start=2):
        vals = [w.get("name"), to_number(w.get("worker")), to_number(w.get("brigade")), to_number(w.get("engineer"))]
        for c_i, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=c_i, value=v)
            c.number_format = heads[c_i - 1][3]
            c.border = border
            c.font = Font(name=FONT_NAME, size=9)
            c.alignment = Alignment(horizontal="left" if c_i == 1 else "center", vertical="center")
            if r % 2 == 1:
                c.fill = _fill(BAND)
    ws.freeze_panes = "A2"
    ws.cell(row=len(cfg.get("work_types", [])) + 3, column=1,
            value="* მიმდინარე ტარიფები. ჩანაწერებში გამოიყენება შენახვის მომენტის ტარიფი.").font = \
        Font(name=FONT_NAME, size=8, italic=True, color="7F7F7F")


def generate_excel_from_records(records, cfg=None):
    cfg = cfg or load_config()
    rows, records = build_rows(records, cfg)
    wb = openpyxl.Workbook()
    build_main_sheet(wb.active, rows)
    build_summary_sheet(wb, rows, records)
    build_tariff_sheet(wb, cfg)
    wb.active = 0
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------------------------
# Static
# ---------------------------------------------------------------------------
@app.route("/logo.png")
@app.route("/favicon.ico")
def serve_logo():
    logo_path = os.path.join(BASE_DIR, "logo.png")
    if os.path.exists(logo_path):
        resp = send_file(logo_path, mimetype="image/png")
        resp.headers["Cache-Control"] = "public, max-age=86400"
        return resp
    return "", 404


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/admin")
def admin():
    return render_template("admin.html")


# ---------------------------------------------------------------------------
# Public read API
# ---------------------------------------------------------------------------
@app.route("/get_brigades")
def get_brigades():
    config = load_config()
    return jsonify(sorted(config.get("brigades", {}).keys(), key=brigade_sort_key))


@app.route("/get_brigade_cities")
def get_brigade_cities():
    brigade = request.args.get("brigade")
    if not brigade:
        return jsonify([])
    config = load_config()
    members = config.get("brigades", {}).get(brigade, {}).get("members", [])
    return jsonify(sorted({m["city"] for m in members if m.get("city")}))


@app.route("/get_brigade_members")
def get_brigade_members():
    brigade = request.args.get("brigade")
    city = request.args.get("city", "")
    if not brigade:
        return jsonify([])
    config = load_config()
    members = config.get("brigades", {}).get(brigade, {}).get("members", [])
    if city:
        members = [m for m in members if m.get("city") == city]
    return jsonify(members)


@app.route("/get_all_members")
def get_all_members():
    config = load_config()
    all_members = []
    for brigade_num, data in config.get("brigades", {}).items():
        for m in data.get("members", []):
            all_members.append({
                "name": m.get("name", ""),
                "position": m.get("position", ""),
                "personal_id": m.get("personal_id", ""),
                "city": m.get("city", ""),
                "brigade": brigade_num,
            })
    all_members.sort(key=lambda m: (brigade_sort_key(m["brigade"]), m["name"]))
    return jsonify(all_members)


@app.route("/get_work_types")
def get_work_types():
    return jsonify([w["name"] for w in load_config().get("work_types", []) if isinstance(w, dict)])


@app.route("/get_options")
def get_options():
    """ავტომობილები, ინჟინრები და ბრიგადების ნაგულისხმევი მნიშვნელობები ფორმისთვის."""
    cfg = load_config()
    return jsonify({
        "cars": cfg.get("cars", []),
        "engineers": cfg.get("engineers", []),
        # {"1": {"თბილისი": {"car": "...", "engineer": "..."}}}
        "brigade_defaults": {k: v.get("city_defaults", {}) for k, v in cfg.get("brigades", {}).items()},
    })


@app.route("/get_cities")
def get_cities():
    """ყველა ქალაქი, დაყოფილი თბილისად და რეგიონებად (Excel-ის ფილტრისთვის)."""
    cities = all_cities(load_config())
    return jsonify({"tbilisi": TBILISI, "regions": [c for c in cities if c != TBILISI]})


# ---------------------------------------------------------------------------
# ოპერატორის შესვლა
# ---------------------------------------------------------------------------
@app.route("/api/me")
def api_me():
    u = current_user()
    return jsonify({
        "login_required": login_required(),
        "user": {"username": u["username"], "display_name": u.get("display_name") or u["username"]} if u else None,
    })


@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json(silent=True) or {}
    username = str(data.get("username") or "").strip()[:100]
    password = str(data.get("password") or "")
    if too_many_failures("login_failed"):
        return jsonify({"error": "ძალიან ბევრი წარუმატებელი მცდელობა. სცადეთ 15 წუთში."}), 429
    u = find_user(username)
    if not u or not u.get("active", True) or not check_password_hash(u.get("password_hash", ""), password):
        log_activity("login_failed", details="არასწორი მომხმარებელი ან პაროლი", username=username)
        return jsonify({"error": "მომხმარებელი ან პაროლი არასწორია"}), 401
    session.clear()
    session.permanent = True
    session["user"] = {"username": u["username"], "v": _session_version(u)}
    g.pop("cu", None)
    log_activity("login", details=u.get("display_name") or "", username=u["username"])
    return jsonify({"success": True, "user": {"username": u["username"], "display_name": u.get("display_name") or u["username"]}})


@app.route("/api/logout", methods=["POST"])
def api_logout():
    if current_user():
        log_activity("logout")
    session.clear()
    return jsonify({"success": True})


# ---------------------------------------------------------------------------
# Records API
# ---------------------------------------------------------------------------
@app.route("/api/records", methods=["GET"])
def api_get_records():
    try:
        with db_cursor() as cur:
            cur.execute("SELECT data FROM records ORDER BY created_at DESC, db_id DESC;")
            rows = cur.fetchall()
        return jsonify({"records": [r["data"] for r in rows]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/records/save", methods=["POST"])
def api_save_record():
    data = request.get_json(silent=True)
    if not data or not str(data.get("id", "")).strip():
        return jsonify({"error": "არასწორი მონაცემები"}), 400

    for field in ("brigade", "city", "date", "address"):
        if not str(data.get(field, "")).strip():
            return jsonify({"error": f"აკლია სავალდებულო ველი: {field}"}), 400
    if not data.get("works"):
        return jsonify({"error": "დაამატეთ მინიმუმ ერთი სამუშაო"}), 400

    rec_id = str(data["id"]).strip()
    db_id = str(data.get("db_id") or f"{rec_id}_{int(datetime.now().timestamp() * 1000)}")
    data["id"] = rec_id
    data["db_id"] = db_id
    confirm_overlap = bool(data.pop("confirm_overlap", False))

    try:
        rec_date = datetime.strptime(str(data.get("date")), "%Y-%m-%d").date()
    except ValueError:
        return jsonify({"error": "თარიღის ფორმატი არასწორია"}), 400

    try:
        # ტარიფის „გაყინვა“ შენახვის მომენტში — ტარიფის შემდგომი ცვლილება ძველ ჩანაწერებს არ ცვლის.
        # რედაქტირებისას იმავე სამუშაოს ძველი ტარიფი რჩება.
        # სამუშაო დროის თანხვედრა — არ ბლოკავს, მაგრამ საჭიროებს დადასტურებას
        conflicts = find_overlaps(data, exclude_db_id=db_id)
        if conflicts and not confirm_overlap:
            return jsonify({"success": False, "overlap": True, "conflicts": conflicts}), 409

        old_tariffs = {}
        with db_cursor() as cur:
            cur.execute("SELECT data FROM records WHERE db_id = %s;", (db_id,))
            old = cur.fetchone()
        old_data = old["data"] if old and isinstance(old.get("data"), dict) else None
        now_iso = datetime.now(timezone.utc).isoformat(timespec="seconds")
        if old_data:
            data["created_by"] = old_data.get("created_by", "")
            data["created_at_ts"] = old_data.get("created_at_ts", "")
        else:
            data["created_by"] = actor_name()
            data["created_at_ts"] = now_iso
        data["updated_by"] = actor_name()
        data["updated_at_ts"] = now_iso
        if old_data:
            for w in old_data.get("works", []) or []:
                if isinstance(w, dict) and w.get("tariff"):
                    old_tariffs[wt_key(w.get("work_type"))] = w["tariff"]
        current = tariff_map(load_config())
        for w in data.get("works", []):
            if isinstance(w, dict):
                k = wt_key(w.get("work_type"))
                w["tariff"] = old_tariffs.get(k) or current.get(k) or {"worker": 0, "brigade": 0, "engineer": 0}

        with db_cursor(commit=True) as cur:
            cur.execute("""
                INSERT INTO records (db_id, id, data, created_at)
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (db_id) DO UPDATE
                SET id = EXCLUDED.id, data = EXCLUDED.data, created_at = EXCLUDED.created_at;
            """, (db_id, rec_id, Json(data), rec_date))

        details = describe_record_changes(old_data, data) if old_data else describe_record(data)
        if conflicts:
            details += " | ⚠ დროის თანხვედრა დადასტურდა: " + _short_list(
                [f"{c['name']} — ID {c['other_id']} (ბრ. {c['other_brigade']}, {', '.join(c['overlaps'])})" for c in conflicts], 3)
        log_activity("record_update" if old_data else "record_create", rec_id, db_id, details)
        return jsonify({"success": True, "db_id": db_id, "overlaps_confirmed": len(conflicts)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/records/delete/<path:record_id>", methods=["DELETE"])
def api_delete_record(record_id):
    try:
        with db_cursor(commit=True) as cur:
            cur.execute("SELECT id, data FROM records WHERE db_id = %s;", (str(record_id),))
            old = cur.fetchone()
            cur.execute("DELETE FROM records WHERE db_id = %s;", (str(record_id),))
            deleted = cur.rowcount
        if not deleted:
            return jsonify({"error": "ჩანაწერი ვერ მოიძებნა"}), 404
        old_data = old["data"] if old and isinstance(old.get("data"), dict) else {}
        log_activity("record_delete", old.get("id") if old else "", record_id, describe_record(old_data) if old_data else "")
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ჩანაწერებში გაცდენის მიზეზის მასობრივი ცვლილება (ერთი პიროვნება, რამდენიმე ჩანაწერი)
ABSENCE_REASONS = ["შვებულება", "ბიულეტინი", "სავარაუდო ბიულეტინი", "უხელფასო შვებულება", "არასაპატიო", "საპატიო"]
PRESENT_STATUS = "__present__"


@app.route("/api/records/bulk-absence", methods=["POST"])
def api_bulk_absence():
    data = request.get_json(silent=True) or {}
    db_ids = list(dict.fromkeys(str(x) for x in (data.get("db_ids") or []) if str(x).strip()))
    pid = str(data.get("personal_id") or "").strip()
    name = str(data.get("name") or "").strip()
    status = str(data.get("status") or "").strip()

    if not db_ids:
        return jsonify({"error": "მონიშნეთ მინიმუმ ერთი ჩანაწერი"}), 400
    if len(db_ids) > 3000:
        return jsonify({"error": "ერთდროულად მაქსიმუმ 3000 ჩანაწერი"}), 400
    if not (pid or name):
        return jsonify({"error": "პიროვნება არ არის მითითებული"}), 400
    if not status or len(status) > 100:
        return jsonify({"error": "აირჩიეთ ახალი სტატუსი"}), 400

    def same_person(m):
        mp = str(m.get("personal_id") or "").strip()
        if pid and mp:
            return mp == pid
        return str(m.get("name") or "").strip() == name

    updated, unchanged, missing = [], 0, 0
    changes = []   # (rec_id, db_id, old_status)
    try:
        with db_cursor(commit=True) as cur:
            for db_id in db_ids:
                cur.execute("SELECT data FROM records WHERE db_id = %s FOR UPDATE;", (db_id,))
                row = cur.fetchone()
                rec = row and row.get("data")
                if not isinstance(rec, dict):
                    missing += 1
                    continue
                changed = False
                found = False
                old_status = None
                for m in rec.get("members", []) or []:
                    if not isinstance(m, dict) or not same_person(m):
                        continue
                    found = True
                    old_status = _member_status(m)
                    if status == PRESENT_STATUS:
                        new_absent, new_note = False, ""
                    else:
                        new_absent, new_note = True, status
                    if bool(m.get("absent")) != new_absent or str(m.get("note") or "").strip() != new_note:
                        m["absent"], m["note"] = new_absent, new_note
                        changed = True
                if not found:
                    missing += 1
                elif not changed:
                    unchanged += 1
                else:
                    rec["updated_by"] = actor_name()
                    rec["updated_at_ts"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
                    cur.execute("UPDATE records SET data = %s WHERE db_id = %s;", (Json(rec), db_id))
                    updated.append(db_id)
                    changes.append((rec.get("id", ""), db_id, old_status))
        new_label = "იმუშავა" if status == PRESENT_STATUS else status
        for rid, dbid, old_status in changes:
            log_activity("bulk_absence", rid, dbid, f"{name or pid}: {old_status} → {new_label} (მასობრივი ცვლილება)")
        return jsonify({"success": True, "updated": len(updated), "unchanged": unchanged,
                        "missing": missing, "updated_ids": updated})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "მონაცემები არასწორია"}), 400

    records = data.get("records", [])
    if not records:
        return jsonify({"error": "ჩანაწერები ვერ მოიძებნა"}), 400

    for rec in records:
        for field in ("id", "brigade", "city", "date", "address"):
            if not rec.get(field):
                return jsonify({"error": f"ჩანაწერს ID {rec.get('id', '?')} აკლია ველი: {field}"}), 400
        if not rec.get("works"):
            return jsonify({"error": f"ჩანაწერს ID {rec.get('id')} არ აქვს სამუშაოები"}), 400

    try:
        output = generate_excel_from_records(records)
    except Exception as exc:
        return jsonify({"error": f"Excel-ის გენერირება ვერ მოხერხდა: {exc}"}), 500

    label = str(data.get("label", "")).strip().replace("/", "-")[:40]
    dates = sorted(str(r.get("date", "")) for r in records)
    log_activity("excel_export", details=f"{len(records)} ჩანაწერი · {dates[0]} – {dates[-1]}" + (f" · {label}" if label else ""))
    filename = f"ანგარიში{('_' + label) if label else ''}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        output,
        download_name=filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Admin API
# ---------------------------------------------------------------------------
@app.route("/api/admin/login", methods=["POST"])
def admin_login():
    if too_many_failures("admin_login_failed"):
        return jsonify({"error": "ძალიან ბევრი წარუმატებელი მცდელობა. სცადეთ 15 წუთში."}), 429
    err = check_admin()
    if err:
        if err[1] == 401:
            log_activity("admin_login_failed", details="არასწორი პაროლი", username=ADMIN_ACTOR)
        return err
    log_activity("admin_login", username=ADMIN_ACTOR)
    return jsonify({"success": True})


@app.route("/api/admin/config", methods=["GET"])
def admin_get_config():
    err = check_admin()
    if err:
        return err
    cfg = load_config()
    positions = set(DEFAULT_POSITIONS)
    for data in cfg.get("brigades", {}).values():
        for m in data.get("members", []):
            if m.get("position"):
                positions.add(m["position"])
    return jsonify({"config": cfg, "cities": all_cities(cfg), "positions": sorted(positions)})


@app.route("/api/admin/config", methods=["PUT"])
def admin_save_config():
    err = check_admin()
    if err:
        return err
    try:
        cfg = normalize_config(request.get_json(silent=True))
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    try:
        old_cfg = load_config()
        save_config(cfg)
    except Exception as e:
        return jsonify({"error": f"შენახვა ვერ მოხერხდა: {e}"}), 500
    log_activity("config_save", details=summarize_config_change(old_cfg, cfg), username=ADMIN_ACTOR)
    return jsonify({"success": True, "config": cfg})


@app.route("/api/admin/reset", methods=["POST"])
def admin_reset_config():
    """ბაზის პარამეტრების დაბრუნება config.json-ის მდგომარეობაზე."""
    err = check_admin()
    if err:
        return err
    try:
        cfg = normalize_config(migrate_config(load_file_config())[0])
        save_config(cfg)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    log_activity("config_reset", details="პარამეტრები დაბრუნდა config.json-ზე", username=ADMIN_ACTOR)
    return jsonify({"success": True, "config": cfg})


@app.route("/api/admin/backup", methods=["GET"])
def admin_backup():
    """სრული სარეზერვო ასლი: ყველა ჩანაწერი + პარამეტრები (ერთი JSON ფაილი)."""
    err = check_admin()
    if err:
        return err
    try:
        with db_cursor() as cur:
            cur.execute("SELECT db_id, id, data, created_at FROM records ORDER BY created_at, db_id;")
            rows = cur.fetchall()
        payload = {
            "created": datetime.now().isoformat(timespec="seconds"),
            "records_count": len(rows),
            "records": [{"db_id": r["db_id"], "id": r["id"], "created_at": str(r["created_at"]), "data": r["data"]} for r in rows],
            "config": load_config(),
        }
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    log_activity("backup_download", details=f"{len(rows)} ჩანაწერი", username=ADMIN_ACTOR)
    buf = io.BytesIO(json.dumps(payload, ensure_ascii=False, indent=1).encode("utf-8"))
    name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
    return send_file(buf, download_name=name, as_attachment=True, mimetype="application/json")


@app.route("/api/admin/export", methods=["GET"])
def admin_export_config():
    err = check_admin()
    if err:
        return err
    buf = io.BytesIO(json.dumps(load_config(), ensure_ascii=False, indent=2).encode("utf-8"))
    return send_file(buf, download_name="config.json", as_attachment=True, mimetype="application/json")


# ---------------------------------------------------------------------------
# Admin: ოპერატორები
# ---------------------------------------------------------------------------
USERNAME_RE = re.compile(r"^[A-Za-z0-9._\-ა-ჰ]{3,40}$")


def _public_user(u, last_login=None):
    return {"username": u["username"], "display_name": u.get("display_name") or "",
            "active": u.get("active", True), "created": u.get("created", ""), "last_login": last_login}


@app.route("/api/admin/users", methods=["GET"])
def admin_list_users():
    err = check_admin()
    if err:
        return err
    users = load_users()
    last = {}
    try:
        with db_cursor() as cur:
            cur.execute("SELECT username, MAX(ts) AS ts FROM activity_log WHERE action = 'login' GROUP BY username;")
            for r in cur.fetchall():
                last[str(r["username"]).lower()] = str(r["ts"])
    except Exception:
        pass
    return jsonify({"users": [_public_user(u, last.get(u["username"].lower())) for u in users],
                    "login_required": login_required()})


@app.route("/api/admin/users", methods=["POST"])
def admin_create_user():
    err = check_admin()
    if err:
        return err
    data = request.get_json(silent=True) or {}
    username = str(data.get("username") or "").strip()
    display = str(data.get("display_name") or "").strip()[:80]
    password = str(data.get("password") or "")
    if not USERNAME_RE.match(username):
        return jsonify({"error": "მომხმარებლის სახელი: 3–40 სიმბოლო (ასოები, ციფრები, . _ -), ჰარის გარეშე"}), 400
    if len(password) < 6:
        return jsonify({"error": "პაროლი მინიმუმ 6 სიმბოლო"}), 400
    users = list(load_users())
    if find_user(username, users):
        return jsonify({"error": "ასეთი მომხმარებელი უკვე არსებობს"}), 400
    users.append({"username": username, "display_name": display, "active": True,
                  "password_hash": generate_password_hash(password, method="pbkdf2:sha256"),
                  "created": datetime.now(timezone.utc).isoformat(timespec="seconds")})
    save_users(users)
    log_activity("user_create", details=f"{username}" + (f" ({display})" if display else ""), username=ADMIN_ACTOR)
    return jsonify({"success": True})


@app.route("/api/admin/users/<username>", methods=["PUT"])
def admin_update_user(username):
    err = check_admin()
    if err:
        return err
    data = request.get_json(silent=True) or {}
    users = [dict(u) for u in load_users()]
    u = find_user(username, users)
    if not u:
        return jsonify({"error": "მომხმარებელი ვერ მოიძებნა"}), 404
    notes = []
    if "display_name" in data:
        u["display_name"] = str(data.get("display_name") or "").strip()[:80]
        notes.append("სახელი შეიცვალა")
    if "active" in data:
        u["active"] = bool(data.get("active"))
        notes.append("გააქტიურდა" if u["active"] else "დაიბლოკა")
    if data.get("password"):
        if len(str(data["password"])) < 6:
            return jsonify({"error": "პაროლი მინიმუმ 6 სიმბოლო"}), 400
        u["password_hash"] = generate_password_hash(str(data["password"]), method="pbkdf2:sha256")
        notes.append("პაროლი შეიცვალა")
    save_users(users)
    log_activity("user_update", details=f"{u['username']}: " + ", ".join(notes or ["—"]), username=ADMIN_ACTOR)
    return jsonify({"success": True})


@app.route("/api/admin/users/<username>", methods=["DELETE"])
def admin_delete_user(username):
    err = check_admin()
    if err:
        return err
    users = list(load_users())
    u = find_user(username, users)
    if not u:
        return jsonify({"error": "მომხმარებელი ვერ მოიძებნა"}), 404
    save_users([x for x in users if x is not u])
    log_activity("user_delete", details=u["username"], username=ADMIN_ACTOR)
    return jsonify({"success": True})


# ---------------------------------------------------------------------------
# Admin: აქტივობის ჟურნალი (მხოლოდ წაკითხვა — ცვლილება/წაშლა არ არსებობს)
# ---------------------------------------------------------------------------
@app.route("/api/admin/activity", methods=["GET"])
def admin_activity():
    err = check_admin()
    if err:
        return err
    a = request.args
    where, params = [], []
    if a.get("user"):
        where.append("LOWER(username) = LOWER(%s)")
        params.append(a["user"].strip())
    if a.get("ip"):
        where.append("ip LIKE %s")
        params.append("%" + a["ip"].strip() + "%")
    if a.get("action"):
        where.append("action = %s")
        params.append(a["action"].strip())
    if a.get("record"):
        where.append("(record_id LIKE %s OR db_id LIKE %s)")
        params += ["%" + a["record"].strip() + "%"] * 2
    try:
        tz_min = int(a.get("tz", "240"))          # ბრაუზერის დროის სარტყელი (თბილისი = +240)
    except ValueError:
        tz_min = 240
    tz = timezone(timedelta(minutes=tz_min))
    for key, op, add in (("from", ">=", 0), ("to", "<", 1)):
        if a.get(key):
            try:
                d = datetime.strptime(a[key], "%Y-%m-%d") + timedelta(days=add)
                where.append(f"ts {op} %s")
                params.append(d.replace(tzinfo=tz).astimezone(timezone.utc))
            except ValueError:
                pass
    sql_where = (" WHERE " + " AND ".join(where)) if where else ""
    try:
        limit = max(1, min(int(a.get("limit", 100)), 500))
        offset = max(0, int(a.get("offset", 0)))
    except ValueError:
        limit, offset = 100, 0
    try:
        with db_cursor() as cur:
            cur.execute("SELECT COUNT(*) AS n FROM activity_log" + sql_where + ";", tuple(params))
            total = int(cur.fetchone()["n"])
            cur.execute("SELECT id, ts, username, ip, action, record_id, db_id, details, user_agent FROM activity_log"
                        + sql_where + " ORDER BY ts DESC, id DESC LIMIT %s OFFSET %s;", tuple(params) + (limit, offset))
            rows = cur.fetchall()
            cur.execute("SELECT DISTINCT username FROM activity_log WHERE username <> '' ORDER BY username;")
            users = [r["username"] for r in cur.fetchall()]
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    out = []
    for r in rows:
        ts = r["ts"]
        if isinstance(ts, datetime):
            ts = (ts if ts.tzinfo else ts.replace(tzinfo=timezone.utc)).isoformat()
        out.append({"id": r["id"], "ts": str(ts), "username": r["username"], "ip": r["ip"], "action": r["action"],
                    "record_id": r["record_id"], "db_id": r["db_id"], "details": r["details"],
                    "user_agent": r["user_agent"]})
    return jsonify({"rows": out, "total": total, "users": users, "client_ip": client_ip()})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
