#!/usr/bin/env python3
"""
نفس منطق app.js لكن على السيرفر (GitHub Actions).

المدخلات:
  STR Loss.xlsx            (جذر المستودع)  أو templates/Structure.xlsx
  data/UTL.xlsx            Hold + Other + AUX + ACW  (UL_lo / UL_Date)
  data/IR.xlsx             assigned_to / added_by / added_on
  data/Compensation.xlsx   Comp_ID / Comp_Da / Comp_Du
  data/Schedule.xlsx       تقرير Scheduled Time per Agent الخام (B=ID, C=تاريخ/كود, J=Duration)
  config.json              teleSchFactor / secondsPerTicket / scheduleCodes

المخرجات:
  output/Final_Report.xlsx  (شيتين: Matrix و Daily_Long)

المعادلة:
  Loss Time = Tele-SCH × 90% − (System + Talk Time + Comp)   (لا يقل عن صفر)
  System    = TKT × 90 ثانية  (= 0.00104166666666667 يوم)
"""
import json
import math
import re
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output" / "Final_Report.xlsx"
STRUCTURE_CANDIDATES = [ROOT / "STR Loss.xlsx", ROOT / "templates" / "Structure.xlsx"]

DEFAULT_CONFIG = {"teleSchFactor": 0.9, "secondsPerTicket": 90, "scheduleCodes": []}

ALIASES = {
    "structureId": ["Teleopti ID", "Teleopti", "ST_ID"],
    "loginId": ["Login ID", "Login", "UL_lo", "Username"],
    "perm": ["Perm"],
    "ttsUser": ["TTS User", "TTS"],
    "bssUser": ["BSS User", "BSS"],
    "group": ["Group"],
    "agentName": ["Agent Name", "Agent"],
    "status": ["Status"],
    "tlId": ["TL ID", "TL Id"],
    "tlName": ["TL Name", "Team Leader"],
    "structureDate": ["ST_D"],
    "structureDuration": ["ST_Du", "ST Duration"],
    "irAssigned": ["assigned_to"],
    "irAdded": ["added_by", "IR_L_E"],
    "irDate": ["added_on", "Date"],
    "utlUser": ["UL_lo", "Login ID", "Login"],
    "utlDate": ["UL_Date", "Date"],
    "compId": ["Comp_ID", "Comp ID", "Teleopti ID", "ST_ID"],
    "compDate": ["Comp_Da", "Date"],
    "compDuration": ["Comp_Du", "Comp Duration", "Duration"],
}
IR_COL = {"assigned": 23, "added": 24, "date": 25}  # X, Y, Z (احتياطي)
SCHEDULE_COL = {"id": 1, "date_or_code": 2, "duration": 9}  # B, C, J
TOTALS_WORDS = {"totals", "total", "الاجمالي", "الإجمالي", "اجمالي", "إجمالي", "المجموع"}
MONTHS = {m: i + 1 for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split())}
PARTIAL_MIN = 5


# ─── أدوات ────────────────────────────────────────────────────────────────
def is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    try:
        return bool(pd.isna(v)) if not isinstance(v, (list, tuple, dict)) else False
    except (TypeError, ValueError):
        return False


def norm(v):
    if is_blank(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"[^a-z0-9\u0600-\u06ff]+", "", str(v).strip().lower())


def find_col(columns, aliases):
    cols = list(columns)
    wanted = [norm(a) for a in aliases if norm(a)]
    for w in wanted:
        for c in cols:
            if norm(c) == w:
                return c
    for w in wanted:
        if len(w) < PARTIAL_MIN:
            continue
        for c in cols:
            if w in norm(c):
                return c
    return None


def text(v):
    return "" if is_blank(v) else str(v).strip()


def parse_seconds(v):
    if is_blank(v) or v == "":
        return 0
    if isinstance(v, datetime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, time):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, timedelta):
        return round(v.total_seconds())
    if isinstance(v, (int, float)):
        return round(v * 86400) if 0 < v < 1 else v
    s = str(v).strip()
    if not s or s.lower() in {"unpaid", "maternity", "planned sick"}:
        return 0
    try:
        parts = [float(p) for p in s.split(":")]
    except ValueError:
        return 0
    if len(parts) == 3:
        return parts[0] * 3600 + parts[1] * 60 + parts[2]
    if len(parts) == 2:
        return parts[0] * 3600 + parts[1] * 60
    return parts[0] if len(parts) == 1 else 0


def fmt_time(seconds):
    if not seconds or seconds <= 0:
        return "0:00:00"
    total = int(round(seconds))
    return f"{total // 3600}:{(total % 3600) // 60:02d}:{total % 60:02d}"


def date_key(v, default_year=None):
    """Month-First (9/1/2026 = 1 سبتمبر). أي نص مش تاريخ صريح يرجع ''."""
    if is_blank(v) or v == "":
        return ""
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)):
        if 20000 < v < 80000:
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).date().isoformat()
        return ""
    s = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$", s)
    if m:
        return f"{int(m[1]):04d}-{int(m[2]):02d}-{int(m[3]):02d}"
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})(?:\s.*)?$", s)
    if m:
        a, b, y = int(m[1]), int(m[2]), int(m[3])
        y = y + 2000 if y < 100 else y
        month, day = (b, a) if (a > 12 and b <= 12) else (a, b)
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return ""
        return f"{y:04d}-{month:02d}-{day:02d}"
    m = re.match(r"^(\d{1,2})[\s\-/]([A-Za-z]{3})[A-Za-z]*(?:[\s\-/,]+(\d{2,4}))?$", s)
    if m:
        month = MONTHS.get(m[2].lower())
        if not month:
            return ""
        y = int(m[3]) if m[3] else (default_year or datetime.now().year)
        y = y + 2000 if y < 100 else y
        return f"{y:04d}-{month:02d}-{int(m[1]):02d}"
    return ""


def lookup_key(value, day):
    n = norm(value)
    return f"{n}|{day}" if n and day else ""


def add(idx, value, day, amount):
    k = lookup_key(value, day)
    if k:
        idx[k] = idx.get(k, 0) + amount


def lookup_first(idx, candidates, day):
    for c in candidates:
        k = lookup_key(c, day)
        if k and k in idx:
            return idx[k]
    return 0


# ─── القراءة ──────────────────────────────────────────────────────────────
def find_data_file(stem):
    for ext in (".xlsx", ".xls", ".csv"):
        p = DATA / f"{stem}{ext}"
        if p.exists():
            return p
    return None


def read_sheets(path, header=0):
    if path.suffix.lower() == ".csv":
        return {"csv": pd.read_csv(path, header=header, dtype=object)}
    return pd.read_excel(path, sheet_name=None, header=header, dtype=object)


def read_rows(stem, hints, validate=None):
    path = find_data_file(stem)
    if not path:
        return []
    sheets = read_sheets(path)
    names = sorted(sheets, key=lambda n: 0 if any(h in norm(n) for h in hints) else 1)
    for name in names:
        df = sheets[name].dropna(how="all")
        if df.empty:
            continue
        rows = df.to_dict("records")
        if not validate or validate(rows):
            return rows
    return []


def read_structure():
    for path in STRUCTURE_CANDIDATES:
        if path.exists():
            sheets = read_sheets(path)
            names = sorted(sheets, key=lambda n: 0 if any(h in norm(n) for h in ("sep", "updated", "structure")) else 1)
            for name in names:
                rows = sheets[name].dropna(how="all").to_dict("records")
                if rows and find_col(rows[0].keys(), ALIASES["agentName"]) is not None:
                    return rows
    return []


def val(row, key, default=""):
    col = find_col(row.keys(), ALIASES[key])
    return row[col] if col is not None else default


# ─── الفهارس ──────────────────────────────────────────────────────────────
def sum_index(rows, user_key, date_key_name, value_key):
    idx = {}
    for r in rows:
        add(idx, val(r, user_key), date_key(val(r, date_key_name)), parse_seconds(val(r, value_key, 0)))
    return idx


def talk_index(rows):
    idx = {}
    for r in rows:
        def g(*names):
            c = find_col(r.keys(), list(names))
            return parse_seconds(r[c]) if c is not None else 0

        total = g("Hold Time", "HoldTime") + g("Other Time", "OtherTime") + g("AUXOUTOFFTIME") + g("ACWOUTOFFTIME")
        add(idx, val(r, "utlUser"), date_key(val(r, "utlDate")), total)
    return idx


def ir_indexes(rows):
    assigning, tkt = {}, {}
    if not rows:
        return assigning, tkt
    keys = list(rows[0].keys())
    c_assigned = find_col(keys, ALIASES["irAssigned"]) or (keys[IR_COL["assigned"]] if len(keys) > IR_COL["assigned"] else None)
    c_added = find_col(keys, ALIASES["irAdded"]) or (keys[IR_COL["added"]] if len(keys) > IR_COL["added"] else None)
    c_date = find_col(keys, ALIASES["irDate"]) or (keys[IR_COL["date"]] if len(keys) > IR_COL["date"] else None)
    for r in rows:
        day = date_key(r.get(c_date))
        if not day:
            continue
        add(assigning, r.get(c_assigned), day, 1)
        add(tkt, r.get(c_added), day, 1)
    return assigning, tkt


def parse_schedule():
    """يرجّع (index, days, code_totals). index[key] = {total, has_total, codes{lower: secs}}"""
    path = find_data_file("Schedule")
    index, days, code_totals = {}, set(), {}
    if not path:
        return index, days, code_totals

    for _, df in read_sheets(path, header=None).items():
        login, day = "", ""
        for row in df.itertuples(index=False):
            cells = [text(row[i]) if len(row) > i else "" for i in range(3)]
            if any(c and norm(c) in {norm(w) for w in TOTALS_WORDS} for c in cells):
                login, day = "", ""
                continue
            m = re.search(r"\b(\d{5,6})\b", cells[1])
            if m:
                login, day = m.group(1), ""
            if not login or cells[2] == "":
                continue

            raw = row[SCHEDULE_COL["duration"]] if len(row) > SCHEDULE_COL["duration"] else None
            has_time = not is_blank(raw) and str(raw).strip() != ""
            d = date_key(row[SCHEDULE_COL["date_or_code"]])
            if d:
                day = d
                days.add(d)
                e = index.setdefault(lookup_key(login, d), {"total": 0, "has_total": False, "codes": {}})
                if has_time:
                    e["total"] += parse_seconds(raw)
                    e["has_total"] = True
                continue
            if not day or not has_time:
                continue
            secs = parse_seconds(raw)
            if not secs:
                continue
            lower = cells[2].lower()
            e = index.setdefault(lookup_key(login, day), {"total": 0, "has_total": False, "codes": {}})
            e["codes"][lower] = e["codes"].get(lower, 0) + secs
            code_totals[lower] = code_totals.get(lower, 0) + secs
        if index:
            break
    return index, days, code_totals


def schedule_seconds(entry, codes):
    if codes:
        return sum(entry["codes"].get(c, 0) for c in codes)
    return entry["total"] if entry["has_total"] else sum(entry["codes"].values())


# ─── الحساب ───────────────────────────────────────────────────────────────
def load_config():
    cfg = dict(DEFAULT_CONFIG)
    p = ROOT / "config.json"
    if p.exists():
        cfg.update(json.loads(p.read_text(encoding="utf-8")))
    return cfg


def build(structure, utl, ir, comp, sched_index, sched_days, codes, cfg):
    days = sorted({d for d in (date_key(val(r, "utlDate")) for r in utl) if d})
    if not days:
        days = set(sched_days)
        for r in ir:
            days.add(date_key(val(r, "irDate")))
        for r in comp:
            days.add(date_key(val(r, "compDate")))
        for r in structure:
            days.add(date_key(val(r, "structureDate")))
        days = sorted(d for d in days if d)

    ir_assigning, ir_tkt = ir_indexes(ir)
    talk = talk_index(utl)
    st_dur = sum_index(structure, "structureId", "structureDate", "structureDuration")
    comp_idx = sum_index(comp, "compId", "compDate", "compDuration")

    long_rows = []
    for r in structure:
        a = {k: text(val(r, k)) for k in
             ("structureId", "loginId", "perm", "ttsUser", "bssUser", "group", "agentName", "status", "tlId", "tlName")}
        a["status"] = a["status"] or "Active"
        if not a["agentName"] and not a["loginId"]:
            continue
        active = a["status"].lower() == "active"

        for day in days:
            assigning = lookup_first(ir_assigning, [a["ttsUser"]], day)
            tkt = lookup_first(ir_tkt, [a["ttsUser"]], day)
            system = round(tkt * cfg["secondsPerTicket"])
            talk_s = lookup_first(talk, [a["loginId"]], day)
            comp_s = lookup_first(comp_idx, [a["structureId"], a["loginId"]], day)

            sch = None
            for cand in (a["structureId"], a["loginId"], a["agentName"]):
                e = sched_index.get(lookup_key(cand, day))
                if e:
                    sch = schedule_seconds(e, codes)
                    break
            if sch is None:
                sch = lookup_first(st_dur, [a["structureId"]], day)

            loss = max(0, sch * cfg["teleSchFactor"] - (system + talk_s + comp_s))
            long_rows.append({
                "Teleopti ID": a["structureId"], "Login ID": a["loginId"], "Perm": a["perm"],
                "TTS User": a["ttsUser"], "BSS User": a["bssUser"], "Group": a["group"],
                "Agent Name": a["agentName"], "Status": a["status"], "TL ID": a["tlId"], "TL Name": a["tlName"],
                "Date": day, "Assigning Tkts": assigning, "TKT": tkt, "System": fmt_time(system),
                "Talk Time": fmt_time(talk_s), "Tele-SCH": fmt_time(sch), "Comp": fmt_time(comp_s),
                "Loss Time": fmt_time(loss) if active else a["status"],
            })
    return days, long_rows


def to_matrix(long_df):
    fixed = ["Teleopti ID", "Login ID", "Perm", "TTS User", "BSS User", "Group",
             "Agent Name", "Status", "TL ID", "TL Name"]
    fixed = [c for c in fixed if c in long_df and (c in ("Teleopti ID", "Login ID", "TTS User", "Agent Name",
                                                         "Status", "TL Name") or (long_df[c] != "").any())]
    metrics = ["Assigning Tkts", "TKT", "System", "Talk Time", "Tele-SCH", "Comp", "Loss Time"]
    wide = long_df.pivot_table(index=fixed, columns="Date", values=metrics, aggfunc="first")
    wide = wide.swaplevel(0, 1, axis=1)
    days = sorted(long_df["Date"].unique())
    wide = wide.reindex(columns=pd.MultiIndex.from_product([days, metrics]))
    wide.columns = [f"{datetime.strptime(d, '%Y-%m-%d').strftime('%d-%b')} - {m}" for d, m in wide.columns]
    return wide.reset_index()


def process_pipeline():
    cfg = load_config()
    structure = read_structure()
    if not structure:
        print("⚠️  لا يوجد ملف Structure (STR Loss.xlsx) صالح - تم الإيقاف")
        return 0

    utl = read_rows("UTL", ("utl", "log"))
    ir = read_rows("IR", ("ir", "ticket"))
    comp = read_rows("Compensation", ("comp",))
    sched_index, sched_days, code_totals = parse_schedule()
    codes = {str(c).strip().lower() for c in cfg.get("scheduleCodes", []) if str(c).strip()}

    print(f"Structure={len(structure)} UTL={len(utl)} IR={len(ir)} Comp={len(comp)} "
          f"Schedule(agent-days)={len(sched_index)} Codes={sorted(codes) or 'إجمالي اليوم'}")
    if code_totals:
        print("الأكواد الموجودة في Schedule:", ", ".join(sorted(code_totals)))

    days, long_rows = build(structure, utl, ir, comp, sched_index, sched_days, codes, cfg)
    if not days or not long_rows:
        print("⚠️  لا توجد تواريخ داخل ملفات data/ - لا يوجد ما يُحدَّث")
        return 0

    long_df = pd.DataFrame(long_rows).fillna("")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        to_matrix(long_df).to_excel(writer, sheet_name="Matrix", index=False)
        long_df.to_excel(writer, sheet_name="Daily_Long", index=False)
    print(f"✅ تم حفظ التقرير في: {OUTPUT}  ({len(days)} يوم)")
    return 0


if __name__ == "__main__":
    sys.exit(process_pipeline())
