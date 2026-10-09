#!/usr/bin/env python3
# =====================================================================
#  1_gateway_clean.py  -  Gateway LMS data cleaning + progression build
#  (Python port of gateway.R.  RUN THIS FIRST, then 2_dashboard.py)
#
#  One script, start to finish:
#    PART 1 - CLEANING
#      1. cleans the users export and builds a stable unique_key per email
#      2. cleans the six course progress exports (bm, fm, ggw, grp, goc, gocg)
#         and attaches unique_key + registration id by email
#      3. builds all_df (one row per participant, flag per course enrolled)
#    PART 2 - PROGRESSION BUILD
#      4. builds the progression tables and writes output/progression_data.xlsx
#         (read by 2_dashboard.py)
#
#  Participant name and email are KEPT in the Participants sheet so the
#  dashboard can search by them and show them in follow-up lists.
#  The xlsx therefore contains personal data: store and share it carefully.
#
#  Each course export is a wide file: name, email, then PAIRS of columns
#  (<activity> = completion status, <time column> = completion time).
#  The pairs are read by POSITION, so no manual renames are needed.
#
#  Needs: pandas>=2.0, numpy, openpyxl (and xlsxwriter, optional but faster)
# =====================================================================
from __future__ import annotations

import os
import re
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------- settings
# Folder that holds data/ and receives output/.  Defaults to this script's folder;
# set the GATEWAY_BASE_DIR environment variable (or edit the line below) to point elsewhere.
BASE_DIR = Path(os.environ.get(
    "GATEWAY_BASE_DIR",
    Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()))
os.chdir(BASE_DIR)

DATA_DIR = Path("data")        # folder holding the raw LMS exports
OUT_DIR = Path("output")       # 2_dashboard.py reads output/progression_data.xlsx
OUT_DIR.mkdir(parents=True, exist_ok=True)
OUT_FILE = OUT_DIR / "progression_data.xlsx"

EXPORT_CLEAN_CSVS = False   # True = also write all_df.csv and per-course csvs to ./clean
INCLUDE_PHONE = False       # True = also keep "phone number" from the users export (High-PII, off by default)

# GATEWAY participant glossary (optional): copied into the xlsx so the dashboard's glossary can use it
_gloss = [p for d in (DATA_DIR, Path(".")) if Path(d).exists() for p in sorted(Path(d).glob("*.xlsx"))
          if re.match(r"^gateway.*gloss.*\.xlsx$", p.name, re.I)]
GLOSSARY_FILE = _gloss[0] if _gloss else None

# newest users export is picked automatically (file name carries the date)
_users = sorted(DATA_DIR.glob("users_export_*.csv"))
if not _users:
    raise SystemExit(f"No users_export_*.csv found in {DATA_DIR}")
USER_FILE = _users[-1]
print("users export:", USER_FILE.name)

COURSE_FILES = {
    "bm": "progress.bm.csv",
    "fm": "progress.fm (2).csv",
    "ggw": "progress.ggw (1).csv",
    "grp": "progress.grp (2).csv",
    "goc": "progress.goc.csv",
    "gocg": "progress.goc-g (1).csv",
}
COURSES = {
    "bm": "Business Management",
    "fm": "Financial Management",
    "ggw": "General Gig Work",
    "grp": "Graphic Design",
    "goc": "GATEWAY Orientation (GOC)",
    "gocg": "GATEWAY Orientation (GOCG)",
}

# ---------------------------------------------------------------- helpers
EPOCH = pd.Timestamp("1970-01-01", tz="UTC")
DAY = 86400.0


def clean_label(x) -> str:
    """drop quotes, decode &amp;, squish whitespace"""
    return re.sub(r"\s+", " ", str(x).replace('"', "").replace("&amp;", "&")).strip()


def to_secs(x) -> np.ndarray:
    """timestamp text -> numeric seconds (UTC). Copes with '2026-03-14T09:21:55Z', '2026-03-14 09:21:55',
    fractional seconds and +03:00 offsets. 1970 placeholders and anything before 2000 -> NaN."""
    s = pd.Series(x).reset_index(drop=True)
    dt = pd.to_datetime(s, utc=True, errors="coerce", format="ISO8601")
    if dt.isna().all() and s.notna().any():                       # fallback for non-ISO text
        dt = pd.to_datetime(s, utc=True, errors="coerce", format="mixed")
    out = (dt - EPOCH).dt.total_seconds().to_numpy(dtype=float, copy=True)
    out[np.isfinite(out) & (out <= 946684800.0)] = np.nan         # 2000-01-01 UTC
    return out


def as_dt(s) -> pd.Series:
    """numeric seconds -> naive UTC datetime (Excel cannot store tz-aware datetimes)"""
    return pd.to_datetime(pd.Series(np.asarray(s, dtype=float)), unit="s")


def repair_names(names: list[str]) -> list[str]:
    """like readr's 'unique' name repair: blanks -> ...N, duplicated names get ...N (N = column position)"""
    counts = pd.Series(names[2:]).value_counts()
    out = []
    for i, nm in enumerate(names, 1):
        if i <= 2:
            out.append(nm)
        elif nm == "":
            out.append(f"...{i}")
        elif counts[nm] > 1:
            out.append(f"{nm}...{i}")
        else:
            out.append(nm)
    return out


def r_round(x, nd):
    return np.round(np.asarray(x, dtype=float), nd)


def read_csv_str(path, **kw) -> pd.DataFrame:
    # same NA rules as readr: only "" and "NA" are missing
    return pd.read_csv(path, dtype=str, keep_default_na=False, na_values=["", "NA"], **kw)


# =====================================================================
#  PART 1 - CLEANING
# =====================================================================

# ---- users: unique_key is built from the cleaned email
user = read_csv_str(USER_FILE)
user.columns = [c.lower() for c in user.columns]
user["email"] = user["email"].str.strip().str.lower()
_rank = {e: i + 1 for i, e in enumerate(sorted(user["email"].dropna().unique()))}
user["unique_key"] = user["email"].map(_rank).map(lambda r: f"KEY{int(r):06d}" if pd.notna(r) else np.nan)

_first = user["first name"].str.strip().str.title()
_last = user["last name"].str.strip().str.title()
user["name"] = (_first.fillna("") + " " + _last.fillna("")).str.replace(r"\s+", " ", regex=True).str.strip()
user = user.drop(columns=[c for c in ("id", "last name", "first name") if c in user.columns])
_front = ["unique_key", "registration id", "name"]
user = user[_front + [c for c in user.columns if c not in _front]]

# glossary rule: flag duplicate emails (same person registered more than once)
user["duplicate_email"] = np.where(user["email"].notna() & user["email"].duplicated(keep=False), "Yes", "No")

print(f"users: {len(user)} rows, {user['registration id'].duplicated().sum()} repeated registration ids, "
      f"{(user['duplicate_email'] == 'Yes').sum()} rows sharing an email")
if user["email"].isna().any():
    print(f"  warning: {user['email'].isna().sum()} users have no email, so they get no unique_key")

# one row per email, so a repeated email cannot double a learner's course rows in the join
user_keys = user[["unique_key", "registration id", "email"]].drop_duplicates(subset="email")


# ---- course exports: columns 1-2 are name + email, then status/time pairs
def read_course(fname: str) -> pd.DataFrame:
    raw = read_csv_str(DATA_DIR / fname, header=None)
    header = ["" if pd.isna(h) else h for h in raw.iloc[0].tolist()]
    d = raw.iloc[1:].reset_index(drop=True)
    d.columns = repair_names(["name", "email"] + [clean_label(h) for h in header[2:]])
    d["name"] = d["name"].str.strip().str.lower()
    d["email"] = d["email"].str.strip().str.lower()
    d = d.merge(user_keys, on="email", how="left")
    first = ["unique_key", "registration id"]
    return d[first + [c for c in d.columns if c not in first]]


course_frames = {code: read_course(f) for code, f in COURSE_FILES.items()}

# gocg: drop one known bad row, then exact duplicates
_g = course_frames["gocg"]
_bad = (_g["email"] == "iguodalaefosa@gmail.com") & (_g["Announcements"].eq("Not completed") | _g["Announcements"].isna())
course_frames["gocg"] = _g[~_bad].drop_duplicates().reset_index(drop=True)

for code, d in course_frames.items():
    print(f"{code:<5} {len(d):>6} rows, {d['unique_key'].isna().sum()} without a matching user")

# ---- all_df: one row per participant + a yes/NA flag per course
JOIN_BY = ["unique_key", "registration id", "email"]


def flag(d: pd.DataFrame, code: str) -> pd.DataFrame:
    f = d[JOIN_BY].drop_duplicates().copy()
    f[code] = "yes"
    return f


_want = ["date joined", "unique_key", "registration id", "name", "email", "duplicate_email", "gender", "age",
         "state", "lga", "learning area", "accept consent", "skill assessment score", "skill assessment status",
         "skill assessment (learning area)", "computer literacy score", "computer literacy status", "status",
         "registration stage", "email verified", "has children", "pwd status", "displaced youth", "has computer",
         "has power supply", "internet access", "heard from"]
_missing = [c for c in _want if c not in user.columns]
if _missing:
    print("  warning: users export lacks", _missing, "- filled with blanks")
    for c in _missing:
        user[c] = np.nan
_extra = [c for c in (["pathway"] + (["phone number"] if INCLUDE_PHONE else [])) if c in user.columns]

all_df = user[_want + _extra]
for code in ("bm", "fm", "ggw", "goc", "gocg", "grp"):
    all_df = all_df.merge(flag(course_frames[code], code), on=JOIN_BY, how="left")
all_df = all_df.drop_duplicates().reset_index(drop=True)

print("all_df:", len(all_df), "participants")
if all_df["unique_key"].dropna().duplicated().any():
    print("WARNING: all_df has repeated unique_key values (repeated emails in the users export); "
          "check before trusting counts.")

if EXPORT_CLEAN_CSVS:
    Path("clean").mkdir(exist_ok=True)
    all_df.to_csv("clean/all_df.csv", index=False, na_rep="")
    for code, d in course_frames.items():
        d.drop(columns=["name", "email"], errors="ignore").to_csv(f"clean/{code}.csv", index=False, na_rep="")

# =====================================================================
#  PART 2 - PROGRESSION BUILD
# =====================================================================

# ---- stage (module) mapping ----------------------------------------
# The exports have no module field, so stages are inferred from activity names.
# Edit these functions (or the Course_Structure sheet) if the mapping is wrong.
def stage_bm(labels):
    out, mod = [], None
    for l in labels:
        m = re.search(r"^Welcome to Module (\d)", l)
        if m:
            mod = m.group(1)
        if re.search(r"^(Portfolio Task|Quick Check-in) \d", l):
            d = re.search(r"\d$", l)
            out.append("Module " + (d.group() if d else "NA"))
        elif l.startswith("Mid-Course"):
            out.append("Mid-Course Check-In")
        elif re.search(r"^(Capstone|Course Completion|Skill Growth|Share Your)", l):
            out.append("Capstone & Completion")
        elif mod is None or l.startswith("Welcome to the Course Q&A"):
            out.append("Orientation")
        else:
            out.append("Module " + mod)
    return out


def stage_fm(labels):
    out = []
    for l in labels:
        if re.search(r"^(Welcome|Download|Self-Check)", l):
            out.append("Orientation")
        elif l.startswith("Mid-Course"):
            out.append("Mid-Course Check-In")
        elif re.search(r"^0\.\d", l):
            out.append("Module 0")
        elif re.search(r"^5.4", l):
            out.append("Module 5")
        elif re.search(r"^Module \d:", l):
            out.append("Module " + re.search(r"^Module (\d)", l).group(1))
        else:
            out.append("Capstone & Completion")
    return out


GGW_RULES = [
    (r"Welcome & Course Overview|Download your Course Guide|Self-Check: General|Course Q&A", "Orientation"),
    (r"Nigerian Gig Landscape|Mindset & The Lifestyle|Mapping Your Skills|Readiness Audit|Niche Mapping|Module 1$|Quiz Check-in 1$", "Module 1"),
    (r"Craft Your Professional Edge|Building Your Portfolio|Multi-Platform Storefront|Tutorial on (How to set up|Setting up|Terawork)|Digital Storefront|Bio-Builder|Module 2$|Quiz Check-in 2$", "Module 2"),
    (r"Navigating The Marketplaces|Writing To Win|Pricing Your Value|Finishing the Job|Profit Map|Module 3$|Quiz Check-in 3$", "Module 3"),
    (r"Mid-Course", "Mid-Course Check-In"),
    (r"Managing the Client Relationship|Increasing Your Project Value|Business Financial Management|Business Sustainability Growth Plan|Value ladder|Module 4$|Quiz Check-in 4$", "Module 4"),
    (r"Global Platforms|Global Communications|Global Payments|Global Launch|Global Clock|Module 5$|Quiz Check-in 5$", "Module 5"),
    (r"Protecting Your Brand|One-Offs to Monthly|Protecting Your Business|Professional Sustainability|Dispute|Module 6$|Quiz Check-in 6$", "Module 6"),
    (r"AI Strategy|AI Tools|Strong Prompts|AI Workflow|AI-Powered|Power Prompt|Module 7$|Quiz Check-in 7$", "Module 7"),
]


def stage_ggw(labels):
    out = []
    for l in labels:
        out.append(next((stg for rx, stg in GGW_RULES if re.search(rx, l, re.I)), "Capstone & Completion"))
    return out


def stage_grp(labels):
    out, mod = [], 1
    for l in labels:
        mm = re.search(r"Module ?(\d+) Course Materials|PortfolioTask_Module(\d+)_", l)
        mq = re.search(r"^Module (\d+) Quiz", l)
        if re.search(r"^(Welcome|Download|Self-Check)", l):
            out.append("Orientation")
        elif mm:
            out.append("Module " + (mm.group(1) or mm.group(2)))
        elif l.startswith("Mid-Course"):
            out.append("Mid-Course Check-In")
        elif mq:
            mod = int(mq.group(1)) + 1
            out.append("Module " + mq.group(1))
        elif re.search(r"^(Practical|Skill Growth|Share Your|Submit Your Capstone|Course Completion|Further Learning)", l):
            out.append("Capstone & Completion")
        else:
            out.append(f"Module {mod}")      # lessons belong to the module whose quiz closes them
    return out


def stage_orient(labels):
    out = []
    for l in labels:
        if re.search(r"Announcements|Navigate the LMS|Say Hello|Introduce Yourself|Welcome to GATEWAY|Pledge|Weekly Learning Plan", l):
            out.append("1. Getting Started")
        elif re.search(r"Learning Journey|Gig Work & Enhanced|Maximizing|Gigs & Financial", l):
            out.append("2. Learning Journey & Opportunities")
        elif re.search(r"Navigating the Platform|Community|Safeguarding|FAQs", l):
            out.append("3. Platform, Community & Support")
        else:
            out.append("4. Orientation Wrap-up")
    return out


STAGERS = {"bm": stage_bm, "fm": stage_fm, "ggw": stage_ggw, "grp": stage_grp,
           "goc": stage_orient, "gocg": stage_orient}
MID_AFTER = {"bm": "Module 3", "fm": "Module 3", "ggw": "Module 3", "grp": "Module 6"}


def order_stages(course: str, stages: list[str]) -> list[str]:
    uniq = list(dict.fromkeys(stages))                      # unique, first-seen order
    if course in ("goc", "gocg"):
        return sorted(uniq)
    num = lambda s: int(m.group()) if (m := re.search(r"\d+", s)) else 10 ** 9
    mods = sorted([s for s in uniq if s.startswith("Module")], key=num)
    out = []
    if "Orientation" in uniq:
        out.append("Orientation")
    for m_ in mods:
        out.append(m_)
        if m_ == MID_AFTER[course] and "Mid-Course Check-In" in uniq:
            out.append("Mid-Course Check-In")
    if "Mid-Course Check-In" in uniq and "Mid-Course Check-In" not in out:
        out.append("Mid-Course Check-In")           # the module it should follow is absent: keep the stage anyway
    if "Capstone & Completion" in uniq:
        out.append("Capstone & Completion")
    return out


# ---- participants (name and email are kept) -------------------------
A = all_df.copy()
A.columns = [re.sub(r"[()]", "", c.lower().replace(" ", "_")) for c in A.columns]
for cl in ("pathway", "phone_number"):
    if cl not in A.columns:
        A[cl] = np.nan

_dj = to_secs(A["date_joined"])                               # seconds, kept for the course build
A["date_joined"] = as_dt(_dj)
A["age"] = pd.to_numeric(A["age"], errors="coerce")
A["day_joined"] = A["date_joined"].dt.day.astype("Int64")     # glossary: derived Day Joined
A["month_joined"] = A["date_joined"].dt.strftime("%Y-%m")     # glossary: derived Month Joined
# eligibility is 18-35 (glossary), so out-of-range ages get their own band instead of vanishing
A["age_band"] = pd.cut(A["age"], [-np.inf, 17, 21, 25, 29, 35, np.inf],
                       labels=["Under 18", "18-21", "22-25", "26-29", "30-35", "36+"]).astype(object)
_state = A["state"].str.replace(r"\s*State$", "", regex=True).str.strip()
_fct = _state.str.lower().str.contains(r"federal capital|^fct|abuja", regex=True, na=False)
A["state"] = _state.where(~_fct, "FCT (Abuja)")
A["learning_area"] = A["learning_area"].fillna("Unknown")
A["registration_complete"] = (A["registration_stage"].fillna("").str.lower()
                              .str.contains("complete|fully", regex=True)).astype(int)
A["skill_assessment_score"] = pd.to_numeric(A["skill_assessment_score"], errors="coerce")
A["computer_literacy_score"] = pd.to_numeric(A["computer_literacy_score"], errors="coerce")
A["courses_enrolled"] = A[list(COURSES)].notna().sum(axis=1)

_pcols = (["unique_key", "registration_id", "name", "email"] + (["phone_number"] if INCLUDE_PHONE else []) +
          ["date_joined", "day_joined", "month_joined", "gender", "age", "age_band", "state", "lga",
           "learning_area", "pathway", "status", "registration_stage", "registration_complete",
           "accept_consent", "email_verified", "duplicate_email",
           "skill_assessment_score", "skill_assessment_status", "skill_assessment_learning_area",
           "computer_literacy_score", "computer_literacy_status",
           "has_children", "pwd_status", "displaced_youth",
           "has_computer", "has_power_supply", "internet_access", "heard_from", "courses_enrolled"])
participants = A[_pcols].rename(columns={"phone_number": "phone"})
# date joined in seconds per key (first row wins for repeated keys, like R's named-vector lookup)
joined = pd.Series(_dj, index=A["unique_key"].to_numpy())
joined = joined[~joined.index.duplicated(keep="first")]

# ---- unpack every course once to get the snapshot date ---------------
raw: dict = {}
snap = -np.inf
for code in COURSES:
    d = course_frames[code].drop(columns=["name", "email", "dup"], errors="ignore").reset_index(drop=True)
    npairs = (d.shape[1] - 2) // 2                      # any trailing stray column is ignored
    lab_idx = [2 + 2 * j for j in range(npairs)]         # positions of the status columns
    labels = [clean_label(re.sub(r"\.\.\.\d+$", "", d.columns[i])) for i in lab_idx]
    st = d.iloc[:, lab_idx].to_numpy(dtype=object)
    tm = np.column_stack([to_secs(d.iloc[:, i + 1]) for i in lab_idx]) if npairs else np.empty((len(d), 0))
    if not np.isfinite(tm).any():
        ex = pd.Series(d.iloc[:, [i + 1 for i in lab_idx]].to_numpy().ravel()).dropna().head(5).tolist()
        raise ValueError(f"Course '{code}': no completion timestamps could be parsed. Example raw values: {' | '.join(map(str, ex))}")
    print(f"{code:<5} {int(np.isfinite(tm).sum()):,} timestamps parsed")
    raw[code] = dict(d=d, labels=labels, st=st, tm=tm)
    snap = max(snap, np.nanmax(tm))
SNAPSHOT = (np.floor(snap / DAY) + 1) * DAY              # day after the last logged activity

# ---- course-level build ---------------------------------------------
structure_l, progress_l, stage_l, item_l, sfun_l = [], [], [], [], []

for code, cname in COURSES.items():
    d, labels, st, tm = raw[code]["d"], np.array(raw[code]["labels"], dtype=object), raw[code]["st"], raw[code]["tm"]
    n, n_part = len(labels), len(d)
    keys = d["unique_key"].to_numpy(dtype=object)
    n_enrolled_keys = pd.Series(keys).nunique(dropna=False)

    stages = np.array(STAGERS[code](list(labels)), dtype=object)
    ordered = order_stages(code, list(stages))
    s_order = {s: i + 1 for i, s in enumerate(ordered)}
    stage_no = np.array([s_order[s] for s in stages])
    idx_sorted = np.lexsort((np.arange(n), stage_no))    # item order = stage order, then column order

    structure_l.append(pd.DataFrame({
        "course": code, "course_name": cname, "item_order": np.arange(1, n + 1),
        "item_label": labels[idx_sorted], "stage": stages[idx_sorted], "stage_order": stage_no[idx_sorted]}))

    st_df = pd.DataFrame(st)
    done = st_df.apply(lambda c: c.str.startswith("Completed", na=False)).to_numpy(bool)
    failed = st_df.apply(lambda c: c.str.contains("did not achieve", regex=False, na=False)).to_numpy(bool)
    Tn = np.where(done, tm, np.nan)                      # keep times only for completed activities

    # ---- stage level
    S = len(ordered)
    COMP = np.zeros((n_part, S), bool)
    TOUCH = np.zeros((n_part, S), bool)
    sp_list = []
    for k in range(S):
        cols = np.where(stages == ordered[k])[0]
        n_done = done[:, cols].sum(axis=1)
        first = np.fmin.reduce(Tn[:, cols], axis=1)
        last = np.fmax.reduce(Tn[:, cols], axis=1)
        COMP[:, k] = n_done == len(cols)
        TOUCH[:, k] = n_done > 0
        span = r_round((last - first) / DAY, 2)
        sp_list.append(pd.DataFrame({
            "course": code, "unique_key": keys, "stage": ordered[k], "stage_order": k + 1,
            "items_total": len(cols), "items_done": n_done.astype(int),
            "first_completed": as_dt(first), "last_completed": as_dt(last),
            "stage_complete": COMP[:, k].astype(int), "span_days": span}))
        # stage funnel row (one per stage)
        spans_done = pd.Series(span[COMP[:, k]])
        sfun_l.append({
            "course": code, "course_name": cname, "stage": ordered[k], "stage_order": k + 1,
            "enrolled": n_enrolled_keys, "started": int((n_done > 0).sum()), "completed": int(COMP[:, k].sum()),
            "avg_items_done": float(n_done.mean()) if n_part else np.nan,
            "median_span_days_completed": round(spans_done.median(), 2) if spans_done.notna().any() else np.nan,
            "mean_span_days_completed": round(spans_done.mean(), 2) if spans_done.notna().any() else np.nan})
    sp = pd.concat(sp_list, ignore_index=True)
    stage_l.append(sp[sp["items_done"] > 0])             # zero-activity stages are rebuilt from Course_Structure in the app

    # ---- participant level
    items_done = done.sum(axis=1)
    ranks = np.arange(1, S + 1)
    last_comp_i = (COMP * ranks).max(axis=1)
    furth_i = (TOUCH * ranks).max(axis=1)
    done_sorted = done[:, idx_sorted]
    done_items = [",".join(str(i + 1) for i in np.flatnonzero(row)) for row in done_sorted]
    first_act = np.fmin.reduce(Tn, axis=1)
    last_act = np.fmax.reduce(Tn, axis=1)
    active_days = pd.DataFrame(np.floor(Tn / DAY)).nunique(axis=1).to_numpy()
    dj = joined.reindex(keys).to_numpy(dtype=float)
    pct = r_round(items_done / n * 100, 1)
    is_done = items_done == n

    progress_l.append(pd.DataFrame({
        "course": code, "course_name": cname, "unique_key": keys, "registration_id": d["registration id"].to_numpy(),
        "items_total": n, "items_done": items_done.astype(int), "items_remaining": (n - items_done).astype(int),
        "pct_complete": pct, "distance_to_completion_pct": r_round((1 - items_done / n) * 100, 1),
        "stages_total": S, "stages_completed": COMP.sum(axis=1).astype(int),
        "stages_remaining": (S - COMP.sum(axis=1)).astype(int),
        "last_completed_stage": np.array(["None"] + ordered, dtype=object)[last_comp_i],
        "furthest_stage": np.array(["Not started"] + ordered, dtype=object)[furth_i],
        "furthest_stage_order": furth_i.astype(int),
        "first_activity": as_dt(first_act), "last_activity": as_dt(last_act),
        "failed_assessments": failed.sum(axis=1).astype(int), "done_items": done_items,
        "date_joined": as_dt(dj),
        "days_since_last_activity": r_round((SNAPSHOT - last_act) / DAY, 1),
        "days_since_joined": r_round((SNAPSHOT - dj) / DAY, 1),
        "joined_to_first_days": r_round((first_act - dj) / DAY, 1),
        "total_span_days": r_round((last_act - first_act) / DAY, 2),
        "active_days": active_days.astype(int),
        "completed": is_done.astype(int),
        "completion_date": as_dt(np.where(is_done, last_act, np.nan)),
        "days_to_complete": np.where(is_done, r_round((last_act - dj) / DAY, 1), np.nan),
        "progress_band": pd.cut(pct, [-1, 0, 24.9, 49.9, 74.9, 99.9, 100],
                                labels=["Not started", "1-24%", "25-49%", "50-74%", "75-99%", "Completed"]).astype(object)}))

    # ---- item funnel
    comp_n = done.sum(axis=0)[idx_sorted]
    item_l.append(pd.DataFrame({
        "course": code, "course_name": cname, "item_order": np.arange(1, n + 1), "item_label": labels[idx_sorted],
        "stage": stages[idx_sorted], "stage_order": stage_no[idx_sorted],
        "enrolled": n_part, "completed": comp_n.astype(int), "pct_completed": r_round(comp_n / n_part * 100, 1),
        "failed_pass_grade": failed.sum(axis=0)[idx_sorted].astype(int),
        "drop_vs_prev_item_pp": np.r_[np.nan, r_round((comp_n[:-1] - comp_n[1:]) / n_part * 100, 1)]}))
    print(f"built {code} ({n_part} learners, {n} activities, {S} stages)")

structure_df = pd.concat(structure_l, ignore_index=True)
progress = pd.concat(progress_l, ignore_index=True)
stage_prog = pd.concat(stage_l, ignore_index=True)
item_funnel = pd.concat(item_l, ignore_index=True)
stage_funnel = pd.DataFrame(sfun_l)
stage_funnel["pct_started"] = r_round(stage_funnel["started"] / stage_funnel["enrolled"] * 100, 1)
stage_funnel["pct_completed"] = r_round(stage_funnel["completed"] / stage_funnel["enrolled"] * 100, 1)

notes = pd.DataFrame({"note": [
    f"Snapshot date used for 'days since last activity': {as_dt([SNAPSHOT])[0]:%Y-%m-%d} (day after the last logged completion).",
    "Course CSVs are paired columns: activity status + completion timestamp. Pairs are read by position (some time headers are mislabelled).",
    "An item counts as done when status starts with 'Completed' (includes 'achieved pass grade' and 'did not achieve pass grade').",
    "'did not achieve pass grade' is kept as failed_assessments so learners who completed but did not pass can be followed up.",
    "1970-01-01 timestamps are LMS placeholders and were set to missing; 'Not completed' rows that carry a real time (viewed, not finished) are ignored.",
    "Stages: Orientation, Modules, Mid-Course Check-In and Capstone & Completion. Portfolio tasks / quizzes / check-ins are filed under the module named in the activity.",
    "BM portfolio tasks and quick check-ins N are filed under Module N; Mid-Course Check-In sits after Module 3 (Module 6 for Graphic Design). Edit Course_Structure to change.",
    "Orientation courses (GOC, GOCG) have no modules; activities are grouped into 4 stages by theme.",
    "All activities in a course are treated as required, including the four track-specific capstone briefs and prompt tutorials in ggw (learners of every track complete all of them).",
    "Stalled = not completed and no activity for N days (N is set in the dashboard; default 30, the glossary's Inactive rule: no LMS activity for 30+ days). Not started = enrolled, zero completed activities.",
    "Duration per stage = days between first and last completed activity in the stage (active span). Overall = first to last completed activity.",
    "'done_items' in Progress is a comma list of item_order numbers (see Course_Structure) so the tracker can show done vs remaining without a huge item table.",
    "Course display names for goc / gocg are labels only; rename in COURSES if they stand for something else.",
    "Participant name (title-cased first + last name) and email come from the users export; learners are matched to course exports by lower-cased, trimmed email.",
    "Glossary: name, email and phone are High-PII and should be excluded from analytical datasets; use Registration ID for reporting. They are kept here only for follow-up and can be hidden with SHOW_PII in 2_dashboard.py.",
    "Glossary-derived fields: Day Joined / Month Joined (from date joined), Actual Completion Date (last completed activity of a finished course), duplicate_email flag, age bands aligned to the 18-35 eligibility range.",
    "Training Start Date is approximated by the first completed activity; the registration-to-first-activity gap is the onboarding lag.",
]})

# ---- GATEWAY glossary (one row per data point; section header rows dropped) ----
glossary = None
if GLOSSARY_FILE is not None:
    g = pd.read_excel(GLOSSARY_FILE, sheet_name="Data_Glossary", dtype=str)
    if g.shape[1] >= 16:
        glossary = g.iloc[:, [1, 2, 3, 4, 5, 6, 7, 14, 15]].copy()
        glossary.columns = ["Field", "Category", "Status", "Source", "Definition",
                            "Data type", "Allowed values", "Privacy level", "Notes"]
        glossary = glossary[glossary["Field"].notna() & glossary["Category"].notna()]
        print(f"glossary: {len(glossary)} data points from {GLOSSARY_FILE.name}")
else:
    print("glossary file not found; the glossary table will be skipped")

# ---- write ----------------------------------------------------------
out_sheets = {"Participants": participants, "Progress": progress, "Stage_Progress": stage_prog,
              "Course_Structure": structure_df, "Stage_Funnel": stage_funnel,
              "Item_Funnel": item_funnel, "Data_Notes": notes}
if glossary is not None:
    out_sheets["Data_Glossary"] = glossary

try:
    import xlsxwriter  # noqa: F401
    _engine = "xlsxwriter"
except ImportError:
    _engine = "openpyxl"
with pd.ExcelWriter(OUT_FILE, engine=_engine) as xw:
    for sheet, df in out_sheets.items():
        df.to_excel(xw, sheet_name=sheet, index=False)

print("written", OUT_FILE)
for nm, df in [("participants", participants), ("progress", progress), ("stage_prog", stage_prog),
               ("structure_df", structure_df), ("stage_funnel", stage_funnel), ("item_funnel", item_funnel)]:
    print(f"{nm:<14} {len(df)} rows x {df.shape[1]} cols")
print(f"snapshot {as_dt([SNAPSHOT])[0]:%Y-%m-%d}")



# cd "/Users/moseskioko/Documents/files/CcHUB Workflow/R Codes/Gateway/python_dashboard"
# source "/Users/moseskioko/Documents/files/CcHUB Workflow/Python Codes/vscode trials/vscode_env/bin/activate"
# python gateway_clean.py
