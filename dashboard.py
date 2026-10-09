# =====================================================================
#  Student Progress Dashboard  -  Streamlit port of app.R
#  Run 1_gateway_clean.py FIRST (it writes output/progression_data.xlsx), then:
#      streamlit run 2_dashboard.py
#  Needs: streamlit>=1.50, pandas, numpy, plotly, openpyxl
#  Data : progression_data.xlsx in ./output or next to this file
# =====================================================================
from __future__ import annotations

import inspect
import pickle
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

st.set_page_config(page_title="Student Progress Dashboard", page_icon="📊", layout="wide")

# ---- dashboard settings (values come from the GATEWAY glossary) -------------
SHOW_PII = True        # False hides learner name / email / phone everywhere (glossary: High PII)
DEFAULT_STALL = 30     # glossary rule: Inactive = no LMS activity for 30+ days
REACH_TARGET = 400000  # five-year reach target
TARGETS = {"Female": 0.75, "Persons with disability": 0.05, "Displaced youth": 0.01}  # GESI targets

STATUS_LEVELS = ["Completed", "Active", "Stalled", "Not started"]
STATUS_COLS = {"Completed": "#10b981", "Active": "#6366f1", "Stalled": "#f43f5e", "Not started": "#94a3b8"}
BAND_LEVELS = ["Not started", "1-24%", "25-49%", "50-74%", "75-99%", "Completed"]

# ---------------------------------------------------------------- data
HERE = Path(__file__).resolve().parent
XLSX = next((p for p in [HERE / "output" / "progression_data.xlsx", HERE / "progression_data.xlsx",
                         Path("output/progression_data.xlsx"), Path("progression_data.xlsx")] if p.exists()), None)
if XLSX is None:
    st.error("progression_data.xlsx not found (looked in ./output and the app folder). "
             "Run 1_gateway_clean.py first.")
    st.stop()

SHEETS = ["Participants", "Progress", "Stage_Progress", "Course_Structure", "Item_Funnel", "Data_Notes", "Data_Glossary"]
PART_FILL = ["name", "email", "phone", "pathway", "pwd_status", "displaced_youth", "has_children", "has_power_supply",
             "accept_consent", "email_verified", "duplicate_email", "skill_assessment_status",
             "skill_assessment_learning_area", "computer_literacy_status", "gender", "age_band", "state",
             "learning_area", "internet_access", "has_computer", "heard_from"]
PROG_COLS = ["unique_key", "name", "email", "phone", "gender", "age_band", "state", "learning_area", "pathway",
             "pwd_status", "displaced_youth", "has_children", "has_power_supply", "internet_access", "has_computer",
             "skill_assessment_status", "computer_literacy_status", "heard_from"]


@st.cache_resource(show_spinner="Loading data...")
def load_all(path_str: str, mtime: float) -> dict:
    path = Path(path_str)
    cache = path.with_name("progression_cache.pkl")        # same idea as the .rds cache in app.R
    dat = None
    if cache.exists() and cache.stat().st_mtime > mtime:
        try:
            dat = pickle.loads(cache.read_bytes())
        except Exception:
            dat = None
    if dat is None:
        xl = pd.ExcelFile(path)
        dat = {s: xl.parse(s) for s in SHEETS if s in xl.sheet_names}
        try:
            cache.write_bytes(pickle.dumps(dat))
        except Exception:
            pass

    P = dat["Participants"].copy()
    for cl in PART_FILL:                                    # older exports may lack some columns
        if cl not in P.columns:
            P[cl] = np.nan
    if not SHOW_PII:
        P[["name", "email", "phone"]] = np.nan
    for df in (P, dat["Progress"]):
        df["registration_id"] = df["registration_id"].map(lambda x: None if pd.isna(x) else str(x))
    dat["Participants"] = P

    struct = dat["Course_Structure"]
    course_tbl = struct.drop_duplicates(["course", "course_name"])[["course", "course_name"]]
    course_only = dict(zip(course_tbl["course"], course_tbl["course_name"]))     # code -> display name

    part1 = P.drop_duplicates("unique_key")[PROG_COLS]
    prog_base = dat["Progress"].merge(part1, on="unique_key", how="left")

    rid = (prog_base[prog_base["registration_id"].notna()].drop_duplicates("registration_id")
           [["registration_id", "name", "email"]].sort_values("registration_id").reset_index(drop=True))
    rid["label"] = [" | ".join([r] + ([n] if pd.notna(n) else []) + ([e] if pd.notna(e) else []))
                    for r, n, e in zip(rid["registration_id"], rid["name"], rid["email"])]
    return dict(dat=dat, struct=struct, course_only=course_only, prog_base=prog_base,
                area_choices=["All"] + sorted(P["learning_area"].dropna().unique().tolist()), rid=rid)


_D = load_all(str(XLSX), XLSX.stat().st_mtime)
dat, struct, course_only = _D["dat"], _D["struct"], _D["course_only"]
prog_base, area_choices, rid_lookup = _D["prog_base"], _D["area_choices"], _D["rid"]
PARTS = dat["Participants"]


def add_status(d: pd.DataFrame, days: float) -> pd.DataFrame:
    d = d.copy()
    status = np.select([d["completed"] == 1, d["items_done"] == 0, d["days_since_last_activity"] >= days],
                       ["Completed", "Not started", "Stalled"], default="Active")
    d["status"] = pd.Categorical(status, categories=STATUS_LEVELS)
    return d


# ---------------------------------------------------------- helpers
def comma(x) -> str:
    return f"{int(round(float(x))):,}"


def pct(x, acc: float = 0.1) -> str:
    return f"{float(x) * 100:.{0 if acc >= 1 else 1}f}%"


def r1(x) -> str:
    return f"{x:.1f}".rstrip("0").rstrip(".")


def is_yes(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.lower().isin(["yes", "true", "1", "y", "verified"])


def pct_of(s: pd.Series, f) -> float:
    return np.nan if s.isna().all() else float(f(s).mean())     # NaN when the export lacks the field


def fmt_pct(v, acc: float = 0.1) -> str:
    return "n/a" if pd.isna(v) else pct(v, acc)


def drop_empty_pii(df: pd.DataFrame) -> pd.DataFrame:
    pii = [c for c in df.columns if c in ("Name", "Email", "Phone", "name", "email", "phone")]
    return df.drop(columns=[c for c in pii if df[c].isna().all()])


CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root { --ink:#0f172a; --muted:#64748b; --line:#e8ebf3; --bg:#f3f5fb; }
html, body, [class*="css"], .stApp { font-family:'Inter',sans-serif; }
.stApp { background:var(--bg); }
.block-container { max-width:1560px; padding-top:1rem; }
.topbar { background:linear-gradient(120deg,#1e1b4b 0%,#3730a3 100%); color:#fff; margin:0 0 16px;
          padding:16px 30px; border-radius:14px; box-shadow:0 2px 10px rgba(30,27,75,.25); }
.topbar h1 { font-size:22px; font-weight:700; margin:0; color:#fff; padding:0; }
.topbar span { font-size:12.5px; opacity:.8; }
.scope { background:#eef2ff; border-radius:10px; padding:6px 12px; text-align:center; margin-top:4px; }
.scope-n { font-size:22px; font-weight:700; color:#3730a3; line-height:1.15; }
.scope-l { font-size:11.5px; color:var(--muted); }
.card-title { font-weight:600; font-size:15px; margin-bottom:2px; }
.sub { color:var(--muted); font-size:12.5px; margin-bottom:8px; }
.kpi-strip { display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:14px; margin-bottom:16px; }
.kpi { background:#fff; border-left:5px solid #6366f1; border-radius:12px; padding:12px 16px;
       box-shadow:0 1px 3px rgba(15,23,42,.06), 0 6px 18px rgba(15,23,42,.05); }
.kpi-lbl { font-size:11px; text-transform:uppercase; letter-spacing:.07em; color:var(--muted); font-weight:600; }
.kpi-val { font-size:27px; font-weight:700; line-height:1.2; }
.kpi-sub { font-size:12.5px; color:var(--muted); }
.profile { line-height:1.8; font-size:14px; overflow-wrap:anywhere; }
.profile-id { font-size:17px; font-weight:700; color:#312e81; }
.pill { display:inline-block; padding:2px 11px; border-radius:999px; color:#fff; font-size:12px; font-weight:600; }
.insights { columns:2; column-gap:44px; padding-left:18px; margin:4px 0; }
.insights li { margin-bottom:10px; font-size:14px; line-height:1.5; break-inside:avoid; }
@media (max-width: 900px) { .insights { columns:1; } }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def kpi_card(label, value, sub="", colour="#6366f1") -> str:
    return (f'<div class="kpi" style="border-left-color:{colour};"><div class="kpi-lbl">{label}</div>'
            f'<div class="kpi-val">{value}</div><div class="kpi-sub">{sub}</div></div>')


def kpi_strip(cards: list[str]):
    st.markdown(f'<div class="kpi-strip">{"".join(cards)}</div>', unsafe_allow_html=True)


def card_head(title: str | None = None, sub: str | None = None):
    if title:
        st.markdown(f'<div class="card-title">{title}</div>', unsafe_allow_html=True)
    if sub:
        st.markdown(f'<div class="sub">{sub}</div>', unsafe_allow_html=True)


def pl_style(fig: go.Figure, legend: bool = True, height: int | None = None) -> go.Figure:
    ax = dict(title_text="", gridcolor="#eef0f4", zeroline=False, automargin=True)
    fig.update_layout(font=dict(family="Inter, sans-serif", size=12, color="#334155"),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=legend,
                      legend=dict(orientation="h", x=0, y=1, yanchor="bottom"),
                      margin=dict(l=10, r=10, t=40, b=10), hoverlabel=dict(font_family="Inter, sans-serif"))
    fig.update_xaxes(**ax)
    fig.update_yaxes(**ax)
    if height:
        fig.update_layout(height=height)
    return fig


# Streamlit >= 1.50 takes width="stretch"; older versions take use_container_width=True.
# Passing the wrong one lands in **kwargs and triggers the "keyword arguments have been deprecated" warning.
_WIDTH_KW = ({"width": "stretch"} if "width" in inspect.signature(st.plotly_chart).parameters
             else {"use_container_width": True})


def show(fig: go.Figure):
    st.plotly_chart(fig, config={"displayModeBar": False}, **_WIDTH_KW)


def table(df: pd.DataFrame, **kw):
    st.dataframe(df, hide_index=True, **kw)


def stacked_status_bars(d: pd.DataFrame, cat: str, share: str = "share", horizontal: bool = True,
                        count: str | None = None, hover_extra: str = "") -> go.Figure:
    """one trace per status; d has columns cat, status, share (and optionally a count column)"""
    fig = go.Figure()
    for s in STATUS_LEVELS:
        sub = d[d["status"] == s]
        if sub.empty:
            continue
        kw = dict(name=s, marker_color=STATUS_COLS[s])
        if horizontal:
            kw.update(x=sub[share], y=sub[cat], orientation="h",
                      text=[f"{v:.0%}" if v >= 0.07 else "" for v in sub[share]],
                      textposition="inside", insidetextanchor="middle", textfont=dict(color="white", size=12))
            if count:
                kw.update(customdata=sub[count], hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:.1%} "
                                                               "(%{customdata:,} learners)<extra></extra>")
            else:
                kw.update(hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:.1%}<extra></extra>")
        fig.add_trace(go.Bar(**kw))
    return fig


# ---------------------------------------------------------------- sidebar: global filters
def _reset():
    st.session_state.update(f_course="all", f_area="All", stall_days=DEFAULT_STALL)


st.session_state.setdefault("f_course", "all")
st.session_state.setdefault("f_area", "All")
st.session_state.setdefault("stall_days", DEFAULT_STALL)

with st.sidebar:
    st.markdown("### Filters")
    course_labels = {"all": "All courses", **course_only}
    f_course = st.selectbox("Course", list(course_labels), format_func=lambda x: course_labels.get(x, x), key="f_course")
    f_area = st.selectbox("Learning area", area_choices, key="f_area")
    stall_days = st.slider("Stalled after (days idle)", 7, 90, key="stall_days")
    st.button("Reset", on_click=_reset)
    scope_slot = st.empty()
    st.caption("Stalled = unfinished and idle for the chosen days (default 30, the glossary's Inactive rule). "
               "The Course filter applies to Overview, Enrolment & Reach, Segments and Follow-up; "
               "Funnel and Learner pages have their own course selector.")

# ---------------- shared data (pf = learning-area filter, p = + course filter)
pf = prog_base if f_area == "All" else prog_base[prog_base["learning_area"] == f_area]
pf = add_status(pf, stall_days)
p = pf if f_course == "all" else pf[pf["course"] == f_course]
scope_slot.markdown(f'<div class="scope"><div class="scope-n">{comma(len(p))}</div>'
                    f'<div class="scope-l">enrolments in view</div></div>', unsafe_allow_html=True)

st.markdown('<div class="topbar"><h1>Student Progress Dashboard</h1>'
            '<span>Learner status, stage funnels, segments and follow-up lists</span></div>', unsafe_allow_html=True)

tab_over, tab_reach, tab_fun, tab_seg, tab_fu, tab_trk = st.tabs(
    ["Overview", "Enrolment & Reach", "Course Funnel", "Segments", "Follow-up List", "Learner Tracker"])

# ================================================================ Overview
with tab_over:
    d, n = p, len(p)
    if n == 0:
        st.info("No learners match the current filters.")
    else:
        s = lambda x: int((d["status"] == x).sum())
        pc = lambda x: pct(s(x) / n)
        kpi_strip([
            kpi_card("Enrolments", comma(n), f"avg. {r1(d['pct_complete'].mean())}% of activities done", "#0ea5e9"),
            kpi_card("Completed", pc("Completed"), f"{comma(s('Completed'))} finished every activity", STATUS_COLS["Completed"]),
            kpi_card("Active", pc("Active"), f"{comma(s('Active'))} still moving", STATUS_COLS["Active"]),
            kpi_card("Stalled", pc("Stalled"), f"{comma(s('Stalled'))} idle {stall_days}+ days", STATUS_COLS["Stalled"]),
            kpi_card("Not started", pc("Not started"), f"{comma(s('Not started'))} enrolled, nothing done", STATUS_COLS["Not started"]),
        ])
        c1, c2 = st.columns([7, 5])
        with c1, st.container(border=True):
            card_head("Where each course stands",
                      "Share of enrolled learners by status (learning-area filter applies). Click a legend item to hide a status.")
            g = pf.groupby(["course_name", "status"], observed=True).size().reset_index(name="n")
            g["share"] = g["n"] / g.groupby("course_name")["n"].transform("sum")
            order = (g[g["status"] == "Completed"].set_index("course_name")["share"]
                     .reindex(g["course_name"].unique()).fillna(0).sort_values().index.tolist())
            fig = stacked_status_bars(g, "course_name", count="n")
            pl_style(fig, height=400).update_layout(barmode="stack")
            fig.update_xaxes(tickformat=".0%", range=[0, 1])
            fig.update_yaxes(categoryorder="array", categoryarray=order)
            show(fig)
        with c2, st.container(border=True):
            card_head("How far learners have got",
                      "Learners by share of activities completed; same colours as the chart on the left.")
            g = d.groupby(["progress_band", "status"], observed=True).size().reset_index(name="n")
            fig = go.Figure([go.Bar(x=g[g.status == s_]["progress_band"], y=g[g.status == s_]["n"], name=s_,
                                    marker_color=STATUS_COLS[s_],
                                    hovertemplate="%{x}<br>%{fullData.name}: %{y:,} learners<extra></extra>")
                             for s_ in STATUS_LEVELS])
            pl_style(fig, legend=False, height=400).update_layout(barmode="stack")
            fig.update_xaxes(categoryorder="array", categoryarray=BAND_LEVELS)
            fig.update_yaxes(tickformat=",")
            show(fig)

        with st.container(border=True):
            card_head("What stands out")
            items = []
            near = int(((d["status"] == "Stalled") & (d["pct_complete"] >= 75)).sum())
            if near > 0:
                items.append(f"<b>{comma(near)} stalled learners are 75%+ through</b> their course: the quickest wins for a reminder.")
            failed = int((d["failed_assessments"] > 0).sum())
            if failed > 0:
                items.append(f"<b>{comma(failed)} learners</b> finished an assessment without reaching the pass grade.")
            stuck = int(((d["completed"] == 0) & (d["pct_complete"] >= 95)).sum())
            if stuck > 0:
                items.append(f"<b>{comma(stuck)} learners sit at 95%+ without finishing</b>: worth checking for a "
                             "course-completion trigger or access problem (flagged in the Town Hall).")
            if f_course == "all":
                w = d.groupby("course_name", observed=True).agg(s=("status", lambda x: (x == "Stalled").mean()),
                                                                c=("status", lambda x: (x == "Completed").mean()))
                items.append(f"Highest completion: <b>{w['c'].idxmax()}</b> ({pct(w['c'].max())}). "
                             f"Highest stalled share: <b>{w['s'].idxmax()}</b> ({pct(w['s'].max())}).")
            else:
                stc = (d[d["status"] == "Stalled"].groupby("furthest_stage").size()
                       .sort_values(ascending=False, kind="stable"))
                if len(stc):
                    items.append(f"Most stalled learners stopped at <b>{stc.index[0]}</b> ({pct(stc.iloc[0] / stc.sum(), 1)} of the stalled).")
            gap = d.loc[d["items_done"] > 0, "joined_to_first_days"].median()
            if np.isfinite(gap):
                items.append(f"Median wait between registering and the first completed activity: <b>{r1(gap)} days</b>.")
            dc = ((d["last_activity"] - d["date_joined"]).dt.total_seconds() / 86400)[d["status"] == "Completed"]
            if dc.notna().any():
                items.append(f"Median time from registering to completing the course: <b>{r1(dc.median())} days</b>.")
            if items:
                st.markdown('<ul class="insights">' + "".join(f"<li>{x}</li>" for x in items) + "</ul>",
                            unsafe_allow_html=True)
            else:
                st.info("Nothing notable for this selection.")

# ======================================================= Enrolment & Reach
with tab_reach:
    ppl = PARTS.drop_duplicates("unique_key")
    if f_area != "All":
        ppl = ppl[ppl["learning_area"] == f_area]
    if f_course != "all":
        ppl = ppl[ppl["unique_key"].isin(p["unique_key"])]
    n = len(ppl)
    if n == 0:
        st.info("No participants match the current filters.")
    else:
        v = [pct_of(ppl["gender"], lambda x: x.fillna("").str.lower().eq("female")),
             pct_of(ppl["pwd_status"], is_yes), pct_of(ppl["displaced_youth"], is_yes)]
        tg = list(TARGETS.values())

        def tgt(i):
            if pd.isna(v[i]):
                return "not in the export"
            return f"target {pct(tg[i], 1)}" + (" (met)" if v[i] >= tg[i] else f" ({r1((tg[i] - v[i]) * 100)} pp short)")

        col = lambda i: STATUS_COLS["Completed"] if (pd.notna(v[i]) and v[i] >= tg[i]) else "#f59e0b"
        whole = f_course == "all" and f_area == "All"
        kpi_strip([
            kpi_card("Registered", comma(n),
                     f"{pct(n / REACH_TARGET)} of the {comma(REACH_TARGET)} five-year target" if whole
                     else "participants in the current view", "#0ea5e9"),
            kpi_card("Female", fmt_pct(v[0]), tgt(0), col(0)),
            kpi_card("Persons with disability", fmt_pct(v[1]), tgt(1), col(1)),
            kpi_card("Displaced youth", fmt_pct(v[2]), tgt(2), col(2)),
            kpi_card("Enrolled in a course", fmt_pct((ppl["courses_enrolled"] > 0).mean()), "of those registered", "#6366f1"),
        ])
        c1, c2 = st.columns([7, 5])
        with c1, st.container(border=True):
            card_head("Registrations over time",
                      "Participants by month joined (bars) and cumulative total (line). Follows the Course and Learning-area filters.")
            m = (ppl[ppl["date_joined"].notna()].assign(month=lambda x: x["date_joined"].dt.strftime("%Y-%m"))
                 .groupby("month").size().reset_index(name="n").sort_values("month"))
            m["cum"] = m["n"].cumsum()
            if m.empty:
                st.info("No join dates for this selection.")
            else:
                fig = make_subplots(specs=[[{"secondary_y": True}]])
                fig.add_trace(go.Bar(x=m["month"], y=m["n"], name="Joined in month", marker_color="#6366f1",
                                     hovertemplate="%{x}<br>%{y:,} joined<extra></extra>"), secondary_y=False)
                fig.add_trace(go.Scatter(x=m["month"], y=m["cum"], name="Cumulative", mode="lines",
                                         line=dict(color="#0ea5e9", width=3),
                                         hovertemplate="%{x}<br>%{y:,} registered to date<extra></extra>"), secondary_y=True)
                pl_style(fig, height=380)
                fig.update_xaxes(type="category")
                fig.update_yaxes(tickformat=",")
                fig.update_yaxes(showgrid=False, secondary_y=True)
                show(fig)
        with c2, st.container(border=True):
            card_head("Inclusion (GESI) vs programme targets",
                      "Share of registered participants; diamond = programme target.")
            g = pd.DataFrame({"group": list(TARGETS), "target": tg, "actual": v}).dropna(subset=["actual"])
            if g.empty:
                st.info("These fields are not in the export.")
            else:
                colors = np.where(g["actual"] >= g["target"], "#10b981", "#f59e0b")
                fig = go.Figure()
                fig.add_trace(go.Bar(x=g["actual"], y=g["group"], orientation="h", name="Actual", marker_color=colors,
                                     text=[pct(a) for a in g["actual"]], textposition="outside",
                                     hovertemplate="%{y}: %{x:.1%}<extra></extra>"))
                fig.add_trace(go.Scatter(x=g["target"], y=g["group"], mode="markers", name="Programme target",
                                         marker=dict(symbol="diamond", size=13, color="#0f172a"),
                                         hovertemplate="Target: %{x:.0%}<extra></extra>"))
                pl_style(fig, height=380)
                fig.update_xaxes(tickformat=".0%", range=[0, min(1, max(g["actual"].max(), g["target"].max()) * 1.3)])
                fig.update_yaxes(categoryorder="array", categoryarray=g["group"].tolist()[::-1])
                show(fig)

        c3, c4 = st.columns(2)
        with c3, st.container(border=True):
            card_head("Onboarding lag: registration to first activity",
                      "Enrolments by days between date joined and first completed activity.")
            lv = ["Within 1 day", "2-7 days", "8-14 days", "15-30 days", "Over 30 days", "Never started"]
            j = p["joined_to_first_days"]
            band = np.select([p["items_done"] == 0, j.isna(), j <= 1, j <= 7, j <= 14, j <= 30],
                             ["Never started", "", "Within 1 day", "2-7 days", "8-14 days", "15-30 days"],
                             default="Over 30 days")
            b = pd.Series(band[band != ""]).value_counts().reindex(lv).dropna().reset_index()
            b.columns = ["band", "n"]
            if b.empty:
                st.info("No data for this selection.")
            else:
                b["share"] = b["n"] / b["n"].sum()
                palette = dict(zip(lv, ["#10b981", "#84cc16", "#f59e0b", "#f97316", "#f43f5e", "#94a3b8"]))
                fig = go.Figure(go.Bar(x=b["band"], y=b["n"], marker_color=[palette[x] for x in b["band"]],
                                       text=[pct(s_, 1) for s_ in b["share"]], textposition="outside",
                                       hovertemplate="%{x}<br>%{y:,} enrolments (%{text})<extra></extra>"))
                pl_style(fig, legend=False, height=340)
                fig.update_xaxes(categoryorder="array", categoryarray=lv)
                fig.update_yaxes(tickformat=",")
                show(fig)
        with c4, st.container(border=True):
            card_head("Data-quality checks", "Rules taken from the notes in the GATEWAY glossary.")
            nan_all = lambda s_: pd.Series(np.nan, index=s_.index)
            not_yes = lambda s_: nan_all(s_) if s_.isna().all() else ~is_yes(s_)
            area_mis = (nan_all(ppl["learning_area"]) if ppl["skill_assessment_learning_area"].isna().all() else
                        ppl["learning_area"].fillna("").str.strip().str.lower()
                        != ppl["skill_assessment_learning_area"].fillna("").str.strip().str.lower())
            checks = [
                ("Consent not recorded as Yes", not_yes(ppl["accept_consent"]), "No record should be active without consent"),
                ("Email not verified", not_yes(ppl["email_verified"]), "Low verification explains communication gaps"),
                ("Duplicate email", nan_all(ppl["duplicate_email"]) if ppl["duplicate_email"].isna().all()
                 else ppl["duplicate_email"] == "Yes", "Same person may be registered twice"),
                ("Learning area differs from skill-assessment area", area_mis, "Flag mismatches or track changes"),
                ("Age outside 18-35", nan_all(ppl["age_band"]) if ppl["age_band"].isna().all()
                 else ppl["age_band"].isin(["Under 18", "36+"]), "Programme eligibility range"),
            ]
            rows = []
            for label, f, why in checks:
                na_all = f.isna().all()
                k = 0 if na_all else int(f.fillna(False).astype(bool).sum())
                rows.append({"Check": label, "Records": "n/a" if na_all else comma(k),
                             "Share": "n/a" if na_all else pct(k / n), "Why it matters": why})
            table(pd.DataFrame(rows))

# ============================================================ Course Funnel
with tab_fun:
    c1, c2 = st.columns([4, 8])
    codes = list(course_only)
    with c1, st.container(border=True):
        fun_course = st.selectbox("Course to inspect", codes, format_func=lambda x: course_only.get(x, x),
                                  index=codes.index("ggw") if "ggw" in codes else 0)
    fc = pf[pf["course"] == fun_course]
    keys = fc["unique_key"]
    base = (struct[struct["course"] == fun_course].drop_duplicates(["stage", "stage_order"])
            [["stage", "stage_order"]].sort_values("stage_order"))
    stage_levels = base["stage"].tolist()
    sprog = dat["Stage_Progress"]
    sprog = sprog[(sprog["course"] == fun_course) & sprog["unique_key"].isin(keys)]
    _sp = sprog.assign(_span=sprog["span_days"].where(sprog["stage_complete"] == 1))
    agg = _sp.groupby("stage").agg(started=("stage", "size"), completed=("stage_complete", "sum"),
                                   median_span_days=("_span", "median")).reset_index()
    stg = base.merge(agg, on="stage", how="left")
    stg[["started", "completed"]] = stg[["started", "completed"]].fillna(0).astype(int)
    stg["enrolled"] = len(keys)
    stg["pct_started"] = stg["started"] / max(len(keys), 1)
    stg["pct_completed"] = stg["completed"] / max(len(keys), 1)
    stg = stg.sort_values("stage_order")

    with c2:
        if len(fc) == 0:
            st.info("No learners for this course and learning area.")
        else:
            drops = stg.assign(drop_pp=(stg["pct_started"].shift(1) - stg["pct_started"]) * 100).dropna(subset=["drop_pp"])
            big = drops.loc[drops["drop_pp"].idxmax()] if len(drops) else None
            kpi_strip([
                kpi_card("Enrolled", comma(len(fc)), "in this course", "#0ea5e9"),
                kpi_card("Completed course", pct((fc["status"] == "Completed").mean()),
                         f"{comma((fc['status'] == 'Completed').sum())} learners", STATUS_COLS["Completed"]),
                kpi_card("Never started", pct((fc["status"] == "Not started").mean()),
                         f"{comma((fc['status'] == 'Not started').sum())} learners", STATUS_COLS["Not started"]),
                kpi_card("Biggest drop-off", big["stage"] if big is not None else "n/a",
                         f"{r1(big['drop_pp'])} pp fewer start it than the stage before" if big is not None else "",
                         STATUS_COLS["Stalled"]),
            ])

    with st.container(border=True):
        card_head("Stage progression: who got through, and where the rest stopped",
                  "Left: share of enrolled learners who started / completed each stage (hover for median days to finish). "
                  "Right: unfinished learners by the furthest stage reached.")
        stop_d = (fc[(fc["status"] != "Completed") & fc["furthest_stage"].isin(stage_levels)]
                  .groupby(["furthest_stage", "status"], observed=True).size().reset_index(name="n"))
        if stg.empty or stop_d.empty:
            st.info("No unfinished learners for this selection.")
        else:
            stg["still_open"] = (stg["pct_started"] - stg["pct_completed"]).clip(lower=0)
            stg["hover"] = [f"<b>{r.stage}</b><br>Started: {pct(r.pct_started)} of enrolled<br>"
                            f"Completed: {pct(r.pct_completed)} of enrolled<br>Median time to finish: "
                            f"{'n/a' if pd.isna(r.median_span_days) else r1(r.median_span_days) + ' days'}"
                            for r in stg.itertuples()]
            fig = make_subplots(rows=1, cols=2, shared_yaxes=True, column_widths=[0.62, 0.38], horizontal_spacing=0.025)
            fig.add_trace(go.Bar(x=stg["pct_completed"], y=stg["stage"], orientation="h", name="Completed stage",
                                 marker_color="#10b981", hoverinfo="text", hovertext=stg["hover"],
                                 text=[f"{v:.0%}" if v >= 0.06 else "" for v in stg["pct_completed"]],
                                 textposition="inside", insidetextanchor="middle", textfont=dict(color="white")), 1, 1)
            fig.add_trace(go.Bar(x=stg["still_open"], y=stg["stage"], orientation="h", name="Started, not finished",
                                 marker_color="#a7f3d0", hoverinfo="text", hovertext=stg["hover"]), 1, 1)
            for s_ in STATUS_LEVELS:
                sub = stop_d[stop_d["status"] == s_]
                if len(sub):
                    fig.add_trace(go.Bar(x=sub["n"], y=sub["furthest_stage"], orientation="h", name=s_,
                                         marker_color=STATUS_COLS[s_],
                                         hovertemplate="%{y}<br>%{fullData.name}: %{x:,} learners<extra></extra>"), 1, 2)
            pl_style(fig, height=560).update_layout(barmode="stack", bargap=0.3)
            fig.update_xaxes(tickformat=".0%", range=[0, 1], title_text="Share of enrolled learners", col=1)
            fig.update_xaxes(tickformat=",", title_text="Learners who stopped here", col=2)
            fig.update_yaxes(categoryorder="array", categoryarray=stage_levels[::-1])
            show(fig)

    with st.container(border=True):
        card_head("Activity-level drop-off",
                  "All learners in the course (learning-area filter does not apply). Sorted by the biggest fall versus the previous activity.")
        it = dat["Item_Funnel"]
        it = it[it["course"] == fun_course].rename(columns={
            "item_order": "#", "stage": "Stage", "item_label": "Activity", "pct_completed": "% completed",
            "drop_vs_prev_item_pp": "Drop vs previous (pp)", "failed_pass_grade": "Did not pass"})
        it = (it[["#", "Stage", "Activity", "% completed", "Drop vs previous (pp)", "Did not pass"]]
              .sort_values("Drop vs previous (pp)", ascending=False, na_position="last"))

        def _drop_colour(x):
            if pd.isna(x):
                return ""
            return "color:#e11d48;font-weight:bold" if x > 10 else ("color:#d97706;font-weight:bold" if x > 5
                                                                    else "color:#64748b;font-weight:bold")
        sty = it.style.map(_drop_colour, subset=["Drop vs previous (pp)"])
        table(sty, column_config={"% completed": st.column_config.ProgressColumn(
            "% completed", min_value=0, max_value=100, format="%.1f%%")})

# ================================================================ Segments
with tab_seg:
    DIMS = {"Learning area": "learning_area", "Gender": "gender", "Age band": "age_band", "State": "state",
            "Internet access": "internet_access", "Has a computer": "has_computer",
            "Has power supply": "has_power_supply", "Pathway": "pathway", "Disability status": "pwd_status",
            "Displaced youth": "displaced_youth", "Has children": "has_children",
            "Skill assessment status": "skill_assessment_status",
            "Computer literacy status": "computer_literacy_status", "Heard about us from": "heard_from"}
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        seg_dim = DIMS[c1.selectbox("Break down by", list(DIMS))]
        seg_min = c2.slider("Minimum group size", 10, 500, 50, step=10)
        seg_sort = {"Completion rate": "Completed", "Stalled share": "Stalled", "Group size": "n"}[
            c3.selectbox("Sort groups by", ["Completion rate", "Stalled share", "Group size"])]

    g = p.assign(grp=p[seg_dim].astype("object")).dropna(subset=["grp"]).astype({"grp": str})
    seg = (g.groupby("grp").agg(n=("grp", "size"),
                                Completed=("status", lambda x: (x == "Completed").mean()),
                                Active=("status", lambda x: (x == "Active").mean()),
                                Stalled=("status", lambda x: (x == "Stalled").mean()),
                                **{"Not started": ("status", lambda x: (x == "Not started").mean())},
                                avg_pct=("pct_complete", "mean")).reset_index())
    seg = seg[seg["n"] >= seg_min].sort_values("n", ascending=False).head(15)

    with st.container(border=True):
        t_mix, t_num = st.tabs(["Status mix", "Numbers"])
        with t_mix:
            st.markdown('<div class="sub" style="margin-top:8px;">The 15 largest groups that meet the minimum size; '
                        'group size in brackets.</div>', unsafe_allow_html=True)
            if seg.empty:
                st.info("No group reaches the minimum size.")
            else:
                seg["label"] = [f"{a}  (n={comma(b)})" for a, b in zip(seg["grp"], seg["n"])]
                order = seg.sort_values(seg_sort)["label"].tolist()
                long = seg.melt(id_vars=["label"], value_vars=STATUS_LEVELS, var_name="status", value_name="share")
                fig = stacked_status_bars(long, "label")
                pl_style(fig, height=520).update_layout(barmode="stack")
                fig.update_xaxes(tickformat=".0%", range=[0, 1])
                fig.update_yaxes(categoryorder="array", categoryarray=order)
                show(fig)
        with t_num:
            if seg.empty:
                st.info("No group reaches the minimum size.")
            else:
                num = pd.DataFrame({"Group": seg["grp"], "Participants": seg["n"],
                                    "Completed %": (seg["Completed"] * 100).round(1),
                                    "Active %": (seg["Active"] * 100).round(1),
                                    "Stalled %": (seg["Stalled"] * 100).round(1),
                                    "Not started %": (seg["Not started"] * 100).round(1),
                                    "Avg % complete": seg["avg_pct"].round(1)})
                table(num, column_config={
                    "Completed %": st.column_config.ProgressColumn("Completed %", min_value=0, max_value=100, format="%.1f"),
                    "Stalled %": st.column_config.ProgressColumn("Stalled %", min_value=0, max_value=100, format="%.1f")})

# ============================================================== Follow-up
FU_TYPES = ["Stalled: near completion (75%+)", "Stalled: mid-way (25-74%)", "Stalled: early (under 25%)",
            "All stalled", "Stuck at 95%+ (unfinished)", "Not started", "Did not pass an assessment"]
with tab_fu:
    with st.container(border=True):
        c1, c2, c3 = st.columns([5, 3, 4])
        fu_type = c1.selectbox("Who needs a nudge?", FU_TYPES)
        stalled = p["status"] == "Stalled"
        mask = {
            FU_TYPES[0]: stalled & (p["pct_complete"] >= 75),
            FU_TYPES[1]: stalled & (p["pct_complete"] >= 25) & (p["pct_complete"] < 75),
            FU_TYPES[2]: stalled & (p["pct_complete"] < 25),
            FU_TYPES[3]: stalled,
            FU_TYPES[4]: (p["completed"] == 0) & (p["pct_complete"] >= 95),
            FU_TYPES[5]: p["status"] == "Not started",
            FU_TYPES[6]: p["failed_assessments"] > 0}[fu_type]
        fu = (p[mask].sort_values(["pct_complete", "days_since_last_activity"], ascending=[False, True])
              .assign(course=lambda x: x["course_name"], last_stage_reached=lambda x: x["furthest_stage"],
                      status=lambda x: x["status"].astype(str))
              [["registration_id", "name", "email", "phone", "course", "learning_area", "state", "gender",
                "pct_complete", "items_remaining", "stages_remaining", "last_stage_reached",
                "days_since_last_activity", "days_since_joined", "failed_assessments", "status"]])
        c2.markdown(f'<div style="margin-top:26px;"><span class="pill" style="background:#4338ca;font-size:14px;'
                    f'padding:6px 16px;">{comma(len(fu))} learners</span></div>', unsafe_allow_html=True)
        c3.download_button("Download list (CSV)", drop_empty_pii(fu).to_csv(index=False, na_rep="").encode("utf-8"),
                           file_name=f"follow_up_{date.today()}.csv", mime="text/csv")

    c1, c2 = st.columns([8, 4])
    with c1, st.container(border=True):
        card_head("Learners in this list",
                  "Ordered by progress, so the learners closest to finishing come first. The CSV has a few extra columns. "
                  "Contains personal data, so share with care.")
        q = st.text_input("Search registration ID, name or email", key="fu_search").strip().lower()
        view = fu
        if q:
            hay = view[["registration_id", "name", "email"]].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
            view = view[hay.str.contains(q, regex=False)]
        tbl = drop_empty_pii(pd.DataFrame({
            "Registration ID": view["registration_id"], "Name": view["name"], "Email": view["email"],
            "Phone": view["phone"], "Course": view["course"], "Learning area": view["learning_area"],
            "State": view["state"], "% complete": view["pct_complete"], "Activities left": view["items_remaining"],
            "Last stage reached": view["last_stage_reached"],
            "Days idle": view["days_since_last_activity"].round(0), "Failed assessments": view["failed_assessments"]}))
        table(tbl, column_config={"% complete": st.column_config.ProgressColumn(
            "% complete", min_value=0, max_value=100, format="%.1f")})
    with c2, st.container(border=True):
        card_head("Where they stopped", "Last stage reached by the learners in the list (top 12).")
        d = (fu.groupby(["course", "last_stage_reached"]).size().reset_index(name="n")
             .sort_values("n", ascending=False, kind="stable").head(12))
        if d.empty:
            st.info("Nobody in this list.")
        else:
            d["label"] = (d["last_stage_reached"] + " \u00b7 " + d["course"]) if f_course == "all" else d["last_stage_reached"]
            fig = go.Figure(go.Bar(x=d["n"], y=d["label"], orientation="h", marker_color="#6366f1",
                                   hovertemplate="%{y}<br>%{x:,} learners<extra></extra>"))
            pl_style(fig, legend=False, height=520)
            fig.update_xaxes(tickformat=",")
            fig.update_yaxes(categoryorder="array", categoryarray=d["label"].tolist()[::-1])
            show(fig)

# ========================================================== Learner Tracker
with tab_trk:
    with st.container(border=True):
        c1, c2, c3 = st.columns([5, 4, 3])
        search = c1.text_input("Find a learner (registration ID, name or email)", key="trk_search").strip().lower()
        pool = rid_lookup if not search else rid_lookup[rid_lookup["label"].str.lower().str.contains(search, regex=False)]
        pool = pool.head(500)                                   # keep the dropdown light for big cohorts
        if pool.empty:
            c1.warning("No learner matches that search.")
            rid = None
        else:
            lab = dict(zip(pool["registration_id"], pool["label"]))
            rid = c1.selectbox("Learner", list(lab), format_func=lambda x: lab.get(x, x))
            if len(pool) == 500:
                c1.caption("Showing the first 500 matches: type more of the ID, name or email to narrow down.")
        lrn = add_status(prog_base[prog_base["registration_id"] == rid], stall_days) if rid else prog_base.iloc[0:0]
        if len(lrn):
            lrn = lrn.sort_values("items_done", ascending=False)
            cl = dict(zip(lrn["course"], lrn["course_name"]))
            trk_course = c2.selectbox("Course", list(cl), format_func=lambda x: cl.get(x, x))
        else:
            trk_course = None
        act_show = c3.radio("Activity list", ["To do", "Done", "All"], horizontal=True)

    if trk_course is not None:
        r = lrn[lrn["course"] == trk_course].iloc[0]
        pr = PARTS[PARTS["unique_key"] == r["unique_key"]].iloc[0]
        na_txt = lambda x: "n/a" if pd.isna(x) else str(x)
        stt = str(r["status"])

        # ---- profile + KPIs
        c1, c2 = st.columns([4, 8])
        with c1, st.container(border=True):
            age_txt = "age n/a" if pd.isna(pr["age"]) else f"{int(pr['age'])} yrs"
            joined_txt = "n/a" if pd.isna(pr["date_joined"]) else f"{pr['date_joined']:%d %b %Y}"
            phone = f'<div>Phone: {pr["phone"]}</div>' if pd.notna(pr["phone"]) else ""
            st.markdown(
                f'<div class="profile"><div><span class="profile-id">{r["registration_id"]}</span> '
                f'<span class="pill" style="background:{STATUS_COLS[stt]};">{stt}</span></div>'
                f'<div><b>{na_txt(pr["name"])}</b> \u00b7 {na_txt(pr["email"])}</div>{phone}'
                f'<div>{na_txt(pr["gender"])} \u00b7 {age_txt} \u00b7 {na_txt(pr["state"])}</div>'
                f'<div>Learning area: <b>{na_txt(pr["learning_area"])}</b></div>'
                f'<div>Joined {joined_txt} \u00b7 enrolled in <b>{len(lrn)}</b> course(s)</div></div>',
                unsafe_allow_html=True)
        with c2:
            span_txt = "n/a" if pd.isna(r["total_span_days"]) else f"{r1(r['total_span_days'])} d"
            idle_txt = "none" if pd.isna(r["days_since_last_activity"]) else f"{r1(r['days_since_last_activity'])} d ago"
            first_txt = "n/a" if pd.isna(r["joined_to_first_days"]) else f"{r1(r['joined_to_first_days'])} d"
            kpi_strip([
                kpi_card("Course progress", f"{r1(r['pct_complete'])}%",
                         f"{r['items_done']} of {r['items_total']} activities \u00b7 {r['items_remaining']} to go",
                         STATUS_COLS["Completed"]),
                kpi_card("Stages completed", f"{r['stages_completed']} / {r['stages_total']}",
                         f"Furthest: {r['furthest_stage']}", STATUS_COLS["Active"]),
                kpi_card("Last activity", idle_txt, f"Joined to first activity: {first_txt}",
                         STATUS_COLS["Stalled"] if stt == "Stalled" else "#f59e0b"),
                kpi_card("Time in course", span_txt, f"Active on {r['active_days']} day(s)", "#06b6d4"),
            ])

        # ---- stage progress + checklist
        sc = (struct[struct["course"] == r["course"]].groupby(["stage", "stage_order"]).size()
              .reset_index(name="items_total"))
        sp = dat["Stage_Progress"]
        sp = sp[(sp["course"] == r["course"]) & (sp["unique_key"] == r["unique_key"])][
            ["stage", "items_done", "first_completed", "last_completed", "span_days"]]
        ls = sc.merge(sp, on="stage", how="left").sort_values("stage_order")
        ls["items_done"] = ls["items_done"].fillna(0)
        ls["pct"] = ls["items_done"] / ls["items_total"]
        ls["stage_status"] = np.select([ls["items_done"] >= ls["items_total"], ls["items_done"] > 0],
                                       ["Complete", "In progress"], default="Not started")

        c1, c2 = st.columns([5, 7])
        with c1, st.container(border=True):
            card_head("Progress through the stages", "Hover a bar for dates and days spent in that stage.")
            f_d = lambda x: "-" if pd.isna(x) else f"{x:%d %b %Y}"
            ls["hover"] = [f"<b>{x.stage}</b><br>{int(x.items_done)} of {x.items_total} activities done ({pct(x.pct, 1)})"
                           f"<br>Time in stage: {'n/a' if pd.isna(x.span_days) else r1(x.span_days) + ' days'}"
                           f"<br>First: {f_d(x.first_completed)} \u00b7 Last: {f_d(x.last_completed)}"
                           for x in ls.itertuples()]
            cols = {"Complete": "#10b981", "In progress": "#f59e0b", "Not started": "#cbd5e1"}
            fig = go.Figure()
            for s_, colr in cols.items():
                sub = ls[ls["stage_status"] == s_]
                if len(sub):
                    fig.add_trace(go.Bar(x=sub["pct"], y=sub["stage"], orientation="h", name=s_, marker_color=colr,
                                         text=[f"{int(a)}/{b}" for a, b in zip(sub["items_done"], sub["items_total"])],
                                         textposition="outside", hoverinfo="text", hovertext=sub["hover"]))
            pl_style(fig, height=520).update_layout(barmode="overlay", bargap=0.3)
            fig.update_xaxes(tickvals=[0, .25, .5, .75, 1], tickformat=".0%", range=[0, 1.18])
            fig.update_yaxes(categoryorder="array", categoryarray=ls["stage"].tolist()[::-1])
            show(fig)
        with c2, st.container(border=True):
            card_head("Activity checklist",
                      "Use the radio buttons above to switch between what is still to do and what is done.")
            ids = ([] if pd.isna(r["done_items"]) or r["done_items"] == ""
                   else [int(i) for i in str(r["done_items"]).split(",")])
            items = struct[struct["course"] == r["course"]].sort_values("item_order").copy()
            items["Status"] = np.where(items["item_order"].isin(ids), "Done", "To do")
            if act_show != "All":
                items = items[items["Status"] == act_show]
            out = items.rename(columns={"item_order": "#", "stage": "Stage", "item_label": "Activity"})[
                ["#", "Stage", "Activity", "Status"]]
            sty = out.style.map(lambda x: "color:#059669;font-weight:bold" if x == "Done"
                                else ("color:#d97706;font-weight:bold" if x == "To do" else ""), subset=["Status"])
            table(sty)


# # =====================================================================
# #  Student Progress Dashboard  -  Streamlit port of app.R
# #  Run 1_gateway_clean.py FIRST (it writes output/progression_data.xlsx), then:
# #      streamlit run 2_dashboard.py
# #  Needs: streamlit>=1.50, pandas, numpy, plotly, openpyxl
# #  Data : progression_data.xlsx in ./output or next to this file
# # =====================================================================
# from __future__ import annotations

# import pickle
# from datetime import date
# from pathlib import Path

# import numpy as np
# import pandas as pd
# import plotly.graph_objects as go
# import streamlit as st
# from plotly.subplots import make_subplots

# st.set_page_config(page_title="Student Progress Dashboard", page_icon="📊", layout="wide")

# # ---- dashboard settings (values come from the GATEWAY glossary) -------------
# SHOW_PII = True        # False hides learner name / email / phone everywhere (glossary: High PII)
# DEFAULT_STALL = 30     # glossary rule: Inactive = no LMS activity for 30+ days
# REACH_TARGET = 400000  # five-year reach target
# TARGETS = {"Female": 0.75, "Persons with disability": 0.05, "Displaced youth": 0.01}  # GESI targets

# STATUS_LEVELS = ["Completed", "Active", "Stalled", "Not started"]
# STATUS_COLS = {"Completed": "#10b981", "Active": "#6366f1", "Stalled": "#f43f5e", "Not started": "#94a3b8"}
# BAND_LEVELS = ["Not started", "1-24%", "25-49%", "50-74%", "75-99%", "Completed"]

# # ---------------------------------------------------------------- data
# HERE = Path(__file__).resolve().parent
# XLSX = next((p for p in [HERE / "output" / "progression_data.xlsx", HERE / "progression_data.xlsx",
#                          Path("output/progression_data.xlsx"), Path("progression_data.xlsx")] if p.exists()), None)
# if XLSX is None:
#     st.error("progression_data.xlsx not found (looked in ./output and the app folder). "
#              "Run 1_gateway_clean.py first.")
#     st.stop()

# SHEETS = ["Participants", "Progress", "Stage_Progress", "Course_Structure", "Item_Funnel", "Data_Notes", "Data_Glossary"]
# PART_FILL = ["name", "email", "phone", "pathway", "pwd_status", "displaced_youth", "has_children", "has_power_supply",
#              "accept_consent", "email_verified", "duplicate_email", "skill_assessment_status",
#              "skill_assessment_learning_area", "computer_literacy_status", "gender", "age_band", "state",
#              "learning_area", "internet_access", "has_computer", "heard_from"]
# PROG_COLS = ["unique_key", "name", "email", "phone", "gender", "age_band", "state", "learning_area", "pathway",
#              "pwd_status", "displaced_youth", "has_children", "has_power_supply", "internet_access", "has_computer",
#              "skill_assessment_status", "computer_literacy_status", "heard_from"]


# @st.cache_resource(show_spinner="Loading data...")
# def load_all(path_str: str, mtime: float) -> dict:
#     path = Path(path_str)
#     cache = path.with_name("progression_cache.pkl")        # same idea as the .rds cache in app.R
#     dat = None
#     if cache.exists() and cache.stat().st_mtime > mtime:
#         try:
#             dat = pickle.loads(cache.read_bytes())
#         except Exception:
#             dat = None
#     if dat is None:
#         xl = pd.ExcelFile(path)
#         dat = {s: xl.parse(s) for s in SHEETS if s in xl.sheet_names}
#         try:
#             cache.write_bytes(pickle.dumps(dat))
#         except Exception:
#             pass

#     P = dat["Participants"].copy()
#     for cl in PART_FILL:                                    # older exports may lack some columns
#         if cl not in P.columns:
#             P[cl] = np.nan
#     if not SHOW_PII:
#         P[["name", "email", "phone"]] = np.nan
#     for df in (P, dat["Progress"]):
#         df["registration_id"] = df["registration_id"].map(lambda x: None if pd.isna(x) else str(x))
#     dat["Participants"] = P

#     struct = dat["Course_Structure"]
#     course_tbl = struct.drop_duplicates(["course", "course_name"])[["course", "course_name"]]
#     course_only = dict(zip(course_tbl["course"], course_tbl["course_name"]))     # code -> display name

#     part1 = P.drop_duplicates("unique_key")[PROG_COLS]
#     prog_base = dat["Progress"].merge(part1, on="unique_key", how="left")

#     rid = (prog_base[prog_base["registration_id"].notna()].drop_duplicates("registration_id")
#            [["registration_id", "name", "email"]].sort_values("registration_id").reset_index(drop=True))
#     rid["label"] = [" | ".join([r] + ([n] if pd.notna(n) else []) + ([e] if pd.notna(e) else []))
#                     for r, n, e in zip(rid["registration_id"], rid["name"], rid["email"])]
#     return dict(dat=dat, struct=struct, course_only=course_only, prog_base=prog_base,
#                 area_choices=["All"] + sorted(P["learning_area"].dropna().unique().tolist()), rid=rid)


# _D = load_all(str(XLSX), XLSX.stat().st_mtime)
# dat, struct, course_only = _D["dat"], _D["struct"], _D["course_only"]
# prog_base, area_choices, rid_lookup = _D["prog_base"], _D["area_choices"], _D["rid"]
# PARTS = dat["Participants"]


# def add_status(d: pd.DataFrame, days: float) -> pd.DataFrame:
#     d = d.copy()
#     status = np.select([d["completed"] == 1, d["items_done"] == 0, d["days_since_last_activity"] >= days],
#                        ["Completed", "Not started", "Stalled"], default="Active")
#     d["status"] = pd.Categorical(status, categories=STATUS_LEVELS)
#     return d


# # ---------------------------------------------------------- helpers
# def comma(x) -> str:
#     return f"{int(round(float(x))):,}"


# def pct(x, acc: float = 0.1) -> str:
#     return f"{float(x) * 100:.{0 if acc >= 1 else 1}f}%"


# def r1(x) -> str:
#     return f"{x:.1f}".rstrip("0").rstrip(".")


# def is_yes(s: pd.Series) -> pd.Series:
#     return s.fillna("").astype(str).str.strip().str.lower().isin(["yes", "true", "1", "y", "verified"])


# def pct_of(s: pd.Series, f) -> float:
#     return np.nan if s.isna().all() else float(f(s).mean())     # NaN when the export lacks the field


# def fmt_pct(v, acc: float = 0.1) -> str:
#     return "n/a" if pd.isna(v) else pct(v, acc)


# def drop_empty_pii(df: pd.DataFrame) -> pd.DataFrame:
#     pii = [c for c in df.columns if c in ("Name", "Email", "Phone", "name", "email", "phone")]
#     return df.drop(columns=[c for c in pii if df[c].isna().all()])


# CSS = """
# <style>
# @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
# :root { --ink:#0f172a; --muted:#64748b; --line:#e8ebf3; --bg:#f3f5fb; }
# html, body, [class*="css"], .stApp { font-family:'Inter',sans-serif; }
# .stApp { background:var(--bg); }
# .block-container { max-width:1560px; padding-top:1rem; }
# .topbar { background:linear-gradient(120deg,#1e1b4b 0%,#3730a3 100%); color:#fff; margin:0 0 16px;
#           padding:16px 30px; border-radius:14px; box-shadow:0 2px 10px rgba(30,27,75,.25); }
# .topbar h1 { font-size:22px; font-weight:700; margin:0; color:#fff; padding:0; }
# .topbar span { font-size:12.5px; opacity:.8; }
# .scope { background:#eef2ff; border-radius:10px; padding:6px 12px; text-align:center; margin-top:4px; }
# .scope-n { font-size:22px; font-weight:700; color:#3730a3; line-height:1.15; }
# .scope-l { font-size:11.5px; color:var(--muted); }
# .card-title { font-weight:600; font-size:15px; margin-bottom:2px; }
# .sub { color:var(--muted); font-size:12.5px; margin-bottom:8px; }
# .kpi-strip { display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:14px; margin-bottom:16px; }
# .kpi { background:#fff; border-left:5px solid #6366f1; border-radius:12px; padding:12px 16px;
#        box-shadow:0 1px 3px rgba(15,23,42,.06), 0 6px 18px rgba(15,23,42,.05); }
# .kpi-lbl { font-size:11px; text-transform:uppercase; letter-spacing:.07em; color:var(--muted); font-weight:600; }
# .kpi-val { font-size:27px; font-weight:700; line-height:1.2; }
# .kpi-sub { font-size:12.5px; color:var(--muted); }
# .profile { line-height:1.8; font-size:14px; overflow-wrap:anywhere; }
# .profile-id { font-size:17px; font-weight:700; color:#312e81; }
# .pill { display:inline-block; padding:2px 11px; border-radius:999px; color:#fff; font-size:12px; font-weight:600; }
# .insights { columns:2; column-gap:44px; padding-left:18px; margin:4px 0; }
# .insights li { margin-bottom:10px; font-size:14px; line-height:1.5; break-inside:avoid; }
# @media (max-width: 900px) { .insights { columns:1; } }
# </style>
# """
# st.markdown(CSS, unsafe_allow_html=True)


# def kpi_card(label, value, sub="", colour="#6366f1") -> str:
#     return (f'<div class="kpi" style="border-left-color:{colour};"><div class="kpi-lbl">{label}</div>'
#             f'<div class="kpi-val">{value}</div><div class="kpi-sub">{sub}</div></div>')


# def kpi_strip(cards: list[str]):
#     st.markdown(f'<div class="kpi-strip">{"".join(cards)}</div>', unsafe_allow_html=True)


# def card_head(title: str | None = None, sub: str | None = None):
#     if title:
#         st.markdown(f'<div class="card-title">{title}</div>', unsafe_allow_html=True)
#     if sub:
#         st.markdown(f'<div class="sub">{sub}</div>', unsafe_allow_html=True)


# def pl_style(fig: go.Figure, legend: bool = True, height: int | None = None) -> go.Figure:
#     ax = dict(title_text="", gridcolor="#eef0f4", zeroline=False, automargin=True)
#     fig.update_layout(font=dict(family="Inter, sans-serif", size=12, color="#334155"),
#                       paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", showlegend=legend,
#                       legend=dict(orientation="h", x=0, y=1, yanchor="bottom"),
#                       margin=dict(l=10, r=10, t=40, b=10), hoverlabel=dict(font_family="Inter, sans-serif"))
#     fig.update_xaxes(**ax)
#     fig.update_yaxes(**ax)
#     if height:
#         fig.update_layout(height=height)
#     return fig


# def show(fig: go.Figure):
#     st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


# def table(df: pd.DataFrame, **kw):
#     st.dataframe(df, hide_index=True, **kw)


# def stacked_status_bars(d: pd.DataFrame, cat: str, share: str = "share", horizontal: bool = True,
#                         count: str | None = None, hover_extra: str = "") -> go.Figure:
#     """one trace per status; d has columns cat, status, share (and optionally a count column)"""
#     fig = go.Figure()
#     for s in STATUS_LEVELS:
#         sub = d[d["status"] == s]
#         if sub.empty:
#             continue
#         kw = dict(name=s, marker_color=STATUS_COLS[s])
#         if horizontal:
#             kw.update(x=sub[share], y=sub[cat], orientation="h",
#                       text=[f"{v:.0%}" if v >= 0.07 else "" for v in sub[share]],
#                       textposition="inside", insidetextanchor="middle", textfont=dict(color="white", size=12))
#             if count:
#                 kw.update(customdata=sub[count], hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:.1%} "
#                                                                "(%{customdata:,} learners)<extra></extra>")
#             else:
#                 kw.update(hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:.1%}<extra></extra>")
#         fig.add_trace(go.Bar(**kw))
#     return fig


# # ---------------------------------------------------------------- sidebar: global filters
# def _reset():
#     st.session_state.update(f_course="all", f_area="All", stall_days=DEFAULT_STALL)


# st.session_state.setdefault("f_course", "all")
# st.session_state.setdefault("f_area", "All")
# st.session_state.setdefault("stall_days", DEFAULT_STALL)

# with st.sidebar:
#     st.markdown("### Filters")
#     course_labels = {"all": "All courses", **course_only}
#     f_course = st.selectbox("Course", list(course_labels), format_func=course_labels.get, key="f_course")
#     f_area = st.selectbox("Learning area", area_choices, key="f_area")
#     stall_days = st.slider("Stalled after (days idle)", 7, 90, key="stall_days")
#     st.button("Reset", on_click=_reset)
#     scope_slot = st.empty()
#     st.caption("Stalled = unfinished and idle for the chosen days (default 30, the glossary's Inactive rule). "
#                "The Course filter applies to Overview, Enrolment & Reach, Segments and Follow-up; "
#                "Funnel and Learner pages have their own course selector.")

# # ---------------- shared data (pf = learning-area filter, p = + course filter)
# pf = prog_base if f_area == "All" else prog_base[prog_base["learning_area"] == f_area]
# pf = add_status(pf, stall_days)
# p = pf if f_course == "all" else pf[pf["course"] == f_course]
# scope_slot.markdown(f'<div class="scope"><div class="scope-n">{comma(len(p))}</div>'
#                     f'<div class="scope-l">enrolments in view</div></div>', unsafe_allow_html=True)

# st.markdown('<div class="topbar"><h1>Student Progress Dashboard</h1>'
#             '<span>Learner status, stage funnels, segments and follow-up lists</span></div>', unsafe_allow_html=True)

# tab_over, tab_reach, tab_fun, tab_seg, tab_fu, tab_trk = st.tabs(
#     ["Overview", "Enrolment & Reach", "Course Funnel", "Segments", "Follow-up List", "Learner Tracker"])

# # ================================================================ Overview
# with tab_over:
#     d, n = p, len(p)
#     if n == 0:
#         st.info("No learners match the current filters.")
#     else:
#         s = lambda x: int((d["status"] == x).sum())
#         pc = lambda x: pct(s(x) / n)
#         kpi_strip([
#             kpi_card("Enrolments", comma(n), f"avg. {r1(d['pct_complete'].mean())}% of activities done", "#0ea5e9"),
#             kpi_card("Completed", pc("Completed"), f"{comma(s('Completed'))} finished every activity", STATUS_COLS["Completed"]),
#             kpi_card("Active", pc("Active"), f"{comma(s('Active'))} still moving", STATUS_COLS["Active"]),
#             kpi_card("Stalled", pc("Stalled"), f"{comma(s('Stalled'))} idle {stall_days}+ days", STATUS_COLS["Stalled"]),
#             kpi_card("Not started", pc("Not started"), f"{comma(s('Not started'))} enrolled, nothing done", STATUS_COLS["Not started"]),
#         ])
#         c1, c2 = st.columns([7, 5])
#         with c1, st.container(border=True):
#             card_head("Where each course stands",
#                       "Share of enrolled learners by status (learning-area filter applies). Click a legend item to hide a status.")
#             g = pf.groupby(["course_name", "status"], observed=True).size().reset_index(name="n")
#             g["share"] = g["n"] / g.groupby("course_name")["n"].transform("sum")
#             order = (g[g["status"] == "Completed"].set_index("course_name")["share"]
#                      .reindex(g["course_name"].unique()).fillna(0).sort_values().index.tolist())
#             fig = stacked_status_bars(g, "course_name", count="n")
#             pl_style(fig, height=400).update_layout(barmode="stack")
#             fig.update_xaxes(tickformat=".0%", range=[0, 1])
#             fig.update_yaxes(categoryorder="array", categoryarray=order)
#             show(fig)
#         with c2, st.container(border=True):
#             card_head("How far learners have got",
#                       "Learners by share of activities completed; same colours as the chart on the left.")
#             g = d.groupby(["progress_band", "status"], observed=True).size().reset_index(name="n")
#             fig = go.Figure([go.Bar(x=g[g.status == s_]["progress_band"], y=g[g.status == s_]["n"], name=s_,
#                                     marker_color=STATUS_COLS[s_],
#                                     hovertemplate="%{x}<br>%{fullData.name}: %{y:,} learners<extra></extra>")
#                              for s_ in STATUS_LEVELS])
#             pl_style(fig, legend=False, height=400).update_layout(barmode="stack")
#             fig.update_xaxes(categoryorder="array", categoryarray=BAND_LEVELS)
#             fig.update_yaxes(tickformat=",")
#             show(fig)

#         with st.container(border=True):
#             card_head("What stands out")
#             items = []
#             near = int(((d["status"] == "Stalled") & (d["pct_complete"] >= 75)).sum())
#             if near > 0:
#                 items.append(f"<b>{comma(near)} stalled learners are 75%+ through</b> their course: the quickest wins for a reminder.")
#             failed = int((d["failed_assessments"] > 0).sum())
#             if failed > 0:
#                 items.append(f"<b>{comma(failed)} learners</b> finished an assessment without reaching the pass grade.")
#             stuck = int(((d["completed"] == 0) & (d["pct_complete"] >= 95)).sum())
#             if stuck > 0:
#                 items.append(f"<b>{comma(stuck)} learners sit at 95%+ without finishing</b>: worth checking for a "
#                              "course-completion trigger or access problem (flagged in the Town Hall).")
#             if f_course == "all":
#                 w = d.groupby("course_name", observed=True).agg(s=("status", lambda x: (x == "Stalled").mean()),
#                                                                 c=("status", lambda x: (x == "Completed").mean()))
#                 items.append(f"Highest completion: <b>{w['c'].idxmax()}</b> ({pct(w['c'].max())}). "
#                              f"Highest stalled share: <b>{w['s'].idxmax()}</b> ({pct(w['s'].max())}).")
#             else:
#                 stc = (d[d["status"] == "Stalled"].groupby("furthest_stage").size()
#                        .sort_values(ascending=False, kind="stable"))
#                 if len(stc):
#                     items.append(f"Most stalled learners stopped at <b>{stc.index[0]}</b> ({pct(stc.iloc[0] / stc.sum(), 1)} of the stalled).")
#             gap = d.loc[d["items_done"] > 0, "joined_to_first_days"].median()
#             if np.isfinite(gap):
#                 items.append(f"Median wait between registering and the first completed activity: <b>{r1(gap)} days</b>.")
#             dc = ((d["last_activity"] - d["date_joined"]).dt.total_seconds() / 86400)[d["status"] == "Completed"]
#             if dc.notna().any():
#                 items.append(f"Median time from registering to completing the course: <b>{r1(dc.median())} days</b>.")
#             if items:
#                 st.markdown('<ul class="insights">' + "".join(f"<li>{x}</li>" for x in items) + "</ul>",
#                             unsafe_allow_html=True)
#             else:
#                 st.info("Nothing notable for this selection.")

# # ======================================================= Enrolment & Reach
# with tab_reach:
#     ppl = PARTS.drop_duplicates("unique_key")
#     if f_area != "All":
#         ppl = ppl[ppl["learning_area"] == f_area]
#     if f_course != "all":
#         ppl = ppl[ppl["unique_key"].isin(p["unique_key"])]
#     n = len(ppl)
#     if n == 0:
#         st.info("No participants match the current filters.")
#     else:
#         v = [pct_of(ppl["gender"], lambda x: x.fillna("").str.lower().eq("female")),
#              pct_of(ppl["pwd_status"], is_yes), pct_of(ppl["displaced_youth"], is_yes)]
#         tg = list(TARGETS.values())

#         def tgt(i):
#             if pd.isna(v[i]):
#                 return "not in the export"
#             return f"target {pct(tg[i], 1)}" + (" (met)" if v[i] >= tg[i] else f" ({r1((tg[i] - v[i]) * 100)} pp short)")

#         col = lambda i: STATUS_COLS["Completed"] if (pd.notna(v[i]) and v[i] >= tg[i]) else "#f59e0b"
#         whole = f_course == "all" and f_area == "All"
#         kpi_strip([
#             kpi_card("Registered", comma(n),
#                      f"{pct(n / REACH_TARGET)} of the {comma(REACH_TARGET)} five-year target" if whole
#                      else "participants in the current view", "#0ea5e9"),
#             kpi_card("Female", fmt_pct(v[0]), tgt(0), col(0)),
#             kpi_card("Persons with disability", fmt_pct(v[1]), tgt(1), col(1)),
#             kpi_card("Displaced youth", fmt_pct(v[2]), tgt(2), col(2)),
#             kpi_card("Enrolled in a course", fmt_pct((ppl["courses_enrolled"] > 0).mean()), "of those registered", "#6366f1"),
#         ])
#         c1, c2 = st.columns([7, 5])
#         with c1, st.container(border=True):
#             card_head("Registrations over time",
#                       "Participants by month joined (bars) and cumulative total (line). Follows the Course and Learning-area filters.")
#             m = (ppl[ppl["date_joined"].notna()].assign(month=lambda x: x["date_joined"].dt.strftime("%Y-%m"))
#                  .groupby("month").size().reset_index(name="n").sort_values("month"))
#             m["cum"] = m["n"].cumsum()
#             if m.empty:
#                 st.info("No join dates for this selection.")
#             else:
#                 fig = make_subplots(specs=[[{"secondary_y": True}]])
#                 fig.add_trace(go.Bar(x=m["month"], y=m["n"], name="Joined in month", marker_color="#6366f1",
#                                      hovertemplate="%{x}<br>%{y:,} joined<extra></extra>"), secondary_y=False)
#                 fig.add_trace(go.Scatter(x=m["month"], y=m["cum"], name="Cumulative", mode="lines",
#                                          line=dict(color="#0ea5e9", width=3),
#                                          hovertemplate="%{x}<br>%{y:,} registered to date<extra></extra>"), secondary_y=True)
#                 pl_style(fig, height=380)
#                 fig.update_xaxes(type="category")
#                 fig.update_yaxes(tickformat=",")
#                 fig.update_yaxes(showgrid=False, secondary_y=True)
#                 show(fig)
#         with c2, st.container(border=True):
#             card_head("Inclusion (GESI) vs programme targets",
#                       "Share of registered participants; diamond = programme target.")
#             g = pd.DataFrame({"group": list(TARGETS), "target": tg, "actual": v}).dropna(subset=["actual"])
#             if g.empty:
#                 st.info("These fields are not in the export.")
#             else:
#                 colors = np.where(g["actual"] >= g["target"], "#10b981", "#f59e0b")
#                 fig = go.Figure()
#                 fig.add_trace(go.Bar(x=g["actual"], y=g["group"], orientation="h", name="Actual", marker_color=colors,
#                                      text=[pct(a) for a in g["actual"]], textposition="outside",
#                                      hovertemplate="%{y}: %{x:.1%}<extra></extra>"))
#                 fig.add_trace(go.Scatter(x=g["target"], y=g["group"], mode="markers", name="Programme target",
#                                          marker=dict(symbol="diamond", size=13, color="#0f172a"),
#                                          hovertemplate="Target: %{x:.0%}<extra></extra>"))
#                 pl_style(fig, height=380)
#                 fig.update_xaxes(tickformat=".0%", range=[0, min(1, max(g["actual"].max(), g["target"].max()) * 1.3)])
#                 fig.update_yaxes(categoryorder="array", categoryarray=g["group"].tolist()[::-1])
#                 show(fig)

#         c3, c4 = st.columns(2)
#         with c3, st.container(border=True):
#             card_head("Onboarding lag: registration to first activity",
#                       "Enrolments by days between date joined and first completed activity.")
#             lv = ["Within 1 day", "2-7 days", "8-14 days", "15-30 days", "Over 30 days", "Never started"]
#             j = p["joined_to_first_days"]
#             band = np.select([p["items_done"] == 0, j.isna(), j <= 1, j <= 7, j <= 14, j <= 30],
#                              ["Never started", "", "Within 1 day", "2-7 days", "8-14 days", "15-30 days"],
#                              default="Over 30 days")
#             b = pd.Series(band[band != ""]).value_counts().reindex(lv).dropna().reset_index()
#             b.columns = ["band", "n"]
#             if b.empty:
#                 st.info("No data for this selection.")
#             else:
#                 b["share"] = b["n"] / b["n"].sum()
#                 palette = dict(zip(lv, ["#10b981", "#84cc16", "#f59e0b", "#f97316", "#f43f5e", "#94a3b8"]))
#                 fig = go.Figure(go.Bar(x=b["band"], y=b["n"], marker_color=[palette[x] for x in b["band"]],
#                                        text=[pct(s_, 1) for s_ in b["share"]], textposition="outside",
#                                        hovertemplate="%{x}<br>%{y:,} enrolments (%{text})<extra></extra>"))
#                 pl_style(fig, legend=False, height=340)
#                 fig.update_xaxes(categoryorder="array", categoryarray=lv)
#                 fig.update_yaxes(tickformat=",")
#                 show(fig)
#         with c4, st.container(border=True):
#             card_head("Data-quality checks", "Rules taken from the notes in the GATEWAY glossary.")
#             nan_all = lambda s_: pd.Series(np.nan, index=s_.index)
#             not_yes = lambda s_: nan_all(s_) if s_.isna().all() else ~is_yes(s_)
#             area_mis = (nan_all(ppl["learning_area"]) if ppl["skill_assessment_learning_area"].isna().all() else
#                         ppl["learning_area"].fillna("").str.strip().str.lower()
#                         != ppl["skill_assessment_learning_area"].fillna("").str.strip().str.lower())
#             checks = [
#                 ("Consent not recorded as Yes", not_yes(ppl["accept_consent"]), "No record should be active without consent"),
#                 ("Email not verified", not_yes(ppl["email_verified"]), "Low verification explains communication gaps"),
#                 ("Duplicate email", nan_all(ppl["duplicate_email"]) if ppl["duplicate_email"].isna().all()
#                  else ppl["duplicate_email"] == "Yes", "Same person may be registered twice"),
#                 ("Learning area differs from skill-assessment area", area_mis, "Flag mismatches or track changes"),
#                 ("Age outside 18-35", nan_all(ppl["age_band"]) if ppl["age_band"].isna().all()
#                  else ppl["age_band"].isin(["Under 18", "36+"]), "Programme eligibility range"),
#             ]
#             rows = []
#             for label, f, why in checks:
#                 na_all = f.isna().all()
#                 k = 0 if na_all else int(f.fillna(False).astype(bool).sum())
#                 rows.append({"Check": label, "Records": "n/a" if na_all else comma(k),
#                              "Share": "n/a" if na_all else pct(k / n), "Why it matters": why})
#             table(pd.DataFrame(rows))

# # ============================================================ Course Funnel
# with tab_fun:
#     c1, c2 = st.columns([4, 8])
#     codes = list(course_only)
#     with c1, st.container(border=True):
#         fun_course = st.selectbox("Course to inspect", codes, format_func=course_only.get,
#                                   index=codes.index("ggw") if "ggw" in codes else 0)
#     fc = pf[pf["course"] == fun_course]
#     keys = fc["unique_key"]
#     base = (struct[struct["course"] == fun_course].drop_duplicates(["stage", "stage_order"])
#             [["stage", "stage_order"]].sort_values("stage_order"))
#     stage_levels = base["stage"].tolist()
#     sprog = dat["Stage_Progress"]
#     sprog = sprog[(sprog["course"] == fun_course) & sprog["unique_key"].isin(keys)]
#     _sp = sprog.assign(_span=sprog["span_days"].where(sprog["stage_complete"] == 1))
#     agg = _sp.groupby("stage").agg(started=("stage", "size"), completed=("stage_complete", "sum"),
#                                    median_span_days=("_span", "median")).reset_index()
#     stg = base.merge(agg, on="stage", how="left")
#     stg[["started", "completed"]] = stg[["started", "completed"]].fillna(0).astype(int)
#     stg["enrolled"] = len(keys)
#     stg["pct_started"] = stg["started"] / max(len(keys), 1)
#     stg["pct_completed"] = stg["completed"] / max(len(keys), 1)
#     stg = stg.sort_values("stage_order")

#     with c2:
#         if len(fc) == 0:
#             st.info("No learners for this course and learning area.")
#         else:
#             drops = stg.assign(drop_pp=(stg["pct_started"].shift(1) - stg["pct_started"]) * 100).dropna(subset=["drop_pp"])
#             big = drops.loc[drops["drop_pp"].idxmax()] if len(drops) else None
#             kpi_strip([
#                 kpi_card("Enrolled", comma(len(fc)), "in this course", "#0ea5e9"),
#                 kpi_card("Completed course", pct((fc["status"] == "Completed").mean()),
#                          f"{comma((fc['status'] == 'Completed').sum())} learners", STATUS_COLS["Completed"]),
#                 kpi_card("Never started", pct((fc["status"] == "Not started").mean()),
#                          f"{comma((fc['status'] == 'Not started').sum())} learners", STATUS_COLS["Not started"]),
#                 kpi_card("Biggest drop-off", big["stage"] if big is not None else "n/a",
#                          f"{r1(big['drop_pp'])} pp fewer start it than the stage before" if big is not None else "",
#                          STATUS_COLS["Stalled"]),
#             ])

#     with st.container(border=True):
#         card_head("Stage progression: who got through, and where the rest stopped",
#                   "Left: share of enrolled learners who started / completed each stage (hover for median days to finish). "
#                   "Right: unfinished learners by the furthest stage reached.")
#         stop_d = (fc[(fc["status"] != "Completed") & fc["furthest_stage"].isin(stage_levels)]
#                   .groupby(["furthest_stage", "status"], observed=True).size().reset_index(name="n"))
#         if stg.empty or stop_d.empty:
#             st.info("No unfinished learners for this selection.")
#         else:
#             stg["still_open"] = (stg["pct_started"] - stg["pct_completed"]).clip(lower=0)
#             stg["hover"] = [f"<b>{r.stage}</b><br>Started: {pct(r.pct_started)} of enrolled<br>"
#                             f"Completed: {pct(r.pct_completed)} of enrolled<br>Median time to finish: "
#                             f"{'n/a' if pd.isna(r.median_span_days) else r1(r.median_span_days) + ' days'}"
#                             for r in stg.itertuples()]
#             fig = make_subplots(rows=1, cols=2, shared_yaxes=True, column_widths=[0.62, 0.38], horizontal_spacing=0.025)
#             fig.add_trace(go.Bar(x=stg["pct_completed"], y=stg["stage"], orientation="h", name="Completed stage",
#                                  marker_color="#10b981", hoverinfo="text", hovertext=stg["hover"],
#                                  text=[f"{v:.0%}" if v >= 0.06 else "" for v in stg["pct_completed"]],
#                                  textposition="inside", insidetextanchor="middle", textfont=dict(color="white")), 1, 1)
#             fig.add_trace(go.Bar(x=stg["still_open"], y=stg["stage"], orientation="h", name="Started, not finished",
#                                  marker_color="#a7f3d0", hoverinfo="text", hovertext=stg["hover"]), 1, 1)
#             for s_ in STATUS_LEVELS:
#                 sub = stop_d[stop_d["status"] == s_]
#                 if len(sub):
#                     fig.add_trace(go.Bar(x=sub["n"], y=sub["furthest_stage"], orientation="h", name=s_,
#                                          marker_color=STATUS_COLS[s_],
#                                          hovertemplate="%{y}<br>%{fullData.name}: %{x:,} learners<extra></extra>"), 1, 2)
#             pl_style(fig, height=560).update_layout(barmode="stack", bargap=0.3)
#             fig.update_xaxes(tickformat=".0%", range=[0, 1], title_text="Share of enrolled learners", col=1)
#             fig.update_xaxes(tickformat=",", title_text="Learners who stopped here", col=2)
#             fig.update_yaxes(categoryorder="array", categoryarray=stage_levels[::-1])
#             show(fig)

#     with st.container(border=True):
#         card_head("Activity-level drop-off",
#                   "All learners in the course (learning-area filter does not apply). Sorted by the biggest fall versus the previous activity.")
#         it = dat["Item_Funnel"]
#         it = it[it["course"] == fun_course].rename(columns={
#             "item_order": "#", "stage": "Stage", "item_label": "Activity", "pct_completed": "% completed",
#             "drop_vs_prev_item_pp": "Drop vs previous (pp)", "failed_pass_grade": "Did not pass"})
#         it = (it[["#", "Stage", "Activity", "% completed", "Drop vs previous (pp)", "Did not pass"]]
#               .sort_values("Drop vs previous (pp)", ascending=False, na_position="last"))

#         def _drop_colour(x):
#             if pd.isna(x):
#                 return ""
#             return "color:#e11d48;font-weight:bold" if x > 10 else ("color:#d97706;font-weight:bold" if x > 5
#                                                                     else "color:#64748b;font-weight:bold")
#         sty = it.style.map(_drop_colour, subset=["Drop vs previous (pp)"])
#         table(sty, column_config={"% completed": st.column_config.ProgressColumn(
#             "% completed", min_value=0, max_value=100, format="%.1f%%")})

# # ================================================================ Segments
# with tab_seg:
#     DIMS = {"Learning area": "learning_area", "Gender": "gender", "Age band": "age_band", "State": "state",
#             "Internet access": "internet_access", "Has a computer": "has_computer",
#             "Has power supply": "has_power_supply", "Pathway": "pathway", "Disability status": "pwd_status",
#             "Displaced youth": "displaced_youth", "Has children": "has_children",
#             "Skill assessment status": "skill_assessment_status",
#             "Computer literacy status": "computer_literacy_status", "Heard about us from": "heard_from"}
#     with st.container(border=True):
#         c1, c2, c3 = st.columns(3)
#         seg_dim = DIMS[c1.selectbox("Break down by", list(DIMS))]
#         seg_min = c2.slider("Minimum group size", 10, 500, 50, step=10)
#         seg_sort = {"Completion rate": "Completed", "Stalled share": "Stalled", "Group size": "n"}[
#             c3.selectbox("Sort groups by", ["Completion rate", "Stalled share", "Group size"])]

#     g = p.assign(grp=p[seg_dim].astype("object")).dropna(subset=["grp"]).astype({"grp": str})
#     seg = (g.groupby("grp").agg(n=("grp", "size"),
#                                 Completed=("status", lambda x: (x == "Completed").mean()),
#                                 Active=("status", lambda x: (x == "Active").mean()),
#                                 Stalled=("status", lambda x: (x == "Stalled").mean()),
#                                 **{"Not started": ("status", lambda x: (x == "Not started").mean())},
#                                 avg_pct=("pct_complete", "mean")).reset_index())
#     seg = seg[seg["n"] >= seg_min].sort_values("n", ascending=False).head(15)

#     with st.container(border=True):
#         t_mix, t_num = st.tabs(["Status mix", "Numbers"])
#         with t_mix:
#             st.markdown('<div class="sub" style="margin-top:8px;">The 15 largest groups that meet the minimum size; '
#                         'group size in brackets.</div>', unsafe_allow_html=True)
#             if seg.empty:
#                 st.info("No group reaches the minimum size.")
#             else:
#                 seg["label"] = [f"{a}  (n={comma(b)})" for a, b in zip(seg["grp"], seg["n"])]
#                 order = seg.sort_values(seg_sort)["label"].tolist()
#                 long = seg.melt(id_vars=["label"], value_vars=STATUS_LEVELS, var_name="status", value_name="share")
#                 fig = stacked_status_bars(long, "label")
#                 pl_style(fig, height=520).update_layout(barmode="stack")
#                 fig.update_xaxes(tickformat=".0%", range=[0, 1])
#                 fig.update_yaxes(categoryorder="array", categoryarray=order)
#                 show(fig)
#         with t_num:
#             if seg.empty:
#                 st.info("No group reaches the minimum size.")
#             else:
#                 num = pd.DataFrame({"Group": seg["grp"], "Participants": seg["n"],
#                                     "Completed %": (seg["Completed"] * 100).round(1),
#                                     "Active %": (seg["Active"] * 100).round(1),
#                                     "Stalled %": (seg["Stalled"] * 100).round(1),
#                                     "Not started %": (seg["Not started"] * 100).round(1),
#                                     "Avg % complete": seg["avg_pct"].round(1)})
#                 table(num, column_config={
#                     "Completed %": st.column_config.ProgressColumn("Completed %", min_value=0, max_value=100, format="%.1f"),
#                     "Stalled %": st.column_config.ProgressColumn("Stalled %", min_value=0, max_value=100, format="%.1f")})

# # ============================================================== Follow-up
# FU_TYPES = ["Stalled: near completion (75%+)", "Stalled: mid-way (25-74%)", "Stalled: early (under 25%)",
#             "All stalled", "Stuck at 95%+ (unfinished)", "Not started", "Did not pass an assessment"]
# with tab_fu:
#     with st.container(border=True):
#         c1, c2, c3 = st.columns([5, 3, 4])
#         fu_type = c1.selectbox("Who needs a nudge?", FU_TYPES)
#         stalled = p["status"] == "Stalled"
#         mask = {
#             FU_TYPES[0]: stalled & (p["pct_complete"] >= 75),
#             FU_TYPES[1]: stalled & (p["pct_complete"] >= 25) & (p["pct_complete"] < 75),
#             FU_TYPES[2]: stalled & (p["pct_complete"] < 25),
#             FU_TYPES[3]: stalled,
#             FU_TYPES[4]: (p["completed"] == 0) & (p["pct_complete"] >= 95),
#             FU_TYPES[5]: p["status"] == "Not started",
#             FU_TYPES[6]: p["failed_assessments"] > 0}[fu_type]
#         fu = (p[mask].sort_values(["pct_complete", "days_since_last_activity"], ascending=[False, True])
#               .assign(course=lambda x: x["course_name"], last_stage_reached=lambda x: x["furthest_stage"],
#                       status=lambda x: x["status"].astype(str))
#               [["registration_id", "name", "email", "phone", "course", "learning_area", "state", "gender",
#                 "pct_complete", "items_remaining", "stages_remaining", "last_stage_reached",
#                 "days_since_last_activity", "days_since_joined", "failed_assessments", "status"]])
#         c2.markdown(f'<div style="margin-top:26px;"><span class="pill" style="background:#4338ca;font-size:14px;'
#                     f'padding:6px 16px;">{comma(len(fu))} learners</span></div>', unsafe_allow_html=True)
#         c3.download_button("Download list (CSV)", drop_empty_pii(fu).to_csv(index=False, na_rep="").encode("utf-8"),
#                            file_name=f"follow_up_{date.today()}.csv", mime="text/csv")

#     c1, c2 = st.columns([8, 4])
#     with c1, st.container(border=True):
#         card_head("Learners in this list",
#                   "Ordered by progress, so the learners closest to finishing come first. The CSV has a few extra columns. "
#                   "Contains personal data, so share with care.")
#         q = st.text_input("Search registration ID, name or email", key="fu_search").strip().lower()
#         view = fu
#         if q:
#             hay = view[["registration_id", "name", "email"]].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
#             view = view[hay.str.contains(q, regex=False)]
#         tbl = drop_empty_pii(pd.DataFrame({
#             "Registration ID": view["registration_id"], "Name": view["name"], "Email": view["email"],
#             "Phone": view["phone"], "Course": view["course"], "Learning area": view["learning_area"],
#             "State": view["state"], "% complete": view["pct_complete"], "Activities left": view["items_remaining"],
#             "Last stage reached": view["last_stage_reached"],
#             "Days idle": view["days_since_last_activity"].round(0), "Failed assessments": view["failed_assessments"]}))
#         table(tbl, column_config={"% complete": st.column_config.ProgressColumn(
#             "% complete", min_value=0, max_value=100, format="%.1f")})
#     with c2, st.container(border=True):
#         card_head("Where they stopped", "Last stage reached by the learners in the list (top 12).")
#         d = (fu.groupby(["course", "last_stage_reached"]).size().reset_index(name="n")
#              .sort_values("n", ascending=False, kind="stable").head(12))
#         if d.empty:
#             st.info("Nobody in this list.")
#         else:
#             d["label"] = (d["last_stage_reached"] + " \u00b7 " + d["course"]) if f_course == "all" else d["last_stage_reached"]
#             fig = go.Figure(go.Bar(x=d["n"], y=d["label"], orientation="h", marker_color="#6366f1",
#                                    hovertemplate="%{y}<br>%{x:,} learners<extra></extra>"))
#             pl_style(fig, legend=False, height=520)
#             fig.update_xaxes(tickformat=",")
#             fig.update_yaxes(categoryorder="array", categoryarray=d["label"].tolist()[::-1])
#             show(fig)

# # ========================================================== Learner Tracker
# with tab_trk:
#     with st.container(border=True):
#         c1, c2, c3 = st.columns([5, 4, 3])
#         search = c1.text_input("Find a learner (registration ID, name or email)", key="trk_search").strip().lower()
#         pool = rid_lookup if not search else rid_lookup[rid_lookup["label"].str.lower().str.contains(search, regex=False)]
#         pool = pool.head(500)                                   # keep the dropdown light for big cohorts
#         if pool.empty:
#             c1.warning("No learner matches that search.")
#             rid = None
#         else:
#             lab = dict(zip(pool["registration_id"], pool["label"]))
#             rid = c1.selectbox("Learner", list(lab), format_func=lab.get)
#             if len(pool) == 500:
#                 c1.caption("Showing the first 500 matches: type more of the ID, name or email to narrow down.")
#         lrn = add_status(prog_base[prog_base["registration_id"] == rid], stall_days) if rid else prog_base.iloc[0:0]
#         if len(lrn):
#             lrn = lrn.sort_values("items_done", ascending=False)
#             cl = dict(zip(lrn["course"], lrn["course_name"]))
#             trk_course = c2.selectbox("Course", list(cl), format_func=cl.get)
#         else:
#             trk_course = None
#         act_show = c3.radio("Activity list", ["To do", "Done", "All"], horizontal=True)

#     if trk_course is not None:
#         r = lrn[lrn["course"] == trk_course].iloc[0]
#         pr = PARTS[PARTS["unique_key"] == r["unique_key"]].iloc[0]
#         na_txt = lambda x: "n/a" if pd.isna(x) else str(x)
#         stt = str(r["status"])

#         # ---- profile + KPIs
#         c1, c2 = st.columns([4, 8])
#         with c1, st.container(border=True):
#             age_txt = "age n/a" if pd.isna(pr["age"]) else f"{int(pr['age'])} yrs"
#             joined_txt = "n/a" if pd.isna(pr["date_joined"]) else f"{pr['date_joined']:%d %b %Y}"
#             phone = f'<div>Phone: {pr["phone"]}</div>' if pd.notna(pr["phone"]) else ""
#             st.markdown(
#                 f'<div class="profile"><div><span class="profile-id">{r["registration_id"]}</span> '
#                 f'<span class="pill" style="background:{STATUS_COLS[stt]};">{stt}</span></div>'
#                 f'<div><b>{na_txt(pr["name"])}</b> \u00b7 {na_txt(pr["email"])}</div>{phone}'
#                 f'<div>{na_txt(pr["gender"])} \u00b7 {age_txt} \u00b7 {na_txt(pr["state"])}</div>'
#                 f'<div>Learning area: <b>{na_txt(pr["learning_area"])}</b></div>'
#                 f'<div>Joined {joined_txt} \u00b7 enrolled in <b>{len(lrn)}</b> course(s)</div></div>',
#                 unsafe_allow_html=True)
#         with c2:
#             span_txt = "n/a" if pd.isna(r["total_span_days"]) else f"{r1(r['total_span_days'])} d"
#             idle_txt = "none" if pd.isna(r["days_since_last_activity"]) else f"{r1(r['days_since_last_activity'])} d ago"
#             first_txt = "n/a" if pd.isna(r["joined_to_first_days"]) else f"{r1(r['joined_to_first_days'])} d"
#             kpi_strip([
#                 kpi_card("Course progress", f"{r1(r['pct_complete'])}%",
#                          f"{r['items_done']} of {r['items_total']} activities \u00b7 {r['items_remaining']} to go",
#                          STATUS_COLS["Completed"]),
#                 kpi_card("Stages completed", f"{r['stages_completed']} / {r['stages_total']}",
#                          f"Furthest: {r['furthest_stage']}", STATUS_COLS["Active"]),
#                 kpi_card("Last activity", idle_txt, f"Joined to first activity: {first_txt}",
#                          STATUS_COLS["Stalled"] if stt == "Stalled" else "#f59e0b"),
#                 kpi_card("Time in course", span_txt, f"Active on {r['active_days']} day(s)", "#06b6d4"),
#             ])

#         # ---- stage progress + checklist
#         sc = (struct[struct["course"] == r["course"]].groupby(["stage", "stage_order"]).size()
#               .reset_index(name="items_total"))
#         sp = dat["Stage_Progress"]
#         sp = sp[(sp["course"] == r["course"]) & (sp["unique_key"] == r["unique_key"])][
#             ["stage", "items_done", "first_completed", "last_completed", "span_days"]]
#         ls = sc.merge(sp, on="stage", how="left").sort_values("stage_order")
#         ls["items_done"] = ls["items_done"].fillna(0)
#         ls["pct"] = ls["items_done"] / ls["items_total"]
#         ls["stage_status"] = np.select([ls["items_done"] >= ls["items_total"], ls["items_done"] > 0],
#                                        ["Complete", "In progress"], default="Not started")

#         c1, c2 = st.columns([5, 7])
#         with c1, st.container(border=True):
#             card_head("Progress through the stages", "Hover a bar for dates and days spent in that stage.")
#             f_d = lambda x: "-" if pd.isna(x) else f"{x:%d %b %Y}"
#             ls["hover"] = [f"<b>{x.stage}</b><br>{int(x.items_done)} of {x.items_total} activities done ({pct(x.pct, 1)})"
#                            f"<br>Time in stage: {'n/a' if pd.isna(x.span_days) else r1(x.span_days) + ' days'}"
#                            f"<br>First: {f_d(x.first_completed)} \u00b7 Last: {f_d(x.last_completed)}"
#                            for x in ls.itertuples()]
#             cols = {"Complete": "#10b981", "In progress": "#f59e0b", "Not started": "#cbd5e1"}
#             fig = go.Figure()
#             for s_, colr in cols.items():
#                 sub = ls[ls["stage_status"] == s_]
#                 if len(sub):
#                     fig.add_trace(go.Bar(x=sub["pct"], y=sub["stage"], orientation="h", name=s_, marker_color=colr,
#                                          text=[f"{int(a)}/{b}" for a, b in zip(sub["items_done"], sub["items_total"])],
#                                          textposition="outside", hoverinfo="text", hovertext=sub["hover"]))
#             pl_style(fig, height=520).update_layout(barmode="overlay", bargap=0.3)
#             fig.update_xaxes(tickvals=[0, .25, .5, .75, 1], tickformat=".0%", range=[0, 1.18])
#             fig.update_yaxes(categoryorder="array", categoryarray=ls["stage"].tolist()[::-1])
#             show(fig)
#         with c2, st.container(border=True):
#             card_head("Activity checklist",
#                       "Use the radio buttons above to switch between what is still to do and what is done.")
#             ids = ([] if pd.isna(r["done_items"]) or r["done_items"] == ""
#                    else [int(i) for i in str(r["done_items"]).split(",")])
#             items = struct[struct["course"] == r["course"]].sort_values("item_order").copy()
#             items["Status"] = np.where(items["item_order"].isin(ids), "Done", "To do")
#             if act_show != "All":
#                 items = items[items["Status"] == act_show]
#             out = items.rename(columns={"item_order": "#", "stage": "Stage", "item_label": "Activity"})[
#                 ["#", "Stage", "Activity", "Status"]]
#             sty = out.style.map(lambda x: "color:#059669;font-weight:bold" if x == "Done"
#                                 else ("color:#d97706;font-weight:bold" if x == "To do" else ""), subset=["Status"])
#             table(sty)


# cd "/Users/moseskioko/Documents/files/CcHUB Workflow/R Codes/Gateway/python_dashboard"
# source "/Users/moseskioko/Documents/files/CcHUB Workflow/Python Codes/vscode trials/vscode_env/bin/activate"
# streamlit run gateway_clean.py