"""
Dental Center Production Dashboard (single-file Streamlit app)

Run:  streamlit run app.py
Data: an insurance-claims / production export (.xlsx/.xls). Patient-identifying
columns are dropped or one-way hashed immediately after load and are never displayed.
"""
import hashlib
import io
import json
import re
import secrets
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Dental Production Dashboard", page_icon="🦷", layout="wide")

DEFAULT_PATH = r"C:\Users\christian Mar\Downloads\Production.xlsx"
CHAIR_MAP_FILE = Path(__file__).with_name("chair_map.json")
REDACTED = "[PATIENT_ID_REDACTED]"

# Palette (validated reference palette; single-hue sequential blue)
BLUE = "#2a78d6"
SEQ_BLUE = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]
MUTED = "#898781"
GRID = "rgba(137,135,129,0.18)"
FONT = 'system-ui, -apple-system, "Segoe UI", sans-serif'

# Columns that identify patients or policies: dropped right after load.
PHI_COLUMNS = [
    "Patient Name", "Member ID", "ID Payer", "Claim ID", "Prior Approval Auth ID",
    "Authorization ID", "Payment Reference", "Ordering Clinician", "Broker ID",
    "Date of Birth", "DOB", "Policy ID", "Policy Number", "Insurance ID",
]

# Candidate source names for each logical field (first match wins).
FIELDS = {
    "date": ["Date", "Appointment Date", "Transaction Date", "Service Date"],
    "start": ["Start Time"],
    "end": ["End Time"],
    "provider": ["Completed Doctor", "Claim Clinician", "Provider", "Clinician", "Doctor"],
    "operatory": ["Operatory", "Chair", "Room", "Op"],
    "code": ["Activity Code", "Procedure Code", "USCLS Code", "ADA Code", "Code"],
    "fee": ["Activity Price", "Fee", "Standard Fee"],
    "gross": ["Gross", "Gross Fees", "Gross Production"],
    "patient_share": ["Patient Share"],
    "paid": ["Payment Amount", "Paid"],
    "rejected": ["Rejected Amount", "Rejection", "Rejected"],
    "settled": ["Date Settelment", "Date Settlement", "RA Received Date"],
    "payer": ["Insurance Company", "Payer"],
    "status": ["Status"],
    "denial": ["Denial Reason"],
    "chart": ["Chart #", "Chart Number", "Patient ID"],
    "visit": ["Encounter ID", "Visit ID", "Appointment ID"],
}
REQUIRED = ["date", "provider", "code", "gross"]

# USCLS (Canada, 5-digit) and ADA CDT (US, Dxxxx) share the same leading-digit categories.
CATEGORIES = {
    "0": "Diagnostic", "1": "Preventive", "2": "Restorative", "3": "Endodontics",
    "4": "Periodontics", "5": "Prosthodontics (removable)", "6": "Prosthodontics (fixed)",
    "7": "Oral Surgery", "8": "Orthodontics", "9": "Adjunctive",
}
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


# --------------------------------------------------------------------------- privacy
def mask(text: str, known: list[str] | None = None) -> str:
    """Scrub patient identifiers out of any text that may reach the screen."""
    text = str(text)
    for term in sorted(known or [], key=len, reverse=True):
        if term and len(term) > 2:
            text = text.replace(term, REDACTED)
    return re.sub(r"\d{5,}", REDACTED, text)


def hash_key(value, salt: str) -> str:
    return hashlib.sha256(f"{salt}|{value}".encode()).hexdigest()[:12]


# --------------------------------------------------------------------------- ingest
def _find_header_row(raw: pd.DataFrame) -> int:
    for i in range(min(25, len(raw))):
        cells = {str(v).strip() for v in raw.iloc[i].dropna()}
        if any(n in cells for n in FIELDS["gross"]) and any(n in cells for n in FIELDS["date"]):
            return i
    return 0


def _pick(df: pd.DataFrame, field: str):
    for name in FIELDS[field]:
        if name in df.columns:
            return name
    return None


def _tidy_provider(name: str) -> str:
    name = re.sub(r"\s+", " ", str(name)).strip()
    name = re.sub(r"^Dr\.?\s*", "Dr. ", name)
    return " ".join(w.capitalize() if w.isupper() and len(w) > 2 else w for w in name.split())


def _category(code: str) -> str:
    c = code.upper().lstrip("D") if code.upper().startswith("D") else code
    return CATEGORIES.get(c[:1], "Other") if c[:1].isdigit() else "Other"


@st.cache_data(show_spinner="Loading and de-identifying data…")
def load_clean(data: bytes, salt: str) -> pd.DataFrame:
    patient_terms: list[str] = []
    try:
        raw = pd.read_excel(io.BytesIO(data), header=None, nrows=30)
        df = pd.read_excel(io.BytesIO(data), header=_find_header_row(raw))
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all").dropna(axis=1, how="all")

        missing = [f for f in REQUIRED if _pick(df, f) is None]
        if missing:
            raise ValueError(f"Required fields not found: {', '.join(missing)}. "
                             f"Columns present: {', '.join(df.columns)}")

        # De-identify first: hash patient & visit keys, drop every PHI column.
        if "Patient Name" in df.columns:
            patient_terms = df["Patient Name"].dropna().astype(str).unique().tolist()
        chart_col, visit_col = _pick(df, "chart"), _pick(df, "visit")
        out = pd.DataFrame(index=df.index)
        out["patient"] = df[chart_col].map(lambda v: hash_key(v, salt)) if chart_col else "n/a"
        out["visit"] = (df[visit_col].map(lambda v: hash_key(v, salt)) if visit_col
                        else out["patient"] + df[_pick(df, "date")].astype(str))
        if "Patient Age" in df.columns:
            out["age_band"] = pd.cut(df["Patient Age"], [-1, 12, 17, 39, 64, 200],
                                     labels=["0-12", "13-17", "18-39", "40-64", "65+"])
        df = df.drop(columns=[c for c in PHI_COLUMNS + [chart_col, visit_col] if c and c in df.columns])

        col = {f: _pick(df, f) for f in FIELDS}
        out["date"] = pd.to_datetime(df[col["date"]], dayfirst=True, errors="coerce").dt.normalize()
        for f in ("start", "end"):
            if col[f]:
                out[f] = pd.to_datetime(out["date"].dt.strftime("%Y-%m-%d") + " " + df[col[f]].astype(str),
                                        errors="coerce")
        out["provider"] = df[col["provider"]].fillna("Unknown Provider").map(_tidy_provider)
        out["provider"] = out["provider"].replace({"Dr. Nan": "Unknown Provider", "": "Unknown Provider"})
        out["operatory"] = df[col["operatory"]].astype(str).str.strip() if col["operatory"] else np.nan

        code = df[col["code"]].astype(str).str.strip()
        out["code"] = code.str.extract(r"^(D?\d+)", expand=False).fillna(code)
        out["code"] = out["code"].where(out["code"].str.startswith("D"), out["code"].str.zfill(5))
        out["category"] = out["code"].map(_category)

        num = lambda f: pd.to_numeric(df[col[f]], errors="coerce").fillna(0.0) if col[f] else 0.0
        out["billed"] = num("gross")
        out["fee"] = num("fee") if col["fee"] else out["billed"]
        out["adjustment"] = out["fee"] - out["billed"]  # discounts / write-offs off standard fee
        out["patient_share"] = num("patient_share")
        out["paid"] = num("paid")
        rejected = num("rejected")
        settled = df[col["settled"]].notna() if col["settled"] else pd.Series(True, index=df.index)
        out["denied"] = np.where(settled, rejected, 0.0)    # adjudicated & not paid
        out["pending"] = np.where(settled, 0.0, rejected)   # still awaiting remittance
        out["net"] = out["fee"] - out["adjustment"] - out["denied"]
        out["payer"] = df[col["payer"]].fillna("Unknown") if col["payer"] else "Unknown"
        out["status"] = df[col["status"]].fillna("Unknown") if col["status"] else "Unknown"
        out["denial_reason"] = df[col["denial"]] if col["denial"] else np.nan
        return out.dropna(subset=["date"]).reset_index(drop=True)
    except Exception as exc:  # never let raw values leak through an error message
        raise RuntimeError(f"{type(exc).__name__}: {mask(exc, patient_terms)}") from None


# --------------------------------------------------------------------------- metrics
def build_visits(df: pd.DataFrame) -> pd.DataFrame:
    agg = {"date": "first", "provider": "first", "chair": "first", "patient": "first", "net": "sum"}
    if "start" in df.columns:
        agg |= {"start": "min", "end": "max"}
    v = df.groupby("visit").agg(agg).reset_index()
    if "start" in v.columns:
        v["minutes"] = ((v["end"] - v["start"]).dt.total_seconds() / 60).clip(lower=0).fillna(0)
    else:
        v["minutes"] = 0.0
    return v


def peak_concurrency(visits: pd.DataFrame) -> int:
    v = visits.dropna(subset=["start", "end"]) if "start" in visits.columns else visits.iloc[0:0]
    if v.empty:
        return 1
    ev = pd.concat([pd.DataFrame({"t": v["start"], "d": 1}), pd.DataFrame({"t": v["end"], "d": -1})])
    ev = ev.sort_values(["t", "d"])
    return int(max(1, ev.groupby(ev["t"].dt.date)["d"].cumsum().max()))


def occupancy_grid(visits: pd.DataFrame) -> pd.DataFrame:
    """Average chairs in use, by weekday × hour (booked minutes / 60 / days of that weekday)."""
    v = visits[(visits["minutes"] > 0) & visits["start"].notna()]
    if v.empty:
        return pd.DataFrame()
    n = v["minutes"].astype(int).to_numpy()
    starts = v["start"].to_numpy().astype("datetime64[m]")
    mins = pd.DatetimeIndex(np.repeat(starts, n) + np.concatenate([np.arange(k) for k in n]).astype("timedelta64[m]"))
    per = pd.Series(1.0, index=mins)
    tot = per.groupby([mins.day_name(), mins.hour]).sum().unstack(fill_value=0) / 60
    days = visits["date"].drop_duplicates().dt.day_name().value_counts()
    grid = tot.div(days, axis=0)
    return grid.reindex([d for d in WEEKDAYS if d in grid.index])


def provider_table(df: pd.DataFrame, visits: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("provider")
    t = pd.DataFrame({
        "Days worked": g["date"].nunique(),
        "Visits": g["visit"].nunique(),
        "Patients": g["patient"].nunique(),
        "Procedures": g.size(),
        "Gross production": g["fee"].sum(),
        "Adjustments": g["adjustment"].sum(),
        "Denied": g["denied"].sum(),
        "Net production": g["net"].sum(),
        "Pending insurance": g["pending"].sum(),
    })
    t["Daily avg (net)"] = t["Net production"] / t["Days worked"]
    hours = visits.groupby("provider")["minutes"].sum() / 60
    t["Chair hours"] = hours.reindex(t.index).fillna(0)
    t["Net / chair hour"] = (t["Net production"] / t["Chair hours"]).replace([np.inf], np.nan)
    t["Denial rate"] = (t["Denied"] / g["billed"].sum()).fillna(0)
    prev = df[df["category"] == "Preventive"].groupby("provider")["net"].sum()
    t["Preventive share"] = (prev.reindex(t.index).fillna(0) / t["Net production"]).fillna(0)
    return t.sort_values("Net production", ascending=False)


def chair_table(df: pd.DataFrame, visits: pd.DataFrame, open_minutes: float) -> pd.DataFrame:
    g = df.groupby("chair")
    t = pd.DataFrame({
        "Providers": g["provider"].nunique(),
        "Days used": g["date"].nunique(),
        "Visits": g["visit"].nunique(),
        "Net production": g["net"].sum(),
    })
    booked = visits.groupby("chair")["minutes"].sum()
    t["Booked hours"] = (booked / 60).reindex(t.index).fillna(0)
    t["Net / booked hour"] = (t["Net production"] / t["Booked hours"]).replace([np.inf], np.nan)
    t["Net / day used"] = t["Net production"] / t["Days used"]
    t["Utilization"] = (booked.reindex(t.index).fillna(0) / (t["Days used"] * open_minutes)).fillna(0)
    t["Share of net"] = t["Net production"] / t["Net production"].sum()
    return t.sort_values("Net production", ascending=False)


# --------------------------------------------------------------------------- charts
def style(fig: go.Figure, height: int = 360) -> go.Figure:
    fig.update_layout(
        height=height, margin=dict(l=4, r=16, t=8, b=4), showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color=MUTED, size=12), barcornerradius=4, bargap=0.3,
        hoverlabel=dict(font=dict(family=FONT, size=12)),
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, linecolor=GRID, ticks="")
    fig.update_yaxes(gridcolor=GRID, zeroline=False, linecolor=GRID, ticks="")
    return fig


def hbar(series: pd.Series, money_fmt: str, label_fmt=None, height=None) -> go.Figure:
    s = series.sort_values()
    fig = go.Figure(go.Bar(
        x=s.values, y=s.index, orientation="h", marker_color=BLUE,
        text=[label_fmt(v) for v in s.values] if label_fmt else None,
        textposition="outside", textfont=dict(color=MUTED), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>" + money_fmt + "<extra></extra>",
    ))
    fig.update_yaxes(showgrid=False)
    return style(fig, height or max(240, 30 * len(s) + 40))


def heatmap(grid: pd.DataFrame, hover: str, height=320) -> go.Figure:
    scale = [[i / (len(SEQ_BLUE) - 1), c] for i, c in enumerate(SEQ_BLUE)]
    fig = go.Figure(go.Heatmap(
        z=grid.values, x=[str(c) for c in grid.columns], y=list(grid.index), colorscale=scale,
        xgap=2, ygap=2, hovertemplate=hover + "<extra></extra>",
        colorbar=dict(thickness=10, outlinewidth=0, tickfont=dict(color=MUTED)),
    ))
    fig.update_yaxes(autorange="reversed", showgrid=False)
    fig.update_xaxes(showgrid=False, type="category")
    return style(fig, height)


# --------------------------------------------------------------------------- app
if "salt" not in st.session_state:
    st.session_state.salt = secrets.token_hex(16)

with st.sidebar:
    st.header("Data source")
    upload = st.file_uploader("Upload production export", type=["xlsx", "xls"])
    path = st.text_input("…or file path", DEFAULT_PATH)
    currency = st.text_input("Currency label", "AED")
    st.caption("Patient names, chart numbers, member/policy IDs and authorizations are "
               "removed or one-way hashed on load and never displayed.")

try:
    data = upload.getvalue() if upload else Path(path).read_bytes()
    df_all = load_clean(data, st.session_state.salt)
except FileNotFoundError:
    st.info("Upload a production export or enter a valid file path in the sidebar.")
    st.stop()
except Exception as exc:
    st.error(f"Could not process the file. {mask(exc)}")
    st.stop()

money = lambda v: f"{currency} {v:,.0f}"
money_short = lambda v: f"{v / 1e6:.2f}M" if abs(v) >= 1e6 else f"{v / 1e3:.1f}k" if abs(v) >= 1e3 else f"{v:.0f}"
HOVER_MONEY = f"{currency} %{{x:,.0f}}"

# ---- chair source: real operatory column if present, else editable provider→chair map
providers_all = sorted(df_all["provider"].unique())
has_operatory = df_all["operatory"].notna().any() and (df_all["operatory"].str.lower() != "nan").any()
if has_operatory:
    df_all["chair"] = df_all["operatory"]
else:
    saved = json.loads(CHAIR_MAP_FILE.read_text()) if CHAIR_MAP_FILE.exists() else {}
    chair_map = {p: saved.get(p, f"Op · {p}") for p in providers_all}
    with st.sidebar.expander("Chair assignment", expanded=not saved):
        st.caption("This export has no operatory column. Assign each provider to the chair they "
                   "work in; providers sharing a chair can use the same name.")
        edited = st.data_editor(
            pd.DataFrame({"Provider": list(chair_map), "Chair": list(chair_map.values())}),
            hide_index=True, disabled=["Provider"], width="stretch", key="chair_editor",
        )
        chair_map = dict(zip(edited["Provider"], edited["Chair"].fillna("").replace("", "Unassigned")))
        if st.button("Save chair map"):
            CHAIR_MAP_FILE.write_text(json.dumps(chair_map, indent=2))
            st.success("Saved.")
    df_all["chair"] = df_all["provider"].map(chair_map)

visits_all = build_visits(df_all)
inferred_chairs = peak_concurrency(visits_all)
open_from = int(visits_all["start"].dt.hour.min()) if "start" in visits_all else 8
open_to = int(min(24, visits_all["end"].dt.hour.max() + 1)) if "end" in visits_all else 18

with st.sidebar:
    st.header("Operations")
    chairs = st.number_input("Operational chairs", 1, 50, inferred_chairs,
                             help=f"Defaults to peak simultaneous appointments in the data ({inferred_chairs}).")
    hours = st.slider("Clinic hours", 0, 24, (open_from, open_to))
open_minutes = max(60, (hours[1] - hours[0]) * 60)

# ---- header + filters (one row)
d0, d1 = df_all["date"].min().date(), df_all["date"].max().date()
st.title("Dental Production Dashboard")
st.caption(f"{d0:%d %b %Y} – {d1:%d %b %Y} · {len(df_all):,} procedures · "
           f"{df_all['visit'].nunique():,} visits · {df_all['patient'].nunique():,} patients")

f1, f2, f3 = st.columns([1.2, 2, 1.4])
rng = f1.date_input("Date range", (d0, d1), min_value=d0, max_value=d1)
sel_prov = f2.multiselect("Providers", providers_all, placeholder="All providers")
sel_payer = f3.multiselect("Payers", sorted(df_all["payer"].unique()), placeholder="All payers")

start_d, end_d = (rng if isinstance(rng, (tuple, list)) and len(rng) == 2 else (d0, d1))
mask_rows = df_all["date"].between(pd.Timestamp(start_d), pd.Timestamp(end_d))
if sel_prov:
    mask_rows &= df_all["provider"].isin(sel_prov)
if sel_payer:
    mask_rows &= df_all["payer"].isin(sel_payer)
df = df_all[mask_rows]
if df.empty:
    st.warning("No records match the current filters.")
    st.stop()
visits = build_visits(df)

# ---- KPIs
gross_prod, adj, denied = df["fee"].sum(), df["adjustment"].sum(), df["denied"].sum()
net_prod, pending = df["net"].sum(), df["pending"].sum()
op_days = df["date"].nunique()
prov_t = provider_table(df, visits)
chair_t = chair_table(df, visits, open_minutes)
prev_net = df.loc[df["category"] == "Preventive", "net"].sum()
rest_net = df.loc[df["category"] == "Restorative", "net"].sum()

k = st.columns(6)
k[0].metric("Gross production", money(gross_prod), help="Σ standard fees (Activity Price).")
k[1].metric("Net production", money(net_prod),
            delta=f"-{money(adj + denied)} adj. & denials", delta_color="off",
            help="Gross − discounts/write-offs − adjudicated insurance denials.")
k[2].metric("Chair yield", money(net_prod / chairs), help=f"Net production ÷ {chairs} operational chairs.")
k[3].metric("Chair yield / day", money(net_prod / chairs / op_days), help=f"Over {op_days} operating days.")
k[4].metric("Top chair", chair_t.index[0], money(chair_t["Net production"].iloc[0]), delta_color="off")
k[5].metric("Preventive : Restorative", f"1 : {rest_net / prev_net:.1f}" if prev_net else "n/a",
            help="Net production ratio, USCLS 1xxxx vs 2xxxx (ADA D1xxx vs D2xxx).")

k2 = st.columns(6)
k2[0].metric("Discounts / write-offs", money(adj))
k2[1].metric("Insurance denied", money(denied))
k2[2].metric("Pending insurance", money(pending), help="Submitted claims not yet remitted.")
k2[3].metric("Insurance paid", money(df["paid"].sum()))
k2[4].metric("Patient share", money(df["patient_share"].sum()))
k2[5].metric("Top provider (daily avg)", prov_t["Daily avg (net)"].idxmax(),
             money(prov_t["Daily avg (net)"].max()), delta_color="off")

tab_chairs, tab_prov, tab_mix, tab_rcm, tab_def = st.tabs(
    ["Chairs", "Providers", "Service mix", "Revenue cycle", "Definitions"])

# ---- Chairs
with tab_chairs:
    if not has_operatory:
        st.caption("Chairs come from the provider → chair assignment in the sidebar "
                   "(the export has no operatory column).")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Net production by chair")
        st.plotly_chart(hbar(chair_t["Net production"], HOVER_MONEY, money_short), width="stretch")
    with c2:
        st.subheader("Net per booked chair hour")
        st.plotly_chart(hbar(chair_t["Net / booked hour"].dropna(), HOVER_MONEY + " / hr", money_short),
                        width="stretch")
    st.subheader("Average chairs in use, by weekday and hour")
    grid = occupancy_grid(visits)
    if not grid.empty:
        st.plotly_chart(heatmap(grid, "%{y} %{x}:00<br>%{z:.1f} chairs in use"), width="stretch")
        st.caption(f"Peak simultaneous appointments in this selection: {peak_concurrency(visits)} "
                   f"· overall utilization {visits['minutes'].sum() / (chairs * op_days * open_minutes):.0%} "
                   f"of {chairs} chairs × {hours[1] - hours[0]} h × {op_days} days")
    st.dataframe(chair_t, width="stretch", column_config={
        "Net production": st.column_config.NumberColumn(format="localized"),
        "Booked hours": st.column_config.NumberColumn(format="%.1f"),
        "Net / booked hour": st.column_config.NumberColumn(format="localized"),
        "Net / day used": st.column_config.NumberColumn(format="localized"),
        "Utilization": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
        "Share of net": st.column_config.NumberColumn(format="percent"),
    })

# ---- Providers
with tab_prov:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Daily provider average (net)")
        st.plotly_chart(hbar(prov_t["Daily avg (net)"], HOVER_MONEY + " / day", money_short), width="stretch")
    with c2:
        st.subheader("Net per chair hour")
        st.plotly_chart(hbar(prov_t["Net / chair hour"].dropna(), HOVER_MONEY + " / hr", money_short),
                        width="stretch")
    daily = df.groupby("date")["net"].sum()
    st.subheader("Daily net production")
    fig = go.Figure(go.Scatter(x=daily.index, y=daily.values, mode="lines+markers",
                               line=dict(color=BLUE, width=2), marker=dict(size=8, color=BLUE),
                               hovertemplate=f"%{{x|%a %d %b}}<br>{currency} %{{y:,.0f}}<extra></extra>"))
    fig.update_layout(hovermode="x")
    st.plotly_chart(style(fig, 300), width="stretch")
    money_cols = ["Gross production", "Adjustments", "Denied", "Net production", "Pending insurance",
                  "Daily avg (net)", "Net / chair hour"]
    st.dataframe(prov_t, width="stretch", column_config={
        **{c: st.column_config.NumberColumn(format="localized") for c in money_cols},
        "Chair hours": st.column_config.NumberColumn(format="%.1f"),
        "Denial rate": st.column_config.NumberColumn(format="percent"),
        "Preventive share": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
    })

# ---- Service mix
with tab_mix:
    cat = df.groupby("category").agg(net=("net", "sum"), procedures=("code", "size"))
    c1, c2 = st.columns([1, 1.3])
    with c1:
        st.subheader("Net production by category")
        st.plotly_chart(hbar(cat["net"], HOVER_MONEY, money_short), width="stretch")
        st.metric("Preventive : Restorative (procedures)",
                  f"1 : {cat['procedures'].get('Restorative', 0) / cat['procedures'].get('Preventive', 1):.2f}")
    with c2:
        st.subheader("Provider mix (share of each provider's net)")
        mix = pd.crosstab(df["provider"], df["category"], values=df["net"], aggfunc="sum").fillna(0)
        mix = mix.div(mix.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
        mix = mix.reindex(prov_t.index)[[c for c in CATEGORIES.values() if c in mix.columns]]
        st.plotly_chart(heatmap(mix * 100, "%{y}<br>%{x}: %{z:.0f}%", height=max(320, 28 * len(mix) + 120)),
                        width="stretch")
    st.subheader("Top procedure codes")
    codes = df.groupby(["code", "category"]).agg(Procedures=("code", "size"), Net=("net", "sum"),
                                                  **{"Avg fee": ("fee", "mean")})
    st.dataframe(codes.sort_values("Net", ascending=False).head(25), width="stretch", column_config={
        "Net": st.column_config.NumberColumn(format="localized"),
        "Avg fee": st.column_config.NumberColumn(format="localized")})

# ---- Revenue cycle
with tab_rcm:
    payer = df.groupby("payer").agg(Billed=("billed", "sum"), Paid=("paid", "sum"), Denied=("denied", "sum"),
                                    Pending=("pending", "sum"), Lines=("code", "size"))
    payer["Denial rate"] = payer["Denied"] / payer["Billed"].replace(0, np.nan)
    c1, c2 = st.columns([1.6, 1])
    with c1:
        st.subheader("By payer")
        st.dataframe(payer.sort_values("Billed", ascending=False), width="stretch", column_config={
            **{c: st.column_config.NumberColumn(format="localized") for c in ["Billed", "Paid", "Denied", "Pending"]},
            "Denial rate": st.column_config.NumberColumn(format="percent")})
    with c2:
        st.subheader("Denial reasons")
        reasons = df.dropna(subset=["denial_reason"]).groupby("denial_reason").agg(
            Lines=("code", "size"), Denied=("denied", "sum")).sort_values("Denied", ascending=False)
        st.dataframe(reasons, width="stretch",
                     column_config={"Denied": st.column_config.NumberColumn(format="localized")})
    st.subheader("Claim status")
    st.dataframe(df.groupby("status")[["billed", "paid", "denied", "pending"]].sum(), width="stretch",
                 column_config={c: st.column_config.NumberColumn(format="localized")
                                for c in ["billed", "paid", "denied", "pending"]})

# ---- Definitions
with tab_def:
    st.markdown(f"""
| Metric | Definition |
|---|---|
| **Gross production** | Σ standard fee per procedure (Activity Price). |
| **Discounts / write-offs** | Standard fee − billed amount (Gross). |
| **Insurance denied** | Rejected amount on claims with a settlement/remittance date. |
| **Pending insurance** | Rejected amount on claims not yet settled — awaiting payer, not a write-off. |
| **Net production** | Gross production − discounts/write-offs − insurance denied. |
| **Daily provider average** | Provider net production ÷ distinct days that provider had procedures. |
| **Chair yield** | Net production ÷ operational chairs ({chairs}; default = peak simultaneous appointments). |
| **Net / chair hour** | Net production ÷ booked appointment hours (end − start time). |
| **Utilization** | Booked minutes ÷ (days used × clinic hours). |
| **Service mix** | First digit of USCLS code (0 Diagnostic … 9 Adjunctive); ADA D-codes map the same way. |
""")
