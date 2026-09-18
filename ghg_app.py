"""Corporate GHG Inventory Dashboard — Scope 1, 2 and 3.

Reads the GHG_Inventory_Scope123.xlsx workbook (FactEmissions + DimEmissionFactor + Targets + DimSite), a CSV in the
same fact layout, or a plain activity-data file (date, site, scope, category, activity type, quantity, unit) which the
app converts to emissions using its factor library. Uploaded files stay in the visitor's session and are never stored.
"""
import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="GHG Inventory — Scope 1, 2, 3", page_icon="🌍", layout="wide")

DATA = Path(__file__).parent / "data"
SCOPE_COLORS = {"Scope 1": "#1E5631", "Scope 2": "#2A9D8F", "Scope 3": "#A8C686"}
ACCENT = "#C9A227"; WARN = "#C0504D"; GREY = "#8A8F98"; INK = "#1B1F2A"
CAT_PALETTE = ["#1E5631", "#2A9D8F", "#A8C686", "#C9A227", "#7A5C99", "#3E7CB1", "#D98E5F", "#5B8C5A", "#8A8F98", "#B56576", "#6A994E", "#386641"]
px.defaults.template = "simple_white"
FACT_COLS = ["Date", "SiteID", "Scope", "CategoryID", "ActivityType", "Quantity", "Unit", "FactorID", "EmissionFactor", "EmissionFactorMarket", "FactorUnit",
             "Method", "DataSource", "DataQualityScore", "tCO2e_Location", "tCO2e_Market", "Note"]
CAT_NAMES = {"1": "Direct emissions", "2": "Purchased energy", "3.1": "Purchased goods and services", "3.2": "Capital goods", "3.3": "Fuel- and energy-related activities",
             "3.4": "Upstream transportation and distribution", "3.5": "Waste generated in operations", "3.6": "Business travel", "3.7": "Employee commuting",
             "3.8": "Upstream leased assets", "3.9": "Downstream transportation and distribution", "3.10": "Processing of sold products", "3.11": "Use of sold products",
             "3.12": "End-of-life treatment of sold products", "3.13": "Downstream leased assets", "3.14": "Franchises", "3.15": "Investments"}
S3_ORDER = [f"3.{i}" for i in range(1, 16)]
DQ_LABEL = {1: "1 · Metered / invoice", 2: "2 · Supplier-specific", 3: "3 · Distance / average data", 4: "4 · Spend-based", 5: "5 · Estimated"}


# --------------------------------------------------------------------------------------- data in
@st.cache_data(show_spinner=False)
def load_sample():
    fact = pd.read_csv(DATA / "sample_ghg_fact.csv.gz", dtype={"CategoryID": str, "SiteID": str}, parse_dates=["Date"])
    factors = pd.read_csv(DATA / "sample_ghg_factors.csv", dtype={"CategoryID": str})
    targets = pd.read_csv(DATA / "sample_ghg_targets.csv")
    sites = pd.read_csv(DATA / "sample_ghg_sites.csv", dtype={"SiteID": str})
    cats = pd.read_csv(DATA / "sample_ghg_categories.csv", dtype={"CategoryID": str})
    return fact, factors, targets, sites, cats


def read_upload(b: bytes, name: str):
    """Returns (fact, factors, targets, sites, cats, kind)."""
    lower = name.lower()
    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        xl = pd.ExcelFile(io.BytesIO(b))
        if "FactEmissions" in xl.sheet_names:
            fact = xl.parse("FactEmissions", dtype={"CategoryID": str, "SiteID": str})
            factors = xl.parse("DimEmissionFactor", dtype={"CategoryID": str}) if "DimEmissionFactor" in xl.sheet_names else None
            targets = xl.parse("Targets").dropna(subset=["Year"]) if "Targets" in xl.sheet_names else None
            if targets is not None: targets = targets[pd.to_numeric(targets.Year, errors="coerce").notna()].copy(); targets["Year"] = targets.Year.astype(int)
            sites = xl.parse("DimSite", dtype={"SiteID": str}) if "DimSite" in xl.sheet_names else None
            cats = xl.parse("DimScopeCategory", dtype={"CategoryID": str}) if "DimScopeCategory" in xl.sheet_names else None
            return fact, factors, targets, sites, cats, "workbook"
        df = xl.parse(xl.sheet_names[0], dtype=str)
    else:
        text = b.decode("utf-8-sig", errors="replace")
        sep = ";" if text[:2000].count(";") > text[:2000].count(",") else ","
        df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str)
    cols = {c.lower().replace(" ", "").replace("_", ""): c for c in df.columns}
    if {"scope", "activitytype", "quantity"} <= set(cols) and ("tco2elocation" in cols or "tco2emarket" in cols):
        return df, None, None, None, None, "fact"
    return df, None, None, None, None, "activity"


def normalise_fact(fact: pd.DataFrame, factors: pd.DataFrame | None, sites: pd.DataFrame | None) -> pd.DataFrame:
    f = fact.copy()
    for c in FACT_COLS:
        if c not in f.columns: f[c] = None
    f["Date"] = pd.to_datetime(f["Date"], errors="coerce")
    f["CategoryID"] = f["CategoryID"].astype(str).str.strip().str.replace(r"^(\d+)\.0$", r"\1", regex=True)
    f["Scope"] = f["Scope"].astype(str).str.strip().str.replace(r"^(?i)scope\s*", "Scope ", regex=True)
    f["Quantity"] = pd.to_numeric(f["Quantity"], errors="coerce").fillna(0)
    for c in ["EmissionFactor", "EmissionFactorMarket", "tCO2e_Location", "tCO2e_Market", "DataQualityScore"]:
        f[c] = pd.to_numeric(f[c], errors="coerce")
    f["EmissionFactorMarket"] = f["EmissionFactorMarket"].fillna(f["EmissionFactor"])
    if f["tCO2e_Location"].isna().all():
        f["tCO2e_Location"] = f.Quantity * f.EmissionFactor / 1000
    if f["tCO2e_Market"].isna().all():
        f["tCO2e_Market"] = f.Quantity * f.EmissionFactorMarket / 1000
    f["tCO2e_Market"] = f["tCO2e_Market"].fillna(f["tCO2e_Location"])
    f["DataQualityScore"] = f["DataQualityScore"].fillna(3).clip(1, 5).astype(int)
    f["Year"] = f.Date.dt.year; f["YearMonth"] = f.Date.dt.to_period("M").astype(str)
    f["CategoryName"] = f.CategoryID.map(CAT_NAMES).fillna(f.get("CategoryName"))
    f["ScopeCategory"] = np.where(f.Scope == "Scope 3", "Cat " + f.CategoryID.str[2:] + " · " + f.CategoryName.astype(str), f.Scope)
    if sites is not None and "SiteName" in sites.columns:
        f["SiteName"] = f.SiteID.astype(str).map(sites.set_index(sites.SiteID.astype(str)).SiteName).fillna(f.SiteID.astype(str))
    elif "SiteName" not in f.columns or f["SiteName"].isna().all():
        f["SiteName"] = f.SiteID.astype(str)
    f["Method"] = f["Method"].fillna("Unspecified")
    return f


def activity_to_fact(df: pd.DataFrame, mapping: dict, factors: pd.DataFrame, choices: dict) -> tuple[pd.DataFrame, list]:
    """Convert a plain activity file into the fact layout using the factor library."""
    out = pd.DataFrame({
        "Date": pd.to_datetime(df[mapping["Date"]], errors="coerce"),
        "SiteID": df[mapping["Site"]].astype(str) if mapping.get("Site") else "Site",
        "Scope": df[mapping["Scope"]].astype(str),
        "CategoryID": df[mapping["Category"]].astype(str) if mapping.get("Category") else None,
        "ActivityType": df[mapping["ActivityType"]].astype(str).str.strip(),
        "Quantity": pd.to_numeric(df[mapping["Quantity"]], errors="coerce"),
        "Unit": df[mapping["Unit"]].astype(str) if mapping.get("Unit") else None,
        "EmissionFactor": pd.to_numeric(df[mapping["Factor"]], errors="coerce") if mapping.get("Factor") else np.nan,
    })
    lib = factors.copy(); lib["key"] = lib.ActivityType.str.lower().str.strip()
    unmatched = []
    for act in out.ActivityType.unique():
        mask = out.ActivityType == act
        if out.loc[mask, "EmissionFactor"].notna().all():
            continue
        fid = choices.get(act)
        if fid is None:
            hit = lib[lib.key == act.lower()]
            if hit.empty: hit = lib[lib.key.str.contains(act.lower()[:12], regex=False)]
            fid = hit.FactorID.iloc[0] if len(hit) else None
        if fid is None:
            unmatched.append(act); continue
        row = lib.set_index("FactorID").loc[fid]
        out.loc[mask, ["EmissionFactor", "FactorID", "FactorUnit", "Method"]] = [row.Factor, fid, row.FactorUnit, row.Method]
        if out.loc[mask, "CategoryID"].isna().all() or (out.loc[mask, "CategoryID"] == "None").all():
            out.loc[mask, "CategoryID"] = str(row.CategoryID)
    out["DataQualityScore"] = out.Method.map({"Fuel-based": 1, "Mass-balance": 1, "Location-based": 1, "Market-based": 1, "Supplier-specific": 2, "Distance-based": 3, "Average-data": 3, "Waste-type-specific": 3, "Spend-based": 4}).fillna(3)
    out = out[out.EmissionFactor.notna()]
    return out, unmatched


def apply_factor_edits(fact: pd.DataFrame, factors: pd.DataFrame) -> pd.DataFrame:
    """Recompute emissions from quantities when the factor library is edited (keeps the market/location ratio)."""
    f = fact.copy()
    lib = factors.set_index("FactorID").Factor
    new = f.FactorID.map(lib)
    has = new.notna() & f.EmissionFactor.notna()
    ratio = np.where(f.EmissionFactor.fillna(0) > 0, f.EmissionFactorMarket / f.EmissionFactor.replace(0, np.nan), 1.0)
    f.loc[has, "EmissionFactor"] = new[has]
    f.loc[has, "EmissionFactorMarket"] = new[has] * pd.Series(ratio, index=f.index)[has].fillna(1.0)
    f["tCO2e_Location"] = f.Quantity * f.EmissionFactor / 1000
    f["tCO2e_Market"] = f.Quantity * f.EmissionFactorMarket / 1000
    return f


# --------------------------------------------------------------------------------------- chart helpers
def layout(fig, title, height=380, legend=True):
    fig.update_layout(title=dict(text=title, x=0, font=dict(size=16, color=INK)), margin=dict(l=10, r=10, t=56, b=10), height=height,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title_text=""), showlegend=legend,
                      font=dict(family="Inter, Segoe UI, Arial", size=12, color=INK), plot_bgcolor="white", paper_bgcolor="white")
    fig.update_yaxes(gridcolor="#EEF0F2", showline=False, zeroline=False)
    fig.update_xaxes(showgrid=False)
    return fig


def gauge(value, target, title, suffix="", good_below=True):
    """Indicator gauge: value vs target with coloured bands."""
    vmax = max(value, target) * 1.25 if max(value, target) > 0 else 1
    bands = [[0, target * 0.9, "#DDEFD9"], [target * 0.9, target * 1.1, "#F5E9C3"], [target * 1.1, vmax, "#F3D6D3"]] if good_below else \
            [[0, target * 0.9, "#F3D6D3"], [target * 0.9, target * 1.1, "#F5E9C3"], [target * 1.1, vmax, "#DDEFD9"]]
    fig = go.Figure(go.Indicator(mode="gauge+number+delta", value=value, number=dict(suffix=suffix, valueformat=",.0f", font=dict(size=30, color=INK)),
                                 delta=dict(reference=target, valueformat=",.0f", increasing=dict(color=WARN if good_below else "#1E5631"), decreasing=dict(color="#1E5631" if good_below else WARN)),
                                 title=dict(text=title, font=dict(size=13, color=GREY)),
                                 gauge=dict(axis=dict(range=[0, vmax], tickformat=",.0f", tickfont=dict(size=10)), bar=dict(color=INK, thickness=0.25),
                                            steps=[dict(range=[a, b], color=c) for a, b, c in bands],
                                            threshold=dict(line=dict(color=ACCENT, width=3), thickness=0.85, value=target))))
    fig.update_layout(height=230, margin=dict(l=20, r=20, t=40, b=0), paper_bgcolor="white", font=dict(family="Inter, Segoe UI, Arial"))
    return fig


def kpis(items):
    cols = st.columns(len(items))
    for c, (label, value, delta, help_) in zip(cols, items):
        c.metric(label, value, delta=delta, delta_color="inverse", help=help_)


def fmt(x, d=0):
    return f"{x:,.{d}f}"


# --------------------------------------------------------------------------------------- sidebar
st.sidebar.title("GHG inventory dashboard")
up = st.sidebar.file_uploader("Upload an inventory or activity file", type=["xlsx", "xlsm", "xls", "csv"],
                              help="The GHG_Inventory_Scope123 workbook, a CSV in its FactEmissions layout, or a plain activity file (date, site, scope, category, activity type, quantity, unit). Files stay in this session and are never stored.")
st.sidebar.caption("No file yet? The dashboard runs on a public sample: Siam Precision Pumps, a fictional manufacturer, 2023–2025.")

s_fact, s_factors, s_targets, s_sites, s_cats = load_sample()
fact_raw, factors, targets, sites, cats, source = s_fact, s_factors, s_targets, s_sites, s_cats, "Sample data — Siam Precision Pumps (synthetic)"
if up is not None:
    f, fa, tg, si, ca, kind = read_upload(up.getvalue(), up.name)
    if kind == "workbook":
        fact_raw, factors, targets, sites, cats = f, (fa if fa is not None else s_factors), tg, si, (ca if ca is not None else s_cats); source = f"{up.name} (workbook)"
    elif kind == "fact":
        fact_raw, factors, targets, sites, cats = f, s_factors, None, None, s_cats; source = f"{up.name} (fact layout)"
    else:
        st.sidebar.warning("Activity file detected — map the columns.")
        with st.sidebar.form("map"):
            opts = ["(none)"] + list(f.columns)
            def pick(label, guesses, required=False):
                g = next((c for c in f.columns if any(k in c.lower() for k in guesses)), None)
                return st.selectbox(label + (" *" if required else ""), opts, index=opts.index(g) if g in opts else 0)
            m = {"Date": pick("Date", ["date", "month", "period"], True), "Site": pick("Site / facility", ["site", "facility", "location", "plant"]),
                 "Scope": pick("Scope", ["scope"], True), "Category": pick("Scope 3 category (e.g. 3.6)", ["categ"]),
                 "ActivityType": pick("Activity type", ["activity", "source", "fuel"], True), "Quantity": pick("Quantity", ["quantity", "amount", "consumption", "kwh", "value"], True),
                 "Unit": pick("Unit", ["unit"]), "Factor": pick("Emission factor (if you have one)", ["factor", "ef"])}
            ok = st.form_submit_button("Apply mapping")
        if ok or st.session_state.get("mapping_applied"):
            st.session_state.mapping_applied = True
            m = {k: (None if v == "(none)" else v) for k, v in m.items()}
            missing = [k for k in ["Date", "Scope", "ActivityType", "Quantity"] if not m.get(k)]
            if missing:
                st.sidebar.error("Still needed: " + ", ".join(missing))
            else:
                choices = st.session_state.get("factor_choices", {})
                conv, unmatched = activity_to_fact(f, m, s_factors, choices)
                if unmatched:
                    with st.sidebar.expander(f"{len(unmatched)} activity types need a factor", expanded=True):
                        lib = s_factors.assign(label=lambda d: d.ActivityType + " — " + d.Factor.map(lambda v: f"{v:g}") + " " + d.FactorUnit)
                        for act in unmatched:
                            sel = st.selectbox(act, ["(skip)"] + lib.label.tolist(), key=f"fc_{act}")
                            if sel != "(skip)":
                                choices[act] = lib.set_index("label").loc[sel, "FactorID"]
                        st.session_state.factor_choices = choices
                        conv, unmatched = activity_to_fact(f, m, s_factors, choices)
                fact_raw, factors, targets, sites, cats = conv, s_factors, None, None, s_cats; source = f"{up.name} (activity file, {len(conv):,} lines)"
        else:
            st.info("Map the required columns in the sidebar to analyse this file. The sample data is shown meanwhile.")

if "factor_edits" not in st.session_state or st.session_state.get("factor_source") != source:
    st.session_state.factor_edits = factors.copy(); st.session_state.factor_source = source
factors_live = st.session_state.factor_edits
fact_all = normalise_fact(fact_raw, factors_live, sites)
fact_all = apply_factor_edits(fact_all, factors_live)
fact_all["DQ_Weighted"] = fact_all.DataQualityScore * fact_all.tCO2e_Location

with st.sidebar.expander("Settings", expanded=True):
    basis = st.radio("Scope 2 basis", ["Market-based", "Location-based"], horizontal=True)
    include_use = st.toggle("Include use of sold products (3.11)", value=True)
    years_all = sorted(fact_all.Year.dropna().astype(int).unique())
    base_year = st.selectbox("Base year", years_all, index=0)
    t12 = st.slider("Scope 1+2 reduction target by 2030", 0, 90, 42, 1, format="%d%%")
    t3 = st.slider("Scope 3 reduction target by 2030", 0, 90, 25, 1, format="%d%%")
VAL = "tCO2e_Market" if basis == "Market-based" else "tCO2e_Location"

st.sidebar.markdown("---")
sel_sites = st.sidebar.multiselect("Sites", sorted(fact_all.SiteName.dropna().unique()), default=sorted(fact_all.SiteName.dropna().unique()))
sel_scopes = st.sidebar.multiselect("Scopes", ["Scope 1", "Scope 2", "Scope 3"], default=["Scope 1", "Scope 2", "Scope 3"])
year_focus = st.sidebar.selectbox("Reporting year", years_all, index=len(years_all) - 1)
page = st.sidebar.radio("Page", ["Overview", "Scope 1 & 2 by site", "Scope 3 categories", "Hotspots", "Targets", "Data quality", "Factors & assumptions"])
st.sidebar.caption(f"Source: {source}")

mask = fact_all.SiteName.isin(sel_sites) & fact_all.Scope.isin(sel_scopes)
if not include_use:
    mask &= fact_all.CategoryID != "3.11"
fact = fact_all[mask].copy()
if fact.empty:
    st.warning("No data matches the current filters."); st.stop()
fy = fact[fact.Year == year_focus]
py = fact[fact.Year == year_focus - 1]
by_scope_year = fact.groupby(["Year", "Scope"])[VAL].sum().unstack(fill_value=0).reindex(columns=["Scope 1", "Scope 2", "Scope 3"], fill_value=0)
tot_year = by_scope_year.sum(axis=1)

# targets: linear path from base year to 2030 using the sliders
base_tot = fact_all[(fact_all.Year == base_year) & fact_all.SiteName.isin(sel_sites)]
base12 = base_tot[base_tot.Scope.isin(["Scope 1", "Scope 2"])][VAL].sum()
base3 = base_tot[(base_tot.Scope == "Scope 3") & (include_use | (base_tot.CategoryID != "3.11"))][VAL].sum()
path = pd.DataFrame({"Year": range(base_year, 2031)})
path["Scope12_Target"] = base12 * (1 - t12 / 100 * (path.Year - base_year) / max(2030 - base_year, 1))
path["Scope3_Target"] = base3 * (1 - t3 / 100 * (path.Year - base_year) / max(2030 - base_year, 1))
actual12 = fact[fact.Scope.isin(["Scope 1", "Scope 2"])].groupby("Year")[VAL].sum()
actual3 = fact[fact.Scope == "Scope 3"].groupby("Year")[VAL].sum()
rev = targets.set_index("Year")["RevenueUSDm"] if targets is not None and "RevenueUSDm" in targets.columns else None
units = targets.set_index("Year")["UnitsSold"] if targets is not None and "UnitsSold" in targets.columns else None


def findings():
    out = []
    s3share = by_scope_year.loc[year_focus, "Scope 3"] / tot_year[year_focus] if tot_year[year_focus] else 0
    out.append(f"Scope 3 is {s3share:.0%} of the {year_focus} footprint ({fmt(by_scope_year.loc[year_focus, 'Scope 3'])} tCO2e); Scope 1 and 2 together are {fmt(by_scope_year.loc[year_focus, ['Scope 1', 'Scope 2']].sum())} tCO2e.")
    top = fy[fy.Scope == "Scope 3"].groupby("ScopeCategory")[VAL].sum().sort_values(ascending=False)
    if len(top):
        out.append(f"The largest Scope 3 category is {top.index[0]} at {top.iloc[0] / max(top.sum(), 1):.0%} of Scope 3 — the first place to look for reductions and for better data.")
    if year_focus - 1 in tot_year.index:
        ch = tot_year[year_focus] / tot_year[year_focus - 1] - 1
        out.append(f"Total emissions {'rose' if ch > 0 else 'fell'} {abs(ch):.1%} versus {year_focus - 1}" + (f"; intensity per USD m revenue {'rose' if (tot_year[year_focus]/rev[year_focus])/(tot_year[year_focus-1]/rev[year_focus-1]) > 1 else 'fell'} {abs((tot_year[year_focus]/rev[year_focus])/(tot_year[year_focus-1]/rev[year_focus-1]) - 1):.1%}." if rev is not None and year_focus in rev.index and year_focus - 1 in rev.index else "."))
    if year_focus in path.Year.values and year_focus in actual12.index:
        gap = actual12[year_focus] / path.set_index("Year").loc[year_focus, "Scope12_Target"] - 1
        out.append(f"Scope 1+2 is {abs(gap):.1%} {'behind' if gap > 0 else 'ahead of'} the linear path to a {t12}% cut by 2030.")
    dq = fy.DQ_Weighted.sum() / max(fy.tCO2e_Location.sum(), 1)
    spend = fy[(fy.Scope == "Scope 3") & (fy.Method == "Spend-based")][VAL].sum() / max(fy[fy.Scope == "Scope 3"][VAL].sum(), 1)
    out.append(f"Emissions-weighted data quality is {dq:.2f} (1 best, 5 worst); {spend:.0%} of Scope 3 still rests on spend-based estimates.")
    loc, mkt = fy[fy.Scope == "Scope 2"].tCO2e_Location.sum(), fy[fy.Scope == "Scope 2"].tCO2e_Market.sum()
    if loc and abs(loc - mkt) / loc > 0.02:
        out.append(f"Contractual instruments (PPA, certified tariffs) cut Scope 2 from {fmt(loc)} location-based to {fmt(mkt)} market-based tCO2e, a {1 - mkt / loc:.0%} difference.")
    return out


def summary_xlsx():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        by_scope_year.assign(Total=tot_year).to_excel(xw, sheet_name="By scope")
        fact.groupby(["Year", "ScopeCategory"])[VAL].sum().unstack(0).to_excel(xw, sheet_name="By category")
        fact.groupby(["Year", "SiteName", "Scope"])[VAL].sum().unstack(0).to_excel(xw, sheet_name="By site")
        path.merge(actual12.rename("Scope12_Actual"), on="Year", how="left").merge(actual3.rename("Scope3_Actual"), on="Year", how="left").to_excel(xw, sheet_name="Targets", index=False)
        factors_live.to_excel(xw, sheet_name="Factors used", index=False)
        pd.DataFrame({"Finding": findings()}).to_excel(xw, sheet_name="Findings", index=False)
    return buf.getvalue()
st.sidebar.download_button("Download summary (Excel)", summary_xlsx(), "ghg_inventory_summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --------------------------------------------------------------------------------------- pages
if page == "Overview":
    st.title(f"GHG inventory {year_focus}")
    st.caption(f"{basis} Scope 2 · {'including' if include_use else 'excluding'} use of sold products · {len(sel_sites)} site(s)")
    prev = lambda s: (by_scope_year.loc[year_focus, s] - by_scope_year.loc[year_focus - 1, s]) if year_focus - 1 in by_scope_year.index else None
    dl = lambda v: None if v is None else f"{v:+,.0f} vs {year_focus - 1}"
    kpis([("Total tCO2e", fmt(tot_year[year_focus]), dl(tot_year[year_focus] - tot_year[year_focus - 1]) if year_focus - 1 in tot_year.index else None, "All scopes, selected sites"),
          ("Scope 1", fmt(by_scope_year.loc[year_focus, "Scope 1"]), dl(prev("Scope 1")), "Direct emissions"),
          ("Scope 2", fmt(by_scope_year.loc[year_focus, "Scope 2"]), dl(prev("Scope 2")), basis),
          ("Scope 3", fmt(by_scope_year.loc[year_focus, "Scope 3"]), dl(prev("Scope 3")), "Value chain")])
    r2 = [("Scope 3 share", f"{by_scope_year.loc[year_focus, 'Scope 3'] / tot_year[year_focus]:.0%}" if tot_year[year_focus] else "—", None, "")]
    if rev is not None and year_focus in rev.index:
        i_now = tot_year[year_focus] / rev[year_focus]; i_prev = tot_year[year_focus - 1] / rev[year_focus - 1] if year_focus - 1 in rev.index and year_focus - 1 in tot_year.index else None
        r2.append(("tCO2e per USD m revenue", fmt(i_now, 1), f"{i_now - i_prev:+,.1f}" if i_prev else None, "Intensity"))
    if units is not None and year_focus in units.index:
        r2.append(("Scope 1+2 kgCO2e per unit", fmt(actual12.get(year_focus, 0) * 1000 / units[year_focus], 1), None, "Product intensity"))
    dq = fy.DQ_Weighted.sum() / max(fy.tCO2e_Location.sum(), 1)
    r2.append(("Data quality (1 best–5)", f"{dq:.2f}", None, "Emissions-weighted score"))
    kpis(r2)

    c1, c2 = st.columns([3, 2])
    m = fact.groupby(["YearMonth", "Scope"])[VAL].sum().reset_index()
    fig = px.area(m, x="YearMonth", y=VAL, color="Scope", color_discrete_map=SCOPE_COLORS, category_orders={"Scope": ["Scope 1", "Scope 2", "Scope 3"]})
    fig.update_traces(hovertemplate="%{x}<br>%{y:,.0f} tCO2e"); fig.update_layout(yaxis_title="tCO2e per month", xaxis_title="")
    c1.plotly_chart(layout(fig, "Monthly emissions by scope"), width="stretch")
    d = by_scope_year.loc[year_focus].reset_index(); d.columns = ["Scope", "tCO2e"]
    fig = px.pie(d, names="Scope", values="tCO2e", hole=0.55, color="Scope", color_discrete_map=SCOPE_COLORS)
    fig.update_traces(textinfo="percent+label", textposition="outside", hovertemplate="%{label}: %{value:,.0f} tCO2e")
    fig.add_annotation(text=f"<b>{fmt(tot_year[year_focus])}</b><br>tCO2e", showarrow=False, font=dict(size=16))
    c2.plotly_chart(layout(fig, f"Share by scope, {year_focus}", legend=False), width="stretch")

    c3, c4 = st.columns(2)
    if year_focus in path.Year.values:
        c3.plotly_chart(gauge(actual12.get(year_focus, 0), path.set_index("Year").loc[year_focus, "Scope12_Target"], f"Scope 1+2 vs target path {year_focus} (tCO2e)"), width="stretch")
        c4.plotly_chart(gauge(actual3.get(year_focus, 0), path.set_index("Year").loc[year_focus, "Scope3_Target"], f"Scope 3 vs target path {year_focus} (tCO2e)"), width="stretch")
        st.caption("Gauges: needle at the gold line means on the linear pathway to the 2030 target; green band is ahead, red band behind.")
    st.markdown("#### Key findings")
    for i, t in enumerate(findings(), 1):
        st.write(f"{i}. {t}")

elif page == "Scope 1 & 2 by site":
    st.title("Scope 1 & 2 by site")
    s12 = fact[fact.Scope.isin(["Scope 1", "Scope 2"])]
    if s12.empty:
        st.info("No Scope 1 or 2 lines in the current selection.")
    else:
        d = s12[s12.Year == year_focus].groupby(["SiteName", "ActivityType"])[VAL].sum().reset_index()
        fig = px.bar(d, x="SiteName", y=VAL, color="ActivityType", color_discrete_sequence=CAT_PALETTE)
        fig.update_layout(yaxis_title="tCO2e", xaxis_title="", barmode="stack"); fig.update_traces(hovertemplate="%{x}<br>%{y:,.0f} tCO2e")
        st.plotly_chart(layout(fig, f"Scope 1 & 2 by site and activity, {year_focus}", height=420), width="stretch")
        c1, c2 = st.columns(2)
        s2 = fact[fact.Scope == "Scope 2"].groupby("YearMonth")[["tCO2e_Location", "tCO2e_Market"]].sum().reset_index()
        if len(s2):
            fig = go.Figure()
            fig.add_scatter(x=s2.YearMonth, y=s2.tCO2e_Location, name="Location-based", line=dict(color=GREY, width=2, dash="dot"))
            fig.add_scatter(x=s2.YearMonth, y=s2.tCO2e_Market, name="Market-based", line=dict(color=SCOPE_COLORS["Scope 2"], width=3), fill="tonexty", fillcolor="rgba(42,157,143,0.12)")
            fig.update_layout(yaxis_title="tCO2e per month", xaxis_title="")
            c1.plotly_chart(layout(fig, "Scope 2: location- vs market-based (gap = contractual instruments)"), width="stretch")
        s1 = fact[fact.Scope == "Scope 1"].groupby(["YearMonth", "ActivityType"])[VAL].sum().reset_index()
        if len(s1):
            fig = px.bar(s1, x="YearMonth", y=VAL, color="ActivityType", color_discrete_sequence=CAT_PALETTE)
            fig.update_layout(yaxis_title="tCO2e per month", xaxis_title="", barmode="stack")
            c2.plotly_chart(layout(fig, "Scope 1 by month and source"), width="stretch")
        st.markdown("#### Site table")
        tbl = s12.groupby(["SiteName", "Scope", "Year"])[VAL].sum().unstack("Year").fillna(0).reset_index()
        if sites is not None and "Employees" in sites.columns:
            tbl = tbl.merge(sites[["SiteName", "Country", "SiteType", "Employees"]], on="SiteName", how="left")
        st.dataframe(tbl.style.format({y: "{:,.0f}" for y in years_all}), width="stretch", hide_index=True)

elif page == "Scope 3 categories":
    st.title("Scope 3 by category")
    s3 = fact[fact.Scope == "Scope 3"]
    if s3.empty:
        st.info("No Scope 3 lines in the current selection.")
    else:
        d = s3[s3.Year == year_focus].groupby(["CategoryID", "ScopeCategory"])[VAL].sum().reset_index()
        d["Status"] = "Included"
        allc = pd.DataFrame({"CategoryID": S3_ORDER}); allc["ScopeCategory"] = "Cat " + allc.CategoryID.str[2:] + " · " + allc.CategoryID.map(CAT_NAMES)
        d = allc.merge(d[["CategoryID", VAL, "Status"]], on="CategoryID", how="left")
        d[VAL] = d[VAL].fillna(0); d["Status"] = d.Status.fillna("Not applicable / not reported")
        d["Share"] = d[VAL] / max(d[VAL].sum(), 1)
        fig = px.bar(d.iloc[::-1], x=VAL, y="ScopeCategory", orientation="h", color="Status", color_discrete_map={"Included": SCOPE_COLORS["Scope 3"], "Not applicable / not reported": "#E4E6E9"},
                     text=d.iloc[::-1].Share.map(lambda v: f"{v:.1%}" if v > 0 else "n/a"))
        fig.update_traces(textposition="outside", hovertemplate="%{y}<br>%{x:,.0f} tCO2e"); fig.update_layout(xaxis_title="tCO2e", yaxis_title="")
        st.plotly_chart(layout(fig, f"All 15 Scope 3 categories, {year_focus} (grey = not applicable, with the reason in the table below)", height=520), width="stretch")
        c1, c2 = st.columns(2)
        tr = s3.groupby(["Year", "ScopeCategory"])[VAL].sum().reset_index()
        fig = px.area(tr, x="Year", y=VAL, color="ScopeCategory", color_discrete_sequence=CAT_PALETTE)
        fig.update_layout(yaxis_title="tCO2e", xaxis=dict(dtick=1, title=""))
        c1.plotly_chart(layout(fig, "Scope 3 trend by category", height=420), width="stretch")
        if year_focus - 1 in s3.Year.values:
            w = s3.groupby(["Year", "ScopeCategory"])[VAL].sum().unstack(0).fillna(0)
            delta = (w[year_focus] - w[year_focus - 1]).sort_values()
            fig = go.Figure(go.Waterfall(orientation="v", measure=["absolute"] + ["relative"] * len(delta) + ["total"],
                                         x=[str(year_focus - 1)] + list(delta.index) + [str(year_focus)], y=[w[year_focus - 1].sum()] + list(delta.values) + [0],
                                         increasing=dict(marker_color=WARN), decreasing=dict(marker_color="#1E5631"), totals=dict(marker_color=GREY), connector=dict(line=dict(color="#CCD0D5"))))
            fig.update_layout(yaxis_title="tCO2e", xaxis_tickangle=-35)
            c2.plotly_chart(layout(fig, f"What changed from {year_focus - 1} to {year_focus}", height=420, legend=False), width="stretch")
        st.markdown("#### Category status and method")
        if cats is not None:
            show = cats[cats.Scope == "Scope 3"].merge(d[["CategoryID", VAL, "Share"]], on="CategoryID", how="left")
            st.dataframe(show.rename(columns={VAL: f"tCO2e {year_focus}"}).style.format({f"tCO2e {year_focus}": "{:,.0f}", "Share": "{:.1%}"}), width="stretch", hide_index=True)

elif page == "Hotspots":
    st.title("Hotspots")
    d = fy.groupby(["Scope", "ScopeCategory", "ActivityType"])[VAL].sum().reset_index()
    d = d[d[VAL] > 0]
    fig = px.treemap(d, path=[px.Constant("All scopes"), "Scope", "ScopeCategory", "ActivityType"], values=VAL, color="Scope", color_discrete_map={**SCOPE_COLORS, "(?)": "#EEF0F2"})
    fig.update_traces(hovertemplate="%{label}<br>%{value:,.0f} tCO2e<br>%{percentRoot:.1%} of total", textinfo="label+percent root")
    st.plotly_chart(layout(fig, f"Where the {year_focus} footprint comes from (click a box to zoom)", height=520, legend=False), width="stretch")
    c1, c2 = st.columns([3, 2])
    p = fy.groupby("ActivityType")[VAL].sum().sort_values(ascending=False).reset_index()
    p["Cumulative"] = p[VAL].cumsum() / p[VAL].sum()
    n80 = int((p.Cumulative <= 0.8).sum()) + 1
    fig = go.Figure()
    fig.add_bar(x=p.ActivityType, y=p[VAL], name="tCO2e", marker_color=[SCOPE_COLORS["Scope 1"] if i < n80 else "#C9CDD2" for i in range(len(p))])
    fig.add_scatter(x=p.ActivityType, y=p.Cumulative, name="Cumulative share", yaxis="y2", line=dict(color=ACCENT, width=3), mode="lines+markers")
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", tickformat=".0%", range=[0, 1.02], showgrid=False), xaxis_tickangle=-40, yaxis_title="tCO2e")
    c1.plotly_chart(layout(fig, f"Pareto: {n80} activity types make up 80% of emissions", height=460), width="stretch")
    top = p.head(10).copy(); top["Share"] = top[VAL] / p[VAL].sum()
    c2.markdown("#### Top 10 activity types")
    c2.dataframe(top.rename(columns={VAL: "tCO2e"}).style.format({"tCO2e": "{:,.0f}", "Cumulative": "{:.0%}", "Share": "{:.1%}"}), width="stretch", hide_index=True)

elif page == "Targets":
    st.title("Targets")
    st.caption(f"Linear pathways from base year {base_year}: Scope 1+2 −{t12}% and Scope 3 −{t3}% by 2030. Change the sliders in Settings to test other ambitions.")
    tp = path.set_index("Year")
    fig = go.Figure()
    fig.add_scatter(x=tp.index, y=tp.Scope12_Target, name="Scope 1+2 target path", line=dict(color=SCOPE_COLORS["Scope 1"], dash="dash"))
    fig.add_scatter(x=actual12.index, y=actual12.values, name="Scope 1+2 actual", line=dict(color=SCOPE_COLORS["Scope 1"], width=4), mode="lines+markers", marker=dict(size=9))
    fig.update_layout(yaxis_title="tCO2e", xaxis=dict(dtick=1, title=""))
    c1, c2 = st.columns(2)
    c1.plotly_chart(layout(fig, "Scope 1+2: actual vs pathway to 2030", height=400), width="stretch")
    fig = go.Figure()
    fig.add_scatter(x=tp.index, y=tp.Scope3_Target, name="Scope 3 target path", line=dict(color=SCOPE_COLORS["Scope 2"], dash="dash"))
    fig.add_scatter(x=actual3.index, y=actual3.values, name="Scope 3 actual", line=dict(color=SCOPE_COLORS["Scope 2"], width=4), mode="lines+markers", marker=dict(size=9))
    fig.update_layout(yaxis_title="tCO2e", xaxis=dict(dtick=1, title=""))
    c2.plotly_chart(layout(fig, "Scope 3: actual vs pathway to 2030", height=400), width="stretch")
    latest = max(actual12.index)
    rows = []
    for label, act, tcol, red in [("Scope 1+2", actual12, "Scope12_Target", t12), ("Scope 3", actual3, "Scope3_Target", t3)]:
        a = act.get(latest, 0); tgt2030 = tp.loc[2030, tcol]; onpath = tp.loc[latest, tcol] if latest in tp.index else np.nan
        yrs = max(2030 - latest, 1); req = 1 - (tgt2030 / a) ** (1 / yrs) if a > 0 and tgt2030 > 0 else np.nan
        rows.append({"Target": label, f"Actual {latest}": a, f"Pathway {latest}": onpath, "Gap to pathway": a - onpath, "2030 target": tgt2030, "Cut still needed": a - tgt2030, "Required reduction per year": req,
                     "Change since base year": a / act.get(base_year, np.nan) - 1 if act.get(base_year, 0) else np.nan})
    st.dataframe(pd.DataFrame(rows).style.format({f"Actual {latest}": "{:,.0f}", f"Pathway {latest}": "{:,.0f}", "Gap to pathway": "{:+,.0f}", "2030 target": "{:,.0f}", "Cut still needed": "{:,.0f}", "Required reduction per year": "{:.1%}", "Change since base year": "{:+.1%}"}), width="stretch", hide_index=True)
    st.caption("A positive gap means emissions are above the linear pathway. 'Required reduction per year' is the compound annual cut needed from the latest year to reach the 2030 target.")

elif page == "Data quality":
    st.title("Data quality")
    st.caption("Score per line: 1 metered/invoice · 2 supplier-specific · 3 distance or average data · 4 spend-based · 5 estimated. Weighted by emissions, so big categories count more.")
    dq_cat = fy.groupby("ScopeCategory").apply(lambda d: pd.Series({"tCO2e": d[VAL].sum(), "Weighted DQ": d.DQ_Weighted.sum() / max(d.tCO2e_Location.sum(), 1)}), include_groups=False).reset_index()
    dq_cat["Improvement potential"] = dq_cat["tCO2e"] * (dq_cat["Weighted DQ"] - 1)
    c1, c2 = st.columns(2)
    fig = px.bar(dq_cat.sort_values("Weighted DQ"), x="Weighted DQ", y="ScopeCategory", orientation="h", color="Weighted DQ", color_continuous_scale=["#1E5631", "#C9A227", "#C0504D"], range_color=[1, 5], text=dq_cat.sort_values("Weighted DQ")["Weighted DQ"].map(lambda v: f"{v:.2f}"))
    fig.update_traces(textposition="outside"); fig.update_layout(xaxis=dict(range=[0, 5.5], title="Weighted data-quality score (1 best)"), yaxis_title="", coloraxis_showscale=False)
    c1.plotly_chart(layout(fig, f"Data quality by category, {year_focus}", height=480, legend=False), width="stretch")
    meth = fact.groupby(["Year", "Method"])[VAL].sum().reset_index()
    meth["Share"] = meth[VAL] / meth.groupby("Year")[VAL].transform("sum")
    fig = px.bar(meth, x="Year", y="Share", color="Method", color_discrete_sequence=CAT_PALETTE, barmode="stack")
    fig.update_layout(yaxis_tickformat=".0%", yaxis_title="Share of emissions", xaxis=dict(dtick=1, title=""))
    c2.plotly_chart(layout(fig, "Share of emissions by calculation method", height=480), width="stretch")
    st.markdown("#### Where better data would help most")
    st.caption("Improvement potential = tCO2e × (score − 1): large, poorly measured categories rank first.")
    st.dataframe(dq_cat.sort_values("Improvement potential", ascending=False).style.format({"tCO2e": "{:,.0f}", "Weighted DQ": "{:.2f}", "Improvement potential": "{:,.0f}"}), width="stretch", hide_index=True)

elif page == "Factors & assumptions":
    st.title("Factors & assumptions")
    st.caption("Every factor here is approximate and labelled with the source it approximates. Edit a value and click Apply — every page recalculates from the activity quantities. Market-based Scope 2 keeps its contractual ratio to the location factor.")
    edited = st.data_editor(factors_live, hide_index=True, width="stretch", disabled=[c for c in factors_live.columns if c != "Factor"], key="factor_editor",
                            column_config={"Factor": st.column_config.NumberColumn(format="%.4f", min_value=0.0)})
    c1, c2 = st.columns([1, 5])
    if c1.button("Apply factors"):
        st.session_state.factor_edits = edited; st.rerun()
    if c2.button("Reset to original"):
        st.session_state.factor_edits = factors.copy(); st.rerun()
    st.markdown("#### Category coverage")
    if cats is not None:
        st.dataframe(cats, width="stretch", hide_index=True)
    st.markdown("#### Boundary and methods")
    st.write("Consolidation: operational control. Scope 2 reported location- and market-based. Scope 3.3 derived from Scope 1 and 2 activity (well-to-tank and T&D losses). "
             "Use of sold products (3.11) = units sold × lifetime energy × grid factor of the sales market. GWP100 per IPCC AR4 for refrigerants; switch in the factor table if AR5/AR6 is required.")
