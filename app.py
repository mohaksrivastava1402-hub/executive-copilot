import pandas as pd, numpy as np
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Co-Pilot", page_icon="📊", layout="wide")
st.title("📊 Executive Co-Pilot – Mining")

EXCEL_FILE = "jl25pg108_mohak_srivastava.xlsx"  # keep same name as your repo file

def load_best_sheet(path):
    # choose first sheet that has at least 5 non-empty header cells
    xls = pd.ExcelFile(path)
    for sh in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sh)
        if df.shape[1] >= 5:
            return df
    return pd.read_excel(path)  # fallback first

def clean_headers(df):
    # try first 5 rows as potential header rows; pick the one with fewest "Unnamed"
    best = None; best_score = 10**9
    for hdr in range(0, 5):
        try:
            d = pd.read_excel(EXCEL_FILE, header=hdr)
        except Exception:
            continue
        cols = [str(c).strip() for c in d.columns]
        score = sum([c.lower().startswith("unnamed") or c == "" for c in cols])
        if score < best_score:
            best, best_score = d, score
    if best is None:
        best = df.copy()
    best.columns = [str(c).strip() for c in best.columns]
    return best

def find_datetime_col(df):
    # find a column already datetime or convertible with low NaT rate
    candidates = []
    for c in df.columns:
        s = pd.to_datetime(df[c], errors="coerce", infer_datetime_format=True)
        nat_rate = s.isna().mean()
        if nat_rate < 0.2:  # most rows parse -> treat as date
            candidates.append((c, nat_rate))
    if candidates:
        candidates.sort(key=lambda x: x[1])
        return candidates[0][0]
    return None

raw = load_best_sheet(EXCEL_FILE)
df = clean_headers(raw)
cols = list(df.columns)

def guess(colnames, names):
    low = {c.lower(): c for c in colnames}
    for n in names:
        if n in low: return low[n]
    return None

g = {
    "date":    guess(cols, ["date","order date","month","period"]),
    "company": guess(cols, ["company","brand","player"]),
    "region":  guess(cols, ["region","zone","area"]),
    "units":   guess(cols, ["units_sold","units","quantity","qty"]),
    "revenue": guess(cols, ["revenue","sales","amount","net sales"]),
    "share":   guess(cols, ["market_share_%","market share %","market share","share"]),
    "csat":    guess(cols, ["customer_satisfaction_%","csat","satisfaction %","customer satisfaction %"]),
}

# If "Date" wasn't guessed, auto-detect a datetime-like column
if g["date"] is None:
    auto_date = find_datetime_col(df)
    g["date"] = auto_date

st.sidebar.header("Map your columns")
date_options = ["<none>"] + cols
date_default = date_options.index(g["date"]) if g["date"] in cols else 0
date_col    = st.sidebar.selectbox("Date column (optional)", options=date_options, index=date_default)

company_col = st.sidebar.selectbox("Company column", options=cols, index=cols.index(g["company"]) if g["company"] else 0)
region_col  = st.sidebar.selectbox("Region column",  options=cols, index=cols.index(g["region"])  if g["region"]  else 0)
units_col   = st.sidebar.selectbox("Units column",   options=cols, index=cols.index(g["units"])   if g["units"]   else 0)
rev_col     = st.sidebar.selectbox("Revenue column", options=cols, index=cols.index(g["revenue"]) if g["revenue"] else 0)
share_col   = st.sidebar.selectbox("Market Share % (optional)", options=["<none>"]+cols, index=(["<none>"]+cols).index(g["share"]) if g["share"] else 0)
csat_col    = st.sidebar.selectbox("CSAT % (optional)", options=["<none>"]+cols, index=(["<none>"]+cols).index(g["csat"]) if g["csat"] else 0)

# distinct validations
required = [company_col, region_col, units_col, rev_col]
if len(set(required)) != len(required):
    st.error("Two or more required dropdowns point to the **same column**. Map each to a different column.")
    st.stop()

# normalize
work = df.rename(columns={company_col:"Company", region_col:"Region", units_col:"Units_Sold", rev_col:"Revenue"})
if date_col != "<none>": work = work.rename(columns={date_col:"Date"})
if share_col != "<none>": work = work.rename(columns={share_col:"Market_Share_%"})
else: work["Market_Share_%"] = np.nan
if csat_col != "<none>": work = work.rename(columns={csat_col:"Customer_Satisfaction_%"})
else: work["Customer_Satisfaction_%"] = np.nan

# types
if "Date" in work.columns:
    work["Date"] = pd.to_datetime(work["Date"], errors="coerce", infer_datetime_format=True)
work["Units_Sold"] = pd.to_numeric(work["Units_Sold"], errors="coerce")
work["Revenue"]    = pd.to_numeric(work["Revenue"], errors="coerce")
work["Market_Share_%"] = pd.to_numeric(work["Market_Share_%"], errors="coerce")
work["Customer_Satisfaction_%"] = pd.to_numeric(work["Customer_Satisfaction_%"], errors="coerce")
work["Rev_per_Unit"] = work["Revenue"] / work["Units_Sold"].replace(0, np.nan)
if "Date" in work.columns:
    work["Month"] = work["Date"].dt.to_period("M").dt.to_timestamp()

# filters
st.sidebar.header("Filters")
companies = sorted(work["Company"].dropna().unique()); regions = sorted(work["Region"].dropna().unique())
sel_comp = st.sidebar.multiselect("Company", companies, default=companies or [])
sel_reg  = st.sidebar.multiselect("Region", regions,   default=regions or [])
mask = (work["Company"].isin(sel_comp) if sel_comp else True) & (work["Region"].isin(sel_reg) if sel_reg else True)
f = work.loc[mask].copy()

# KPIs
fmt = lambda x: f"₹{x:,.0f}" if pd.notnull(x) else "—"
c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("Revenue", fmt(f["Revenue"].sum()))
c2.metric("Units", int(f["Units_Sold"].sum()) if f["Units_Sold"].notna().any() else 0)
c3.metric("Avg Share %", round(f["Market_Share_%"].mean(),2) if f["Market_Share_%"].notna().any() else 0.0)
c4.metric("Avg CSAT %",  round(f["Customer_Satisfaction_%"].mean(),2) if f["Customer_Satisfaction_%"].notna().any() else 0.0)
c5.metric("Avg Rev/Unit", fmt(f["Rev_per_Unit"].mean()))

# charts
a,b = st.columns(2)
with a:
    piv = f.pivot_table(index="Region", columns="Company", values="Revenue", aggfunc="sum", fill_value=0)
    st.plotly_chart(px.bar(piv, barmode="group", title="Revenue by Region × Company"), use_container_width=True)
with b:
    if "Date" in f.columns and f["Date"].notna().any():
        mt = f.groupby(f["Date"].dt.to_period("M").dt.to_timestamp(), as_index=False).agg({"Revenue":"sum","Units_Sold":"sum"})
        st.plotly_chart(px.line(mt, x="Date", y=["Revenue","Units_Sold"], title="Monthly Trend", markers=True), use_container_width=True)
    else:
        st.info("No Date column mapped — skipping monthly trend.")

a,b = st.columns(2)
with a:
    cs = f.groupby("Company", as_index=False)["Customer_Satisfaction_%"].mean().dropna()
    if not cs.empty:
        st.plotly_chart(px.bar(cs, x="Customer_Satisfaction_%", y="Company", orientation="h", title="Avg CSAT by Company"), use_container_width=True)
    else:
        st.info("CSAT column not provided.")
with b:
    comp = f.groupby("Company", as_index=False).agg(Revenue=("Revenue","sum"), Share=("Market_Share_%","mean"))
    if comp["Share"].notna().any():
        fig = px.scatter(comp, x="Share", y="Revenue", text="Company", title="Market Share vs Revenue")
        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Market Share column not provided.")
