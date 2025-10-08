# app.py — Executive Co-Pilot (clean v2)
import glob, os
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Co-Pilot", page_icon="📊", layout="wide")
st.title("📊 Executive Co-Pilot – Mining")

# ----------------------------
# 0) Find your Excel file
# ----------------------------
DEFAULT_FILE = "jl25pg108_mohak_srivastava.xlsx"
cands = [p for p in glob.glob("*.xlsx") if not os.path.basename(p).startswith("~$")]
if DEFAULT_FILE in cands:
    EXCEL_FILE = DEFAULT_FILE
elif cands:
    EXCEL_FILE = cands[0]
else:
    st.error("No .xlsx found. Upload your Excel to the repo and redeploy.")
    st.stop()
st.sidebar.info(f"Using file: **{EXCEL_FILE}**")

# ----------------------------
# 1) Sheet & header picker
# ----------------------------
xls = pd.ExcelFile(EXCEL_FILE)
sheet = st.sidebar.selectbox("Choose sheet", xls.sheet_names, index=0)

def load_with_header(h):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=h, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df

# auto-pick header with fewest "Unnamed"
scores, opts = [], []
raw0 = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None, engine="openpyxl")
for h in range(0, min(8, len(raw0))):
    try:
        d = load_with_header(h)
        score = sum([c.lower().startswith("unnamed") or c == "" for c in d.columns])
        scores.append(score); opts.append(h)
    except Exception:
        pass
auto_header = opts[scores.index(min(scores))] if opts else 0
header_row = st.sidebar.number_input("Header row (0 = first row)", min_value=0, max_value=50,
                                     value=auto_header, step=1)

df_raw = load_with_header(header_row)
df_raw = df_raw.loc[:, ~df_raw.columns.str.lower().str.startswith("unnamed")]
df_raw = df_raw.dropna(how="all", axis=1)
df_raw.columns = [str(c).strip() for c in df_raw.columns]
if df_raw.shape[1] < 4:
    st.error("Selected sheet/header doesn’t look like a data table. Try another sheet/header row.")
    st.stop()

cols = list(df_raw.columns)

# ----------------------------
# 2) Filters FIRST (nice UX)
# ----------------------------
def uniq(series):
    return sorted(pd.Series(series).dropna().unique().tolist())

# We’ll map columns later; until then, show a tiny preview:
with st.expander("Preview current header (first 10 rows)"):
    st.dataframe(df_raw.head(10), use_container_width=True)

# ----------------------------
# 3) Column mapping (in expander)
# ----------------------------
def guess(colnames, names):
    low = {c.lower(): c for c in colnames}
    for n in names:
        if n in low: return low[n]
    return None

g = {
    "date":    guess(cols, ["date","order date","month","period"]),
    "company": guess(cols, ["company","brand","player"]),
    "region":  guess(cols, ["region","zone","area","state"]),
    "units":   guess(cols, ["units_sold","units","quantity","qty"]),
    "revenue": guess(cols, ["revenue","sales","amount","net sales","turnover"]),
    "share":   guess(cols, ["market_share_%","market share %","market share","share"]),
    "csat":    guess(cols, ["customer_satisfaction_%","csat","satisfaction %","customer satisfaction %"]),
}

with st.sidebar.expander("Advanced: Map columns (use only if headers look wrong)", expanded=False):
    date_col    = st.selectbox("Date (optional)", ["<none>"] + cols,
                               index=(["<none>"]+cols).index(g["date"]) if g["date"] else 0)
    company_col = st.selectbox("Company", cols, index=cols.index(g["company"]) if g["company"] else 0)
    region_col  = st.selectbox("Region",  cols, index=cols.index(g["region"])  if g["region"]  else 0)
    units_col   = st.selectbox("Units",   cols, index=cols.index(g["units"])   if g["units"]   else 0)
    rev_col     = st.selectbox("Revenue", cols, index=cols.index(g["revenue"]) if g["revenue"] else 0)
    share_col   = st.selectbox("Market Share % (optional)", ["<none>"] + cols,
                               index=(["<none>"]+cols).index(g["share"]) if g["share"] else 0)
    csat_col    = st.selectbox("CSAT % (optional)", ["<none>"] + cols,
                               index=(["<none>"]+cols).index(g["csat"]) if g["csat"] else 0)

# If user didn’t open the expander, set defaults now:
date_col    = locals().get("date_col", "<none>")
company_col = locals().get("company_col", g["company"] or cols[0])
region_col  = locals().get("region_col",  g["region"]  or cols[min(1, len(cols)-1)])
units_col   = locals().get("units_col",   g["units"]   or cols[min(2, len(cols)-1)])
rev_col     = locals().get("rev_col",     g["revenue"] or cols[min(3, len(cols)-1)])
share_col   = locals().get("share_col",   g["share"]   or "<none>")
csat_col    = locals().get("csat_col",    g["csat"]    or "<none>")

# Validate required mappings
required = [company_col, region_col, units_col, rev_col]
if len(set(required)) != len(required):
    st.error("Two or more required dropdowns point to the **same column**. Map each to a different column.")
    st.stop()

# ----------------------------
# 4) Normalize + types
# ----------------------------
work = df_raw.rename(columns={
    company_col: "Company",
    region_col:  "Region",
    units_col:   "Units_Sold",
    rev_col:     "Revenue",
})
if date_col != "<none>":
    work = work.rename(columns={date_col: "Date"})
if share_col != "<none>":
    work = work.rename(columns={share_col: "Market_Share_%"})
else:
    work["Market_Share_%"] = np.nan
if csat_col != "<none>":
    work = work.rename(columns={csat_col: "Customer_Satisfaction_%"})
else:
    work["Customer_Satisfaction_%"] = np.nan

def to_num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(",", ""), errors="coerce")

if "Date" in work.columns:
    work["Date"] = pd.to_datetime(work["Date"], errors="coerce", infer_datetime_format=True)
work["Units_Sold"] = to_num(work["Units_Sold"])
work["Revenue"]    = to_num(work["Revenue"])
work["Market_Share_%"] = to_num(work["Market_Share_%"])
work["Customer_Satisfaction_%"] = to_num(work["Customer_Satisfaction_%"])
work["Rev_per_Unit"] = work["Revenue"] / work["Units_Sold"].replace(0, np.nan)
if "Date" in work.columns:
    work["Month"] = work["Date"].dt.to_period("M").dt.to_timestamp()

# ----------------------------
# 5) Filters (now using normalized columns)
# ----------------------------
st.sidebar.subheader("Filters")
companies = uniq(work["Company"])
regions   = uniq(work["Region"])

sel_companies = st.sidebar.multiselect("Company", companies, default=companies)
sel_regions   = st.sidebar.multiselect("Region", regions, default=regions)

mask = (work["Company"].isin(sel_companies) if sel_companies else True) & \
       (work["Region"].isin(sel_regions) if sel_regions else True)
f = work.loc[mask].copy()

# ----------------------------
# 6) KPIs
# ----------------------------
fmt_money = lambda x: f"₹{x:,.0f}" if pd.notnull(x) else "—"
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Revenue", fmt_money(f["Revenue"].sum()))
c2.metric("Units", int(f["Units_Sold"].sum()) if f["Units_Sold"].notna().any() else 0)
c3.metric("Avg Share %", round(f["Market_Share_%"].mean(), 2) if f["Market_Share_%"].notna().any() else 0.0)
c4.metric("Avg CSAT %",  round(f["Customer_Satisfaction_%"].mean(), 2) if f["Customer_Satisfaction_%"].notna().any() else 0.0)
c5.metric("Avg Rev/Unit", fmt_money(f["Rev_per_Unit"].mean()))

# ----------------------------
# 7) Charts (robust)
# ----------------------------
a, b = st.columns(2)
with a:
    rev_piv = f.pivot_table(index="Region", columns="Company", values="Revenue", aggfunc="sum", fill_value=0)
    st.plotly_chart(px.bar(rev_piv, barmode="group", title="Revenue by Region × Company"),
                    use_container_width=True)

with b:
    # Guarded monthly trend in long format
    if "Date" in f.columns and f["Date"].notna().any():
        mt = f.dropna(subset=["Date"]).copy()
        mt["Date"] = pd.to_datetime(mt["Date"], errors="coerce")
        mt = mt.groupby(pd.Grouper(key="Date", freq="MS")).agg(
            Revenue=("Revenue","sum"),
            Units_Sold=("Units_Sold","sum")
        ).reset_index()
        # keep only numeric columns that actually have data
        value_cols = [c for c in ["Revenue","Units_Sold"]
                      if pd.to_numeric(mt[c], errors="coerce").notna().any()]
        if len(value_cols) >= 1 and mt["Date"].notna().any():
            mt_long = mt.melt(id_vars="Date", value_vars=value_cols,
                              var_name="Metric", value_name="Value")
            st.plotly_chart(
                px.line(mt_long, x="Date", y="Value", color="Metric",
                        title="Monthly Trend", markers=True),
                use_container_width=True
            )
        else:
            st.info("Monthly trend has no numeric data to plot.")
    else:
        st.info("No Date column mapped — skipping monthly trend.")

a, b = st.columns(2)
with a:
    cs = f.groupby("Company", as_index=False)["Customer_Satisfaction_%"].mean().dropna()
    if not cs.empty:
        st.plotly_chart(px.bar(cs, x="Customer_Satisfaction_%", y="Company",
                               orientation="h", title="Avg CSAT by Company"),
                        use_container_width=True)
    else:
        st.info("CSAT column not provided.")
with b:
    comp = f.groupby("Company", as_index=False)\
            .agg(Revenue=("Revenue","sum"), Share=("Market_Share_%","mean"))
    if comp["Share"].notna().any():
        fig = px.scatter(comp, x="Share", y="Revenue", text="Company",
                         title="Market Share vs Revenue")
        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Market Share column not provided.")

# ----------------------------
# 8) Download
# ----------------------------
st.subheader("⬇️ Download")
st.download_button("Download filtered CSV",
                   data=f.to_csv(index=False).encode("utf-8"),
                   file_name="filtered_view.csv",
                   mime="text/csv")
