# app.py  — Executive Co-Pilot (robust version)
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
DEFAULT_FILE = "jl25pg108_mohak_srivastava.xlsx"  # change if you want
candidates = [p for p in glob.glob("*.xlsx") if not os.path.basename(p).startswith("~$")]
if DEFAULT_FILE in candidates:
    EXCEL_FILE = DEFAULT_FILE
elif candidates:
    EXCEL_FILE = candidates[0]
else:
    st.error("No .xlsx file found in the repo. Please upload your Excel and redeploy.")
    st.stop()

st.sidebar.info(f"Using file: **{EXCEL_FILE}**")

# ----------------------------
# 1) Pick sheet & header row
# ----------------------------
xls = pd.ExcelFile(EXCEL_FILE)
sheet = st.sidebar.selectbox("Choose sheet", xls.sheet_names, index=0)

# Try header rows 0..7 and score by how few 'Unnamed' there are
def load_with_header(h):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=h, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df

scores, options = [], []
for h in range(0, min(8, 1 + len(pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)))):
    try:
        df_try = load_with_header(h)
        score = sum([c.lower().startswith("unnamed") or c == "" for c in df_try.columns])
        scores.append(score); options.append(h)
    except Exception:
        pass

auto_header = options[scores.index(min(scores))] if options else 0
header_row = st.sidebar.number_input("Header row (0 = first row)", min_value=0, max_value=50,
                                     value=auto_header, step=1)

df_raw = load_with_header(header_row)

# Drop completely empty columns and trim spaces
df_raw = df_raw.loc[:, ~df_raw.columns.str.lower().str.startswith("unnamed")]
df_raw = df_raw.dropna(how="all", axis=1)
df_raw.columns = [str(c).strip() for c in df_raw.columns]

if df_raw.shape[1] < 4:
    st.error("Selected sheet/header doesn’t look like a data table. Try a different header row or sheet.")
    st.stop()

# ----------------------------
# 2) Column guessing + mapping
# ----------------------------
def guess(colnames, candidates):
    low = {c.lower(): c for c in colnames}
    for name in candidates:
        if name in low: return low[name]
    return None

cols = list(df_raw.columns)

guesses = {
    "date":    guess(cols, ["date","order date","month","period"]),
    "company": guess(cols, ["company","brand","player"]),
    "region":  guess(cols, ["region","zone","area","state"]),
    "units":   guess(cols, ["units_sold","units","quantity","qty"]),
    "revenue": guess(cols, ["revenue","sales","amount","net sales","turnover"]),
    "share":   guess(cols, ["market_share_%","market share %","market share","share"]),
    "csat":    guess(cols, ["customer_satisfaction_%","csat","satisfaction %","customer satisfaction %"]),
}

st.sidebar.subheader("Map your columns")
date_col    = st.sidebar.selectbox("Date (optional)", ["<none>"] + cols,
                                   index=(["<none>"]+cols).index(guesses["date"]) if guesses["date"] else 0)
company_col = st.sidebar.selectbox("Company", cols, index=cols.index(guesses["company"]) if guesses["company"] else 0)
region_col  = st.sidebar.selectbox("Region",  cols, index=cols.index(guesses["region"])  if guesses["region"]  else 0)
units_col   = st.sidebar.selectbox("Units",   cols, index=cols.index(guesses["units"])   if guesses["units"]   else 0)
rev_col     = st.sidebar.selectbox("Revenue", cols, index=cols.index(guesses["revenue"]) if guesses["revenue"] else 0)
share_col   = st.sidebar.selectbox("Market Share % (optional)", ["<none>"] + cols,
                                   index=(["<none>"]+cols).index(guesses["share"]) if guesses["share"] else 0)
csat_col    = st.sidebar.selectbox("CSAT % (optional)", ["<none>"] + cols,
                                   index=(["<none>"]+cols).index(guesses["csat"]) if guesses["csat"] else 0)

required = [company_col, region_col, units_col, rev_col]
if len(set(required)) != len(required):
    st.error("Two or more required dropdowns point to the **same column**. Map each to a different column.")
    st.stop()

# ----------------------------
# 3) Normalize + type coercion
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

# Clean numeric-like strings (e.g., "1,234")
def to_num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", ""), errors="coerce")

# Coerce
if "Date" in work.columns:
    work["Date"] = pd.to_datetime(work["Date"], errors="coerce", infer_datetime_format=True)
work["Units_Sold"] = to_num(work["Units_Sold"])
work["Revenue"]    = to_num(work["Revenue"])
work["Market_Share_%"] = to_num(work["Market_Share_%"])
work["Customer_Satisfaction_%"] = to_num(work["Customer_Satisfaction_%"])

work["Rev_per_Unit"] = work["Revenue"] / work["Units_Sold"].replace(0, np.nan)
if "Date" in work.columns:
    work["Month"] = work["Date"].dt.to_period("M").dt.to_timestamp()

# Show a quick preview so you know we read the right thing
with st.expander("Preview data (first 20 rows)"):
    st.dataframe(work.head(20), use_container_width=True)

# ----------------------------
# 4) Filters
# ----------------------------
st.sidebar.subheader("Filters")
companies = sorted(work["Company"].dropna().unique())
regions   = sorted(work["Region"].dropna().unique())

sel_companies = st.sidebar.multiselect("Company", companies, default=companies or [])
sel_regions   = st.sidebar.multiselect("Region", regions, default=regions or [])

mask = (work["Company"].isin(sel_companies) if sel_companies else True) & \
       (work["Region"].isin(sel_regions) if sel_regions else True)
f = work.loc[mask].copy()

# ----------------------------
# 5) KPIs
# ----------------------------
fmt_money = lambda x: f"₹{x:,.0f}" if pd.notnull(x) else "—"
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Revenue", fmt_money(f["Revenue"].sum()))
c2.metric("Units", int(f["Units_Sold"].sum()) if f["Units_Sold"].notna().any() else 0)
c3.metric("Avg Share %", round(f["Market_Share_%"].mean(), 2) if f["Market_Share_%"].notna().any() else 0.0)
c4.metric("Avg CSAT %",  round(f["Customer_Satisfaction_%"].mean(), 2) if f["Customer_Satisfaction_%"].notna().any() else 0.0)
c5.metric("Avg Rev/Unit", fmt_money(f["Rev_per_Unit"].mean()))

# ----------------------------
# 6) Charts
# ----------------------------
a, b = st.columns(2)
with a:
    rev_piv = f.pivot_table(index="Region", columns="Company", values="Revenue", aggfunc="sum", fill_value=0)
    st.plotly_chart(px.bar(rev_piv, barmode="group", title="Revenue by Region × Company"),
                    use_container_width=True)

with b:
    if "Date" in f.columns and f["Date"].notna().any():
        mt = f.groupby(f["Date"].dt.to_period("M").dt.to_timestamp(), as_index=False)\
              .agg({"Revenue":"sum", "Units_Sold":"sum"})
        st.plotly_chart(px.line(mt, x="Date", y=["Revenue","Units_Sold"],
                                title="Monthly Trend", markers=True), use_container_width=True)
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
# 7) Download filtered CSV
# ----------------------------
st.subheader("⬇️ Download")
csv = f.to_csv(index=False).encode("utf-8")
st.download_button("Download filtered CSV", data=csv, file_name="filtered_view.csv", mime="text/csv")

