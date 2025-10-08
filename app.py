import pandas as pd, numpy as np
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Co-Pilot", page_icon="📊", layout="wide")
st.title("📊 Executive Co-Pilot – Mining")

# -----------------------
# 0) Load Excel (your renamed file)
# -----------------------
EXCEL_FILE = "jl25pg108_mohak_srivastava.xlsx"  # change here if you rename again
df = pd.read_excel(EXCEL_FILE)  # first sheet by default
cols = [c.strip() for c in df.columns]
df.columns = cols  # trim spaces

# -----------------------
# 1) Column mapping helper
# -----------------------
def guess(colnames, candidates):
    low = {c.lower(): c for c in colnames}
    for name in candidates:
        if name in low:
            return low[name]
    return None

# Guess common names (add more aliases if needed)
guesses = {
    "date":     guess(cols, ["date", "order date", "month", "period"]),
    "company":  guess(cols, ["company", "brand", "player"]),
    "region":   guess(cols, ["region", "zone", "area"]),
    "units":    guess(cols, ["units_sold", "units", "quantity", "qty"]),
    "revenue":  guess(cols, ["revenue", "sales", "amount", "net sales"]),
    "share":    guess(cols, ["market_share_%", "market share %", "market share", "share"]),
    "csat":     guess(cols, ["customer_satisfaction_%", "csat", "satisfaction %", "customer satisfaction %"]),
}

st.sidebar.header("Map your columns (if needed)")
date_col    = st.sidebar.selectbox("Date column",    options=cols, index=cols.index(guesses["date"]) if guesses["date"] in cols else 0)
company_col = st.sidebar.selectbox("Company column", options=cols, index=cols.index(guesses["company"]) if guesses["company"] in cols else 0)
region_col  = st.sidebar.selectbox("Region column",  options=cols, index=cols.index(guesses["region"]) if guesses["region"] in cols else 0)
units_col   = st.sidebar.selectbox("Units column",   options=cols, index=cols.index(guesses["units"]) if guesses["units"] in cols else 0)
rev_col     = st.sidebar.selectbox("Revenue column", options=cols, index=cols.index(guesses["revenue"]) if guesses["revenue"] in cols else 0)

# Optional fields (not mandatory)
share_col   = st.sidebar.selectbox("Market Share % (optional)", options=["<none>"] + cols,
                                   index=(["<none>"] + cols).index(guesses["share"]) if guesses["share"] else 0)
csat_col    = st.sidebar.selectbox("CSAT % (optional)", options=["<none>"] + cols,
                                   index=(["<none>"] + cols).index(guesses["csat"]) if guesses["csat"] else 0)

# -----------------------
# 2) Normalize dataframe to standard names
# -----------------------
work = df.rename(columns={
    date_col: "Date",
    company_col: "Company",
    region_col: "Region",
    units_col: "Units_Sold",
    rev_col: "Revenue",
})
if share_col != "<none>":
    work = work.rename(columns={share_col: "Market_Share_%"})
else:
    work["Market_Share_%"] = np.nan

if csat_col != "<none>":
    work = work.rename(columns={csat_col: "Customer_Satisfaction_%"})
else:
    work["Customer_Satisfaction_%"] = np.nan

# Coerce types
work["Date"] = pd.to_datetime(work["Date"], errors="coerce")
work["Units_Sold"] = pd.to_numeric(work["Units_Sold"], errors="coerce")
work["Revenue"] = pd.to_numeric(work["Revenue"], errors="coerce")
if "Market_Share_%" in work:
    work["Market_Share_%"] = pd.to_numeric(work["Market_Share_%"], errors="coerce")
if "Customer_Satisfaction_%" in work:
    work["Customer_Satisfaction_%"] = pd.to_numeric(work["Customer_Satisfaction_%"], errors="coerce")

# Derived fields
work["Month"] = work["Date"].dt.to_period("M").dt.to_timestamp()
work["Rev_per_Unit"] = work["Revenue"] / work["Units_Sold"].replace(0, np.nan)

# -----------------------
# 3) Filters
# -----------------------
st.sidebar.header("Filters")
companies = sorted(work["Company"].dropna().unique())
regions = sorted(work["Region"].dropna().unique())
sel_companies = st.sidebar.multiselect("Company", companies, default=companies or [])
sel_regions = st.sidebar.multiselect("Region", regions, default=regions or [])

mask = (work["Company"].isin(sel_companies) if sel_companies else True) & \
       (work["Region"].isin(sel_regions) if sel_regions else True)
f = work.loc[mask].copy()

# -----------------------
# 4) KPIs
# -----------------------
def fmt_money(x): 
    return f"₹{x:,.0f}" if pd.notnull(x) else "—"

c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("Revenue", fmt_money(f["Revenue"].sum()))
c2.metric("Units", int(f["Units_Sold"].sum()) if f["Units_Sold"].notna().any() else 0)
c3.metric("Avg Share %", round(f["Market_Share_%"].mean(),2) if f["Market_Share_%"].notna().any() else 0.0)
c4.metric("Avg CSAT %", round(f["Customer_Satisfaction_%"].mean(),2) if f["Customer_Satisfaction_%"].notna().any() else 0.0)
c5.metric("Avg Rev/Unit", fmt_money(f["Rev_per_Unit"].mean()))

# -----------------------
# 5) Charts
# -----------------------
a,b = st.columns(2)
with a:
    piv = f.pivot_table(index="Region", columns="Company", values="Revenue", aggfunc="sum", fill_value=0)
    st.plotly_chart(px.bar(piv, barmode="group", title="Revenue by Region × Company"),
                    use_container_width=True)
with b:
    mt = f.groupby("Month", as_index=False).agg({"Revenue":"sum","Units_Sold":"sum"})
    st.plotly_chart(px.line(mt, x="Month", y=["Revenue","Units_Sold"], title="Monthly Trend", markers=True),
                    use_container_width=True)

a,b = st.columns(2)
with a:
    cs = f.groupby("Company", as_index=False)["Customer_Satisfaction_%"].mean().dropna()
    if not cs.empty:
        st.plotly_chart(px.bar(cs, x="Customer_Satisfaction_%", y="Company", orientation="h",
                               title="Avg CSAT by Company"), use_container_width=True)
    else:
        st.info("CSAT column not provided.")
with b:
    comp = f.groupby("Company", as_index=False).agg(Revenue=("Revenue","sum"),
                                                    Share=("Market_Share_%","mean"))
    if comp["Share"].notna().any():
        fig = px.scatter(comp, x="Share", y="Revenue", text="Company",
                         title="Market Share vs Revenue")
        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Market Share column not provided.")
