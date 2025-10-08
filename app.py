
import pandas as pd, numpy as np
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Executive Co-Pilot", page_icon="📊", layout="wide")

df = pd.read_excel("jl25pg108_mohak_srivastava.xlsx")

df["Date"] = pd.to_datetime(df["Date"])
df["Month"] = df["Date"].dt.to_period("M").dt.to_timestamp()
df["Rev_per_Unit"] = df["Revenue"] / df["Units_Sold"].replace(0, np.nan)

st.sidebar.header("Filters")
companies = sorted(df["Company"].unique()); regions = sorted(df["Region"].unique())
sel_companies = st.sidebar.multiselect("Company", companies, default=companies)
sel_regions = st.sidebar.multiselect("Region", regions, default=regions)
f = df[df["Company"].isin(sel_companies) & df["Region"].isin(sel_regions)]

c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("Revenue", f"₹{f['Revenue'].sum():,.0f}")
c2.metric("Units", int(f["Units_Sold"].sum()))
c3.metric("Avg Share %", round(f["Market_Share_%"].mean(),2))
c4.metric("Avg CSAT %", round(f["Customer_Satisfaction_%"].mean(),2))
c5.metric("Avg Rev/Unit", f"₹{f['Rev_per_Unit'].mean():,.0f}")

a,b = st.columns(2)
with a:
    piv = f.pivot_table(index="Region", columns="Company", values="Revenue", aggfunc="sum", fill_value=0)
    st.plotly_chart(px.bar(piv, barmode="group", title="Revenue by Region × Company"), use_container_width=True)
with b:
    mt = f.groupby("Month", as_index=False).agg({"Revenue":"sum","Units_Sold":"sum"})
    st.plotly_chart(px.line(mt, x="Month", y=["Revenue","Units_Sold"], title="Monthly Trend", markers=True), use_container_width=True)

a,b = st.columns(2)
with a:
    cs = f.groupby("Company", as_index=False)["Customer_Satisfaction_%"].mean()
    st.plotly_chart(px.bar(cs, x="Customer_Satisfaction_%", y="Company", orientation="h", title="Avg CSAT by Company"), use_container_width=True)
with b:
    comp = f.groupby("Company", as_index=False).agg(Revenue=("Revenue","sum"), Share=("Market_Share_%","mean"))
    fig = px.scatter(comp, x="Share", y="Revenue", text="Company", title="Market Share vs Revenue")
    fig.update_traces(textposition="top center")
    st.plotly_chart(fig, use_container_width=True)
