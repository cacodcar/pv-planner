import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


def to_excel(dfs):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in dfs.items():
            df.to_excel(writer, sheet_name=name)
    return output.getvalue()


def run_model(
    dm_fac_vals,
    pv_fac_vals,
    pv_cost_vals,
    battery_cost_vals,
    prod_cap_max,
    stg_cap_max,
    pv_years_val,
    battery_years_val,
):

    from gana import Prg, I, V, P, inf

    p = Prg()

    p.y = I(size=1)
    p.m = I(size=12)

    p.res_cons = I('solar')
    p.res_dem = I('power')
    p.res_stg = I('charge')
    p.res = p.res_cons | p.res_dem | p.res_stg

    p.pro_var = I('pv')
    p.pro_cer = I('li', 'li_d')
    p.pro = p.pro_var | p.pro_cer

    p.dm_fac = P(p.power, p.m, _=dm_fac_vals)
    max_pv_fac = max(pv_fac_vals)
    p.pv_fac = P(p.pv, p.m, _=[p / max_pv_fac for p in pv_fac_vals])
    p.demand = P(p.power, p.m, _=[1] * 12)

    p.capex = P(
        p.pro,
        p.y,
        _=[pv_cost_vals[0] / pv_years_val, battery_cost_vals[0] / battery_years_val, 0],
    )
    p.fopex = P(p.pro, p.y, _=[pv_cost_vals[1], battery_cost_vals[1], 0])
    p.vopex = P(p.pro, p.y, _=[pv_cost_vals[2], battery_cost_vals[2], 0])

    p.cap_p = V(p.pro, p.y, tag='nameplate production capacity')
    p.cap_s = V(p.res_stg, p.y, tag='nameplate storage capacity')
    p.sell = V(p.res_dem, p.m, tag='amount of power sold')
    p.con = V(p.res_cons, p.m, tag='amount of solar consumed')
    p.inv = V(p.res_stg, p.m, tag='charge inventory')
    p.prod = V(p.pro, p.m, tag='production capacity utilization')
    p.ex_cap = V(p.pro, p.y, tag='capital expenditure')
    p.ex_fop = V(p.pro, p.y, tag='fixed operating expenditure')
    p.ex_vop = V(p.pro, p.y, tag='variable operating expenditure')

    p.con_capmax = p.cap_p(p.pro, p.y) <= prod_cap_max
    p.con_capstg = p.cap_s(p.charge, p.y) <= stg_cap_max
    p.con_consmax = p.con(p.res_cons, p.m) <= prod_cap_max * 100
    p.con_sell = p.sell(p.power, p.m) >= p.dm_fac(p.power, p.m) * p.demand(p.power, p.m)

    p.con_pv = p.prod(p.pv, p.m) <= p.pv_fac(p.pv, p.m) * p.cap_p(p.pv, p.y)

    p.con_prod = p.prod(p.pro_cer, p.m) <= p.cap_p(p.pro_cer, p.y)
    p.con_inv = p.inv(p.charge, p.m) <= p.cap_s(p.charge, p.y)

    p.con_vopex = p.ex_vop(p.pro, p.y) == p.vopex(p.pro, p.y) * sum(
        p.prod(p.pro, m) for m in p.m
    )
    p.con_capex = p.ex_cap(p.pro, p.y) == p.capex(p.pro, p.y) * p.cap_p(p.pro, p.y)
    p.con_fopex = p.ex_fop(p.pro, p.y) == p.fopex(p.pro, p.y) * p.cap_p(p.pro, p.y)

    p.con_solar = p.prod(p.pv, p.m) == p.con(p.solar, p.m)
    p.con_power = (
        sum(p.prod(i, p.m) for i in p.pro_var)
        - p.prod(p.li, p.m)
        + p.prod(p.li_d, p.m)
        - p.sell(p.power, p.m)
        == 0
    )
    p.con_charge = (
        p.prod(p.li, p.m)
        - p.prod(p.li_d, p.m)
        + p.inv(p.charge, p.m - 1)
        - p.inv(p.charge, p.m)
        == 0
    )

    p.o = inf(sum(p.ex_cap) + sum(p.ex_vop) + sum(p.ex_fop))

    p.opt()

    return p


st.title("I want solar ☀️")


st.markdown(
    """
Planning your energy needs can be challenging. Residential solar is cool 
and can help you save costs, but how do you know how big your system should be?
Moreover, there is variability in terms of solar availability, energy demand, and costs.

The ChiEAC Energy Planner is here to help you navigate these complexities 
and make informed decisions!

This planner answers questions such as: 

- How much solar capacity should I install?
- How much storage do I need to meet my demand?
- How much will all of this cost? 

What the planner needs from you is information (data). 
You need to input your monthly energy demand, 
how much solar you expect to get each month, 
and the costs of installing and operating your system.
"""
)

st.info("Adjust parameters in the sidebar and click 'Optimize!'.")

months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]
processes = ["Solar", "Battery"]

costs = ["To install", "To maintain", "To run"]

st.subheader("How much energy do you need?")

st.markdown(
    """
First, let us input your household monthly energy demand. 
This is in kWh and can be found on your electricity bill.
"""
)

dm_fac = st.data_editor(
    pd.DataFrame(
        {"Monthly Demand": [600, 500, 400, 300, 200, 100, 50, 100, 200, 300, 400, 500]},
        index=months,
    )
)


st.subheader("What is the weather like?")


st.markdown(
    """
Next, we are going to need weather data. Direct Normal Irradiance (DNI) is a measure of 
the amount of solar energy available at your location. This can be found easily on the internet. 
Look for databases such as the National Solar Radiation Database (NSRDB) 
or use tools like PVGIS to get this information.
"""
)

pv_fac = st.data_editor(
    pd.DataFrame(
        {"Average DNI": [20, 30, 40, 50, 60, 70, 60, 50, 40, 30, 20, 10]}, index=months
    )
)


st.subheader("Let's Talk PV")


st.markdown("How much does my PV cost (per kWh)?")


pv_cost = st.data_editor(pd.DataFrame({"PV Costs": [50, 0.1, 0]}, index=costs))

st.markdown("Is there a limit to your PV capacity?")

cap_limit = st.number_input("Maximum PV Capacity in kWh", value=5000)

st.markdown("How long do you expect your PV system to last?")

pv_years = st.number_input("PV System Lifetime in Years", value=25)


st.subheader("And the battery?")


st.markdown("How much does my battery cost (per kWh)?")


battery_cost = st.data_editor(
    pd.DataFrame({"Battery Costs": [40, 0.2, 0]}, index=costs)
)

st.markdown("Is there a limit to your storage capacity?")

storage_limit = st.number_input("Maximum Storage Capacity in kWh", value=5000)

st.markdown("How long do you expect your battery system to last?")

battery_years = st.number_input("Battery System Lifetime in Years", value=15)


if st.button("Optimize!"):
    p = run_model(
        dm_fac["Monthly Demand"].tolist(),
        pv_fac["Average DNI"].tolist(),
        pv_cost["PV Costs"].tolist(),
        battery_cost["Battery Costs"].tolist(),
        cap_limit,
        storage_limit,
        pv_years,
        battery_years,
    )

    st.title("⚡ Let's see what we have here! ⚡")

    # -------------------------
    # FORMAT OUTPUTS
    # -------------------------
    prod_vals = p.prod.output(aslist=True)
    cap_vals = p.cap_p.output(aslist=True)[:2]  # Only PV and Battery, ignore CER

    processes_d = ["Solar", "Battery Charge", "Battery Discharge"]

    prod_df = pd.DataFrame(
        np.array(prod_vals).reshape(len(processes_d), len(months)),
        index=processes_d,
        columns=months,
    )

    cap_df = pd.DataFrame({"Capacity": cap_vals}, index=processes)

    # -------------------------
    # TOP METRICS
    # -------------------------

    st.header("How much will this cost me?")

    col1, col2 = st.columns(2)

    col1.metric("Annual Cost", round(p.o.output(asfloat=True), 2))

    # -------------------------
    # TABS
    # -------------------------
    tab1, tab2, tab3 = st.tabs(["📊 Summary", "⚙️ Production", "📈 Trends"])

    with tab1:
        st.subheader("Installed Capacity")
        st.dataframe(cap_df, use_container_width=True)

    with tab2:
        st.subheader("Production by Process")
        st.dataframe(prod_df, use_container_width=True)

    with tab3:
        st.subheader("Production Trends")
        st.bar_chart(prod_df.T)

    # -------------------------
    # DOWNLOAD SECTION
    # -------------------------
    st.subheader("⬇️ Download Results")

    col1, col2 = st.columns(2)

    # Excel (multi-sheet)
    excel_file = to_excel(
        {
            "Capacity": cap_df,
            "Production": prod_df,
        }
    )

    col2.download_button(
        "Download Full Excel Report",
        excel_file,
        "results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


st.subheader("About")

st.markdown(
    """
This work was possible due to the generous support provided by the Chicago Education Advocacy Cooperative (ChiEAC).

The back-end is built using Gana, a general purpose algebraic modeling language for multiscale modeling. 
Refer to the resources on cacodcar.com to understand how to construct and optimize such energy systems models.
Models of this nature, at scale, can be built using Energia. 
Both Gana and Energia are open-source and available as python packages.
"""
)