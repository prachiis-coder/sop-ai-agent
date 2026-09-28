import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import optimizer

st.set_page_config(page_title="Steel S&OP Optimization Agent", page_icon="🏭", layout="wide")

st.title("🏭 Steel S&OP / MPS Optimization Agent")
st.caption("Upload the supplied Excel template, edit demand and Active/Inactive choices, then generate an optimized production plan.")

with st.sidebar:
    st.header("How it works")
    st.write("1. Read plant and rolling-mill constraints from Excel")
    st.write("2. Apply the demand and Active/Inactive choices")
    st.write("3. Find a feasible rolling-mill allocation")
    st.write("4. Minimize setup/changeover hours")
    st.write("5. Generate the optimized MPS and capacity report")

uploaded = st.file_uploader("Upload your Excel template", type=["xlsx"])

if uploaded is None:
    st.info("Start by uploading the Excel template you were given (for example, Step3_Production_Allocation.xlsx).")
    st.stop()

with tempfile.TemporaryDirectory() as tmp:
    template_path = Path(tmp) / "template.xlsx"
    template_path.write_bytes(uploaded.getvalue())

    try:
        source = optimizer.parse_template(template_path)
    except Exception as e:
        st.error(f"I could not read this Excel template: {e}")
        st.stop()

    st.success("Template loaded successfully.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("SKUs", len(source.skus))
    c2.metric("Planning months", len(source.months))
    c3.metric("Sinter annual capacity", f"{source.sinter_annual:,.0f} t")
    c4.metric("Blast Furnace annual capacity", f"{source.blast_furnace_annual:,.0f} t")

    st.subheader("1. Demand and Active / Inactive choices")
    st.write("Yellow-style input in the original Excel model is represented here as an editable table. Change the demand or switch a SKU-month to Inactive if required.")

    rows = []
    for sku in source.skus:
        for month in source.months:
            default = optimizer.TEMP_DEMAND.get(sku, [0] * len(source.months))
            try:
                default_demand = float(default[source.months.index(month)])
            except (IndexError, ValueError):
                default_demand = 0.0
            rows.append({"SKU": sku, "Month": month, "Demand": default_demand, "Active": True})

    input_df = pd.DataFrame(rows)
    edited_df = st.data_editor(
        input_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "SKU": st.column_config.TextColumn(disabled=True),
            "Month": st.column_config.TextColumn(disabled=True),
            "Demand": st.column_config.NumberColumn(min_value=0, step=1),
            "Active": st.column_config.CheckboxColumn("Active?"),
        },
        key="demand_editor",
    )

    run = st.button("🚀 Run Optimization", type="primary", use_container_width=True)

    if run:
        demand = {sku: {month: 0.0 for month in source.months} for sku in source.skus}
        active = {sku: {month: True for month in source.months} for sku in source.skus}

        for _, row in edited_df.iterrows():
            sku = str(row["SKU"])
            month = str(row["Month"])
            demand[sku][month] = max(0.0, float(row["Demand"]))
            active[sku][month] = bool(row["Active"])

        try:
            with st.status("Running optimization...", expanded=True) as status:
                st.write("✓ Reading capacities and setup constraints")
                st.write("✓ Checking monthly upstream capacity")
                st.write("✓ Testing feasible Rolling Mill 1 / Rolling Mill 2 allocations")
                st.write("✓ Minimizing setup/changeover hours")
                plans, campaigns, capacity_rows = optimizer.run_optimization(source, demand, active)
                st.write("✓ Building the optimized MPS and capacity requirements")
                status.update(label="Optimization completed", state="complete", expanded=False)
        except Exception as e:
            st.error(f"Optimization could not be completed: {e}")
            st.stop()

        total_demand = sum(p.effective_demand for p in plans)
        total_production = sum(p.planned_production for p in plans)
        total_unmet = sum(p.unmet_demand for p in plans)
        total_setup = sum(c.setup_hours for c in campaigns)
        bf_capacity = source.blast_furnace_annual / 2.0

        st.subheader("2. Optimization Results")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Effective demand", f"{total_demand:,.1f} t")
        m2.metric("Planned production", f"{total_production:,.1f} t")
        m3.metric("Unmet demand", f"{total_unmet:,.1f} t")
        m4.metric("Setup time", f"{total_setup:,.1f} h")

        st.subheader("Optimized MPS")
        mps_df = pd.DataFrame([
            {
                "Month": p.month,
                "SKU": p.sku,
                "Demand": p.demand,
                "Active": "Y" if p.active else "N",
                "Effective Demand": p.effective_demand,
                "RM1 Qty": p.rm1_qty,
                "RM2 Qty": p.rm2_qty,
                "Planned Production": p.planned_production,
                "Unmet Demand": p.unmet_demand,
            }
            for p in plans
        ])
        st.dataframe(mps_df, use_container_width=True, hide_index=True)

        st.subheader("Capacity and setup")
        cap_df = pd.DataFrame(capacity_rows)
        st.dataframe(cap_df, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            st.write("**Rolling Mill load by month**")
            chart_df = cap_df.set_index("Month")[["RM1 Normalized Load", "RM2 Normalized Load"]]
            st.bar_chart(chart_df)
        with col2:
            st.write("**Upstream utilization by month**")
            util_df = cap_df.set_index("Month")[["Sinter Utilization", "Blast Furnace Utilization"]]
            st.line_chart(util_df)

        st.subheader("Campaign / sequence plan")
        campaign_df = pd.DataFrame([
            {
                "Month": c.month,
                "Mill": c.mill,
                "Active SKUs": ", ".join(c.active_skus) or "No production",
                "Recommended Sequence": " → ".join(c.sequence) or "No production",
                "Setup Hours": c.setup_hours,
                "Normalized Load": c.normalized_load,
            }
            for c in campaigns
        ])
        st.dataframe(campaign_df, use_container_width=True, hide_index=True)

        output_path = Path(tmp) / "SOP_MPS_Optimized_Output.xlsx"
        optimizer.write_output(
            output_path,
            source,
            template_path,
            demand,
            active,
            {},
            plans,
            campaigns,
            capacity_rows,
        )
        optimizer.validate_output(output_path)

        st.subheader("3. Download the Excel result")
        st.download_button(
            "⬇️ Download Optimized MPS Excel",
            data=output_path.read_bytes(),
            file_name="SOP_MPS_Optimized_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.caption("Raw-material MRP is not calculated unless BOM/consumption factors are supplied, because the original model does not invent those factors.")
