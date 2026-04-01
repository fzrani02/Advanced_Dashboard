import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

def render_weekly_tab(df_qty_weekly):
    st.subheader("Quantity and Yield per Week")

    week_order = [f"WW{str(i).zfill(2)}" for i in range(1,53)]
    available_weeks = [w for w in week_order if w in df_qty_weekly.columns]

    customer_week = st.selectbox(
        "Choose Customer",
        sorted(df_qty_weekly["Customer"].unique()),
        key="weekly_customer"
    )
    
    station_list = (
        df_qty_weekly[df_qty_weekly["Customer"] == customer_week]["Station"]
        .dropna()
        .unique()
    )
    
    station_week = st.selectbox(
        "Choose Station",
        sorted(station_list),
        key="weekly_station"
    )
    
    col1, col2 = st.columns(2)

    with col1: 
        week_start = st.selectbox(
            "Week From",
            available_weeks, 
            key="week_start"
        )

    with col2:
        week_end = st.selectbox(
            "Week To",
            available_weeks, 
            index=len(available_weeks)-1,
            key="week_end"
        )
    
    metric = st.selectbox(
        "Choose Metric",
        ["TOTAL QTY", "TOTAL YIELD (%)"],
        key = "weekly_metric"
    )

    start_idx = week_order.index(week_start)
    end_idx = week_order.index(week_end)

    selected_weeks = week_order[start_idx:end_idx+1]
    
    df_station = df_qty_weekly[
        (df_qty_weekly["Customer"] == customer_week) &
        (df_qty_weekly["Station"] == station_week)
    ].copy()

    # ========================
    # HITUNG TOTAL PER PROJECT
    # ========================
    projects = df_station["Project"].unique()
    project_data = []

    for proj in projects:
        df_proj = df_station[df_station["Project"] == proj]

        qty_in = df_proj[df_proj["QTYWeek"] == "QTY IN"][selected_weeks].sum(axis=1).sum()
        qty_pass = df_proj[df_proj["QTYWeek"] == "QTY PASS"][selected_weeks].sum(axis=1).sum()
        qty_fail = df_proj[df_proj["QTYWeek"] == "QTY FAIL"][selected_weeks].sum(axis=1).sum()

        yield_val = (qty_pass / qty_in * 100) if qty_in > 0 else 0

        project_data.append({
            "Project": proj.replace(".xlsx",""),
            "PASS": qty_pass,
            "FAIL": qty_fail,
            "IN": qty_in,
            "YIELD": yield_val
        })

    df_plot = pd.DataFrame(project_data)

    if not df_plot.empty:
        fig, ax= plt.subplots(figsize=(14,6))

        if metric == "TOTAL QTY":
            pass_values = df_plot["PASS"]
            fail_values = df_plot["FAIL"]
            total_values = df_plot["IN"]

            colors = plt.cm.tab20c(range(len(df_plot)))

            ax.barh(
                df_plot["Project"],
                fail_values,
                color="black",
                label="FAIL"
            )
                
            ax.barh(
                df_plot["Project"],
                pass_values,
                left=fail_values,
                color=colors,
                label="PASS"
            )

            for i in range(len(df_plot)):
                if fail_values[i] > 0:
                    ax.text(fail_values[i], i, int(fail_values[i]), va='center')

                if pass_values[i] > 0:
                    ax.text(fail_values[i]+pass_values[i], i, int(pass_values[i]), va='center')

                if total_values[i] > 0:
                    ax.text(total_values[i]+1, i, int(total_values[i]), color='red', fontweight='bold')
            
            ax.set_xlabel("Quantity")
            ax.set_title(f"Total Quantity per Project ({week_start} - {week_end})")
            ax.set_xlim(0, total_values.max()*1.2)
            
        else: 
            colors = plt.cm.tab20c(range(len(df_plot)))
            df_plot["YIELD"] = (df_plot["PASS"] / df_plot["IN"]) * 100

            ax.barh(
                df_plot["Project"],
                df_plot["YIELD"],
                color=colors
            )

            for i, value in enumerate(df_plot["YIELD"]):
                ax.text(
                    value+1, 
                    i, 
                    round(value,2),
                    va='center', 
                    fontsize=8 
                )

            ax.set_xlabel("Yield (%)")
            ax.set_xlim(0,115)
            ax.set_title(f"Total Yield per Project - ({week_start} - {week_end})")
                        
        ax.set_ylabel("Project")
        plt.tight_layout()
        st.pyplot(fig)
        
    else:
        st.warning("No weekly data available.")
