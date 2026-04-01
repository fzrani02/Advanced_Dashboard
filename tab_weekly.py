import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

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

            unique_projects = df_plot["Project"].unique()

            colors = plt.cm.tab20c(range(len(unique_projects)))

            color_map_project = {
                proj: colors[i]
                for i, proj in enumerate(unique_projects)
            }

            pass_colors = df_plot["Project"].map(color_map_project)

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
                color=pass_colors,
                label="PASS"
            )

            n_bars = len(df_plot)
            base_size = max(6, 12 - int(n_bars * 0.4))

            fail_size = base_size - 1
            pass_size = base_size 
            total_size = base_size + 3

            extra = total_values.max() * 0.002
            base_offset = total_values.max() *0.005
            char_width = total_values.max() * 0.012

            for i in range(n_bars):
                fail_val = fail_values.iloc[i] 
                pass_val = pass_values.iloc[i]
                total_val = total_values.iloc[i] 

                # LABEL FAIL (HITAM)
                if fail_val > 0:
                    ax.text(
                        fail_val + extra,
                        i, 
                        int(fail_val),
                        ha='left', va='center',
                        fontsize=fail_size, fontweight='bold', color='black'
                    )

                # LABEL PASS
                if pass_val > 0:
                    ax.text(
                        (fail_val + pass_val) + base_offset,
                        i,
                        int(pass_val),
                        ha='left', va='center',
                        fontsize=pass_size, fontweight='bold', color='blue'
                    )

                # LABEL TOTAL 
                if total_val >0 :
                    digit = len(str(int(pass_val))) if pass_val > 0 else 0
                    dynamic_offset = base_offset + (digit * char_width) + (total_values.max() * 0.005)
                    ax.text(
                        total_val + dynamic_offset, 
                        i, 
                        int(total_val),
                        ha='left', va='center', 
                        fontsize=total_size, fontweight='bold', color='red'
                    )

            #legend_elements = [
                #Patch(facecolor=color_map_project[proj], label=proj)
                #for proj in unique_projects
            #]
            legend_elements.append(Patch(facecolor="red", label="TOTAL"))
            legend_elements.append(Patch(facecolor="blue", label="QTY PASS"))
            legend_elements.append(Patch(facecolor="black", label="QTY FAIL"))   
            ax.legend(handles=legend_elements, title="Project", bbox_to_anchor=(1.02, 1), loc="upper left")
            
            ax.set_xlabel("Quantity")
            ax.set_title(f"Total Quantity (QTY Fail+ QTY Pass) per Project ({week_start} - {week_end})")
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
