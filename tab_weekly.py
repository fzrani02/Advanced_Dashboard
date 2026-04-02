import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

def render_weekly_tab(df_qty_weekly, df_weekly_detail):
    st.subheader("Quantity and Yield per Week")

    week_order = [f"WW{str(i).zfill(2)}" for i in range(1,53)]
    
    available_weeks = [w for w in week_order if w in df_qty_detail["Week"].unique()]

    customer_week = st.selectbox(
        "Choose Customer",
        sorted(df_qty_detail["Customer"].unique()),
        key="weekly_customer"
    )
    
    station_list = (
        df_qty_weekly[df_qty_detail["Customer"] == customer_week]["Station"]
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

    # =============
    # FILTER DATA SHEET 6
    # =============
    df_plot = df_weekly_detail[
        (df_weekly_detail["Customer"] == customer_week) &
        (df_weekly_detail["Station"] == station_week) &
        (df_weekly_detail["Week"].isin(selected_weeks))
    ].copy()

    df_plot['Week_Cat'] = pd.Categorical(df_plot['Week'], categories=week_order, ordered=True)
    df_plot = df_plot.sort_values('Week_Cat').reset_index(drop=True)

    if not df_plot.empty:
        fig, ax = plt.subplots(figsize=(14, max(6, len(df_plot) * 0.6)))

        if metric == "TOTAL QTY":
            pass_values = df_plot["TOTAL QTY PASS"]
            fail_values = df_plot["TOTAL QTY FAIL"]
            total_values = df_plot["TOTAL QTY IN"]
            y_labels = df_plot["Week"]

            #unique_projects = df_plot["Project"].unique()

            #colors = plt.cm.tab20c(range(len(unique_projects)))

            #color_map_project = {
                #proj: colors[i]
                #for i, proj in enumerate(unique_projects)
            #}

            #pass_colors = df_plot["Project"].map(color_map_project)

            ax.barh(
                y_labels,
                #df_plot["Project"],
                fail_values,
                color="black",
                label="QTY FAIL"
            )
                
            ax.barh(
                y_labels, 
                #df_plot["Project"],
                pass_values,
                left=fail_values,
                #color=pass_colors,
                color="darkblue", 
                label="QTY PASS"
            )

            n_bars = len(df_plot)
            base_size = max(8, 12 - int(n_bars * 0.2))

            fail_size = base_size - 1
            pass_size = base_size 
            total_size = base_size + 2

            extra = total_values.max() * 0.002
            base_offset = total_values.max() *0.005
            char_width = total_values.max() * 0.020

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
                        fontsize=pass_size, fontweight='bold', color='darkblue'
                    )

                # LABEL TOTAL 
                if total_val >0 :
                    digit = len(str(int(pass_val))) if pass_val > 0 else 0
                    dynamic_offset = base_offset + (digit * char_width) + (total_values.max() * 0.007)
                    ax.text(
                        total_val + dynamic_offset, 
                        i, 
                        int(total_val),
                        ha='left', va='center', 
                        fontsize=total_size, fontweight='bold', color='red'
                    )

            legend_elements = [
                Patch(facecolor="red", label="QTY IN (TOTAL)"),
                Patch(facecolor="darkblue", label="QTY PASS"),
                Patch(facecolor="black", label="QTY FAIL")
            ]
             
            ax.legend(handles=legend_elements, title="Project", bbox_to_anchor=(1.02, 1), loc="upper left")
            
            ax.set_xlabel("Quantity")
            ax.set_title(f"{station_week} - Weekly Total Quantity ({week_start} - {week_end})", fontsize=14, fontweight="bold")
            ax.set_xlim(0, total_values.max()*1.4 if total_values.max() > 0 else 10)
            
        else:
            y_labels = df_plot["Week"]
            yield_values = df_plot["TOTAL YIELD (%)"]

            

            ###
            #colors = plt.cm.tab20c(range(len(df_plot)))
            #df_plot["YIELD"] = (df_plot["PASS"] / df_plot["IN"]) * 100

            ax.barh(
                y_labels, 
                yield_values, 
                color="darkblue"
            )

            for i, value in enumerate(yield_values):
                ax.text(
                    value+1, 
                    i, 
                    f"{round(value,2)}%",
                    va='center', 
                    fontsize=10
                )

            ax.set_xlabel("Yield (%)")
            ax.set_xlim(0,115)
            ax.set_title(f"{station_week} - Weekly Total Yield ({week_start} - {week_end})", fontsize=14, fontweight="bold")
                        
        ax.set_ylabel("Week")
        ax.invert_yaxis()
        
        plt.tight_layout()
        st.pyplot(fig)
        
    else:
        st.warning("No weekly data available.")
