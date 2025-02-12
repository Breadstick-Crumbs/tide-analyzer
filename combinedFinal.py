import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import io
import numpy as np

def process_excel(file, start_date_str, end_date_str, constant_value):
    df = pd.read_excel(file)
    
    date_col = df.columns[0]
    df.rename(columns={date_col: "Date Time"}, inplace=True)
    
    try:
        df["Date Time"] = pd.to_datetime(df["Date Time"], format="%m/%d/%Y %H:%M:%S")
    except Exception as e:
        raise ValueError(f"Error converting dates in Excel file. Please check that the date column is in the format 'mm/dd/yyyy HH:MM:SS'. Detail: {e}")
    
    try:
        start_date_str = start_date_str.strip()
        end_date_str = end_date_str.strip()
        start_date = datetime.strptime(start_date_str, "%m/%d/%Y")
        end_date = datetime.strptime(end_date_str, "%m/%d/%Y")
    except ValueError:
        raise ValueError("Invalid start/end date format. Please ensure you use the format mm/dd/yyyy exactly. Example: 08/21/2023")
    
    filtered_df = df[(df["Date Time"].dt.date >= start_date.date()) & (df["Date Time"].dt.date <= end_date.date())].copy()
    
    try:
        third_col = df.columns[2]
        filtered_df["Modified value"] = constant_value - filtered_df[third_col]
    except Exception as e:
        raise ValueError(f"Error processing the third column. Please ensure the Excel file has a valid third column. Detail: {e}")
    
    filtered_df.sort_values(by="Date Time", inplace=True)
    filtered_df.reset_index(drop=True, inplace=True)
    
    return filtered_df

def get_extreme_in_window(df, base_time, direction, tide_type, gap_hours=6, tolerance_hours=2):

    if direction == "forward":
        start_time = base_time + timedelta(hours=gap_hours - tolerance_hours)
        end_time = base_time + timedelta(hours=gap_hours + tolerance_hours)
    elif direction == "backward":
        start_time = base_time - timedelta(hours=gap_hours + tolerance_hours)
        end_time = base_time - timedelta(hours=gap_hours - tolerance_hours)
    else:
        raise ValueError("direction must be 'forward' or 'backward'")
    
    candidate_df = df[(df["Date Time"] >= start_time) & (df["Date Time"] <= end_time)]
    if candidate_df.empty:
        return None
    
    if tide_type == "high":
        extreme_row = candidate_df.loc[candidate_df["Modified value"].idxmax()]
    elif tide_type == "low":
        extreme_row = candidate_df.loc[candidate_df["Modified value"].idxmin()]
    else:
        raise ValueError("tide_type must be 'high' or 'low'")
    
    return extreme_row

def select_tide_chain(df, gap_hours=6, tolerance_hours=2):

    absolute_low = df["Modified value"].min()
    
    candidates = df.sort_values(by="Modified value", ascending=False)
    main_high = None
    main_low = None
    
    for _, candidate in candidates.iterrows():
        candidate_time = candidate["Date Time"]
        candidate_low = get_extreme_in_window(df, candidate_time, direction="forward", tide_type="low",
                                               gap_hours=gap_hours, tolerance_hours=tolerance_hours)
        if candidate_low is not None and np.isclose(candidate_low["Modified value"], absolute_low, atol=1e-6):
            main_high = candidate
            main_low = candidate_low
            break
    
    if main_high is None:
        main_high = candidates.iloc[0]
        main_low = get_extreme_in_window(df, main_high["Date Time"], direction="forward", tide_type="low",
                                         gap_hours=gap_hours, tolerance_hours=tolerance_hours)
    
    
    L1 = main_low
    H2 = None
    L3 = None
    if L1 is not None:
        H2 = get_extreme_in_window(df, L1["Date Time"], direction="forward", tide_type="high",
                                   gap_hours=gap_hours, tolerance_hours=tolerance_hours)
        if H2 is not None:
            L3 = get_extreme_in_window(df, H2["Date Time"], direction="forward", tide_type="low",
                                       gap_hours=gap_hours, tolerance_hours=tolerance_hours)
    
    L_minus1 = get_extreme_in_window(df, main_high["Date Time"], direction="backward", tide_type="low",
                                     gap_hours=gap_hours, tolerance_hours=tolerance_hours)
    H_minus2 = None
    L_minus3 = None
    if L_minus1 is not None:
        H_minus2 = get_extreme_in_window(df, L_minus1["Date Time"], direction="backward", tide_type="high",
                                         gap_hours=gap_hours, tolerance_hours=tolerance_hours)
        if H_minus2 is not None:
            L_minus3 = get_extreme_in_window(df, H_minus2["Date Time"], direction="backward", tide_type="low",
                                             gap_hours=gap_hours, tolerance_hours=tolerance_hours)
    
    high_tides = [event for event in [H_minus2, main_high, H2] if event is not None]
    low_tides  = [event for event in [L_minus3, L_minus1, L1, L3] if event is not None]
    
    high_tides_df = pd.DataFrame(high_tides).sort_values(by="Date Time")
    low_tides_df = pd.DataFrame(low_tides).sort_values(by="Date Time")
    
    return {"high_tides": high_tides_df, "low_tides": low_tides_df}

def plot_tide_results(df, high_tides_df, low_tides_df):
    fig, ax = plt.subplots(figsize=(12, 6))
    
    ax.plot(df["Date Time"], df["Modified value"], color="blue", label="Tide Level", linewidth=2)
    
    if not high_tides_df.empty:
        ax.scatter(high_tides_df["Date Time"], high_tides_df["Modified value"],
                   color="red", marker="^", s=100, label="High Tide")
    
    if not low_tides_df.empty:
        ax.scatter(low_tides_df["Date Time"], low_tides_df["Modified value"],
                   color="green", marker="v", s=100, label="Low Tide")
    
    ax.set_xlabel("Date Time")
    ax.set_ylabel("Tide Level")
    ax.set_title("Tide Graph with Selected Extremes")
    ax.legend()
    ax.grid(True)
    
    return fig

def main():
    st.title("Excel Data Processor and Tide Analyzer")
    
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
    if uploaded_file is not None:
        st.write("File uploaded successfully!")
        
        start_date_input = st.text_input("Enter the start date (mm/dd/yyyy):")
        end_date_input = st.text_input("Enter the end date (mm/dd/yyyy):")
        constant_value = st.number_input("Enter the constant value:", value=0.0, format="%.2f")
        
        if st.button("Process File"):
            if not start_date_input or not end_date_input:
                st.error("Please enter both start and end dates in the format mm/dd/yyyy.")
            else:
                try:
                    processed_df = process_excel(uploaded_file, start_date_input, end_date_input, constant_value)
                    st.success("File processed successfully!")
                    
                    st.subheader("Processed Excel Data")
                    st.dataframe(processed_df)
                    
                    output = io.BytesIO()
                    processed_df.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)
                    st.download_button(
                        label="Download Processed Excel",
                        data=output,
                        file_name="modified_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    chain = select_tide_chain(processed_df, gap_hours=6, tolerance_hours=2)
                    high_tides_df = chain["high_tides"]
                    low_tides_df = chain["low_tides"]
                    
                    st.subheader("Tide Analysis Results")
                    st.markdown("**High Tides (Selected):**")
                    st.dataframe(high_tides_df)
                    
                    st.markdown("**Low Tides (Selected):**")
                    st.dataframe(low_tides_df)
                    
                    fig = plot_tide_results(processed_df, high_tides_df, low_tides_df)
                    st.pyplot(fig)
                    
                except Exception as e:
                    st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
