import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import io

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
    
    filtered_df = df[
        (df["Date Time"].dt.date >= start_date.date()) &
        (df["Date Time"].dt.date <= end_date.date())
    ].copy()
    

    try:
        third_col = df.columns[2]
        filtered_df["Modified value"] = constant_value - filtered_df[third_col]
    except Exception as e:
        raise ValueError(f"Error processing the third column. Please ensure the Excel file has a valid third column. Detail: {e}")
    
    return filtered_df



def select_tides(df, value_col, n, time_col, min_gap_hours=24, high=True):

    df_sorted = df.sort_values(by=value_col, ascending=not high)
    selected_rows = []
    
    for idx, row in df_sorted.iterrows():
        candidate_time = row[time_col]
        if all(abs(candidate_time - sel[time_col]) >= timedelta(hours=min_gap_hours) 
               for sel in selected_rows):
            selected_rows.append(row)
        if len(selected_rows) == n:
            break
            
    return pd.DataFrame(selected_rows).sort_values(by=time_col)

def plot_tide_results(df, high_tides, low_tides):

    fig, ax = plt.subplots(figsize=(12, 6))
    
    ax.plot(df["Date Time"], df["Modified value"], color="blue", label="Tide Level", linewidth=2)
    
    ax.scatter(high_tides["Date Time"], high_tides["Modified value"],
               color="red", marker="^", s=100, label="High Tide (Crest)")
    
    ax.scatter(low_tides["Date Time"], low_tides["Modified value"],
               color="green", marker="v", s=100, label="Low Tide (Trough)")
    
    ax.set_xlabel("Date Time")
    ax.set_ylabel("Tide Level")
    ax.set_title("Tide Graph")
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

                    high_tides = select_tides(processed_df, value_col="Modified value", n=3, time_col="Date Time",
                                              min_gap_hours=24, high=True)
                    low_tides = select_tides(processed_df, value_col="Modified value", n=3, time_col="Date Time",
                                             min_gap_hours=24, high=False)
                    
                    st.subheader("Tide Analysis Results")
                    st.markdown("**High Tides (Crests):**")
                    st.dataframe(high_tides)
                    
                    st.markdown("**Low Tides (Troughs):**")
                    st.dataframe(low_tides)
                    
                    # Create and display the scatter plot.
                    fig = plot_tide_results(processed_df, high_tides, low_tides)
                    st.pyplot(fig)
                    
                except Exception as e:
                    st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
