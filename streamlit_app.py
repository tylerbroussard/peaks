import streamlit as st
import pandas as pd
import altair as alt
import importlib.util


def check_openpyxl():
    if importlib.util.find_spec("openpyxl") is None:
        st.error("Missing optional dependency 'openpyxl'. Please install it via pip: pip install openpyxl")
        st.stop()


def main():
    st.title("Peaks Dashboard")
    st.subheader("Data from peaks.xlsx")

    check_openpyxl()
    
    try:
        # Load the Excel file from the same directory
        data_path = "peaks.xlsx"
        df = pd.read_excel(data_path)
        
        st.write("### Data Overview")
        st.dataframe(df)
        
        # Plot a simple line chart if there are at least two numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        if len(numeric_cols) >= 2:
            st.write("### Sample Plot")
            chart = alt.Chart(df).mark_line().encode(
                x=numeric_cols[0],
                y=numeric_cols[1]
            )
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Not enough numeric columns available for a basic chart.")
    except Exception as e:
        st.error(f"An error occurred while loading data: {e}")


if __name__ == "__main__":
    main()
