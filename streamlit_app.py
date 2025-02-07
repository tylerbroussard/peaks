import streamlit as st
import pandas as pd
import altair as alt
import importlib.util


def check_openpyxl():
    if importlib.util.find_spec("openpyxl") is None:
        st.error("Missing optional dependency 'openpyxl'. Please install it via pip: pip install openpyxl")
        st.stop()


def main():
    st.title("Agent Peak Analysis Dashboard")
    st.subheader("Daily Peak Values")

    check_openpyxl()
    
    try:
        # Load the Excel file
        data_path = "peaks.xlsx"
        df = pd.read_excel(data_path)
        
        # Ensure date column is in datetime format
        date_col = df.select_dtypes(include=['datetime64']).columns[0]
        
        # Data Overview
        st.write("### Data Overview")
        st.dataframe(df)
        
        # Time series visualization
        st.write("### Peak Values Over Time")
        
        # Get numeric columns (excluding date)
        numeric_cols = df.select_dtypes(include=['number']).columns
        
        # Create a selector for which peak value to display
        selected_peak = st.selectbox("Select Peak Value to Display:", numeric_cols)
        
        # Create the time series chart
        chart = alt.Chart(df).mark_line(point=True).encode(
            x=alt.X(date_col, title='Date'),
            y=alt.Y(selected_peak, title='Peak Value'),
            tooltip=[date_col, selected_peak]
        ).properties(
            width=800,
            height=400,
            title=f"Daily {selected_peak} Peak Values"
        ).interactive()
        
        st.altair_chart(chart, use_container_width=True)
        
        # Show summary statistics
        st.write("### Summary Statistics")
        stats = df[selected_peak].describe()
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Average Peak", f"{stats['mean']:.2f}")
        with col2:
            st.metric("Maximum Peak", f"{stats['max']:.2f}")
        with col3:
            st.metric("Minimum Peak", f"{stats['min']:.2f}")
        with col4:
            st.metric("Standard Deviation", f"{stats['std']:.2f}")

    except Exception as e:
        st.error(f"An error occurred while loading data: {e}")


if __name__ == "__main__":
    main()
