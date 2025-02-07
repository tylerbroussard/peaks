import streamlit as st
import pandas as pd
import altair as alt


def main():
    st.title("Peaks Dashboard")
    st.subheader("Data from peaks.xlsx")
    
    try:
        # Load the Excel file from the 'peaks' directory
        data_path = "peaks.xlsx"
        df = pd.read_excel(data_path)
        
        st.write("### Data Overview")
        st.dataframe(df)
        
        # If the dataframe contains at least two numeric columns, display a line chart
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
