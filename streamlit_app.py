import streamlit as st
import pandas as pd
import altair as alt
import importlib.util


def check_openpyxl():
    if importlib.util.find_spec("openpyxl") is None:
        st.error("Missing optional dependency 'openpyxl'. Please install it via pip: pip install openpyxl")
        st.stop()


def main():
    st.title("Agent and Call Peak Analysis")
    st.subheader("Daily Peak Analysis")

    check_openpyxl()
    
    try:
        # Load the Excel file
        data_path = "peaks.xlsx"
        df = pd.read_excel(data_path)
        
        # Convert date column with specific format
        df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')
        
        # Ensure we have valid data
        if df.empty:
            st.error("No data found in the Excel file.")
            return
            
        # Data Overview
        st.write("### Data Overview")
        st.dataframe(df)
        
        # Create tabs for different visualizations
        tab1, tab2 = st.tabs(["Agents Peak", "Calls Peak Analysis"])
        
        with tab1:
            st.write("### Agent Peak Values Over Time")
            
            # Create the agents peak chart
            agents_chart = alt.Chart(df).mark_line(
                point=True,
                color='#1f77b4'
            ).encode(
                x=alt.X('Date:T', title='Date'),
                y=alt.Y('Agents Peak:Q', title='Number of Agents'),
                tooltip=['Date:T', 'Agents Peak:Q']
            ).properties(
                width=800,
                height=400,
                title="Daily Agent Peak Values"
            ).interactive()
            
            st.altair_chart(agents_chart, use_container_width=True)
            
            # Show agent peak statistics
            st.write("### Agent Peak Statistics")
            agent_stats = df['Agents Peak'].describe()
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Average Agents", f"{agent_stats['mean']:.0f}")
            with col2:
                st.metric("Maximum Agents", f"{agent_stats['max']:.0f}")
            with col3:
                st.metric("Minimum Agents", f"{agent_stats['min']:.0f}")
            with col4:
                st.metric("Standard Deviation", f"{agent_stats['std']:.1f}")
        
        with tab2:
            st.write("### Call Peak Analysis")
            
            # Create a combined chart for all call types
            call_data = pd.melt(
                df,
                id_vars=['Date'],
                value_vars=['Calls Peak', 'Calls Peak (Inbound)', 'Calls Peak (Outbound)'],
                var_name='Call Type',
                value_name='Peak Value'
            )
            
            calls_chart = alt.Chart(call_data).mark_line(point=True).encode(
                x=alt.X('Date:T', title='Date'),
                y=alt.Y('Peak Value:Q', title='Number of Calls'),
                color=alt.Color('Call Type:N', title='Type'),
                tooltip=['Date:T', 'Call Type:N', 'Peak Value:Q']
            ).properties(
                width=800,
                height=400,
                title="Daily Call Peak Values by Type"
            ).interactive()
            
            st.altair_chart(calls_chart, use_container_width=True)
            
            # Show call peak statistics
            st.write("### Call Peak Statistics")
            
            col1, col2, col3 = st.columns(3)
            metrics = ['Calls Peak', 'Calls Peak (Inbound)', 'Calls Peak (Outbound)']
            cols = [col1, col2, col3]
            
            for col, metric in zip(cols, metrics):
                with col:
                    st.write(f"**{metric}**")
                    stats = df[metric].describe()
                    st.write(f"Average: {stats['mean']:.0f}")
                    st.write(f"Maximum: {stats['max']:.0f}")
                    st.write(f"Minimum: {stats['min']:.0f}")

    except Exception as e:
        st.error(f"An error occurred while loading data: {e}")
        st.write("Please check that:")
        st.write("1. The 'Date' column contains dates in MM/DD/YYYY format")
        st.write("2. There are no missing or invalid values")


if __name__ == "__main__":
    main()
