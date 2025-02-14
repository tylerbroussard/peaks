import streamlit as st
import pandas as pd
import altair as alt
import importlib.util
import re


def check_openpyxl():
    if importlib.util.find_spec("openpyxl") is None:
        st.error("Missing optional dependency 'openpyxl'. Please install it via pip: pip install openpyxl")
        st.stop()


def is_date_format(x):
    # Check if string matches MM/DD/YYYY format
    return bool(re.match(r'\d{2}/\d{2}/\d{4}$', str(x)))


def main():
    st.title("Agent and Call Peaks")

    check_openpyxl()
    
    try:
        # Load the Excel file
        data_path = "peaks.xlsx"
        df = pd.read_excel(data_path)
        
        # Filter out non-date rows and convert to datetime
        df = df[df['Date'].apply(is_date_format)].copy()
        df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')
        
        # Calculate and display date range
        min_date = df['Date'].min().strftime('%m/%d/%Y')
        max_date = df['Date'].max().strftime('%m/%d/%Y')
        st.write(f'**Data Date Range:** {min_date} to {max_date}')
        
        # Ensure we have valid data
        if df.empty:
            st.error("No valid data found in the Excel file.")
            return
            
        # Data Overview
        st.write("### Data")
        st.dataframe(df)
        
        # Create tabs for different visualizations
        tab1, tab2 = st.tabs(["Agent Peaks", "Call Peaks"])
        
        with tab1:
            st.write("### Agent Peaks")
            
            # Create a selection that chooses the nearest point & selects based on x-position
            nearest = alt.selection_point(nearest=True, on='mouseover',
                                        fields=['Date'], empty=False)

            # Create the base line
            line = alt.Chart(df).mark_line(color='#1f77b4').encode(
                x=alt.X('Date:T', title='Date'),
                y=alt.Y('Agents Peak:Q', title='Number of Agents')
            )

            # Create the points
            points = line.mark_point(size=100).encode(
                opacity=alt.value(0)  # make the points transparent
            ).add_params(
                nearest
            )

            # Create a rule that highlights the date at the x-position of the cursor
            rules = alt.Chart(df).mark_rule(color='gray').encode(
                x='Date:T'
            ).transform_filter(
                nearest
            )

            # Create text labels that appear on hover
            text = line.mark_text(
                align='left',
                dx=5,
                dy=-5,
                fontSize=12
            ).encode(
                text=alt.condition(nearest, 'Agents Peak:Q', alt.value(' '), format='.0f')
            ).transform_filter(
                nearest
            )

            # Combine all the layers
            agents_chart = alt.layer(
                line,
                points,
                rules,
                text
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
            
            # Create a combined chart for all call types with improved interaction
            call_data = pd.melt(
                df,
                id_vars=['Date'],
                value_vars=['Calls Peak', 'Calls Peak (Inbound)', 'Calls Peak (Outbound)'],
                var_name='Call Type',
                value_name='Peak Value'
            )
            
            # Create selection for calls chart
            nearest_calls = alt.selection_point(nearest=True, on='mouseover',
                                              fields=['Date'], empty=False)

            # Base lines for calls
            calls_lines = alt.Chart(call_data).mark_line().encode(
                x=alt.X('Date:T', title='Date'),
                y=alt.Y('Peak Value:Q', title='Number of Calls'),
                color=alt.Color('Call Type:N', title='Type')
            )

            # Points for calls
            calls_points = calls_lines.mark_point(size=100).encode(
                opacity=alt.value(0)
            ).add_params(
                nearest_calls
            )

            # Rules for calls
            calls_rules = alt.Chart(call_data).mark_rule(color='gray').encode(
                x='Date:T'
            ).transform_filter(
                nearest_calls
            )

            # Text for calls
            calls_text = calls_lines.mark_text(
                align='left',
                dx=5,
                dy=-5,
                fontSize=12
            ).encode(
                text=alt.condition(nearest_calls, 'Peak Value:Q', alt.value(' '), format='.0f')
            ).transform_filter(
                nearest_calls
            )

            # Combine all layers for calls chart
            calls_chart = alt.layer(
                calls_lines,
                calls_points,
                calls_rules,
                calls_text
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
        st.write("Please check that the data is in the correct format:")
        st.write("1. Dates should be in MM/DD/YYYY format")
        st.write("2. Only numeric values for peak counts")


if __name__ == "__main__":
    main()
