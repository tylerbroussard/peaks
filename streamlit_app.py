import streamlit as st
import pandas as pd
import altair as alt
import importlib.util
import re
import datetime
from zoneinfo import ZoneInfo


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
        
        # Calculate date range and last updated info
        min_date = df['Date'].min().strftime('%m/%d/%Y')
        max_date = df['Date'].max().strftime('%m/%d/%Y')
        est_time = datetime.datetime.now(ZoneInfo("America/New_York")).strftime('%m/%d/%Y %I:%M %p EST')
        
        # Create a visually distinct data information section
        st.markdown("---")
        st.markdown("<h2 style='text-align: center; color: #1f77b4;'>Data Information</h2>", unsafe_allow_html=True)
        
        # Create two columns for better layout
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("<div style='background-color: #f0f2f6; padding: 20px; border-radius: 10px;'>"
                       "<h3 style='color: #1f77b4;'>Data Date Range</h3>"
                       f"<p style='font-size: 18px;'>{min_date} to {max_date}</p>"
                       "</div>", unsafe_allow_html=True)
            
        with col2:
            st.markdown("<div style='background-color: #f0f2f6; padding: 20px; border-radius: 10px;'>"
                       "<h3 style='color: #1f77b4;'>Date Last Updated</h3>"
                       f"<p style='font-size: 18px;'>{est_time}</p>"
                       "</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        
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
            mean_agents = agent_stats['mean']
            std_dev = agent_stats['std']
            median_agents = agent_stats['50%']
            
            # Calculate the range where approximately 68% of values fall
            lower_range = mean_agents - std_dev
            upper_range = mean_agents + std_dev
            
            # Calculate how many days fall outside this range
            days_outside_range = len(df[(df['Agents Peak'] < lower_range) | (df['Agents Peak'] > upper_range)])
            total_days = len(df)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("Median Daily Peak", f"{median_agents:.0f} agents")
                st.metric("Maximum Peak", f"{agent_stats['max']:.0f} agents")
                st.metric("Minimum Peak", f"{agent_stats['min']:.0f} agents")
            
            with col2:
                st.markdown(f"""
                    <div style='background-color: #f0f2f6; padding: 15px; border-radius: 10px;'>
                        <h4 style='color: #1f77b4; margin-top: 0;'>Peak Staffing Patterns</h4>
                        <p style='font-size: 16px;'>
                        • Typical range: <strong>{lower_range:.0f}</strong> to <strong>{upper_range:.0f}</strong> agents<br>
                        • Average (mean): <strong>{mean_agents:.0f}</strong> agents<br>
                        • {days_outside_range} out of {total_days} days fall outside this range</p>
                        <p style='font-size: 14px; color: #666;'>
                        Note: The wide range suggests significant variation in peak staffing needs. 
                        Consider checking specific days of the week or times of year for patterns.</p>
                    </div>
                """, unsafe_allow_html=True)
        
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
            
            # Calculate statistics for each call type
            call_types = {
                'Total Calls': 'Calls Peak',
                'Inbound Calls': 'Calls Peak (Inbound)',
                'Outbound Calls': 'Calls Peak (Outbound)'
            }
            
            for call_type_display, call_type_col in call_types.items():
                stats = df[call_type_col].describe()
                mean_calls = stats['mean']
                std_dev = stats['std']
                median_calls = stats['50%']
                
                # Calculate typical range
                lower_range = mean_calls - std_dev
                upper_range = mean_calls + std_dev
                
                # Calculate days outside range
                days_outside = len(df[(df[call_type_col] < lower_range) | (df[call_type_col] > upper_range)])
                
                st.markdown(f"""
                    <div style='background-color: #f0f2f6; padding: 15px; border-radius: 10px; margin-bottom: 15px;'>
                        <h4 style='color: #1f77b4; margin-top: 0;'>{call_type_display} Peak Patterns</h4>
                        <div style='display: flex; justify-content: space-between;'>
                            <div style='flex: 1;'>
                                <p style='font-size: 16px; margin: 5px 0;'>
                                • Median Peak: <strong>{median_calls:.0f}</strong> calls<br>
                                • Average Peak: <strong>{mean_calls:.0f}</strong> calls<br>
                                • Range: <strong>{stats['min']:.0f}</strong> to <strong>{stats['max']:.0f}</strong> calls
                                </p>
                            </div>
                            <div style='flex: 1;'>
                                <p style='font-size: 16px; margin: 5px 0;'>
                                • Typical Range: <strong>{lower_range:.0f}</strong> to <strong>{upper_range:.0f}</strong> calls<br>
                                • Standard Deviation: ±<strong>{std_dev:.1f}</strong> calls<br>
                                • <strong>{days_outside}</strong> out of <strong>{len(df)}</strong> days outside typical range
                                </p>
                            </div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                # Calculate and display the ratio of inbound/outbound if this is the total calls section
                if call_type_col == 'Calls Peak':
                    inbound_ratio = (df['Calls Peak (Inbound)'] / df['Calls Peak']).mean() * 100
                    outbound_ratio = (df['Calls Peak (Outbound)'] / df['Calls Peak']).mean() * 100
                    st.markdown(f"""
                        <div style='background-color: #e6f3ff; padding: 15px; border-radius: 10px; margin-bottom: 20px;'>
                            <h4 style='color: #1f77b4; margin-top: 0;'>Call Type Distribution</h4>
                            <p style='font-size: 16px;'>
                            On average, during peak times:<br>
                            • <strong>{inbound_ratio:.1f}%</strong> of calls are Inbound<br>
                            • <strong>{outbound_ratio:.1f}%</strong> of calls are Outbound
                            </p>
                        </div>
                    """, unsafe_allow_html=True)
                
    except Exception as e:
        st.error(f"An error occurred while loading data: {e}")
        st.write("Please check that the data is in the correct format:")
        st.write("1. Dates should be in MM/DD/YYYY format")
        st.write("2. Only numeric values for peak counts")


if __name__ == "__main__":
    main()
