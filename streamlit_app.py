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
    st.subheader("Analysis of Peak Data")

    check_openpyxl()
    
    try:
        # Load the Excel file
        data_path = "peaks.xlsx"
        df = pd.read_excel(data_path)
        
        # Data Overview Section
        st.write("### Data Overview")
        st.dataframe(df)

        # Add filters
        st.write("### Filters")
        cols = df.columns.tolist()
        selected_cols = st.multiselect("Select columns to analyze:", cols, default=cols[:2])

        if selected_cols:
            filtered_df = df[selected_cols]
            
            # Create visualizations
            st.write("### Peak Analysis")
            
            # Bar chart for comparing values
            if len(selected_cols) >= 1:
                for col in selected_cols:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        st.write(f"#### Analysis of {col}")
                        chart = alt.Chart(df).mark_bar().encode(
                            x=alt.X(col, bin=True),
                            y='count()',
                            tooltip=[col, 'count()']
                        ).properties(
                            width=600,
                            height=400,
                            title=f"Distribution of {col}"
                        )
                        st.altair_chart(chart, use_container_width=True)

                        # Show basic statistics
                        st.write("Basic Statistics:")
                        stats = df[col].describe()
                        st.write(f"Mean: {stats['mean']:.2f}")
                        st.write(f"Median: {stats['50%']:.2f}")
                        st.write(f"Max: {stats['max']:.2f}")
                        st.write(f"Min: {stats['min']:.2f}")

            # Correlation heatmap for multiple numeric columns
            numeric_cols = df[selected_cols].select_dtypes(include=['number']).columns
            if len(numeric_cols) >= 2:
                st.write("#### Correlation Analysis")
                corr_df = df[numeric_cols].corr().reset_index()
                corr_df = corr_df.melt('index', var_name='variable', value_name='correlation')
                
                heatmap = alt.Chart(corr_df).mark_rect().encode(
                    x='index:O',
                    y='variable:O',
                    color='correlation:Q',
                    tooltip=['index', 'variable', 'correlation']
                ).properties(
                    width=400,
                    height=300,
                    title="Correlation Heatmap"
                )
                st.altair_chart(heatmap, use_container_width=True)

    except Exception as e:
        st.error(f"An error occurred while loading data: {e}")


if __name__ == "__main__":
    main()
