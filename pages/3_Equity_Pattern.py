import streamlit as st
import pandas as pd
import plotly.graph_objects as go
#

# Function to preprocess the data
def preprocess_data_df(excel_file_path):
    df=pd.read_excel(excel_file_path)
    df=df.T
    # Set the first row as the header
    df.columns = df.iloc[0]

    # Drop the first row since it's now the header
    df = df.drop(df.index[0])
    df.reset_index(inplace=True)
    df["Date"]=df["index"]
    # df.columns[1:-1]
    df[df.columns[-1]]= pd.to_datetime(df[df.columns[-1]], format='%Y%m').dt.date 
    return df
st.set_page_config(layout="wide")

@st.cache_data
def load_and_preprocess(filepath):
    processed_df = preprocess_data_df(filepath)
    return processed_df

# Streamlit app title
st.title(" Data Visualization Equity-Pattern ")

# Sidebar for file uploader and date range selection
st.sidebar.header("Upload and Filter Data")

uploaded_files = st.sidebar.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

# Initialize the DataFrame variable
df_full_Agy_SOFR = pd.DataFrame()

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        df = load_and_preprocess(uploaded_file)
        all_dfs.append(df)
        
    # Combine all dataframes into one
    df = pd.concat(all_dfs, ignore_index=True)

    # Convert 'Date' to datetime
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df_full_Agy_SOFR=  df.sort_values(by='Date')
    # Set default dates for the date input
    if not df_full_Agy_SOFR.empty:
        default_start_date = df_full_Agy_SOFR['Date'].min()
        default_end_date = df_full_Agy_SOFR['Date'].max()
    else:
        default_start_date = default_end_date = pd.NaT

    # Select date range to filter
    start_date, end_date = st.sidebar.date_input(
        "Select date range",
        value=(default_start_date, default_end_date),
        min_value=default_start_date,
        max_value=default_end_date
    )

    # Filter the dataframe by selected date range
    try:
        filtered_df = df_full_Agy_SOFR[
            (df_full_Agy_SOFR['Date'] >= pd.to_datetime(start_date)) & 
            (df_full_Agy_SOFR['Date'] <= pd.to_datetime(end_date))
        ]
    except Exception as e:
        st.warning("Error filtering by date. Defaulting to full date range.")
        filtered_df = df_full_Agy_SOFR

    # Select the columns to display
    selected_columns = st.sidebar.multiselect(
        "Select columns to display", 
        filtered_df.columns.tolist(), 
        default=['Portfolio', 'Portfolio difference', 'Portfolio growth rate',
       'Rates', 'Rates difference', 'Rates growth rate', 'Volatility',
       'Volatility difference', 'Volatility growth rate', 'OAS',
       'OAS difference', 'OAS growth rate', 'Model Change',
       'Model Change difference', 'Model Change growth rate', 'Other',
       'Other difference', 'Other growth rate', 'Residual',
       'Residual difference', 'Residual growth rate', 'Total Equity', 'Date']
    )
    num_rows_to_display = st.sidebar.slider("Select number of rows to display", min_value=1, max_value=len(filtered_df), value=10)

    # Display the filtered dataframe
    st.dataframe(filtered_df[selected_columns].head(num_rows_to_display), use_container_width=True)

    # Convert 'Date' to string format for plotting
    filtered_df['Date'] = filtered_df['Date'].dt.strftime('%d-%m-%Y')

    # Define a new theme for the plots
    def create_theme():
        return {
            'title_font': dict(size=24, family='Arial, sans-serif', color="Black"),
            'axis_font': dict(size=16, family='Arial, sans-serif', color="Black"),
            'colors': ['#1B1A55', '#2ca02c', '#EB3678', '#FB773C'],  # Darker color scheme for bars and trend lines
            'bg_color': '#EEEDEB',  # Subtle gray background for the plot
            'grid_color': 'rgba(0, 0, 0, 0.1)',  # Light gridlines
            'paper_bgcolor': 'white',  # White background for the paper
        }

    theme = create_theme()
   # Sort the filtered dataframe by 'Date' (date) in ascending order
    # filtered_df = filtered_df.sort_values(ascending=True)
            # Define colors
    main_color = '#1B1A55'  # Blue for positive values
    change_color = '#2ca02c'  # Red for positive values
    main_negative_color = '#EB3678'  # Red for negative values
    change_negative_color = '#FB773C'  # Dark red for negative values

    # Create individual bar charts for each term
    for col in selected_columns:
        if col == "Date":
            continue

        # Create a bar chart
        fig = go.Figure()

        # Add bars with hover info and custom color
        fig.add_trace(go.Bar(
            x=filtered_df["Date"],
            y=filtered_df[col].round(),
            marker_color=[main_color if value >= 0 else main_negative_color for value in df[col]],  # Use updated color
            # opacity=1,
            # width=0.6,
            text=filtered_df[col].round(2),
            textposition='inside',
            hoverinfo='text',
            hovertext=[f"Date: {filtered_df['Date'].iloc[i]}<br>{col}: {v}" for i, v in enumerate(filtered_df[col])]
        ))

        # Update layout with new theme
        fig.update_layout(
            title=f'<b>{col}</b>',
            title_font=theme['title_font'],
            title_x=.3,
            xaxis_title=f'<b>Change by Date </b>',
            xaxis=dict(
                title_font=theme['axis_font'],
                type='category',
                showgrid=True,
                gridcolor=theme['grid_color'],
                tickangle=45,
                tickfont=dict(size=12, color="Black"),
            ),
            yaxis_title=f'<b>Value of {col}</b>',
            yaxis=dict(
                title_font=theme['axis_font'],
                zeroline=False,
                tickfont=dict(size=12, color="Black"),
                gridcolor=theme['grid_color'],
            ),
            plot_bgcolor=theme['bg_color'],
            paper_bgcolor=theme['paper_bgcolor'],
            margin=dict(l=40, r=40, t=80, b=40),
            font=dict(
                family="Arial, sans-serif",
                size=12,
                color="Black"
            ),
            legend=dict(
                font=dict(
                    size=12,
                    color="Black"
                )
            ),
            width=800,  # Set the width of the figure
            height=600
        )

        # Add a trend line with updated color
        fig.add_trace(go.Scatter(
            x=filtered_df["Date"],
            y=filtered_df[col].round(),
            mode='lines+markers',
            name='Trend Line',
            marker=dict(color=theme['colors'][1]),  # Updated color for the trend line
            line=dict(width=2, color=theme['colors'][1])
        ))

        # Display the chart with full width
        with st.container():
            st.plotly_chart(fig, use_container_width=True)