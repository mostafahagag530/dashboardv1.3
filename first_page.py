import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# Function to set the first row as the header and drop it from the DataFrame
def set_first_row_as_header(df):
    df.columns = df.iloc[0]
    return df[1:]

# Function to preprocess the data
def preprocess_data_df(excel_file_path):
    xls = pd.ExcelFile(excel_file_path)
    sheet_name = xls.sheet_names[0]

    # Load the required data from the Excel file
    date_al = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AL", skiprows=1, nrows=1)
    date_ap = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AP", skiprows=1, nrows=1)

    df_al = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AL", skiprows=6, nrows=1)
    df_ao = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AO", skiprows=6, nrows=1)
    df_ar = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AR", skiprows=6, nrows=1)

    df_al2 = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AL", skiprows=22, nrows=1)
    df_ao2 = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AO", skiprows=22, nrows=1)
    df_ar2 = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AR", skiprows=22, nrows=1)

    df_al3 = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AL", skiprows=61, nrows=1)
    df_ao3 = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AO", skiprows=61, nrows=1)
    df_ar3 = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AR", skiprows=61, nrows=1)

    # Combine data
    df_al_combined = pd.concat([df_al, df_ao, df_ar], axis=1).reset_index(drop=True)
    df_ao_combined = pd.concat([df_al2, df_ao2, df_ar2], axis=1).reset_index(drop=True)
    df_ar_combined = pd.concat([df_al3, df_ao3, df_ar3], axis=1).reset_index(drop=True)

    df_combined = pd.concat([df_al_combined.T, df_ao_combined.T, df_ar_combined.T], axis=0)
    
    df_row_8 = df_combined.iloc[:3, :]
    df_row_8["date_type"] = ["Agy/SOFR".strip(), date_al.columns[0], date_ap.columns[0]]
    df_row_8.reset_index(drop=True, inplace=True)

    df_row_24 = df_combined.iloc[3:6, :]
    df_row_24["date_type"] = ["Agy/SOFR".strip(), date_al.columns[0], date_ap.columns[0]]
    df_row_24.reset_index(drop=True, inplace=True)

    df_row_63 = df_combined.iloc[6:, :]
    df_row_63["date_type"] = ["Agy/SOFR".strip(), date_al.columns[0], date_ap.columns[0]]
    df_row_63.reset_index(drop=True, inplace=True)

    df_full_Agy_SOFR = pd.concat([df_row_63, df_row_24.iloc[:, :1], df_row_8.iloc[:, :1]], axis=1)
    df_full_Agy_SOFR = set_first_row_as_header(df_full_Agy_SOFR)
    df_full_Agy_SOFR.columns = ["Term_30", "Agy/SOFR", "Term_10", "Term_2"]

    return df_full_Agy_SOFR


st.set_page_config(layout="wide")

@st.cache_data
def load_and_preprocess(filepath):
    processed_df = preprocess_data_df(filepath)
    return processed_df

# Streamlit app title
st.title("Excel File Upload and Data Visualization 'Agy/SOFR/Date'  Term_30 , Term_10 , Term_2 ")

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
    df_full_Agy_SOFR = pd.concat(all_dfs, ignore_index=True)

    # Convert 'Agy/SOFR' to datetime
    df_full_Agy_SOFR['Agy/SOFR'] = pd.to_datetime(df_full_Agy_SOFR['Agy/SOFR'], errors='coerce')
    df_full_Agy_SOFR=  df_full_Agy_SOFR.sort_values(by='Agy/SOFR')
    # Set default dates for the date input
    if not df_full_Agy_SOFR.empty:
        default_start_date = df_full_Agy_SOFR['Agy/SOFR'].min()
        default_end_date = df_full_Agy_SOFR['Agy/SOFR'].max()
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
            (df_full_Agy_SOFR['Agy/SOFR'] >= pd.to_datetime(start_date)) & 
            (df_full_Agy_SOFR['Agy/SOFR'] <= pd.to_datetime(end_date))
        ]
    except Exception as e:
        st.warning("Error filtering by date. Defaulting to full date range.")
        filtered_df = df_full_Agy_SOFR

    # Select the columns to display
    selected_columns = st.sidebar.multiselect(
        "Select columns to display", 
        filtered_df.columns.tolist(), 
        default=['Agy/SOFR', 'Term_30', 'Term_10', 'Term_2']
    )
    num_rows_to_display = st.sidebar.slider("Select number of rows to display", min_value=1, max_value=len(filtered_df), value=10)

    # Display the filtered dataframe
    st.dataframe(filtered_df[selected_columns].head(num_rows_to_display), use_container_width=True)

    # Convert 'Agy/SOFR' to string format for plotting
    filtered_df['Agy/SOFR'] = filtered_df['Agy/SOFR'].dt.strftime('%d-%m-%Y')

    # Define a new theme for the plots
    def create_theme():
        return {
            'title_font': dict(size=24, family='Arial, sans-serif', color="Black"),
            'axis_font': dict(size=16, family='Arial, sans-serif', color="Black"),
            'colors': ['#1B1A55', '#9290C3', '#EB3678', '#FB773C'],  # Darker color scheme for bars and trend lines
            'bg_color': '#EEEDEB',  # Subtle gray background for the plot
            'grid_color': 'rgba(0, 0, 0, 0.1)',  # Light gridlines
            'paper_bgcolor': 'white',  # White background for the paper
        }

    theme = create_theme()
   # Sort the filtered dataframe by 'Agy/SOFR' (date) in ascending order
    # filtered_df = filtered_df.sort_values(ascending=True)

    # Create individual bar charts for each term
    for col in selected_columns:
        if col == "Agy/SOFR":
            continue

        # Create a bar chart
        fig = go.Figure()

        # Add bars with hover info and custom color
        fig.add_trace(go.Bar(
            x=filtered_df["Agy/SOFR"],
            y=filtered_df[col].round(),
            marker_color=theme['colors'][0],  # Use updated color
            # opacity=0.7,
            # width=0.6,
            text=filtered_df[col].round(2),
            textposition='outside',
            hoverinfo='text',
            hovertext=[f"Agy/SOFR: {filtered_df['Agy/SOFR'].iloc[i]}<br>{col}: {v}" for i, v in enumerate(filtered_df[col])]
        ))

        # Update layout with new theme
        fig.update_layout(
            title=f'<b>Agy/SOFR vs {col}</b>',
            title_font=theme['title_font'],
            title_x=0.1,
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
            height=600, 
        )

        # Add a trend line with updated color
        fig.add_trace(go.Scatter(
            x=filtered_df["Agy/SOFR"],
            y=filtered_df[col].round(),
            mode='lines+markers',
            name='Trend Line',
            marker=dict(color=theme['colors'][1]),  # Updated color for the trend line
            line=dict(width=2, color=theme['colors'][1])
        ))

        # Display the chart with full width
        with st.container():
            st.plotly_chart(fig, use_container_width=True)