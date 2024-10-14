import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# Function to set the first row as the header and drop it from the DataFrame
def set_first_row_as_header(df):
    df.columns = df.iloc[0]
    return df[1:]

# Function to preprocess the data
def preprocess_data_df(excel_file_path):
    xls = pd.ExcelFile(excel_file_path)
    sheet_name = xls.sheet_names[0]
# Define the path to your Excel file
# Assuming you want the first sheet

    # Read data from ranges with no header
    df_s = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="S", skiprows=1, nrows=4, header=None)
    df_V = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="V", skiprows=1, nrows=4, header=None)
    df_X = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="X", skiprows=1, nrows=4, header=None)
    df_Y = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="Y", skiprows=1, nrows=4, header=None)
    df_Z = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="Z", skiprows=1, nrows=4, header=None)
    df_AA = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AA", skiprows=1, nrows=4, header=None)

    df_s = set_first_row_as_header(df_s)
    df_V = set_first_row_as_header(df_V)
    df_X = set_first_row_as_header(df_X)
    df_Y = set_first_row_as_header(df_Y)
    df_Z = set_first_row_as_header(df_Z)
    df_AA = set_first_row_as_header(df_AA)
    # Combine all DataFrames into one
    SOFR_df = pd.concat([df_s, df_V, df_X, df_Y, df_Z, df_AA], axis=1)

    # Reset index if needed
    SOFR_df.reset_index(drop=True, inplace=True)

    # Display the combined DataFrame
    # print(SOFR_df)



    # Read data from ranges with no header
    df_s = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="S", skiprows=6, nrows=4, header=None)
    df_V = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="V", skiprows=6, nrows=4, header=None)



    df_s = set_first_row_as_header(df_s)
    df_V = set_first_row_as_header(df_V)


    # Combine all DataFrames into one
    Govt_df = pd.concat([df_s, df_V,], axis=1)

    # Reset index if needed
    Govt_df.reset_index(drop=True, inplace=True)

    # Display the combined DataFrame
    # print(Govt_df)





    # Read data from ranges with no header
    df_s = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="S", skiprows=20, nrows=4,)
    df_T= pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="T", skiprows=20, nrows=4,)




    df_s = set_first_row_as_header(df_s)
    df_T = set_first_row_as_header(df_T)


    # Combine all DataFrames into one
    Mtge_Sprd_bps_df = pd.concat([df_s, df_T,], axis=1)

    # Reset index if needed
    Mtge_Sprd_bps_df.reset_index(drop=True, inplace=True)

    # Display the combined DataFrame
    print("Combined DataFrame with the first row as headers:")
    # print(Mtge_Sprd_bps_df)




    # # Read data from ranges with no header
    df_s = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="S", skiprows=25, nrows=4,)
    df_w= pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="W", skiprows=25, nrows=4,)



    df_s = set_first_row_as_header(df_s)
    df_w = set_first_row_as_header(df_w)


    # Combine all DataFrames into one
    Swaption_Vol_df = pd.concat([df_s, df_w,], axis=1)

    # Reset index if needed
    Swaption_Vol_df.reset_index(drop=True, inplace=True)

    # Display the combined DataFrame
    # print(Swaption_Vol_df)






    # # Read data from ranges with no header
    df_s = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="S", skiprows=30, nrows=4,)
    df_V = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="V", skiprows=30, nrows=4,)


    # Set the first row as the header and drop it from the DataFrame

    df_s = set_first_row_as_header(df_s)
    df_V = set_first_row_as_header(df_V)


    # Combine all DataFrames into one
    AGENCY_Vol_df = pd.concat([df_s, df_V,], axis=1)

    # Reset index if needed
    AGENCY_Vol_df.reset_index(drop=True, inplace=True)







    # # Read data from ranges with no header
    df_af = pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AF", skiprows=25, nrows=4,)
    df_ag= pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AG", skiprows=25, nrows=4,)
    df_ah= pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AH", skiprows=25, nrows=4,)
    df_ai= pd.read_excel(excel_file_path, sheet_name=sheet_name, usecols="AI", skiprows=25, nrows=4,)




    df_af= set_first_row_as_header(df_af) 
    df_ag= set_first_row_as_header(df_ag)
    df_ah= set_first_row_as_header(df_ah)
    df_ai= set_first_row_as_header(df_ai)

    # Combine all DataFrames into one
    x10_OIS_x10_SOFR_x10_Agy_df = pd.concat([df_af,df_ag,df_ah,df_ai], axis=1)

    # Reset index if needed
    x10_OIS_x10_SOFR_x10_Agy_df.reset_index(drop=True, inplace=True)
    # x10_OIS_x10_SOFR_x10_Agy_df.astype("object")
    # x10_OIS_x10_SOFR_x10_Agy_df[3:1]="Change"
    # Display the combined DataFrame
    print("Combined DataFrame with the first row as headers:")
    # print(x10_OIS_x10_SOFR_x10_Agy_df)


    df_full=pd.concat([AGENCY_Vol_df,x10_OIS_x10_SOFR_x10_Agy_df.iloc[:,1:],Swaption_Vol_df.iloc[:,1:],Mtge_Sprd_bps_df.iloc[:,1:],Govt_df.iloc[:,1:],SOFR_df.iloc[:,1:]], axis=1)

    return df_full


st.set_page_config(layout="wide")

@st.cache_data
def load_and_preprocess(filepath):
    processed_df = preprocess_data_df(filepath)
    return processed_df

# Streamlit app title
st.title(" Data Visualization AttributionAnalysis Change ")

# Sidebar for file uploader and date range selection
st.sidebar.header("Upload and Filter Data")

uploaded_files = st.sidebar.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

# Initialize the DataFrame variable
df = pd.DataFrame()

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        df = load_and_preprocess(uploaded_file)
        all_dfs.append(df)
    # st.write(all_dfs)
    # Combine all dataframes into one

    df = pd.concat(all_dfs, ignore_index=True)
    df.columns=['Date_Change', 'Agence_10Y', '1x10 OIS', '1x10 SOFR', '1x10 Agy', '1x10', '30Y MBS',
       'Govt_10Y', 'SOFR_10Y', 'OAS Fx', 'OAS Callable', 'OAS MPF', 'YTM']
    df = df[df['Date_Change'] != 'Change']

    # Convert 'Agy/SOFR' to datetime
    df['Date_Change'] = pd.to_datetime(df['Date_Change'], errors='coerce')
    df=df.sort_values(by='Date_Change')
# Calculate the change for each column
    for column in df.columns[1:]:
        print(column,"kjkjkl")
        df[f'Change_of {column}'] = df[column].diff()
    # Set default dates for the date input
    if not df.empty:
        default_start_date = df['Date_Change'].min()
        default_end_date = df['Date_Change'].max()
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
        filtered_df = df[
            (df['Date_Change'] >= pd.to_datetime(start_date)) & 
            (df['Date_Change'] <= pd.to_datetime(end_date))
        ]
    except Exception as e:
        st.warning("Error filtering by date. Defaulting to full date range.")
        filtered_df = df

    # Select the columns to display
    selected_columns = st.sidebar.multiselect(
        "Select columns to display", 
        filtered_df.columns.tolist(), 
        default=['Date_Change', 'Agence_10Y', '1x10 OIS', '1x10 SOFR', '1x10 Agy', '1x10', '30Y MBS',
       'Govt_10Y', 'SOFR_10Y', 'OAS Fx', 'OAS Callable', 'OAS MPF', 'YTM']
    )
    num_rows_to_display = st.sidebar.slider("Select number of rows to display", min_value=1, max_value=len(filtered_df), value=10)

    # Display the filtered dataframe
    st.dataframe(filtered_df[selected_columns].head(num_rows_to_display), use_container_width=True)

    # Convert 'Agy/SOFR' to string format for plotting
    df['Date_Change'] = df['Date_Change'].dt.strftime('%d-%m-%Y')

    # Define a theme for the plots
    def create_theme():
        return {
            'title_font': dict(size=24, family='Arial, sans-serif', color="black"),
            'axis_font': dict(size=14, family='Arial, sans-serif', color="black"),
            'colors': ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728'],
            'bg_color': '#f5f5f5',
            'grid_color': 'rgba(0, 0, 0, 0.1)',
            'paper_bgcolor': 'white',
        }

    theme = create_theme()

    # Create individual bar charts for each term
# Create a plot for each column and its change column
    for column in df.columns[1:13]:
        if column != 'Date_Change':
            # Create a new figure for each column
            fig = go.Figure()

            # Define colors
            main_color = '#1B1A55'  # Blue for positive values
            change_color = '#9290C3'  # Red for positive values
            main_negative_color = '#EB3678'  # Red for negative values
            change_negative_color = '#FB773C'  # Dark red for negative values

            # Add the column values as a bar trace with conditional colors
            fig.add_trace(go.Bar(
                x=df['Date_Change'],
                y=df[column],
                name=column,
                marker_color=[main_color if value >= 0 else main_negative_color for value in df[column]],
                text=df[column],  # Display the values on the bars
                textposition='inside',  # Position the text inside the bars
                hoverinfo='text',
                hovertext=[f"Date_Change: {df['Date_Change'].iloc[i]}<br>{column}: {value}" for i, value in enumerate(df[column])]
            ))

            # Check if the column has a corresponding 'change_of_' column
            change_column = f'Change_of {column}'
            if change_column in df.columns:
                print(change_column,"ffff")
                # Add the change column values as a separate bar trace with conditional colors
                fig.add_trace(go.Bar(
                    x=df['Date_Change'],
                    y=df[change_column],
                    name=change_column,
                    marker_color=[change_color if value >= 0 else change_negative_color for value in df[change_column]],
                    text=df[change_column],  # Display the values on the bars
                    textposition='inside',  # Position the text inside the bars
                    hoverinfo='text',
                    hovertext=[f"Date_Change: {df['Date_Change'].iloc[i]}<br>{change_column}: {value}" for i, value in enumerate(df[change_column])]
                ))

            # Update layout to match the custom theme
            fig.update_layout(
                title=f'<b>  {column} and {change_column} Over Date </b>',
                title_font=dict(size=20, family='Arial, sans-serif',color="Black"),
                title_x=0.1,  # Center the title

                # X-axis title and styling
                xaxis_title='<b>Date</b>',
                xaxis=dict(
                    title_font=dict(size=16, family='Arial, sans-serif',color="Black"),
                    type='category',  # Treat x-axis as categorical
                    showgrid=False,  # Hide gridlines
                    tickangle=-45,  # Rotation for the x-ticks
                    tickfont=dict(size=12,color="Black")
                ),

                # Y-axis title and styling
                yaxis_title=f'<b> Value of {column} and {change_column} </b>',
                yaxis=dict(
                    title_font=dict(size=16, family='Arial, sans-serif',color="Black"),
                    zeroline=False,  # Hide zero line
                    tickfont=dict(size=12,color="Black"),
                    gridcolor='Black'  # Subtle gridlines for the y-axis
                ),

                # Background colors
                plot_bgcolor='#EEEDEB',  # Set plot background color
                paper_bgcolor='white',  # Set paper background color

                # Update margins for better spacing
                margin=dict(l=40, r=40, t=80, b=40),

                # Global font settings for all text
                font=dict(
                    family="Arial, sans-serif",
                    size=12,
                    color="Black"
                ),
                legend=dict(
                font=dict(
                    color="Black"  # Change legend text color to white or any desired color
                ),
                orientation="h",  # Optional: make the legend horizontal
                x=0.8,  # Center the legend horizontally
                xanchor="center",
                y=1.05,  # Position it just above the plot
                yanchor="bottom"
            ),
            width=800,  # Set the width of the figure
            height=600 
            
            )

            # Display the chart with full width inside the loop
            with st.container():
                st.plotly_chart(fig, use_container_width=True)