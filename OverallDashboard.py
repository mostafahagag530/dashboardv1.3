import streamlit as st

# Set up the page configuration
st.set_page_config(page_title="Home Page of Dashboards", layout="wide")

# Title of the page
st.title("Home Page of Dashboards")

# Introduction
st.header("Welcome to the Dashboard")
st.write(
    "This dashboard provides a comprehensive overview of various metrics and KPIs across different departments. "
    "It is designed to help you visualize data and make informed decisions based on real-time insights."
)

# Features Section
st.header("Dashboard Features")
st.write(
    "- **Real-Time Data Monitoring**: Track live data updates across multiple departments.\n"
    "- **Interactive Visualizations**: Explore your data through graphs, charts, and tables that allow for deeper analysis.\n"
    "- **Custom Filters**: Use filters to view specific data subsets based on your needs.\n"
    "- **Export Options**: Easily export reports and data for offline analysis."
)

st.header("How to Navigate")
st.write(
    "Use the sidebar to select the dashboard you wish to view. Each dashboard is equipped with features that allow you to:\n"
    "- Analyze performance metrics\n"
    "- Compare data across time periods\n"
    "- Generate detailed reports\n"
    "If you need further assistance, please refer to the documentation or contact support."
)

# Conclusion
st.write("Thank you for using our dashboard! We hope it helps you gain valuable insights.")

# # Sidebar options
st.sidebar.success(" Please Choose What You Want ? ")
