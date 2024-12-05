import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
import math

# Database connection setup
DATABASE_URI = "mysql+pymysql://root:SubhanAllah%401DB@localhost:3306/TicketActivityDB"  # Replace with your DB URI
engine = create_engine(DATABASE_URI)

# Mapping of table names to friendly display names
TABLE_NAME_MAPPING = {
    "activitysummary": "Activity Summary",
    "datewisesummary": "Date-Wise Summary",
    "extractedactivities": "Extracted Activities",
    "teamwisesummary": "Team-Wise Summary",
    # Add more mappings if needed
}

def fetch_table_names():
    """Fetch all table names from the database."""
    with engine.connect() as connection:
        query = text("SHOW TABLES;")
        result = connection.execute(query)
        tables = [
            row[0] for row in result 
            if row[0] != "configsetup"  # Exclude 'configsetup'
        ]
        return tables

def get_friendly_name(table_name):
    """Get the friendly display name for a table."""
    return TABLE_NAME_MAPPING.get(table_name, table_name) + " Report"

def fetch_column_names(table_name):
    """Fetch column names of the given table."""
    raw_table_name = table_name.replace(" Report", "")  # Remove 'Report' for querying
    query = text(f"DESCRIBE {raw_table_name};")
    with engine.connect() as connection:
        result = connection.execute(query)
        columns = [row[0] for row in result]
    return columns

def fetch_table_data(table_name, offset=0, limit=10, search_query=None):
    """Fetch data from the specified table with pagination and optional search."""
    raw_table_name = table_name.replace(" Report", "")  # Remove 'Report' for querying
    query = f"SELECT * FROM {raw_table_name}"
    
    if search_query:
        column_names = fetch_column_names(raw_table_name)
        if column_names:
            # Safely build the WHERE clause with CONCAT_WS
            where_clause = f" WHERE CONCAT_WS(' ', {', '.join(column_names)}) LIKE :search_query"
            query += where_clause

    query += f" LIMIT {limit} OFFSET {offset};"

    with engine.connect() as connection:
        if search_query:
            df = pd.read_sql(text(query), connection, params={"search_query": f"%{search_query}%"})
        else:
            df = pd.read_sql(query, connection)

    return df


def fetch_total_row_count(table_name, search_query=None):
    """Fetch the total row count of a table, optionally filtered by a search query."""
    raw_table_name = table_name.replace(" Report", "")  # Remove 'Report' for querying
    with engine.connect() as connection:
        if search_query:
            column_names = fetch_column_names(raw_table_name)
            if column_names:
                where_clause = f" WHERE CONCAT_WS(' ', {', '.join(column_names)}) LIKE :search_query"
                query = text(f"SELECT COUNT(*) FROM {raw_table_name} {where_clause}")
                result = connection.execute(query, {"search_query": f"%{search_query}%"})
            else:
                # If no columns, return 0
                return 0
        else:
            query = text(f"SELECT COUNT(*) FROM {raw_table_name}")
            result = connection.execute(query)

        total_count = result.scalar()  # Get the single scalar result (COUNT)
        return total_count


def update_activity_summary_counts():
    """Update the 'count' column in activitySummary table based on extractedActivities table."""
    update_query = """
    UPDATE activitySummary
    SET `count` = (
        SELECT COUNT(*)
        FROM extractedActivities
        WHERE extractedActivities.activityType LIKE CONCAT('%', activitySummary.activityType, '%')
    );
    """
    try:
        with engine.connect() as connection:
            connection.execute(text(update_query))
        st.success("Activity Summary table updated successfully!")
    except Exception as e:
        st.error(f"Error updating Activity Summary table: {e}")

# Streamlit App
st.set_page_config(page_title="Desk Ticket Activity Report Viewer", layout="wide", page_icon="📊")

# Initialize session state for page_number
if "page_number" not in st.session_state:
    st.session_state.page_number = 1

# Title
st.title("📊 Desk Ticket Activity Report Viewer")
st.write("Explore and interact with your database reports dynamically!")

# Sidebar: Report Selector
tables = fetch_table_names()
friendly_table_names = [get_friendly_name(table) for table in tables]  # Convert to friendly names
selected_table = st.sidebar.selectbox("Select a Report", friendly_table_names)

# Map the selected friendly name back to the raw table name
raw_selected_table = tables[friendly_table_names.index(selected_table)]  # Get raw name from mapping

# Search and Pagination
search_query = st.text_input("Search", placeholder="Type to search across all columns...")
rows_per_page = st.selectbox("Rows per page", options=[5, 10, 20, 50], index=1)

# Fetch data
total_rows = fetch_total_row_count(raw_selected_table, search_query)
total_pages = max(1, math.ceil(total_rows / rows_per_page))  # Ensure at least 1 page

# Pagination Logic
col1, col2, col3 = st.columns([1, 2, 1])
with col1:
    if st.button("Previous Page") and st.session_state.page_number > 1:
        st.session_state.page_number -= 1
with col3:
    if st.button("Next Page") and st.session_state.page_number < total_pages:
        st.session_state.page_number += 1

# Display current page number
page_number = st.session_state.page_number
st.write(f"Page {page_number} of {total_pages}")

# Offset calculation for SQL query
offset = (page_number - 1) * rows_per_page

data = fetch_table_data(raw_selected_table, offset=offset, limit=rows_per_page, search_query=search_query)

# Display Data
st.markdown(f"### Report Name: {selected_table} ({total_rows} Records)")
if not data.empty:
    # Remove the "id" column from display
    if "id" in data.columns:
        data = data.drop(columns=["id"])
    st.dataframe(data, use_container_width=True)
else:
    st.warning("No data found!")

# Sidebar: Additional Features
st.sidebar.header("Advanced Options")
st.sidebar.write("Choose features to enable on this page:")
enable_sorting = st.sidebar.checkbox("Enable Sorting", value=True)
enable_filtering = st.sidebar.checkbox("Enable Filtering", value=True)

if enable_sorting or enable_filtering:
    st.info("Sorting and filtering are enabled for displayed data.")

# Button to Update Activity Summary
if raw_selected_table == "activitysummary":
    if st.button("Update Activity Summary Counts"):
        update_activity_summary_counts()

# Footer
st.markdown(
    """
    <style>
        footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div style="text-align: center; margin-top: 20px;">
        <small>Built with ❤️ using Streamlit</small>
    </div>
    """,
    unsafe_allow_html=True,
)
