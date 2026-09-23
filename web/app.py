import os
import sqlite3
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ontario Daily Court Lists Dashboard",
    page_icon="⚖️",
    layout="wide"
)

st.title("⚖️ Ontario Daily Court Lists Search Dashboard")
st.markdown(
    "Search, filter, and inspect daily court listings across Ontario Superior Court of Justice and Ontario Court of Justice."
)

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
DB_PATH = os.path.join(DATA_DIR, "court_dockets.db")
LATEST_CSV = os.path.join(DATA_DIR, "latest.csv")


@st.cache_data(ttl=600)
def load_data():
    if os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        df = pd.read_sql("SELECT * FROM dockets", conn)
        conn.close()
        return df
    elif os.path.exists(LATEST_CSV):
        return pd.read_csv(LATEST_CSV)
    return pd.DataFrame()


df = load_data()

if df.empty:
    st.warning("No court docket data found yet. Run the daily scraper to populate data.")
    st.stop()

# Sidebar Filters
st.sidebar.header("🔍 Filters")

# Municipality Filter
all_cities = sorted([str(c) for c in df["city"].dropna().unique()])
selected_cities = st.sidebar.multiselect("Municipality / City", options=all_cities, default=[])

# Court Filter
all_courts = sorted([str(c) for c in df["court"].dropna().unique()])
selected_courts = st.sidebar.multiselect("Court Level", options=all_courts, default=[])

# Attendance Method Filter
all_methods = sorted([str(m) for m in df["method_of_attendance"].dropna().unique() if str(m).strip()])
selected_methods = st.sidebar.multiselect("Method of Attendance", options=all_methods, default=[])

# Search bar
search_query = st.sidebar.text_input("Search Party Name / Case # / Title", placeholder="e.g. Smith or CR-25-...")

# Filter logic
filtered_df = df.copy()

if selected_cities:
    filtered_df = filtered_df[filtered_df["city"].isin(selected_cities)]

if selected_courts:
    filtered_df = filtered_df[filtered_df["court"].isin(selected_courts)]

if selected_methods:
    filtered_df = filtered_df[filtered_df["method_of_attendance"].isin(selected_methods)]

if search_query.strip():
    q = search_query.strip().lower()
    filtered_df = filtered_df[
        filtered_df["party_name"].astype(str).str.lower().str.contains(q)
        | filtered_df["case_number"].astype(str).str.lower().str.contains(q)
        | filtered_df["short_title"].astype(str).str.lower().str.contains(q)
        | filtered_df["event_type"].astype(str).str.lower().str.contains(q)
    ]

# Summary Metrics
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Cases Listed", f"{len(filtered_df):,}")
col2.metric("Municipalities", len(filtered_df["city"].unique()))
col3.metric("Remote / Video Cases", len(filtered_df[filtered_df["method_of_attendance"].str.contains("Video|Zoom|Teleconference", case=False, na=False)]))
col4.metric("In-Person Cases", len(filtered_df[filtered_df["method_of_attendance"].str.contains("Attendance|In Person|Hybrid", case=False, na=False)]))

st.markdown("---")

# Download options
col_dl1, col_dl2 = st.columns([1, 4])
with col_dl1:
    csv_bytes = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Export Filtered CSV",
        data=csv_bytes,
        file_name="ontario_court_dockets_filtered.csv",
        mime="text/csv"
    )

# Presentation Table
display_cols = [
    "date_text", "city", "court", "time", "room",
    "party_name", "case_number", "short_title",
    "event_type", "method_of_attendance", "docket_line"
]
existing_display_cols = [c for c in display_cols if c in filtered_df.columns]

st.dataframe(
    filtered_df[existing_display_cols],
    use_container_width=True,
    hide_index=True
)
