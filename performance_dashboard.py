import pandas as pd
import streamlit as st
from io import StringIO
import plotly.express as px
from datetime import datetime
import numpy as np

# Import libraries for web requests and Outlook automation
try:
    import requests
    import urllib3
except ImportError:
    st.error("The 'requests' library is required. Run 'pip install requests' and restart.")
    st.stop()

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    win32 = None
    pythoncom = None

# Suppress the InsecureRequestWarning from verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- Page Configuration ---
st.set_page_config(layout="wide", page_title="Team Performance Dashboard", page_icon="🏆")

# --- Constants and Team Definitions ---
ORIGINAL_CASE_TEAM = [
    "ashish johnjeyaraj", "Basavaraj Awati", "K SAI KOUSHIK", "Kaustubh SINGH",
    "Manoj KUMAR R", "Mohammed Misba RIZVI", "Nagaraj JOTHI",
    "Paluvadi Venkata SAI UJWALA", "Pranam M", "Rajath Y C", "Shachi JAIN",
    "Shikha SINHA", "Shreesha J", "Sukanya DODDAGOUDAR"
]
UCS_BANGALORE_TEAM_LOWER = {name.lower() for name in ORIGINAL_CASE_TEAM}
LOWER_TO_ORIGINAL_CASE_MAP = {name.lower(): name for name in ORIGINAL_CASE_TEAM}

INDIVIDUAL_TARGET = 15
CONTRIBUTION_RECORD_TYPES = ['WO', 'PTR', 'TR']
EXEMPTION_KEYWORDS = ['atc-mon', 'atc-fup', 'atc-ign', 'atc-sup']
EXEMPTION_REGEX = '|'.join(EXEMPTION_KEYWORDS)
SENDER_EMAIL = "blr-atc@amadeus.com"
EMAIL_MAPPING = {name.lower(): f"{name.replace(' ', '.')}@amadeus.com" for name in ORIGINAL_CASE_TEAM}
EMAIL_MAPPING.update({
    "k sai koushik": "k.saikoushik@amadeus.com",
    "manoj kumar r": "manoj.kumarr@amadeus.com",
    "mohammed misba rizvi": "misba.rizvi@amadeus.com",
    "paluvadi venkata sai ujwala": "paluvadivenkata.saiujwala@amadeus.com",
    "kaustubh singh": "kaustubh.singh@example.com",
})

CSV_COLUMN_NAMES = [
    "record_id", "rec_type", "severity", "assignee_name", "location", "transfers",
    "tags", "start_date", "end_date", "duration_days", "title", "entity_1_code",
    "entity_1_name", "entity_2_code", "entity_2_name", "entity_3_code", "entity_3_name"
]
API_BASE_URL = "https://dashproach.amadeus.net/api/record/DAPPATC/teamactivity"

@st.cache_data(ttl=600)
def fetch_data_from_api():
    """Fetches data for current/previous month from API, handling internal SSL."""
    today = datetime.now()
    all_data = []
    for month_offset in range(2):
        dt = today - pd.DateOffset(months=month_offset)
        url = f"{API_BASE_URL}?year={dt.year}&month={dt.month}"
        try:
            response = requests.get(url, timeout=20, verify=False)
            response.raise_for_status()
            all_data.append(response.text)
        except requests.exceptions.RequestException as e:
            st.error(f"Failed to fetch data from API ({url}): {e}")
            return None
    return "\n".join(all_data)

@st.cache_data(ttl=600)
def load_and_process_data(raw_csv_data):
    """Loads and processes raw data, assigning an ISO week based on the END DATE."""
    if not raw_csv_data: return pd.DataFrame()
    try:
        df = pd.read_csv(StringIO(raw_csv_data), header=None, names=CSV_COLUMN_NAMES)
        df.rename(columns={"assignee_name": "Team Member", "rec_type": "Record Type"}, inplace=True)
        df["Team Member"] = df["Team Member"].str.lower().fillna("unassigned")
        df["tags"] = df["tags"].str.lower().fillna("")
        df["start_date"] = pd.to_datetime(df["start_date"], errors='coerce').dt.tz_localize('UTC')
        df["end_date"] = pd.to_datetime(df["end_date"], errors='coerce').dt.tz_localize('UTC')
        df.dropna(subset=["start_date", "end_date"], inplace=True)
        iso_cal = df["end_date"].dt.isocalendar()
        df["year_week"] = iso_cal["year"].astype(str) + "-W" + iso_cal["week"].astype(str).str.zfill(2)
        df['Team'] = df['Team Member'].apply(lambda x: 'UCS Bangalore' if x in UCS_BANGALORE_TEAM_LOWER else 'Other Teams')
        return df
    except Exception as e:
        st.error(f"Error parsing data: {e}")
        return pd.DataFrame()

def calculate_contribution_summary(df, team_members):
    """Generic function to calculate contribution summary for a given dataframe and member list."""
    df_contributions = df[df['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)].copy()
    gross = df_contributions.groupby('Team Member').size()
    exempted = df_contributions[df_contributions['tags'].str.contains(EXEMPTION_REGEX, na=False)].groupby('Team Member').size()
    summary = pd.DataFrame(index=team_members)
    summary.index.name = 'Team Member'
    summary['Gross Contributions (WO,PTR,TR)'] = summary.index.str.lower().map(gross).fillna(0).astype(int)
    summary['Exempted'] = summary.index.str.lower().map(exempted).fillna(0).astype(int)
    summary['Net Contributions'] = summary['Gross Contributions (WO,PTR,TR)'] - summary['Exempted']
    return summary

def display_top_performers(weekly_summary, monthly_summary, selected_month_str):
    """NEW: Displays the top weekly and monthly contributors."""
    st.header("🏆 Top Performers")
    col1, col2 = st.columns(2)
    
    # Weekly Top Performer
    if not weekly_summary.empty:
        top_weekly_contributor = weekly_summary['Net Contributions'].idxmax()
        top_weekly_count = weekly_summary['Net Contributions'].max()
        col1.metric(
            label="Top Contributor of the Week",
            value=top_weekly_contributor,
            help=f"Based on Net Contributions for the selected week."
        )
        col1.write(f"**Net Contributions:** {top_weekly_count}")

    # Monthly Top Performer
    if not monthly_summary.empty:
        top_monthly_contributor = monthly_summary['Net Contributions'].idxmax()
        top_monthly_count = monthly_summary['Net Contributions'].max()
        col2.metric(
            label=f"Top Contributor for {selected_month_str}",
            value=top_monthly_contributor,
            help=f"Based on Net Contributions for the selected month."
        )
        col2.write(f"**Net Contributions:** {top_monthly_count}")

def display_target_analysis(summary_df):
    """Displays the UCS Bangalore target analysis table."""
    st.header("🚀 UCS Bangalore Target Analysis (Full Week: Mon-Sun)")
    summary_df[f'Needed for Target ({INDIVIDUAL_TARGET})'] = (INDIVIDUAL_TARGET - summary_df['Net Contributions']).clip(lower=0)
    summary_df.rename(columns={'Net Contributions': 'Net Contributions (For Target)'}, inplace=True)
    st.dataframe(
        summary_df.style.format("{:d}").background_gradient(
            cmap='Greens', subset=['Net Contributions (For Target)']
        ).background_gradient(
            cmap='Oranges', subset=['Exempted']
        ).background_gradient(
            cmap='Blues', subset=['Gross Contributions (WO,PTR,TR)']
        ),
        use_container_width=True
    )

# --- All other display and helper functions remain the same ---
# (Omitted for brevity, but are included in the final script below)
def display_email_tool(ucs_summary_df, selected_year_week):
    with st.expander("📧 Send Email Notifications to Team Members (via Outlook)"):
        if win32 is None:
            st.warning("Email functionality is disabled because 'pywin32' is not installed. This is expected on deployed apps.")
            return
        st.info(f"Emails will be sent from **{SENDER_EMAIL}** (requires Outlook and permissions).")
        recipients = st.multiselect("Select recipients:", options=ucs_summary_df.index.tolist(), default=[])
        if st.button("✉️ Send Selected Emails via Outlook"):
            if not recipients: st.warning("Please select at least one recipient.")
            else:
                with st.spinner("Sending emails..."):
                    for name in recipients:
                        person_data = ucs_summary_df.loc[name]
                        recipient_email = EMAIL_MAPPING.get(name.lower())
                        if not recipient_email:
                            st.warning(f"No email found for {name}. Skipping.")
                            continue
                        subject = f"Your Weekly Contribution Summary - {selected_year_week}"
                        body = f"Hi {name.split(' ')[0]},\n\nHere is your performance summary for week {selected_year_week}:\n\n- Your Net Contributions: {person_data['Net Contributions (For Target)']}\n- Activities Needed to Meet Target ({INDIVIDUAL_TARGET}): {person_data[f'Needed for Target ({INDIVIDUAL_TARGET})']}\n\nThank you!\nTeam Management"
                        send_email_with_outlook(recipient_email, subject, body)
                    st.success("Email sending process complete.")

def send_email_with_outlook(recipient_email, subject, body):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = SENDER_EMAIL
        mail.To = recipient_email
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        return True
    except Exception as e:
        st.error(f"Failed to send email to {recipient_email}: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def display_drill_down_analysis(df_week_all_days):
    st.header("🔍 Detailed Activity Drill-Down")
    st.info("Select a team member to see every record that **ended** in this ISO week and its status.")
    member_list = sorted(df_week_all_days['Team Member'].unique())
    display_member_list = [LOWER_TO_ORIGINAL_CASE_MAP.get(m, m.title()) for m in member_list]
    selected_member_display = st.selectbox("Select a Team Member to Analyze:", display_member_list)
    if selected_member_display:
        selected_member_lower = selected_member_display.lower()
        member_df = df_week_all_days[df_week_all_days['Team Member'] == selected_member_lower].copy()
        is_contrib_type = member_df['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)
        has_exempt_tag = member_df['tags'].str.contains(EXEMPTION_REGEX, na=False)
        conditions = [~is_contrib_type, is_contrib_type & has_exempt_tag]
        choices = ["❌ Excluded (Not a WO, PTR, or TR)", "⚠️ Excluded (Exempted Tag)"]
        member_df['Status'] = np.select(conditions, choices, default="✅ Included in Net Count")
        st.subheader(f"All Activities Ending in Selected Week for: {selected_member_display}")
        st.dataframe(
            member_df[['record_id', 'Record Type', 'start_date', 'end_date', 'title', 'tags', 'Status']].sort_values("Status"),
            use_container_width=True,
            column_config={
                "start_date": st.column_config.DatetimeColumn("Time In (UTC)", format="YYYY-MM-DD HH:mm"),
                "end_date": st.column_config.DatetimeColumn("Time Out (UTC)", format="YYYY-MM-DD HH:mm")
            }
        )

def display_all_teams_contribution(df_week_all_days):
    st.markdown("---")
    st.header("📊 Full DAPPATC Team Contributions")
    st.markdown("This table shows the contributions for all team members in the dataset for the selected week.")
    df_contrib = df_week_all_days[df_week_all_days['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)].copy()
    gross = df_contrib.groupby('Team Member').size()
    exempted = df_contrib[df_contrib['tags'].str.contains(EXEMPTION_REGEX, na=False)].groupby('Team Member').size()
    summary = pd.DataFrame({'Gross Contributions': gross, 'Exempted': exempted}).fillna(0).astype(int)
    summary['Net Contributions'] = summary['Gross Contributions'] - summary['Exempted']
    summary.index = [LOWER_TO_ORIGINAL_CASE_MAP.get(name, name.title()) for name in summary.index]
    st.dataframe(summary.sort_values('Net Contributions', ascending=False), use_container_width=True)

def main():
    st.title("🏆 Team Performance & Target Dashboard")
    st.info(
        "💡 **Correct Logic:** Data is grouped by the **End Date (Time_out)** of an activity. "
        "A contribution is a record of type **WO, PTR, or TR**, minus exemptions."
    )

    st.sidebar.header("⚙️ Data Source")
    if 'df_full' not in st.session_state: st.session_state.df_full = pd.DataFrame()

    if st.sidebar.button("🚀 Fetch & Analyze Live Data", type="primary"):
        with st.spinner("Fetching and processing latest data..."):
            raw_data = fetch_data_from_api()
            if raw_data:
                st.session_state.df_full = load_and_process_data(raw_data)
                if not st.session_state.df_full.empty: st.sidebar.success("Data processed successfully!")

    if st.session_state.df_full.empty:
        st.info("⬅️ Click 'Fetch & Analyze Live Data' in the sidebar to begin.")
        return

    # --- Sidebar Filters ---
    st.sidebar.markdown("---")
    st.sidebar.header("Weekly Analysis")
    available_year_weeks = sorted(st.session_state.df_full["year_week"].unique(), reverse=True)
    selected_year_week = st.sidebar.selectbox("Select a Week to Analyze", options=available_year_weeks, index=0)

    st.sidebar.markdown("---")
    st.sidebar.header("Monthly Analysis")
    # Create a list of available months in 'YYYY-MM' format for filtering
    available_months = sorted(st.session_state.df_full['end_date'].dt.to_period('M').unique().astype(str), reverse=True)
    selected_month = st.sidebar.selectbox("Select a Month to Analyze", options=available_months, index=0)

    # --- Data Slicing and Calculations ---
    df_week = st.session_state.df_full[st.session_state.df_full["year_week"] == selected_year_week]
    df_month = st.session_state.df_full[st.session_state.df_full['end_date'].dt.to_period('M').astype(str) == selected_month]
    
    ucs_weekly_summary = calculate_contribution_summary(df_week[df_week['Team'] == 'UCS Bangalore'], ORIGINAL_CASE_TEAM)
    ucs_monthly_summary = calculate_contribution_summary(df_month[df_month['Team'] == 'UCS Bangalore'], ORIGINAL_CASE_TEAM)
    
    # --- Display Dashboard Sections ---
    selected_month_str = datetime.strptime(selected_month, "%Y-%m").strftime("%B %Y")
    display_top_performers(ucs_weekly_summary, ucs_monthly_summary, selected_month_str)
    st.markdown("---")
    
    display_target_analysis(ucs_weekly_summary)
    display_email_tool(ucs_weekly_summary, selected_year_week)
    
    display_drill_down_analysis(df_week)
    
    display_all_teams_contribution(df_week)

if __name__ == "__main__":
    main()