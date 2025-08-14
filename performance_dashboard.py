import pandas as pd
import streamlit as st
from io import StringIO
import plotly.express as px
from datetime import datetime
import numpy as np
from typing import Optional, Tuple, List

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

# === API Settings (Modified for Office Network) ===
API_HOST = "dashproach.amadeus.net"
API_IP = "10.57.52.6"  # Resolved from nslookup in office network
API_BASE_URL = f"https://{API_IP}/api/record/DAPPATC/teamactivity"  # Use IP directly

# =========================================================
# FETCH HELPERS — On-demand per month, with caching
# =========================================================
@st.cache_data(ttl=600)
def fetch_month_csv(year: int, month: int) -> Optional[str]:
    """Fetch raw CSV for a specific (year, month).
    Tries hostname first, then IP with Host header (office network case).
    Returns CSV text or None on failure.
    """
    # Try hostname
    try:
        url = f"https://{API_HOST}/api/record/DAPPATC/teamactivity?year={year}&month={month}"
        r = requests.get(url, timeout=20, verify=False)
        r.raise_for_status()
        st.caption(f"✅ {year}-{month:02d}: fetched via hostname")
        return r.text
    except Exception as e:
        st.caption(f"⚠️ {year}-{month:02d}: hostname failed ({e}); trying IP fallback")

    # Fallback to IP + Host header
    try:
        url = f"https://{API_IP}/api/record/DAPPATC/teamactivity?year={year}&month={month}"
        r = requests.get(url, timeout=20, verify=False, headers={"Host": API_HOST})
        r.raise_for_status()
        st.caption(f"✅ {year}-{month:02d}: fetched via IP fallback")
        return r.text
    except Exception as e:
        st.error(f"❌ {year}-{month:02d}: failed via hostname and IP: {e}")
        return None

@st.cache_data(ttl=600)
def fetch_months_csv(year: int, months: Tuple[int, ...]) -> Optional[str]:
    """Fetch and join CSV for several months of a year."""
    parts: List[str] = []
    for m in months:
        txt = fetch_month_csv(year, int(m))
        if txt:
            parts.append(txt)
    if not parts:
        return None
    return "\n".join(parts)

# =========================================================
# LOAD & TRANSFORM
# =========================================================
@st.cache_data(ttl=600)
def load_and_process_data(raw_csv_data: str) -> pd.DataFrame:
    """Loads and processes raw data, assigning ISO week based on END DATE,
    plus convenient 'year' and 'month' columns."""
    if not raw_csv_data:
        return pd.DataFrame()
    try:
        df = pd.read_csv(StringIO(raw_csv_data), header=None, names=CSV_COLUMN_NAMES)
        df.rename(columns={"assignee_name": "Team Member", "rec_type": "Record Type"}, inplace=True)

        df["Team Member"] = df["Team Member"].str.lower().fillna("unassigned")
        df["tags"] = df["tags"].str.lower().fillna("")
        df["start_date"] = pd.to_datetime(df["start_date"], errors='coerce', utc=True)
        df["end_date"] = pd.to_datetime(df["end_date"], errors='coerce', utc=True)
        df.dropna(subset=["start_date", "end_date"], inplace=True)

        iso_cal = df["end_date"].dt.isocalendar()
        df["year_week"] = iso_cal["year"].astype(str) + "-W" + iso_cal["week"].astype(str).str.zfill(2)
        df["year"] = df["end_date"].dt.year
        df["month"] = df["end_date"].dt.to_period("M").astype(str)

        df['Team'] = df['Team Member'].apply(
            lambda x: 'UCS Bangalore' if x in UCS_BANGALORE_TEAM_LOWER else 'Other Teams'
        )
        return df
    except Exception as e:
        st.error(f"Error parsing data: {e}")
        return pd.DataFrame()

# =========================================================
# METRICS & DISPLAYS (existing logic preserved)
# =========================================================
def calculate_contribution_summary(df: pd.DataFrame, team_members: List[str]) -> pd.DataFrame:
    df_contributions = df[df['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)].copy()
    gross = df_contributions.groupby('Team Member').size()
    exempted = df_contributions[df_contributions['tags'].str.contains(EXEMPTION_REGEX, na=False)].groupby('Team Member').size()
    summary = pd.DataFrame(index=team_members)
    summary.index.name = 'Team Member'
    summary['Gross Contributions (WO,PTR,TR)'] = summary.index.str.lower().map(gross).fillna(0).astype(int)
    summary['Exempted'] = summary.index.str.lower().map(exempted).fillna(0).astype(int)
    summary['Net Contributions'] = summary['Gross Contributions (WO,PTR,TR)'] - summary['Exempted']
    return summary

def total_net_contributions(df_slice: pd.DataFrame) -> int:
    if df_slice.empty:
        return 0
    dfc = df_slice[df_slice['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)].copy()
    gross = len(dfc)
    exempted = dfc['tags'].str.contains(EXEMPTION_REGEX, na=False).sum()
    return int(gross - exempted)

def display_top_performers(weekly_summary: pd.DataFrame, monthly_summary: pd.DataFrame, selected_month_str: str):
    st.header("🏆 Top Performers")
    col1, col2 = st.columns(2)

    if not weekly_summary.empty and weekly_summary['Net Contributions'].sum() > 0:
        top_weekly_contributor = weekly_summary['Net Contributions'].idxmax()
        top_weekly_count = int(weekly_summary['Net Contributions'].max())
        col1.metric("Top Contributor of the Week", value=top_weekly_contributor,
                    help="Based on Net Contributions for the selected week.")
        col1.write(f"**Net Contributions:** {top_weekly_count}")

    if not monthly_summary.empty and monthly_summary['Net Contributions'].sum() > 0:
        top_monthly_contributor = monthly_summary['Net Contributions'].idxmax()
        top_monthly_count = int(monthly_summary['Net Contributions'].max())
        col2.metric(f"Top Contributor for {selected_month_str}", value=top_monthly_contributor,
                    help="Based on Net Contributions for the selected month.")
        col2.write(f"**Net Contributions:** {top_monthly_count}")

def display_target_analysis(summary_df: pd.DataFrame):
    st.header("🚀 UCS Bangalore Target Analysis (Full Week: Mon–Sun)")
    summary_df[f'Needed for Target ({INDIVIDUAL_TARGET})'] = (INDIVIDUAL_TARGET - summary_df['Net Contributions']).clip(lower=0)
    summary_df.rename(columns={'Net Contributions': 'Net Contributions (For Target)'}, inplace=True)
    st.dataframe(
        summary_df.style.format("{:d}")
            .background_gradient(cmap='Greens', subset=['Net Contributions (For Target)'])
            .background_gradient(cmap='Oranges', subset=['Exempted'])
            .background_gradient(cmap='Blues', subset=['Gross Contributions (WO,PTR,TR)']),
        use_container_width=True
    )

def display_email_tool(ucs_summary_df: pd.DataFrame, selected_year_week: str):
    with st.expander("📧 Send Email Notifications to Team Members (via Outlook)"):
        if win32 is None:
            st.warning("Email functionality is disabled because 'pywin32' is not installed.")
            return
        st.info(f"Emails will be sent from **{SENDER_EMAIL}**.")
        recipients = st.multiselect("Select recipients:", options=ucs_summary_df.index.tolist(), default=[])
        if st.button("✉️ Send Selected Emails via Outlook"):
            if not recipients:
                st.warning("Please select at least one recipient.")
            else:
                with st.spinner("Sending emails..."):
                    for name in recipients:
                        person_data = ucs_summary_df.loc[name]
                        recipient_email = EMAIL_MAPPING.get(name.lower())
                        if not recipient_email:
                            st.warning(f"No email found for {name}. Skipping.")
                            continue
                        subject = f"Your Weekly Contribution Summary - {selected_year_week}"
                        body = (f"Hi {name.split(' ')[0]},\n\nHere is your performance summary for week {selected_year_week}:\n\n"
                                f"- Your Net Contributions: {person_data['Net Contributions (For Target)']}\n"
                                f"- Activities Needed to Meet Target ({INDIVIDUAL_TARGET}): "
                                f"{person_data[f'Needed for Target ({INDIVIDUAL_TARGET})']}\n\nThank you!\nTeam Management")
                        send_email_with_outlook(recipient_email, subject, body)
                    st.success("Email sending process complete.")

def send_email_with_outlook(recipient_email: str, subject: str, body: str):
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

def display_drill_down_analysis(df_week_all_days: pd.DataFrame):
    st.header("🔍 Detailed Activity Drill-Down")
    if df_week_all_days.empty:
        st.info("No records for the selected week.")
        return

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

def display_all_teams_contribution(df_week_all_days: pd.DataFrame):
    st.markdown("---")
    st.header("📊 Full DAPPATC Team Contributions (Selected Week)")
    if df_week_all_days.empty:
        st.info("No records for the selected week.")
        return

    df_contrib = df_week_all_days[df_week_all_days['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)].copy()
    gross = df_contrib.groupby('Team Member').size()
    exempted = df_contrib[df_contrib['tags'].str.contains(EXEMPTION_REGEX, na=False)].groupby('Team Member').size()
    summary = pd.DataFrame({'Gross Contributions': gross, 'Exempted': exempted}).fillna(0).astype(int)
    summary['Net Contributions'] = summary['Gross Contributions'] - summary['Exempted']
    summary.index = [LOWER_TO_ORIGINAL_CASE_MAP.get(name, name.title()) for name in summary.index]
    st.dataframe(summary.sort_values('Net Contributions', ascending=False), use_container_width=True)

def display_ucs_share(df_week: pd.DataFrame, df_month: pd.DataFrame,
                      ucs_weekly_summary: pd.DataFrame, ucs_monthly_summary: pd.DataFrame,
                      selected_month_str: str):
    st.markdown("---")
    st.header("🧭 UCS Contribution Share")

    ucs_net_week = int(ucs_weekly_summary['Net Contributions'].sum()) if not ucs_weekly_summary.empty else 0
    all_net_week = total_net_contributions(df_week)
    pct_week = (ucs_net_week / all_net_week * 100.0) if all_net_week > 0 else 0.0

    ucs_net_month = int(ucs_monthly_summary['Net Contributions'].sum()) if not ucs_monthly_summary.empty else 0
    all_net_month = total_net_contributions(df_month)
    pct_month = (ucs_net_month / all_net_month * 100.0) if all_net_month > 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("UCS Net (Week)", value=ucs_net_week)
    c2.metric("UCS Share (Week)", value=f"{pct_week:.1f}%")
    c3.metric(f"UCS Net ({selected_month_str})", value=ucs_net_month)
    c4.metric(f"UCS Share ({selected_month_str})", value=f"{pct_month:.1f}%")

    pie_week = px.pie(values=[ucs_net_week, max(all_net_week - ucs_net_week, 0)],
                      names=["UCS", "Others"], title="Weekly Share", hole=0.5)
    pie_month = px.pie(values=[ucs_net_month, max(all_net_month - ucs_net_month, 0)],
                       names=["UCS", "Others"], title=f"{selected_month_str} Share", hole=0.5)
    p1, p2 = st.columns(2)
    p1.plotly_chart(pie_week, use_container_width=True)
    p2.plotly_chart(pie_month, use_container_width=True)

# =========================================================
# YEARLY EXPLORER (Monthly tops + Weekly heatmap & trends)
# =========================================================
def _net_flag(df: pd.DataFrame) -> pd.Series:
    """1 if counts toward Net Contributions, else 0."""
    is_contrib = df['Record Type'].isin(CONTRIBUTION_RECORD_TYPES)
    is_exempt = df['tags'].str.contains(EXEMPTION_REGEX, na=False)
    return (is_contrib & ~is_exempt).astype(int)

@st.cache_data(ttl=600)
def compute_yearly_monthly_tops(df_full: pd.DataFrame, year: int) -> pd.DataFrame:
    """Rows per month, showing top contributor (UCS Bangalore only) and net count."""
    if df_full.empty:
        return pd.DataFrame(columns=["month", "top_contributor", "top_net"])

    df_year = df_full[(df_full['year'] == year) & (df_full['Team'] == 'UCS Bangalore')].copy()
    if df_year.empty:
        return pd.DataFrame(columns=["month", "top_contributor", "top_net"])

    df_year["net"] = _net_flag(df_year)
    monthly_member = (df_year.groupby(['month', 'Team Member'], as_index=False)['net'].sum())

    tops = (monthly_member.sort_values(['month', 'net'], ascending=[True, False])
            .groupby('month', as_index=False).first())
    tops.rename(columns={'Team Member': 'top_contributor', 'net': 'top_net'}, inplace=True)

    tops['top_contributor'] = tops['top_contributor'].apply(
        lambda x: LOWER_TO_ORIGINAL_CASE_MAP.get(x, x.title())
    )
    tops['month_dt'] = pd.PeriodIndex(tops['month'], freq="M").to_timestamp()
    tops = tops.sort_values('month_dt')
    return tops[['month', 'top_contributor', 'top_net']]

@st.cache_data(ttl=600)
def compute_yearly_weekly_matrix(df_full: pd.DataFrame, year: int) -> pd.DataFrame:
    """Pivot (rows=member, cols=year_week, values=net) for UCS Bangalore in given year."""
    if df_full.empty:
        return pd.DataFrame()

    df_year = df_full[(df_full['year'] == year) & (df_full['Team'] == 'UCS Bangalore')].copy()
    if df_year.empty:
        return pd.DataFrame()

    df_year["net"] = _net_flag(df_year)
    weekly_member = (df_year.groupby(['year_week', 'Team Member'], as_index=False)['net'].sum())

    # order weeks chronologically by constructing a date anchor (Monday of that ISO week)
    year_int = weekly_member['year_week'].str.slice(0, 4).astype(int)
    week_int = weekly_member['year_week'].str[-2:].astype(int)
    weekly_member['week_start'] = pd.to_datetime(
        year_int.astype(str) + "-W" + week_int.astype(str) + "-1",
        format="%G-W%V-%u"
    )
    weekly_member = weekly_member.sort_values('week_start')

    pivot = weekly_member.pivot_table(index='Team Member', columns='year_week',
                                      values='net', fill_value=0, aggfunc='sum')
    pivot.index = [LOWER_TO_ORIGINAL_CASE_MAP.get(i, i.title()) for i in pivot.index]
    return pivot

# === NEW: Per-member Monthly matrix for trend ===
@st.cache_data(ttl=600)
def compute_yearly_monthly_matrix(df_full: pd.DataFrame, year: int) -> pd.DataFrame:
    """Pivot (rows=member, cols=month 'YYYY-MM', values=net) for UCS Bangalore in given year."""
    if df_full.empty:
        return pd.DataFrame()

    df_year = df_full[(df_full['year'] == year) & (df_full['Team'] == 'UCS Bangalore')].copy()
    if df_year.empty:
        return pd.DataFrame()

    df_year["net"] = _net_flag(df_year)

    monthly_member = (df_year.groupby(['month', 'Team Member'], as_index=False)['net'].sum())
    # Ensure chronological month order
    monthly_member['month_dt'] = pd.PeriodIndex(monthly_member['month'], freq="M").to_timestamp()
    monthly_member = monthly_member.sort_values('month_dt')

    pivot = monthly_member.pivot_table(index='Team Member', columns='month',
                                       values='net', fill_value=0, aggfunc='sum')
    pivot.index = [LOWER_TO_ORIGINAL_CASE_MAP.get(i, i.title()) for i in pivot.index]
    # Reorder columns chronologically
    cols_sorted = sorted(list(pivot.columns), key=lambda m: pd.Period(m, freq="M").to_timestamp())
    pivot = pivot[cols_sorted]
    return pivot

def display_yearly_explorer(df_full: pd.DataFrame):
    st.markdown("---")
    st.header("📅 Yearly Explorer (UCS Bangalore)")

    if df_full.empty:
        st.info("No data available to explore.")
        return

    years = sorted(df_full['year'].dropna().unique().tolist(), reverse=True)
    year = st.selectbox("Select Year", options=years, index=0)

    # --- Monthly Top Contributors ---
    tops = compute_yearly_monthly_tops(df_full, year)
    st.subheader("🏅 Top Contributor by Month")
    if tops.empty:
        st.info("No UCS Bangalore data for the selected year.")
    else:
        st.dataframe(
            tops.rename(columns={
                "month": "Month",
                "top_contributor": "Top Contributor",
                "top_net": "Net Contributions"
            }),
            use_container_width=True
        )

        bar = px.bar(
            tops.sort_values("month"),
            x="month",
            y="top_net",
            text="top_contributor",
            labels={"month": "Month", "top_net": "Net Contributions"},
            title=f"Top Contributor Per Month – {year}"
        )
        bar.update_traces(textposition="outside")
        st.plotly_chart(bar, use_container_width=True)

    # --- Weekly Contributions (Heatmap + Member Trend) ---
    st.subheader("📈 Weekly Contributions (Net) – Heatmap")
    weekly_pivot = compute_yearly_weekly_matrix(df_full, year)
    if weekly_pivot.empty:
        st.info("No weekly UCS Bangalore data for the selected year.")
    else:
        heatmap = px.imshow(
            weekly_pivot.values,
            labels=dict(x="ISO Week", y="Team Member", color="Net"),
            x=list(weekly_pivot.columns),
            y=list(weekly_pivot.index),
            aspect="auto",
            title=f"Weekly Net Contributions – {year}"
        )
        st.plotly_chart(heatmap, use_container_width=True)

        st.subheader("👤 Weekly Trend by Member")
        members_sorted_w = list(weekly_pivot.index)
        chosen_member_w = st.selectbox("Choose a Team Member (Weekly Trend)", options=members_sorted_w, index=0)
        if chosen_member_w:
            member_weekly = weekly_pivot.loc[chosen_member_w]
            trend_df = pd.DataFrame({
                "year_week": member_weekly.index,
                "net": member_weekly.values
            })
            yy = trend_df['year_week'].str.slice(0, 4).astype(int)
            ww = trend_df['year_week'].str[-2:].astype(int)
            trend_df["week_start"] = pd.to_datetime(
                yy.astype(str) + "-W" + ww.astype(str) + "-1",
                format="%G-W%V-%u"
            )
            trend_df = trend_df.sort_values("week_start")

            line = px.line(
                trend_df,
                x="year_week",
                y="net",
                markers=True,
                labels={"year_week": "ISO Week", "net": "Net Contributions"},
                title=f"Weekly Net Contributions – {chosen_member_w} ({year})"
            )
            # --- NEW FEATURE: Target line at 15 on weekly trend chart ---
            line.add_hline(
                y=INDIVIDUAL_TARGET,
                line_dash="dash",
                line_color="darkred",   
                annotation_text=f"Target {INDIVIDUAL_TARGET}",
                annotation_position="top left"
            )
            st.plotly_chart(line, use_container_width=True)

    # --- NEW: Per-member Monthly Trend ---
    st.subheader("📅 Per-member Monthly Trend")
    monthly_pivot = compute_yearly_monthly_matrix(df_full, year)
    if monthly_pivot.empty:
        st.info("No monthly UCS Bangalore data for the selected year.")
        return

    members_sorted_m = list(monthly_pivot.index)
    chosen_member_m = st.selectbox("Choose a Team Member (Monthly Trend)", options=members_sorted_m, index=0)
    if chosen_member_m:
        m_series = monthly_pivot.loc[chosen_member_m]
        m_trend = pd.DataFrame({
            "month": m_series.index,
            "net": m_series.values
        })
        # Chronological sort (columns were already sorted, but be safe)
        m_trend["month_dt"] = pd.PeriodIndex(m_trend["month"], freq="M").to_timestamp()
        m_trend = m_trend.sort_values("month_dt")

        m_line = px.line(
            m_trend,
            x="month",
            y="net",
            markers=True,
            labels={"month": "Month", "net": "Net Contributions"},
            title=f"Monthly Net Contributions – {chosen_member_m} ({year})"
        )
        st.plotly_chart(m_line, use_container_width=True)

        # Optional: show table for quick copy
        with st.expander("Show monthly values table"):
            # --- NEW FEATURE: include total of all months in the same table ---
            table_df = m_trend[["month", "net"]].rename(columns={"month": "Month", "net": "Net"}).copy()
            total_val = int(table_df["Net"].sum())
            table_df = pd.concat(
                [table_df, pd.DataFrame([{"Month": "Total", "Net": total_val}])],
                ignore_index=True
            )
            st.dataframe(table_df, use_container_width=True)

# =========================================================
# MAIN
# =========================================================
def main():
    st.title("🏆 Team Performance & Target Dashboard")

    st.sidebar.header("⚙️ Data Source")
    if 'df_full' not in st.session_state:
        st.session_state.df_full = pd.DataFrame()

    # --- Year & Month picker for on-demand fetch ---
    this_year = datetime.now().year
    year = st.sidebar.number_input("Year", min_value=2020, max_value=this_year + 1,
                                   value=this_year, step=1)

    all_month_labels = [f"{i:02d}" for i in range(1, 13)]
    default_month = datetime.now().month
    months_selected_labels = st.sidebar.multiselect(
        "Months",
        options=all_month_labels,
        default=[f"{default_month:02d}"],
        help="Pick one or more months to fetch"
    )

    fetch_col1, fetch_col2 = st.sidebar.columns([1, 1])
    replace_data = fetch_col1.checkbox(
        "Replace existing data", value=True,
        help="If unchecked, new data will append to current dataset."
    )

    if fetch_col2.button("🚀 Fetch Selected Months", type="primary"):
        if not months_selected_labels:
            st.sidebar.warning("Select at least one month.")
        else:
            with st.spinner("Fetching and processing selected month(s)..."):
                months_tuple = tuple(int(m) for m in months_selected_labels)
                raw_data = fetch_months_csv(year, months_tuple)
                if raw_data:
                    new_df = load_and_process_data(raw_data)
                    if not new_df.empty:
                        if replace_data or st.session_state.df_full.empty:
                            st.session_state.df_full = new_df
                        else:
                            st.session_state.df_full = (
                                pd.concat([st.session_state.df_full, new_df], ignore_index=True)
                                  .drop_duplicates()
                            )
                        st.sidebar.success(f"Loaded {len(new_df)} rows for {year}-{','.join(months_selected_labels)}")
                    else:
                        st.sidebar.error("Parsing returned no rows.")
                else:
                    st.sidebar.error("No data returned for the selected months.")

    # If no data yet, prompt the user
    if st.session_state.df_full.empty:
        st.info("⬅️ Choose Year/Months and click **Fetch Selected Months** to begin.")
        return

    # --- Weekly & Monthly selectors for analysis (based on loaded data) ---
    st.sidebar.markdown("---")
    st.sidebar.header("Weekly Analysis")
    available_year_weeks = sorted(st.session_state.df_full["year_week"].unique(), reverse=True)
    selected_year_week = st.sidebar.selectbox("Select a Week to Analyze", options=available_year_weeks, index=0)

    st.sidebar.markdown("---")
    st.sidebar.header("Monthly Analysis")
    available_months = sorted(st.session_state.df_full['end_date'].dt.to_period('M').unique().astype(str), reverse=True)
    selected_month = st.sidebar.selectbox("Select a Month to Analyze", options=available_months, index=0)

    # --- Filtered frames ---
    df_week = st.session_state.df_full[st.session_state.df_full["year_week"] == selected_year_week]
    df_month = st.session_state.df_full[st.session_state.df_full['end_date'].dt.to_period('M').astype(str) == selected_month]

    # --- UCS summaries ---
    ucs_weekly_summary = calculate_contribution_summary(
        df_week[df_week['Team'] == 'UCS Bangalore'], ORIGINAL_CASE_TEAM
    )
    ucs_monthly_summary = calculate_contribution_summary(
        df_month[df_month['Team'] == 'UCS Bangalore'], ORIGINAL_CASE_TEAM
    )

    selected_month_str = datetime.strptime(selected_month, "%Y-%m").strftime("%B %Y")

    # --- Views ---
    display_top_performers(ucs_weekly_summary, ucs_monthly_summary, selected_month_str)

    st.markdown("---")
    display_ucs_share(df_week, df_month, ucs_weekly_summary, ucs_monthly_summary, selected_month_str)

    display_target_analysis(ucs_weekly_summary)

    display_email_tool(ucs_weekly_summary, selected_year_week)

    display_drill_down_analysis(df_week)

    display_all_teams_contribution(df_week)

    # Yearly explorer at the end (includes Weekly + Monthly trends per member)
    display_yearly_explorer(st.session_state.df_full)


if __name__ == "__main__":
    main()
# Run the Streamlit app
# To run, use: streamlit run performance_dashboard.py
# Ensure you have the required libraries installed:
# pip install streamlit pandas plotly pythoncom pywin32 requests
# Note: The Outlook email functionality requires 'pywin32' to be installed.
# If running outside of the office network, ensure the API_HOST is reachable.