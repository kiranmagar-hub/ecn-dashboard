import pandas as pd
import json
from datetime import datetime

# Read the Excel file
# Check if we're running from schedule_daily_dashboard.py (uses source_data.xlsx)
# or manually (uses BEF ECN FY27.xlsx)
import os
if os.path.exists('source_data.xlsx'):
    excel_file = 'source_data.xlsx'
elif os.path.exists('BEF ECN FY27.xlsx'):
    excel_file = 'BEF ECN FY27.xlsx'
else:
    raise FileNotFoundError("Could not find 'source_data.xlsx' or 'BEF ECN FY27.xlsx'")

df = pd.read_excel(excel_file, sheet_name='Document_TB11')

print(f"Loaded {len(df)} records")

# Load coordinator name mapping
try:
    coord_mapping = pd.read_excel('ECN Coordinators.xlsx')
    # Strip whitespace and convert to uppercase for case-insensitive matching
    coord_mapping['Username'] = coord_mapping['Username'].str.strip().str.upper()
    coord_mapping['Coordinator'] = coord_mapping['Coordinator'].str.strip()
    coord_dict = dict(zip(coord_mapping['Username'], coord_mapping['Coordinator']))
    print(f"Loaded {len(coord_dict)} coordinator name mappings")
except FileNotFoundError:
    print("Warning: ECN Coordinators.xlsx not found. Using user IDs instead of names.")
    coord_dict = {}

# Convert date columns to datetime (already done by pandas, but ensure proper format)
date_columns = ['SubmitDate', 'ProcEndTime', 'FinalDate', 'FinalOrHoldDate', 'HoldDate']
for col in date_columns:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')

# Remove rows with missing critical data
df_clean = df.dropna(subset=['SubmitDate']).copy()

# Calculate cycle time columns if they don't exist (for database source)
if 'ProcCT(days)' not in df_clean.columns:
    print("Calculating ProcCT(days) from date columns...")
    # ProcCT = Processing Cycle Time (Submit to Processing/Close date)
    # Use ProcessingDate or ClosedDate as the end point
    if 'ProcessingDate' in df_clean.columns:
        df_clean['FinalOrHoldDate'] = df_clean['ProcessingDate'].fillna(df_clean['ClosedDate']).fillna(df_clean['HoldDate'])
        df_clean['ProcEndTime'] = df_clean['ProcessingDate'].fillna(df_clean['ClosedDate'])
    else:
        df_clean['FinalOrHoldDate'] = df_clean['ClosedDate'].fillna(df_clean['HoldDate'])
        df_clean['ProcEndTime'] = df_clean['ClosedDate']

    df_clean['ProcCT(days)'] = (df_clean['FinalOrHoldDate'] - df_clean['SubmitDate']).dt.days
    df_clean['ProcCT(days)'] = df_clean['ProcCT(days)'].fillna(0)

if 'TotalCT(Days)' not in df_clean.columns:
    print("Calculating TotalCT(Days) from date columns...")
    # TotalCT = Total Cycle Time (Submit to absolute final date)
    df_clean['FinalDate'] = df_clean['ClosedDate'].fillna(df_clean['VoidDate'])
    df_clean['TotalCT(Days)'] = (df_clean['FinalDate'] - df_clean['SubmitDate']).dt.days
    df_clean['TotalCT(Days)'] = df_clean['TotalCT(Days)'].fillna(0)

# Add other required columns if missing
if 'ECNCoordinatorSite' not in df_clean.columns and 'Site' in df_clean.columns:
    df_clean['ECNCoordinatorSite'] = df_clean['Site']

if 'RequestNum' not in df_clean.columns:
    df_clean['RequestNum'] = 1  # Count column for aggregations

# Calculate PendingDays if not present
if 'PendingDays' not in df_clean.columns:
    # PendingDays = days an ECN has been in Pending state
    pending_mask = df_clean['State'].str.contains('Pending', case=False, na=False)
    df_clean['PendingDays'] = 0
    df_clean.loc[pending_mask, 'PendingDays'] = (pd.Timestamp.now() - df_clean.loc[pending_mask, 'SubmitDate']).dt.days

# Filter data to last 365 days (or use full dataset if already filtered by schedule_daily_dashboard.py)
# Only apply date filter if the data hasn't been pre-filtered
if len(df_clean) > 100000:  # If we have a very large dataset, filter it
    end_date = datetime.now()
    start_date = end_date - pd.Timedelta(days=365)
    df_clean = df_clean[
        (df_clean['SubmitDate'] >= start_date) &
        (df_clean['SubmitDate'] <= end_date)
    ].copy()
    print(f"Filtered to last 365 days: {len(df_clean)} records")
else:
    # Data already filtered by schedule_daily_dashboard.py
    print(f"Using pre-filtered data: {len(df_clean)} records")

# Map coordinator user IDs to names (case-insensitive with whitespace handling)
df_clean['CoordinatorName'] = df_clean['ECNCoordinator'].str.strip().str.upper().map(coord_dict).fillna(df_clean['ECNCoordinator'])

# Parse ECN Topic - split by ~ to get main and sub topics
df_clean['ECN Topic'] = df_clean['EcnTopic'].str.split('~').str[0].str.strip().str.upper()
df_clean['ECN Sub Topic'] = df_clean['EcnTopic'].str.split('~').str[1].str.strip().str.upper() if df_clean['EcnTopic'].str.contains('~', na=False).any() else 'NONE'

# For rows without ~, use the full topic as main topic
df_clean.loc[~df_clean['EcnTopic'].str.contains('~', na=False), 'ECN Topic'] = df_clean.loc[~df_clean['EcnTopic'].str.contains('~', na=False), 'EcnTopic'].str.strip().str.upper()
df_clean.loc[~df_clean['EcnTopic'].str.contains('~', na=False), 'ECN Sub Topic'] = 'NONE'

# Filter out ECN Type 9 (excluding from all metrics)
# Matches: "9", "(9)", "(9) Restricted ECN", "9 - Restricted ECN", etc.
df_before_filter = len(df_clean)
df_clean = df_clean[~df_clean['ECN Topic'].str.contains(r'^\(?\s*9[\)\s\-]|^9$', case=False, na=False, regex=True)].copy()
df_after_filter = len(df_clean)
if df_before_filter > df_after_filter:
    print(f"Filtered out {df_before_filter - df_after_filter} ECN Type 9 records")

# Filter out Test and TestClosed states
df_before_test_filter = len(df_clean)
df_clean = df_clean[~df_clean['State'].isin(['Test', 'TestClosed'])].copy()
df_after_test_filter = len(df_clean)
if df_before_test_filter > df_after_test_filter:
    print(f"Filtered out {df_before_test_filter - df_after_test_filter} Test/TestClosed records")

# Calculate monthly trends
df_clean['YearMonth'] = df_clean['SubmitDate'].dt.to_period('M')

# Monthly trends for cycle times
monthly_trends = df_clean.groupby('YearMonth').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()

monthly_trends['YearMonth'] = monthly_trends['YearMonth'].astype(str)

# Calculate quarterly trends
df_clean['YearQuarter'] = df_clean['SubmitDate'].dt.to_period('Q')

# Cycle time by Site - Separate SYSTEM USERID from other coordinators
# First get overall site stats
site_ct_all = df_clean.groupby('ECNCoordinatorSite').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()

# Site stats for SYSTEM USERID only
site_ct_system = df_clean[df_clean['CoordinatorName'] == 'SYSTEM USERID'].groupby('ECNCoordinatorSite').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
site_ct_system['Coordinator'] = 'SYSTEM USERID'

# Site stats for other coordinators (excluding SYSTEM USERID)
site_ct_others = df_clean[df_clean['CoordinatorName'] != 'SYSTEM USERID'].groupby('ECNCoordinatorSite').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
site_ct_others['Coordinator'] = 'Other Coordinators'

# Combined site comparison with separation
site_ct_combined = pd.concat([site_ct_system, site_ct_others]).sort_values('RequestNum', ascending=False)

# Keep original for backward compatibility
site_ct = site_ct_all.sort_values('RequestNum', ascending=False).head(10)

# Cycle time by ECN Coordinator (with Region)
# For SYSTEM USERID, merge all sites together; for others, keep site-specific
df_coord_temp = df_clean[df_clean['ECNCoordinator'].notna()].copy()

# Separate SYSTEM USERID from others
df_system = df_coord_temp[df_coord_temp['CoordinatorName'] == 'SYSTEM USERID']
df_others = df_coord_temp[df_coord_temp['CoordinatorName'] != 'SYSTEM USERID']

# Aggregate SYSTEM USERID across all sites
if len(df_system) > 0:
    system_agg = df_system.groupby('CoordinatorName').agg({
        'ProcCT(days)': 'mean',
        'TotalCT(Days)': 'mean',
        'RequestNum': 'count'
    }).reset_index()
    system_agg['Region'] = 'ALL SITES'
    system_agg = system_agg.rename(columns={'CoordinatorName': 'ECNCoordinator'})
else:
    system_agg = pd.DataFrame()

# Aggregate others by coordinator and site
if len(df_others) > 0:
    others_agg = df_others.groupby(['CoordinatorName', 'ECNCoordinatorSite']).agg({
        'ProcCT(days)': 'mean',
        'TotalCT(Days)': 'mean',
        'RequestNum': 'count'
    }).reset_index()
    others_agg = others_agg.rename(columns={'ECNCoordinatorSite': 'Region', 'CoordinatorName': 'ECNCoordinator'})
else:
    others_agg = pd.DataFrame()

# Combine and sort
if len(system_agg) > 0 and len(others_agg) > 0:
    coordinator_ct = pd.concat([system_agg, others_agg])
elif len(system_agg) > 0:
    coordinator_ct = system_agg
else:
    coordinator_ct = others_agg

coordinator_ct = coordinator_ct.sort_values('RequestNum', ascending=False).head(10)

# Cycle time by ECN Topic
topic_ct = df_clean.groupby('ECN Topic').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
topic_ct = topic_ct.sort_values('RequestNum', ascending=False).head(10)

# Quarterly trends by ECN Topic (top 5 topics) - Last 1 year only
top_5_topics = topic_ct.head(5)['ECN Topic'].tolist()

# Get the last 4 quarters (1 year)
all_quarters = df_clean['YearQuarter'].unique()
all_quarters_sorted = sorted(all_quarters, reverse=True)
last_4_quarters = all_quarters_sorted[:4] if len(all_quarters_sorted) >= 4 else all_quarters_sorted

# Filter data for last 4 quarters and top 5 topics
df_last_year = df_clean[df_clean['YearQuarter'].isin(last_4_quarters) & df_clean['ECN Topic'].isin(top_5_topics)]

quarterly_topic_trends = df_last_year.groupby(['YearQuarter', 'ECN Topic']).agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
quarterly_topic_trends['YearQuarter'] = quarterly_topic_trends['YearQuarter'].astype(str)

# Cycle time by ECN Sub Topic
subtopic_ct = df_clean.groupby('ECN Sub Topic').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
subtopic_ct = subtopic_ct.sort_values('RequestNum', ascending=False).head(15)

# Cycle time by ECN Sub Topic and Status (for filtering)
subtopic_by_status = df_clean.groupby(['ECN Sub Topic', 'State']).agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
# Get top 15 subtopics overall to maintain consistency
top_15_subtopics = subtopic_ct['ECN Sub Topic'].tolist()
subtopic_by_status = subtopic_by_status[subtopic_by_status['ECN Sub Topic'].isin(top_15_subtopics)]
subtopic_by_status = subtopic_by_status.sort_values(['State', 'RequestNum'], ascending=[True, False])

# State distribution
state_dist = df_clean.groupby('State').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean'
}).reset_index()
state_dist = state_dist.sort_values('RequestNum', ascending=False)

# Rush vs Regular comparison
rush_comparison = df_clean.groupby('RushRequest').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()

# 3Z Residual Disposition Analysis
# Filter all 3z subtopics
df_3z = df_clean[df_clean['ECN Sub Topic'].str.contains('3Z', case=False, na=False)].copy()

# Breakdown by specific 3z subtopic
residual_3z_breakdown = df_3z.groupby('ECN Sub Topic').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
residual_3z_breakdown = residual_3z_breakdown.sort_values('RequestNum', ascending=False)

# 3z by Status
residual_3z_by_status = df_3z.groupby('State').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()

# 3z breakdown by status
residual_3z_breakdown_by_status = df_3z.groupby(['ECN Sub Topic', 'State']).agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
residual_3z_breakdown_by_status = residual_3z_breakdown_by_status.sort_values(['State', 'RequestNum'], ascending=[True, False])

# 3z Monthly trends
df_3z['YearMonth'] = df_3z['SubmitDate'].dt.to_period('M')
residual_3z_monthly = df_3z.groupby('YearMonth').agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
residual_3z_monthly['YearMonth'] = residual_3z_monthly['YearMonth'].astype(str)

# 3z Monthly trends by status
residual_3z_monthly_by_status = df_3z.groupby(['YearMonth', 'State']).agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'RequestNum': 'count'
}).reset_index()
residual_3z_monthly_by_status['YearMonth'] = residual_3z_monthly_by_status['YearMonth'].astype(str)

# Overall 3z stats
residual_3z_total = {
    'total_count': int(len(df_3z)),
    'avg_proc_ct': float(df_3z['ProcCT(days)'].mean()) if len(df_3z) > 0 else 0,
    'avg_total_ct': float(df_3z['TotalCT(Days)'].mean()) if len(df_3z) > 0 else 0,
    'median_proc_ct': float(df_3z['ProcCT(days)'].median()) if len(df_3z) > 0 else 0,
    'median_total_ct': float(df_3z['TotalCT(Days)'].median()) if len(df_3z) > 0 else 0
}

# ========================================
# ADVANCED KPIs
# ========================================

# 1. VOID RATE ANALYSIS
void_count = len(df_clean[df_clean['State'] == 'Void'])
total_count = len(df_clean)
void_rate = (void_count / total_count) * 100 if total_count > 0 else 0

void_by_reason = df_clean[df_clean['State'] == 'Void'].groupby('VoidReason').agg({
    'RequestNum': 'count'
}).reset_index().sort_values('RequestNum', ascending=False)

void_by_topic = df_clean.groupby('ECN Topic').agg({
    'RequestNum': 'count',
    'State': lambda x: (x == 'Void').sum()
}).reset_index()
void_by_topic['VoidRate'] = (void_by_topic['State'] / void_by_topic['RequestNum']) * 100
void_by_topic = void_by_topic.sort_values('VoidRate', ascending=False)

void_by_coordinator = df_clean[df_clean['ECNCoordinator'].notna()].groupby('CoordinatorName').agg({
    'RequestNum': 'count',
    'State': lambda x: (x == 'Void').sum()
}).reset_index()
void_by_coordinator = void_by_coordinator.rename(columns={'CoordinatorName': 'ECNCoordinator'})
void_by_coordinator['VoidRate'] = (void_by_coordinator['State'] / void_by_coordinator['RequestNum']) * 100
void_by_coordinator = void_by_coordinator.sort_values('RequestNum', ascending=False)

void_by_site = df_clean.groupby('ECNCoordinatorSite').agg({
    'RequestNum': 'count',
    'State': lambda x: (x == 'Void').sum()
}).reset_index()
void_by_site['VoidRate'] = (void_by_site['State'] / void_by_site['RequestNum']) * 100

# Void trend over time
df_clean['YearMonth'] = df_clean['SubmitDate'].dt.to_period('M')
void_monthly_trend = df_clean.groupby('YearMonth').agg({
    'RequestNum': 'count',
    'State': lambda x: (x == 'Void').sum()
}).reset_index()
void_monthly_trend['VoidRate'] = (void_monthly_trend['State'] / void_monthly_trend['RequestNum']) * 100
void_monthly_trend['YearMonth'] = void_monthly_trend['YearMonth'].astype(str)

# 2. HOLD ANALYSIS
hold_count = df_clean['HoldReason'].notna().sum()
hold_rate = (hold_count / total_count) * 100 if total_count > 0 else 0

hold_by_reason = df_clean[df_clean['HoldReason'].notna()].groupby('HoldReason').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean'
}).reset_index().sort_values('RequestNum', ascending=False)

hold_by_topic = df_clean.groupby('ECN Topic').agg({
    'RequestNum': 'count',
    'HoldReason': lambda x: x.notna().sum()
}).reset_index()
hold_by_topic['HoldRate'] = (hold_by_topic['HoldReason'] / hold_by_topic['RequestNum']) * 100
hold_by_topic = hold_by_topic.sort_values('HoldRate', ascending=False)

hold_by_requestor = df_clean[df_clean['Requestor'].notna()].groupby('Requestor').agg({
    'RequestNum': 'count',
    'HoldReason': lambda x: x.notna().sum()
}).reset_index()
hold_by_requestor['HoldRate'] = (hold_by_requestor['HoldReason'] / hold_by_requestor['RequestNum']) * 100
hold_by_requestor = hold_by_requestor[hold_by_requestor['RequestNum'] >= 10].sort_values('HoldRate', ascending=False).head(20)

# CT comparison with/without holds
ct_with_holds = df_clean[df_clean['HoldReason'].notna()].agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean'
})
ct_without_holds = df_clean[df_clean['HoldReason'].isna()].agg({
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean'
})

hold_ct_comparison = pd.DataFrame({
    'Category': ['With Holds', 'Without Holds'],
    'ProcCT(days)': [ct_with_holds['ProcCT(days)'], ct_without_holds['ProcCT(days)']],
    'TotalCT(Days)': [ct_with_holds['TotalCT(Days)'], ct_without_holds['TotalCT(Days)']]
})

# 3. FIRST-TIME-RIGHT (FTR) RATE
ftr_count = len(df_clean[(df_clean['HoldReason'].isna()) & (df_clean['State'] == 'Closed')])
ftr_rate = (ftr_count / total_count) * 100 if total_count > 0 else 0

ftr_by_coordinator = df_clean[df_clean['ECNCoordinator'].notna()].groupby('CoordinatorName').agg({
    'RequestNum': 'count',
    'HoldReason': lambda x: x.isna().sum(),
    'State': lambda x: ((x == 'Closed')).sum()
}).reset_index()
# Calculate FTR count using a different method to avoid deprecation warning
ftr_counts_coord = df_clean[df_clean['ECNCoordinator'].notna()].groupby('CoordinatorName')[['HoldReason', 'State']].apply(
    lambda x: ((x['HoldReason'].isna()) & (x['State'] == 'Closed')).sum(), include_groups=False
)
ftr_by_coordinator['FTR_Count'] = ftr_by_coordinator['CoordinatorName'].map(ftr_counts_coord)
ftr_by_coordinator = ftr_by_coordinator.rename(columns={'CoordinatorName': 'ECNCoordinator'})
ftr_by_coordinator['FTR_Rate'] = (ftr_by_coordinator['FTR_Count'] / ftr_by_coordinator['RequestNum']) * 100
ftr_by_coordinator = ftr_by_coordinator.sort_values('RequestNum', ascending=False)

ftr_by_requestor = df_clean[df_clean['Requestor'].notna()].groupby('Requestor').agg({
    'RequestNum': 'count'
}).reset_index()
# Calculate FTR count using a different method to avoid deprecation warning
ftr_counts_req = df_clean[df_clean['Requestor'].notna()].groupby('Requestor')[['HoldReason', 'State']].apply(
    lambda x: ((x['HoldReason'].isna()) & (x['State'] == 'Closed')).sum(), include_groups=False
)
ftr_by_requestor['FTR_Count'] = ftr_by_requestor['Requestor'].map(ftr_counts_req)
ftr_by_requestor['FTR_Rate'] = (ftr_by_requestor['FTR_Count'] / ftr_by_requestor['RequestNum']) * 100
ftr_by_requestor = ftr_by_requestor[ftr_by_requestor['RequestNum'] >= 10].sort_values('FTR_Rate', ascending=False).head(20)

# FTR trend over time
ftr_monthly_trend = df_clean.groupby('YearMonth')[['HoldReason', 'State']].apply(
    lambda x: pd.Series({
        'Total': len(x),
        'FTR_Count': ((x['HoldReason'].isna()) & (x['State'] == 'Closed')).sum()
    }), include_groups=False
).reset_index()
ftr_monthly_trend['FTR_Rate'] = (ftr_monthly_trend['FTR_Count'] / ftr_monthly_trend['Total']) * 100
ftr_monthly_trend['YearMonth'] = ftr_monthly_trend['YearMonth'].astype(str)

# 4. COORDINATOR WORKLOAD & PERFORMANCE
coordinator_workload = df_clean[df_clean['ECNCoordinator'].notna()].groupby('CoordinatorName').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean'
}).reset_index()
coordinator_workload = coordinator_workload.rename(columns={'CoordinatorName': 'ECNCoordinator'})
# Filter to only include coordinators with 100+ ECNs (exclude test/one-off user IDs)
coordinator_workload = coordinator_workload[coordinator_workload['RequestNum'] >= 100]
coordinator_workload['Percentage'] = (coordinator_workload['RequestNum'] / coordinator_workload['RequestNum'].sum()) * 100
coordinator_workload = coordinator_workload.sort_values('RequestNum', ascending=False)

coordinator_monthly_workload = df_clean[df_clean['ECNCoordinator'].notna()].groupby(['YearMonth', 'CoordinatorName']).agg({
    'RequestNum': 'count'
}).reset_index()
coordinator_monthly_workload = coordinator_monthly_workload.rename(columns={'CoordinatorName': 'ECNCoordinator'})
coordinator_monthly_workload['YearMonth'] = coordinator_monthly_workload['YearMonth'].astype(str)

# 5. MANUFACTURING SITE ANALYSIS
if 'MFGSiteCode' in df_clean.columns:
    mfg_site_analysis = df_clean[df_clean['MFGSiteCode'].notna()].groupby('MFGSiteCode').agg({
        'RequestNum': 'count',
        'ProcCT(days)': 'mean',
        'TotalCT(Days)': 'mean'
    }).reset_index().sort_values('RequestNum', ascending=False).head(20)
else:
    # Use Site if MFGSiteCode not available
    if 'Site' in df_clean.columns:
        mfg_site_analysis = df_clean[df_clean['Site'].notna()].groupby('Site').agg({
            'RequestNum': 'count',
            'ProcCT(days)': 'mean',
            'TotalCT(Days)': 'mean'
        }).reset_index().rename(columns={'Site': 'MFGSiteCode'}).sort_values('RequestNum', ascending=False).head(20)
    else:
        # Create empty dataframe if neither column exists
        mfg_site_analysis = pd.DataFrame(columns=['MFGSiteCode', 'RequestNum', 'ProcCT(days)', 'TotalCT(Days)'])

# 6. REQUESTOR ANALYSIS
requestor_analysis = df_clean[df_clean['Requestor'].notna()].groupby('Requestor').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean',
    'HoldReason': lambda x: x.notna().sum(),
    'State': lambda x: (x == 'Void').sum(),
    'RushRequest': lambda x: x.notna().sum()
}).reset_index()
requestor_analysis['HoldRate'] = (requestor_analysis['HoldReason'] / requestor_analysis['RequestNum']) * 100
requestor_analysis['VoidRate'] = (requestor_analysis['State'] / requestor_analysis['RequestNum']) * 100
requestor_analysis['RushRate'] = (requestor_analysis['RushRequest'] / requestor_analysis['RequestNum']) * 100
requestor_analysis = requestor_analysis[requestor_analysis['RequestNum'] >= 10].sort_values('RequestNum', ascending=False).head(20)

# 7. RUSH REQUEST DEEP DIVE
rush_count = df_clean['RushRequest'].notna().sum()
rush_rate = (rush_count / total_count) * 100 if total_count > 0 else 0

rush_by_topic = df_clean.groupby('ECN Topic').agg({
    'RequestNum': 'count',
    'RushRequest': lambda x: x.notna().sum()
}).reset_index()
rush_by_topic['RushRate'] = (rush_by_topic['RushRequest'] / rush_by_topic['RequestNum']) * 100
rush_by_topic = rush_by_topic.sort_values('RushRate', ascending=False)

rush_monthly_trend = df_clean.groupby('YearMonth').agg({
    'RequestNum': 'count',
    'RushRequest': lambda x: x.notna().sum()
}).reset_index()
rush_monthly_trend['RushRate'] = (rush_monthly_trend['RushRequest'] / rush_monthly_trend['RequestNum']) * 100
rush_monthly_trend['YearMonth'] = rush_monthly_trend['YearMonth'].astype(str)

rush_success_rate = df_clean.groupby('RushRequest').agg({
    'RequestNum': 'count',
    'State': lambda x: (x == 'Void').sum(),
    'HoldReason': lambda x: x.notna().sum()
}).reset_index()
rush_success_rate['VoidRate'] = (rush_success_rate['State'] / rush_success_rate['RequestNum']) * 100
rush_success_rate['HoldRate'] = (rush_success_rate['HoldReason'] / rush_success_rate['RequestNum']) * 100
rush_success_rate['RushRequest'] = rush_success_rate['RushRequest'].fillna('Regular')

# 8. PENDING DAYS ANALYSIS
if 'PendingDays' in df_clean.columns:
    pending_analysis = df_clean[df_clean['PendingDays'].notna()].agg({
        'PendingDays': ['count', 'mean', 'median', 'max'],
        'TotalCT(Days)': 'mean'
    }) if df_clean['PendingDays'].notna().sum() > 0 else pd.DataFrame()

    if not pending_analysis.empty:
        pending_by_topic = df_clean[df_clean['PendingDays'].notna()].groupby('ECN Topic').agg({
            'PendingDays': ['count', 'mean'],
            'TotalCT(Days)': 'mean'
        }).reset_index()
        pending_by_topic.columns = ['ECN Topic', 'Count', 'AvgPendingDays', 'AvgTotalCT']
        pending_by_topic = pending_by_topic.sort_values('AvgPendingDays', ascending=False)
    else:
        pending_by_topic = pd.DataFrame()
else:
    # PendingDays column doesn't exist - create empty dataframes
    pending_analysis = pd.DataFrame()
    pending_by_topic = pd.DataFrame()

# Filter for Closed ECNs only for percentile calculations
df_closed = df_clean[df_clean['State'] == 'Closed'].copy()

# Calculate percentile thresholds for closed ECNs
p50 = df_closed['ProcCT(days)'].quantile(0.50)
p75 = df_closed['ProcCT(days)'].quantile(0.75)
p90 = df_closed['ProcCT(days)'].quantile(0.90)

# Top 3 categories for each percentile range
fast_range = df_closed[df_closed['ProcCT(days)'] <= p50]
medium_range = df_closed[(df_closed['ProcCT(days)'] > p50) & (df_closed['ProcCT(days)'] <= p75)]
slower_range = df_closed[(df_closed['ProcCT(days)'] > p75) & (df_closed['ProcCT(days)'] <= p90)]

top_categories_50th = fast_range.groupby('ECN Topic').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean'
}).reset_index().sort_values('RequestNum', ascending=False).head(3)

top_categories_75th = medium_range.groupby('ECN Topic').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean'
}).reset_index().sort_values('RequestNum', ascending=False).head(3)

top_categories_90th = slower_range.groupby('ECN Topic').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean'
}).reset_index().sort_values('RequestNum', ascending=False).head(3)

# 90th Percentile Analysis - All ECN types above 90th percentile
# Get all ECNs in 90th percentile (above threshold)
ecns_90th_percentile = df_closed[df_closed['ProcCT(days)'] > p90].copy()

# Explicitly filter out Type 9 (safety check) - matches any format starting with 9
ecns_90th_percentile = ecns_90th_percentile[~ecns_90th_percentile['ECN Topic'].str.contains(r'^\(?\s*9[\)\s\-]|^9$', case=False, na=False, regex=True)].copy()

# Group by ECN Topic for chart - merge all 3Z sub-types together
ecns_90th_for_chart = ecns_90th_percentile.copy()

# Check if main topic contains "(3)"
type_3_mask = ecns_90th_for_chart['ECN Topic'].str.contains(r'\(3\)', case=False, na=False, regex=True)

# Check if sub-topic contains "3Z"
type_3z_mask = ecns_90th_for_chart['ECN Sub Topic'].str.contains('3Z', case=False, na=False)

# For Type 3 + 3Z sub-types: merge all into one label
type_3_and_3z = type_3_mask & type_3z_mask
ecns_90th_for_chart.loc[type_3_and_3z, 'Chart_Topic'] = '(3) 3Z - Residual Disposition'

# For Type 3 + non-3Z sub-types: group all together
type_3_not_3z = type_3_mask & ~type_3z_mask
ecns_90th_for_chart.loc[type_3_not_3z, 'Chart_Topic'] = '(3) Other Sub-types'

# For all other types, use the main topic
ecns_90th_for_chart.loc[~type_3_mask, 'Chart_Topic'] = ecns_90th_for_chart.loc[~type_3_mask, 'ECN Topic']

# Group by the chart topic
percentile_90th_by_topic = ecns_90th_for_chart.groupby('Chart_Topic').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean',
    'TotalCT(Days)': 'mean'
}).reset_index()

# Add a sort key to keep Type 3 entries together
def get_sort_key(topic):
    if '(3) 3Z' in topic:
        return (0, 1)  # Type 3 3Z - show second
    elif '(3) Other' in topic:
        return (0, 0)  # Type 3 Other - show first
    else:
        return (1, 0)  # All others - show after

percentile_90th_by_topic['sort_group'] = percentile_90th_by_topic['Chart_Topic'].apply(get_sort_key)

# Sort by count descending, but keep Type 3 entries together at the top
percentile_90th_by_topic = percentile_90th_by_topic.sort_values(['sort_group', 'RequestNum'], ascending=[True, False])
percentile_90th_by_topic = percentile_90th_by_topic.drop('sort_group', axis=1)

# Rename columns for clarity in chart
percentile_90th_by_topic = percentile_90th_by_topic.rename(columns={
    'Chart_Topic': 'ECN Topic',
    'RequestNum': 'Count',
    'ProcCT(days)': 'AvgProcCT',
    'TotalCT(Days)': 'AvgTotalCT'
})

# Add a flag for 3Z to help with chart coloring (any sub-topic containing "3Z")
percentile_90th_by_topic['Is3Z'] = percentile_90th_by_topic['ECN Topic'].str.contains('3Z', case=False, na=False)

# Overall statistics with percentiles (Closed ECNs only for percentiles)
overall_stats = {
    'total_requests': int(len(df_clean)),
    'total_closed': int(len(df_closed)),
    'total_void': int(len(df_clean[df_clean['State'] == 'Void'])),
    'avg_proc_ct': float(df_clean['ProcCT(days)'].mean()),
    'avg_total_ct': float(df_clean['TotalCT(Days)'].mean()),
    'median_proc_ct': float(df_clean['ProcCT(days)'].median()),
    'median_total_ct': float(df_clean['TotalCT(Days)'].median()),
    'max_proc_ct': float(df_clean['ProcCT(days)'].max()),
    'max_total_ct': float(df_clean['TotalCT(Days)'].max()),
    # Percentiles for CLOSED ECNs only
    'percentile_50_proc_ct': float(df_closed['ProcCT(days)'].quantile(0.50)),
    'percentile_75_proc_ct': float(df_closed['ProcCT(days)'].quantile(0.75)),
    'percentile_90_proc_ct': float(df_closed['ProcCT(days)'].quantile(0.90)),
    'percentile_50_total_ct': float(df_closed['TotalCT(Days)'].quantile(0.50)),
    'percentile_75_total_ct': float(df_closed['TotalCT(Days)'].quantile(0.75)),
    'percentile_90_total_ct': float(df_closed['TotalCT(Days)'].quantile(0.90)),
    # Top 3 categories by percentile range
    'top_categories_50th': top_categories_50th.to_dict('records'),
    'top_categories_75th': top_categories_75th.to_dict('records'),
    'top_categories_90th': top_categories_90th.to_dict('records'),
    'date_range': {
        'start': df_clean['SubmitDate'].min().strftime('%Y-%m-%d'),
        'end': df_clean['SubmitDate'].max().strftime('%Y-%m-%d')
    }
}

# Prepare data for export
data_export = {
    'overall_stats': overall_stats,
    'monthly_trends': monthly_trends.to_dict('records'),
    'quarterly_topic_trends': quarterly_topic_trends.to_dict('records'),
    'site_comparison': site_ct.to_dict('records'),
    'site_comparison_separated': site_ct_combined.to_dict('records'),
    'coordinator_comparison': coordinator_ct.to_dict('records'),
    'topic_comparison': topic_ct.to_dict('records'),
    'subtopic_comparison': subtopic_ct.to_dict('records'),
    'subtopic_by_status': subtopic_by_status.to_dict('records'),
    'state_distribution': state_dist.to_dict('records'),
    'rush_comparison': rush_comparison.to_dict('records'),
    'percentile_90th_analysis': percentile_90th_by_topic.to_dict('records'),
    'residual_3z': {
        'total_stats': residual_3z_total,
        'breakdown': residual_3z_breakdown.to_dict('records'),
        'by_status': residual_3z_by_status.to_dict('records'),
        'monthly_trends': residual_3z_monthly.to_dict('records'),
        'breakdown_by_status': residual_3z_breakdown_by_status.to_dict('records'),
        'monthly_trends_by_status': residual_3z_monthly_by_status.to_dict('records')
    },
    'advanced_kpis': {
        'void_analysis': {
            'void_rate': float(void_rate),
            'void_count': int(void_count),
            'by_reason': void_by_reason.to_dict('records'),
            'by_topic': void_by_topic.to_dict('records'),
            'by_coordinator': void_by_coordinator.to_dict('records'),
            'by_site': void_by_site.to_dict('records'),
            'monthly_trend': void_monthly_trend.to_dict('records')
        },
        'hold_analysis': {
            'hold_rate': float(hold_rate),
            'hold_count': int(hold_count),
            'by_reason': hold_by_reason.to_dict('records'),
            'by_topic': hold_by_topic.to_dict('records'),
            'by_requestor': hold_by_requestor.to_dict('records'),
            'ct_comparison': hold_ct_comparison.to_dict('records')
        },
        'ftr_analysis': {
            'ftr_rate': float(ftr_rate),
            'ftr_count': int(ftr_count),
            'by_coordinator': ftr_by_coordinator.to_dict('records'),
            'by_requestor': ftr_by_requestor.to_dict('records'),
            'monthly_trend': ftr_monthly_trend.to_dict('records')
        },
        'coordinator_workload': {
            'workload_distribution': coordinator_workload.to_dict('records'),
            'monthly_workload': coordinator_monthly_workload.to_dict('records')
        },
        'mfg_site_analysis': mfg_site_analysis.to_dict('records'),
        'requestor_analysis': requestor_analysis.to_dict('records'),
        'rush_analysis': {
            'rush_rate': float(rush_rate),
            'rush_count': int(rush_count),
            'by_topic': rush_by_topic.to_dict('records'),
            'monthly_trend': rush_monthly_trend.to_dict('records'),
            'success_rate': rush_success_rate.to_dict('records')
        },
        'pending_analysis': {
            'summary': {
                'count': int(pending_analysis['PendingDays']['count']) if not pending_analysis.empty else 0,
                'avg': float(pending_analysis['PendingDays']['mean']) if not pending_analysis.empty else 0,
                'median': float(pending_analysis['PendingDays']['median']) if not pending_analysis.empty else 0,
                'max': float(pending_analysis['PendingDays']['max']) if not pending_analysis.empty else 0
            },
            'by_topic': pending_by_topic.to_dict('records') if not pending_by_topic.empty else []
        }
    },
    'ecn_type_state_distribution': df_clean.groupby(['ECN Topic', 'State']).size().reset_index(name='Count').to_dict('records')
}

# ECN Type by State distribution for OPEN ECNs only (as of today)
# States that indicate ECN is still open/in-progress
open_states = [
    'Hold', 'Hold_Response', 'Submitted', 'ABS Status Pending',
    'PENDING_SAP_CREATION', 'Pending ADPH FG Input', 'Pending Approval',
    'Pending Subcon Input', 'Pending Subcon Input (Close ECN)',
    'Processing', 'Processing (Auto:ABS)', 'Processing (AutoClose:3AA)',
    'Error', 'Pending ADML FG Input', 'Issue_Review_BOM', 'Issue_Review_Material'
]

df_open_ecns = df_clean[df_clean['State'].isin(open_states)].copy()
today_date = pd.Timestamp.now().strftime('%Y-%m-%d')
print(f'\nOpen ECNs as of {today_date}: {len(df_open_ecns):,} ECNs')

if len(df_open_ecns) > 0:
    ecn_type_state_open = df_open_ecns.groupby(['ECN Topic', 'State']).size().reset_index(name='Count')
    data_export['ecn_type_state_open'] = ecn_type_state_open.to_dict('records')
    data_export['open_ecns_info'] = {
        'as_of_date': today_date,
        'total_ecns': int(len(df_open_ecns)),
        'states_included': open_states
    }

    # Summary by state
    open_by_state = df_open_ecns.groupby('State').size().reset_index(name='Count')
    open_by_state = open_by_state.sort_values('Count', ascending=False)
    print('  Open ECNs by State:')
    for _, row in open_by_state.iterrows():
        print(f'    {row["State"]}: {row["Count"]:,}')
else:
    data_export['ecn_type_state_open'] = []
    data_export['open_ecns_info'] = {
        'as_of_date': today_date,
        'total_ecns': 0,
        'states_included': open_states
    }

# Calculate ECNs open for more than 100 days (excluding Closed, Void, Test, TestClosed, Rejected)
print('\nAnalyzing ECNs open for more than 100 days...')
df_open = df_clean[~df_clean['State'].isin(['Closed', 'Void', 'Test', 'TestClosed', 'Rejected'])].copy()
today = pd.Timestamp.now()
df_open['DaysOpen'] = (today - df_open['SubmitDate']).dt.days
df_over_100 = df_open[df_open['DaysOpen'] > 100].copy()

if len(df_over_100) > 0:
    # Group by ECN Topic and State
    ecn_100_days_by_type_state = df_over_100.groupby(['ECN Topic', 'State']).agg({
        'RequestNum': 'count',
        'DaysOpen': 'mean'
    }).reset_index()
    ecn_100_days_by_type_state.columns = ['ECN Topic', 'State', 'Count', 'AvgDaysOpen']

    # Summary by ECN Topic only
    ecn_100_days_by_type = df_over_100.groupby('ECN Topic').agg({
        'RequestNum': 'count',
        'DaysOpen': 'mean'
    }).reset_index()
    ecn_100_days_by_type.columns = ['ECN Topic', 'Count', 'AvgDaysOpen']
    ecn_100_days_by_type = ecn_100_days_by_type.sort_values('Count', ascending=False)

    data_export['ecns_over_100_days'] = {
        'total_count': int(len(df_over_100)),
        'avg_days_open': float(df_over_100['DaysOpen'].mean()),
        'max_days_open': int(df_over_100['DaysOpen'].max()),
        'by_type_and_state': ecn_100_days_by_type_state.to_dict('records'),
        'by_type': ecn_100_days_by_type.to_dict('records')
    }
    print(f'  Found {len(df_over_100):,} ECNs open for more than 100 days')
    print(f'  Average days open: {df_over_100["DaysOpen"].mean():.1f}')
    print(f'  Oldest: {df_over_100["DaysOpen"].max()} days')

    # Export to Excel
    print('  Creating Excel file for ECNs over 100 days...')
    over_100_output_file = 'ECN_Over_100_Days.xlsx'

    # Prepare export columns
    over_100_export = df_over_100[[
        'RequestNum', 'SubmitDate', 'State', 'ECN Topic', 'ECN Sub Topic',
        'CoordinatorName', 'Requestor', 'MFGSiteCode', 'DaysOpen', 'ProcCT(days)', 'TotalCT(Days)'
    ]].copy()

    # Sort by days open descending
    over_100_export = over_100_export.sort_values('DaysOpen', ascending=False)

    with pd.ExcelWriter(over_100_output_file, engine='openpyxl') as writer:
        # Sheet 1: All ECNs over 100 days
        over_100_export.to_excel(writer, sheet_name='All ECNs Over 100 Days', index=False)

        # Sheet 2: By Type and State
        ecn_100_days_by_type_state.to_excel(writer, sheet_name='By Type and State', index=False)

        # Sheet 3: By Type only
        ecn_100_days_by_type.to_excel(writer, sheet_name='By ECN Type', index=False)

        # Sheet 4: By State only
        over_100_by_state = df_over_100.groupby('State').agg({
            'RequestNum': 'count',
            'DaysOpen': 'mean'
        }).reset_index()
        over_100_by_state.columns = ['State', 'Count', 'AvgDaysOpen']
        over_100_by_state = over_100_by_state.sort_values('Count', ascending=False)
        over_100_by_state.to_excel(writer, sheet_name='By State', index=False)

    print(f'  [OK] Excel file created: {over_100_output_file}')
else:
    data_export['ecns_over_100_days'] = {
        'total_count': 0,
        'avg_days_open': 0,
        'max_days_open': 0,
        'by_type_and_state': [],
        'by_type': []
    }
    print('  No ECNs found open for more than 100 days')

# ==================================================================================
# OUTLIER ANALYSIS FOR PROCESS OPTIMIZATION
# ==================================================================================
print("\nCalculating Outlier Analysis for Process Optimization...")

# Calculate statistics by ECN Type for closed ECNs
ecn_type_stats = df_closed.groupby('ECN Topic')['ProcCT(days)'].agg([
    ('Count', 'count'),
    ('Mean', 'mean'),
    ('Median', 'median'),
    ('StdDev', 'std'),
    ('Min', 'min'),
    ('Max', 'max'),
    ('P25', lambda x: x.quantile(0.25)),
    ('P75', lambda x: x.quantile(0.75)),
    ('P90', lambda x: x.quantile(0.90))
]).reset_index()

# Calculate Coefficient of Variation (CV = StdDev / Mean)
ecn_type_stats['CV'] = ecn_type_stats['StdDev'] / ecn_type_stats['Mean']
ecn_type_stats['IQR'] = ecn_type_stats['P75'] - ecn_type_stats['P25']

# Filter to ECN types with at least 50 ECNs for statistical significance
ecn_type_stats_filtered = ecn_type_stats[ecn_type_stats['Count'] >= 50].copy()
ecn_type_stats_filtered = ecn_type_stats_filtered.sort_values('CV', ascending=False)

# Identify outliers (z-score > 3)
df_closed_with_stats = df_closed.merge(
    ecn_type_stats[['ECN Topic', 'Mean', 'StdDev']],
    on='ECN Topic',
    how='left'
)
df_closed_with_stats['ZScore'] = (df_closed_with_stats['ProcCT(days)'] - df_closed_with_stats['Mean']) / df_closed_with_stats['StdDev']
df_outliers = df_closed_with_stats[df_closed_with_stats['ZScore'] > 3].copy()

total_outliers = len(df_outliers)
if total_outliers > 0:
    total_excess_days = (df_outliers['ProcCT(days)'] - df_outliers['Mean']).sum()
    avg_excess_per_outlier = total_excess_days / total_outliers
else:
    total_excess_days = 0
    avg_excess_per_outlier = 0

# Add to data export
data_export['outlier_analysis'] = {
    'total_outliers': int(total_outliers),
    'total_excess_days': float(total_excess_days),
    'avg_excess_per_outlier': float(avg_excess_per_outlier),
    'high_variability_types': ecn_type_stats_filtered.head(10).to_dict('records')
}

print(f"  [OK] Outlier Analysis calculated: {total_outliers:,} outliers representing {total_excess_days:,.0f} excess days")

# Save to JSON
with open('data.json', 'w') as f:
    json.dump(data_export, f, indent=2, default=str)

# Export 90th Percentile ECNs to separate Excel file
print("\nCreating 90th Percentile ECN Excel file...")

# Get all ECNs in the 90th percentile (top 10% slowest)
# Using closed ECNs only for percentile calculation
p90_threshold = df_closed['ProcCT(days)'].quantile(0.90)
df_90th_percentile = df_closed[df_closed['ProcCT(days)'] > p90_threshold].copy()

# Sort by Processing CT descending (slowest first)
df_90th_percentile = df_90th_percentile.sort_values('ProcCT(days)', ascending=False)

# Select relevant columns for the export
columns_to_export = [
    'RequestNum',
    'SubmitDate',
    'FinalDate',
    'ProcCT(days)',
    'TotalCT(Days)',
    'ECN Topic',
    'ECN Sub Topic',
    'CoordinatorName',
    'ECNCoordinatorSite',
    'State',
    'RushRequest',
    'MfgSite',
    'ProcEndTime',
    'FinalOrHoldDate'
]

# Filter to only include columns that exist
available_columns = [col for col in columns_to_export if col in df_90th_percentile.columns]
df_export = df_90th_percentile[available_columns].copy()

# Rename CoordinatorName back to ECNCoordinator for clarity
if 'CoordinatorName' in df_export.columns:
    df_export = df_export.rename(columns={'CoordinatorName': 'ECNCoordinator'})

# Format dates for better readability
date_columns = ['SubmitDate', 'FinalDate', 'ProcEndTime', 'FinalOrHoldDate']
for col in date_columns:
    if col in df_export.columns:
        df_export[col] = pd.to_datetime(df_export[col]).dt.strftime('%Y-%m-%d')

# Export to Excel with formatting and grouping by topic
output_file = 'ECN_90th_Percentile.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Sheet 1: All 90th percentile ECNs (sorted by slowest first)
    df_export.to_excel(writer, sheet_name='All 90th Percentile ECNs', index=False)

    # Sheet 2: Summary by ECN Topic
    p90_by_topic = df_90th_percentile.groupby('ECN Topic').agg({
        'RequestNum': 'count',
        'ProcCT(days)': ['mean', 'median', 'max']
    }).reset_index()
    p90_by_topic.columns = ['ECN Topic', 'Count', 'Avg_ProcCT', 'Median_ProcCT', 'Max_ProcCT']
    p90_by_topic = p90_by_topic.sort_values('Count', ascending=False)
    p90_by_topic.to_excel(writer, sheet_name='Summary by Topic', index=False)

    # Sheet 3: Detailed breakdown by topic (top 5 slowest topics)
    top_5_slow_topics = p90_by_topic.head(5)['ECN Topic'].tolist()
    for topic in top_5_slow_topics:
        # Clean sheet name (Excel has 31 char limit and special char restrictions)
        sheet_name = str(topic)[:28].replace('/', '-').replace('\\', '-').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
        if not sheet_name:
            sheet_name = 'Unknown'

        # Get ECNs for this topic
        df_topic = df_export[df_export['ECN Topic'] == topic].copy()

        # Export to sheet
        if len(df_topic) > 0:
            df_topic.to_excel(writer, sheet_name=sheet_name, index=False)

    # Auto-adjust column widths for all sheets
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for idx, column in enumerate(worksheet.columns):
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"[OK] 90th Percentile Excel file created: {output_file}")
print(f"  - Threshold: {p90_threshold:.2f} days")
print(f"  - Total ECNs in 90th percentile: {len(df_90th_percentile)}")
print(f"  - Percentage of closed ECNs: {(len(df_90th_percentile)/len(df_closed)*100):.1f}%")
print(f"  - Processing CT range: {df_90th_percentile['ProcCT(days)'].min():.2f} to {df_90th_percentile['ProcCT(days)'].max():.2f} days")
print(f"  - Sheets: All ECNs + Summary by Topic + Top 5 Slowest Topics")

# Export Void ECNs by Reason to separate Excel file
print("\nCreating Void ECNs by Reason Excel file...")

# Get all voided ECNs
df_void = df_clean[df_clean['State'] == 'Void'].copy()

# Sort by VoidReason and then by SubmitDate
df_void = df_void.sort_values(['VoidReason', 'SubmitDate'], ascending=[True, False])

# Select relevant columns for the export
void_columns_to_export = [
    'RequestNum',
    'VoidReason',
    'SubmitDate',
    'FinalDate',
    'ProcCT(days)',
    'TotalCT(Days)',
    'ECN Topic',
    'ECN Sub Topic',
    'CoordinatorName',
    'ECNCoordinatorSite',
    'RushRequest',
    'MfgSite',
    'FinalOrHoldDate'
]

# Filter to only include columns that exist
void_available_columns = [col for col in void_columns_to_export if col in df_void.columns]
df_void_export = df_void[void_available_columns].copy()

# Rename CoordinatorName for clarity
if 'CoordinatorName' in df_void_export.columns:
    df_void_export = df_void_export.rename(columns={'CoordinatorName': 'ECNCoordinator'})

# Format dates
void_date_columns = ['SubmitDate', 'FinalDate', 'FinalOrHoldDate']
for col in void_date_columns:
    if col in df_void_export.columns:
        df_void_export[col] = pd.to_datetime(df_void_export[col]).dt.strftime('%Y-%m-%d')

# Create summary by void reason
void_summary = df_void.groupby('VoidReason').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean'
}).reset_index()
void_summary = void_summary.rename(columns={
    'RequestNum': 'Count',
    'ProcCT(days)': 'AvgProcCT'
})
void_summary = void_summary.sort_values('Count', ascending=False)
void_summary['Percentage'] = (void_summary['Count'] / len(df_void) * 100).round(2)

# Export to Excel with multiple sheets
void_output_file = 'ECN_Void_by_Reason.xlsx'
with pd.ExcelWriter(void_output_file, engine='openpyxl') as writer:
    # Sheet 1: Summary by void reason
    void_summary.to_excel(writer, sheet_name='Summary by Void Reason', index=False)

    # Sheet 2: All voided ECNs
    df_void_export.to_excel(writer, sheet_name='All Void ECNs', index=False)

    # Sheet 3: Separate sheet for each top void reason (top 5)
    top_reasons = void_summary.head(5)['VoidReason'].tolist()

    for reason in top_reasons:
        # Clean sheet name (Excel has 31 char limit and special char restrictions)
        sheet_name = str(reason)[:28].replace('/', '-').replace('\\', '-').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
        if not sheet_name:
            sheet_name = 'Unknown'

        # Get ECNs for this reason
        df_reason = df_void_export[df_void_export['VoidReason'] == reason].copy()

        # Export to sheet
        if len(df_reason) > 0:
            df_reason.to_excel(writer, sheet_name=sheet_name, index=False)

    # Auto-adjust column widths for all sheets
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for idx, column in enumerate(worksheet.columns):
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"[OK] Void ECNs Excel file created: {void_output_file}")
print(f"  - Total void ECNs: {len(df_void)}")
print(f"  - Void rate: {void_rate:.2f}%")
print(f"  - Unique void reasons: {df_void['VoidReason'].nunique()}")
print(f"  - Sheets created: Summary + All ECNs + Top 5 reasons")
print(f"  - Top void reason: {void_summary.iloc[0]['VoidReason']} ({void_summary.iloc[0]['Count']} ECNs, {void_summary.iloc[0]['Percentage']:.1f}%)")

# Export Open ECNs to Excel
if len(df_open_ecns) > 0:
    print("\nCreating Open ECNs Excel file...")
    open_ecns_output_file = 'ECN_Open_ECNs.xlsx'

    # Prepare export columns
    open_ecns_export = df_open_ecns[[
        'RequestNum', 'SubmitDate', 'State', 'ECN Topic', 'ECN Sub Topic',
        'CoordinatorName', 'Requestor', 'MFGSiteCode', 'ProcCT(days)', 'TotalCT(Days)'
    ]].copy()

    # Calculate days open
    open_ecns_export['DaysOpen'] = (pd.Timestamp.now() - open_ecns_export['SubmitDate']).dt.days

    # Sort by days open descending
    open_ecns_export = open_ecns_export.sort_values('DaysOpen', ascending=False)

    with pd.ExcelWriter(open_ecns_output_file, engine='openpyxl') as writer:
        # Sheet 1: All Open ECNs
        open_ecns_export.to_excel(writer, sheet_name='All Open ECNs', index=False)

        # Sheet 2: Summary by State
        open_by_state = df_open_ecns.groupby('State').agg({
            'RequestNum': 'count'
        }).reset_index()
        open_by_state.columns = ['State', 'Count']
        open_by_state = open_by_state.sort_values('Count', ascending=False)
        open_by_state.to_excel(writer, sheet_name='Summary by State', index=False)

        # Sheet 3: Summary by ECN Type
        open_by_type = df_open_ecns.groupby('ECN Topic').agg({
            'RequestNum': 'count'
        }).reset_index()
        open_by_type.columns = ['ECN Topic', 'Count']
        open_by_type = open_by_type.sort_values('Count', ascending=False)
        open_by_type.to_excel(writer, sheet_name='Summary by ECN Type', index=False)

        # Sheet 4: By Type and State (chart data)
        type_state_summary = df_open_ecns.groupby(['ECN Topic', 'State']).agg({
            'RequestNum': 'count'
        }).reset_index()
        type_state_summary.columns = ['ECN Topic', 'State', 'Count']
        type_state_summary = type_state_summary.sort_values('Count', ascending=False)
        type_state_summary.to_excel(writer, sheet_name='By Type and State', index=False)

    print(f"[OK] Open ECNs Excel file created: {open_ecns_output_file}")
    print(f"  - Total open ECNs: {len(df_open_ecns):,}")
    print(f"  - States: {df_open_ecns['State'].nunique()}")
    print(f"  - ECN Types: {df_open_ecns['ECN Topic'].nunique()}")

# Export Quarterly Trends to Excel
print("\nCreating Quarterly Trends Excel file...")

quarterly_output_file = 'ECN_Quarterly_Trends.xlsx'
with pd.ExcelWriter(quarterly_output_file, engine='openpyxl') as writer:
    # Sheet 1: Quarterly trends by topic
    quarterly_topic_trends.to_excel(writer, sheet_name='Quarterly by Topic', index=False)

    # Sheet 2: Overall quarterly summary (all ECNs, not just top 5 topics)
    quarterly_summary = df_clean.groupby('YearQuarter').agg({
        'ProcCT(days)': ['mean', 'median'],
        'TotalCT(Days)': ['mean', 'median'],
        'RequestNum': 'count'
    }).reset_index()
    quarterly_summary.columns = ['YearQuarter', 'Avg_ProcCT', 'Median_ProcCT', 'Avg_TotalCT', 'Median_TotalCT', 'ECN_Count']
    quarterly_summary = quarterly_summary.sort_values('YearQuarter')
    quarterly_summary.to_excel(writer, sheet_name='Quarterly Summary', index=False)

    # Sheet 3: Monthly trends (for comparison)
    monthly_trends.to_excel(writer, sheet_name='Monthly Trends', index=False)

    # Sheet 4: All ECN Details (sorted by quarter and processing time)
    ecn_details_columns = [
        'RequestNum',
        'YearQuarter',
        'YearMonth',
        'SubmitDate',
        'FinalDate',
        'ProcCT(days)',
        'TotalCT(Days)',
        'ECN Topic',
        'ECN Sub Topic',
        'CoordinatorName',
        'ECNCoordinatorSite',
        'State',
        'RushRequest',
        'MfgSite'
    ]

    # Filter to only include columns that exist
    available_detail_columns = [col for col in ecn_details_columns if col in df_clean.columns]
    df_ecn_details = df_clean[available_detail_columns].copy()

    # Sort by quarter and processing time (slowest first)
    df_ecn_details = df_ecn_details.sort_values(['YearQuarter', 'ProcCT(days)'], ascending=[False, False])

    # Rename CoordinatorName to ECNCoordinator for clarity
    if 'CoordinatorName' in df_ecn_details.columns:
        df_ecn_details = df_ecn_details.rename(columns={'CoordinatorName': 'ECNCoordinator'})

    # Format dates
    for date_col in ['SubmitDate', 'FinalDate']:
        if date_col in df_ecn_details.columns:
            df_ecn_details[date_col] = pd.to_datetime(df_ecn_details[date_col]).dt.strftime('%Y-%m-%d')

    df_ecn_details.to_excel(writer, sheet_name='All ECN Details', index=False)

    # Sheet 5+: Individual sheets for last 6 months
    last_6_months = monthly_trends.tail(6)['YearMonth'].tolist()

    for month in last_6_months:
        # Clean sheet name (Excel has 31 char limit)
        sheet_name = str(month)[:28]

        # Get ECNs for this month
        df_month = df_ecn_details[df_ecn_details['YearMonth'] == month].copy()

        # Export to sheet if there are ECNs
        if len(df_month) > 0:
            # Sort by processing time (slowest first)
            df_month = df_month.sort_values('ProcCT(days)', ascending=False)
            df_month.to_excel(writer, sheet_name=sheet_name, index=False)

    # Auto-adjust column widths for all sheets
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for idx, column in enumerate(worksheet.columns):
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"[OK] Quarterly Trends Excel file created: {quarterly_output_file}")
print(f"  - Quarters included: {quarterly_summary['YearQuarter'].nunique()}")
print(f"  - Total ECNs in details: {len(df_ecn_details):,}")
print(f"  - Individual month sheets: {len(last_6_months)}")
print(f"  - Sheets: Quarterly by Topic, Quarterly Summary, Monthly Trends, All ECN Details + Last 6 Months")

print("\nData processed successfully!")
print(f"Total requests: {overall_stats['total_requests']}")
print(f"Average Processing CT: {overall_stats['avg_proc_ct']:.2f} days")
print(f"Average Total CT: {overall_stats['avg_total_ct']:.2f} days")
print(f"Date range: {overall_stats['date_range']['start']} to {overall_stats['date_range']['end']}")
