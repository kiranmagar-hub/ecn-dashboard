"""
Identify Outliers and Process Optimization Opportunities by ECN Type
"""
import pyodbc
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Database connection
DB_CONFIG = {
    'server': 'wilmatom1f',
    'database': 'WWECNRequests',
    'username': 'ECNRequestData',
    'password': 'd!iSs5ZHuN',
    'table': 'Qry_EcnRequestGeneralInfoUpdate'
}

print('=' * 100)
print(f'ECN OUTLIER ANALYSIS FOR PROCESS OPTIMIZATION - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
print('=' * 100)

# Connect to database
print('\nConnecting to database...')
conn_str = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={DB_CONFIG['server']};"
    f"DATABASE={DB_CONFIG['database']};"
    f"UID={DB_CONFIG['username']};"
    f"PWD={DB_CONFIG['password']}"
)
conn = pyodbc.connect(conn_str)

# Get last 365 days of data
print('Querying last 365 days of ECN data...')
end_date = datetime.now()
start_date = end_date.replace(year=end_date.year - 1)

query = f"""
SELECT
    RequestNum,
    EcnTopic,
    SubmitDate,
    ClosedDate,
    ECNCoordinator,
    Requestor,
    MFGSiteCode,
    State
FROM {DB_CONFIG['table']}
WHERE SubmitDate >= '{start_date.strftime("%Y-%m-%d")}'
  AND SubmitDate <= '{end_date.strftime("%Y-%m-%d")}'
  AND State IN ('Closed', 'Void')
"""

df = pd.read_sql(query, conn)
conn.close()

print(f'Retrieved {len(df):,} closed/void ECNs')

# Parse ECN Topic
df['ECN Topic'] = df['EcnTopic'].str.split('~').str[0].str.strip()
df['ECN Sub Topic'] = df['EcnTopic'].str.split('~').str[1].str.strip()

# Filter out Type 9 and Test states
df = df[~df['ECN Topic'].str.contains(r'^[\(\{\[]?\s*(9|10)[\)\}\][\s\-]|^(9|10)$', case=False, na=False, regex=True)].copy()
df = df[~df['State'].isin(['Test', 'TestClosed'])].copy()

# Calculate cycle times (only for Closed ECNs with valid dates)
df_closed = df[df['State'] == 'Closed'].copy()
df_closed['SubmitDate'] = pd.to_datetime(df_closed['SubmitDate'])
df_closed['ClosedDate'] = pd.to_datetime(df_closed['ClosedDate'])
df_closed = df_closed[(df_closed['SubmitDate'].notna()) & (df_closed['ClosedDate'].notna())]
df_closed['ProcCT(days)'] = (df_closed['ClosedDate'] - df_closed['SubmitDate']).dt.days

print(f'Analyzing {len(df_closed):,} closed ECNs with valid cycle times')

# ==================================================================================
# 1. IDENTIFY ECN TYPES WITH HIGH VARIABILITY
# ==================================================================================
print('\n' + '=' * 100)
print('1. ECN TYPES WITH HIGH PROCESSING TIME VARIABILITY (Inconsistent Performance)')
print('=' * 100)

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

# Calculate Coefficient of Variation (StdDev / Mean) - higher = more variability
ecn_type_stats['CV'] = ecn_type_stats['StdDev'] / ecn_type_stats['Mean']
ecn_type_stats['IQR'] = ecn_type_stats['P75'] - ecn_type_stats['P25']

# Filter to types with at least 50 ECNs for statistical significance
ecn_type_stats = ecn_type_stats[ecn_type_stats['Count'] >= 50].copy()
ecn_type_stats = ecn_type_stats.sort_values('CV', ascending=False)

print('\nTop 10 ECN Types with HIGHEST VARIABILITY (Inconsistent Processing):')
print(f"{'#':<3} {'ECN Type':<50} {'Count':<8} {'Mean':<8} {'Median':<8} {'StdDev':<8} {'CV':<6}")
print('-' * 100)
for i, row in ecn_type_stats.head(10).iterrows():
    print(f"{ecn_type_stats.index.get_loc(i)+1:<3} {row['ECN Topic'][:48]:<50} {row['Count']:<8.0f} {row['Mean']:<8.1f} {row['Median']:<8.1f} {row['StdDev']:<8.1f} {row['CV']:<6.2f}")

# ==================================================================================
# 2. IDENTIFY OUTLIER ECNs (3 Standard Deviations from Mean)
# ==================================================================================
print('\n' + '=' * 100)
print('2. OUTLIER ECNs (Taking Significantly Longer Than Normal)')
print('=' * 100)

# For each ECN, calculate z-score within its type
df_closed = df_closed.merge(
    ecn_type_stats[['ECN Topic', 'Mean', 'StdDev']],
    on='ECN Topic',
    how='left'
)

# Calculate z-score (how many standard deviations from mean)
df_closed['ZScore'] = (df_closed['ProcCT(days)'] - df_closed['Mean']) / df_closed['StdDev']

# Outliers: z-score > 3 (more than 3 standard deviations above mean)
df_outliers = df_closed[df_closed['ZScore'] > 3].copy()
df_outliers = df_outliers.sort_values('ProcCT(days)', ascending=False)

print(f'\nFound {len(df_outliers):,} outlier ECNs (z-score > 3)')
print(f'These ECNs took significantly longer than typical for their ECN Type')

if len(df_outliers) > 0:
    print('\nTop 20 Slowest Outlier ECNs:')
    print(f"{'RequestNum':<12} {'ECN Type':<40} {'ProcCT':<10} {'TypeAvg':<10} {'Excess':<10} {'Coordinator':<20}")
    print('-' * 100)
    for _, row in df_outliers.head(20).iterrows():
        excess = row['ProcCT(days)'] - row['Mean']
        print(f"{row['RequestNum']:<12} {row['ECN Topic'][:38]:<40} {row['ProcCT(days)']:<10.0f} {row['Mean']:<10.1f} {excess:<10.1f} {str(row['ECNCoordinator'])[:18]:<20}")

# ==================================================================================
# 3. ROOT CAUSE ANALYSIS - BY COORDINATOR
# ==================================================================================
print('\n' + '=' * 100)
print('3. OUTLIERS BY COORDINATOR (Who has the most outlier ECNs?)')
print('=' * 100)

outlier_by_coord = df_outliers.groupby('ECNCoordinator').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean',
    'Mean': 'first'
}).reset_index()
outlier_by_coord.columns = ['Coordinator', 'OutlierCount', 'AvgOutlierCT', 'AvgNormalCT']
outlier_by_coord = outlier_by_coord.sort_values('OutlierCount', ascending=False)

print(f"\n{'Coordinator':<30} {'Outlier Count':<15} {'Avg Outlier CT':<18} {'Avg Normal CT':<15}")
print('-' * 80)
for _, row in outlier_by_coord.head(10).iterrows():
    print(f"{str(row['Coordinator'])[:28]:<30} {row['OutlierCount']:<15.0f} {row['AvgOutlierCT']:<18.1f} {row['AvgNormalCT']:<15.1f}")

# ==================================================================================
# 4. ROOT CAUSE ANALYSIS - BY MFG SITE
# ==================================================================================
print('\n' + '=' * 100)
print('4. OUTLIERS BY MANUFACTURING SITE')
print('=' * 100)

outlier_by_site = df_outliers.groupby('MFGSiteCode').agg({
    'RequestNum': 'count',
    'ProcCT(days)': 'mean'
}).reset_index()
outlier_by_site.columns = ['Site', 'OutlierCount', 'AvgOutlierCT']
outlier_by_site = outlier_by_site.sort_values('OutlierCount', ascending=False)

print(f"\n{'Site':<15} {'Outlier Count':<15} {'Avg Outlier CT':<18}")
print('-' * 50)
for _, row in outlier_by_site.head(10).iterrows():
    print(f"{str(row['Site'])[:13]:<15} {row['OutlierCount']:<15.0f} {row['AvgOutlierCT']:<18.1f}")

# ==================================================================================
# 5. OPTIMIZATION RECOMMENDATIONS
# ==================================================================================
print('\n' + '=' * 100)
print('5. PROCESS OPTIMIZATION RECOMMENDATIONS')
print('=' * 100)

print('\nBased on the analysis:')
print('\n1. HIGH VARIABILITY ECN TYPES (Standardize Processes):')
for i, row in ecn_type_stats.head(5).iterrows():
    print(f'   - {row["ECN Topic"]}: CV={row["CV"]:.2f}, Range={row["Min"]:.0f}-{row["Max"]:.0f} days')
    print(f'     --> High inconsistency suggests need for standardized process/checklist')

print('\n2. FOCUS AREAS FOR IMPROVEMENT:')
total_outliers = len(df_outliers)
total_excess_days = (df_outliers['ProcCT(days)'] - df_outliers['Mean']).sum()
print(f'   - {total_outliers:,} outlier ECNs represent {total_excess_days:,.0f} excess days')
print(f'   - Average excess time per outlier: {total_excess_days/total_outliers:.1f} days')

print('\n3. COORDINATOR-SPECIFIC TRAINING NEEDED:')
for _, row in outlier_by_coord.head(3).iterrows():
    excess = row['AvgOutlierCT'] - row['AvgNormalCT']
    print(f'   - {row["Coordinator"]}: {row["OutlierCount"]:.0f} outliers, avg {excess:.1f} days longer than normal')

print('\n4. SITE-SPECIFIC ISSUES:')
for _, row in outlier_by_site.head(3).iterrows():
    print(f'   - {row["Site"]}: {row["OutlierCount"]:.0f} outliers, avg {row["AvgOutlierCT"]:.1f} days')

# ==================================================================================
# 6. EXPORT TO EXCEL
# ==================================================================================
print('\n' + '=' * 100)
print('EXPORTING RESULTS TO EXCEL')
print('=' * 100)

output_file = f'ECN_Outlier_Analysis_{datetime.now().strftime("%Y%m%d")}.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Sheet 1: ECN Type Variability
    ecn_type_stats.to_excel(writer, sheet_name='ECN Type Variability', index=False)

    # Sheet 2: All Outlier ECNs
    outlier_export = df_outliers[[
        'RequestNum', 'ECN Topic', 'ProcCT(days)', 'Mean', 'StdDev', 'ZScore',
        'ECNCoordinator', 'MFGSiteCode', 'SubmitDate', 'ClosedDate'
    ]].copy()
    outlier_export['ExcessDays'] = outlier_export['ProcCT(days)'] - outlier_export['Mean']
    outlier_export = outlier_export.sort_values('ProcCT(days)', ascending=False)
    outlier_export.to_excel(writer, sheet_name='All Outlier ECNs', index=False)

    # Sheet 3: Outliers by Coordinator
    outlier_by_coord.to_excel(writer, sheet_name='By Coordinator', index=False)

    # Sheet 4: Outliers by Site
    outlier_by_site.to_excel(writer, sheet_name='By Site', index=False)

    # Sheet 5: Outliers by ECN Type
    outlier_by_type = df_outliers.groupby('ECN Topic').agg({
        'RequestNum': 'count',
        'ProcCT(days)': 'mean',
        'Mean': 'first'
    }).reset_index()
    outlier_by_type.columns = ['ECN Topic', 'OutlierCount', 'AvgOutlierCT', 'AvgNormalCT']
    outlier_by_type = outlier_by_type.sort_values('OutlierCount', ascending=False)
    outlier_by_type.to_excel(writer, sheet_name='By ECN Type', index=False)

print(f'\n[OK] Analysis exported to: {output_file}')
print(f'\nSheets created:')
print(f'  1. ECN Type Variability - Types with inconsistent processing times')
print(f'  2. All Outlier ECNs - Complete list of {len(df_outliers):,} outlier ECNs')
print(f'  3. By Coordinator - Outliers grouped by coordinator')
print(f'  4. By Site - Outliers grouped by manufacturing site')
print(f'  5. By ECN Type - Outliers grouped by ECN type')

print('\n' + '=' * 100)
print('ANALYSIS COMPLETE')
print('=' * 100)
