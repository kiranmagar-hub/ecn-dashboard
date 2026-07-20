"""
Export ECN Metrics Data for Power BI
Creates clean datasets optimized for Power BI dashboards
"""
import pandas as pd
import pyodbc
from datetime import datetime, timedelta
import os

# Database configuration
DB_CONFIG = {
    'server': 'wilmatom1f',
    'database': 'WWECNRequests',
    'username': 'ECNRequestData',
    'password': 'd!iSs5ZHuN',
    'table': 'Qry_EcnRequestGeneralInfoUpdate'
}

# Date range (last 365 days by default)
DAYS_BACK = 365

print('='*80)
print('EXPORTING ECN DATA FOR POWER BI')
print('='*80)

# Calculate date range
end_date = datetime.now()
start_date = end_date - timedelta(days=DAYS_BACK)

print(f'\nDate Range: {start_date.strftime("%Y-%m-%d")} to {end_date.strftime("%Y-%m-%d")}')
print(f'Database: {DB_CONFIG["server"]} / {DB_CONFIG["database"]}')

# Connect to database
print('\nConnecting to database...')
conn_str = (
    f"DRIVER={{SQL Server}};"
    f"SERVER={DB_CONFIG['server']};"
    f"DATABASE={DB_CONFIG['database']};"
    f"UID={DB_CONFIG['username']};"
    f"PWD={DB_CONFIG['password']}"
)

try:
    conn = pyodbc.connect(conn_str)
    print('Connected successfully!')

    # Query data
    print('\nQuerying data...')
    query = f"""
    SELECT * FROM {DB_CONFIG['table']}
    WHERE SubmitDate >= '{start_date.strftime("%Y-%m-%d")}'
    AND SubmitDate <= '{end_date.strftime("%Y-%m-%d")}'
    """

    df = pd.read_sql(query, conn)
    conn.close()

    print(f'Retrieved {len(df):,} records')

    # Data cleaning and transformation
    print('\nCleaning and transforming data...')

    # Parse ECN Topic
    df['ECN Topic'] = df['EcnTopic'].str.split('~').str[0].str.strip().str.upper()
    df.loc[~df['EcnTopic'].str.contains('~', na=False), 'ECN Topic'] = \
        df.loc[~df['EcnTopic'].str.contains('~', na=False), 'EcnTopic'].str.strip().str.upper()

    # Filter out ECN Type 9 (Restricted ECNs)
    df_filtered = df[
        ~df['ECN Topic'].str.contains(r'^[\(\{\[]?\s*(9|10)[\)\}\][\s\-]|^(9|10)$', case=False, na=False, regex=True)
    ].copy()

    print(f'After filtering Type 9: {len(df_filtered):,} records')

    # Convert dates
    date_columns = ['SubmitDate', 'ClosedDate', 'HoldDate', 'FinalDate']
    for col in date_columns:
        if col in df_filtered.columns:
            df_filtered[col] = pd.to_datetime(df_filtered[col], errors='coerce')

    # Calculate cycle times
    print('Calculating cycle times...')

    # Calculate FinalOrHoldDate
    df_filtered['FinalOrHoldDate'] = df_filtered['ClosedDate'].fillna(df_filtered['HoldDate'])

    # Calculate ProcCT(days)
    df_filtered['ProcCT(days)'] = (
        df_filtered['FinalOrHoldDate'] - df_filtered['SubmitDate']
    ).dt.days

    # Calculate TotalCT(days) if ClosedDate exists
    if 'ClosedDate' in df_filtered.columns:
        df_filtered['TotalCT(Days)'] = (
            df_filtered['ClosedDate'] - df_filtered['SubmitDate']
        ).dt.days

    # Add time dimensions for Power BI
    df_filtered['SubmitYear'] = df_filtered['SubmitDate'].dt.year
    df_filtered['SubmitMonth'] = df_filtered['SubmitDate'].dt.month
    df_filtered['SubmitMonthName'] = df_filtered['SubmitDate'].dt.strftime('%Y-%m')
    df_filtered['SubmitQuarter'] = df_filtered['SubmitDate'].dt.quarter
    df_filtered['SubmitWeek'] = df_filtered['SubmitDate'].dt.isocalendar().week

    # Load coordinator mapping
    coordinator_file = r'C:\Users\kmagar\Downloads\ECN Coordinators.xlsx'
    coordinator_map = {}

    if os.path.exists(coordinator_file):
        print('Loading coordinator names...')
        df_coord = pd.read_excel(coordinator_file)
        if 'ECN Coordinators' in df_coord.columns and 'Name' in df_coord.columns:
            coordinator_map = dict(zip(
                df_coord['ECN Coordinators'].str.upper(),
                df_coord['Name']
            ))
            print(f'Loaded {len(coordinator_map)} coordinator names')

    # Map coordinator names
    df_filtered['CoordinatorName'] = df_filtered['ECNCoordinator'].str.upper().map(coordinator_map)
    df_filtered['CoordinatorName'] = df_filtered['CoordinatorName'].fillna(df_filtered['ECNCoordinator'])

    # Use Site as fallback for ECNCoordinatorSite if missing
    if 'ECNCoordinatorSite' not in df_filtered.columns and 'Site' in df_filtered.columns:
        df_filtered['ECNCoordinatorSite'] = df_filtered['Site']

    # Export to Excel files for Power BI
    output_file = f'ECN_Metrics_PowerBI_Export_{datetime.now().strftime("%Y%m%d")}.xlsx'

    print(f'\nExporting to {output_file}...')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # 1. Main fact table - All ECN records
        main_cols = [
            'RequestNum', 'State', 'ECN Topic', 'EcnTopic',
            'SubmitDate', 'ClosedDate', 'HoldDate', 'FinalOrHoldDate',
            'ProcCT(days)',
            'ECNCoordinator', 'CoordinatorName', 'ECNCoordinatorSite',
            'SubmitYear', 'SubmitMonth', 'SubmitMonthName', 'SubmitQuarter', 'SubmitWeek'
        ]

        # Add optional columns if they exist
        if 'TotalCT(Days)' in df_filtered.columns:
            main_cols.insert(main_cols.index('ProcCT(days)') + 1, 'TotalCT(Days)')

        if 'MFGSiteCode' in df_filtered.columns:
            main_cols.append('MFGSiteCode')
        elif 'Site' in df_filtered.columns:
            main_cols.append('Site')

        df_main = df_filtered[main_cols].copy()
        df_main.to_excel(writer, sheet_name='ECN_FactTable', index=False)
        print(f'  - ECN_FactTable: {len(df_main):,} rows')

        # 2. Closed ECNs only (for cycle time analysis)
        closed_cols = [
            'RequestNum', 'ECN Topic', 'EcnTopic',
            'SubmitDate', 'ClosedDate', 'FinalOrHoldDate',
            'ProcCT(days)',
            'CoordinatorName', 'ECNCoordinatorSite',
            'SubmitYear', 'SubmitMonth', 'SubmitMonthName'
        ]

        if 'TotalCT(Days)' in df_filtered.columns:
            closed_cols.insert(closed_cols.index('ProcCT(days)') + 1, 'TotalCT(Days)')

        df_closed = df_filtered[df_filtered['State'] == 'Closed'][closed_cols].copy()
        df_closed.to_excel(writer, sheet_name='ECN_Closed', index=False)
        print(f'  - ECN_Closed: {len(df_closed):,} rows')

        # 3. Coordinator dimension table
        df_coordinators = df_filtered[['ECNCoordinator', 'CoordinatorName', 'ECNCoordinatorSite']].drop_duplicates()
        df_coordinators.to_excel(writer, sheet_name='Dim_Coordinators', index=False)
        print(f'  - Dim_Coordinators: {len(df_coordinators):,} rows')

        # 4. ECN Topic dimension table
        df_topics = df_filtered[['ECN Topic', 'EcnTopic']].drop_duplicates()
        df_topics.to_excel(writer, sheet_name='Dim_Topics', index=False)
        print(f'  - Dim_Topics: {len(df_topics):,} rows')

        # 5. Site dimension table
        site_col = 'MFGSiteCode' if 'MFGSiteCode' in df_filtered.columns else 'Site'
        df_sites = df_filtered[[site_col]].drop_duplicates()
        df_sites.to_excel(writer, sheet_name='Dim_Sites', index=False)
        print(f'  - Dim_Sites: {len(df_sites):,} rows')

        # 6. Monthly summary (pre-aggregated for performance)
        df_monthly = df_closed.groupby('SubmitMonthName').agg({
            'RequestNum': 'count',
            'ProcCT(days)': ['median', 'mean', lambda x: x.quantile(0.75), lambda x: x.quantile(0.90)]
        }).reset_index()
        df_monthly.columns = ['Month', 'ECN_Count', 'Median_ProcCT', 'Mean_ProcCT', 'P75_ProcCT', 'P90_ProcCT']
        df_monthly.to_excel(writer, sheet_name='Summary_Monthly', index=False)
        print(f'  - Summary_Monthly: {len(df_monthly):,} rows')

        # 7. Coordinator summary (pre-aggregated)
        df_coord_summary = df_closed.groupby('CoordinatorName').agg({
            'RequestNum': 'count',
            'ProcCT(days)': 'median'
        }).reset_index()
        df_coord_summary.columns = ['Coordinator', 'ECN_Count', 'Median_ProcCT']
        df_coord_summary = df_coord_summary.sort_values('ECN_Count', ascending=False)
        df_coord_summary.to_excel(writer, sheet_name='Summary_Coordinators', index=False)
        print(f'  - Summary_Coordinators: {len(df_coord_summary):,} rows')

    print(f'\n{"="*80}')
    print('EXPORT COMPLETE!')
    print('='*80)
    print(f'\nFile: {output_file}')
    print(f'\nSheets created:')
    print(f'  1. ECN_FactTable - Main table with all ECN records')
    print(f'  2. ECN_Closed - Closed ECNs only (for cycle time analysis)')
    print(f'  3. Dim_Coordinators - Coordinator lookup table')
    print(f'  4. Dim_Topics - ECN Topic lookup table')
    print(f'  5. Dim_Sites - Site lookup table')
    print(f'  6. Summary_Monthly - Pre-aggregated monthly metrics')
    print(f'  7. Summary_Coordinators - Pre-aggregated coordinator performance')

    print(f'\n{"="*80}')
    print('POWER BI SETUP INSTRUCTIONS')
    print('='*80)
    print(f'''
OPTION A: Use Excel File (Easiest)
----------------------------------
1. Open Power BI Desktop
2. Get Data > Excel > Select: {output_file}
3. Load these tables:
   - ECN_FactTable (main data)
   - Dim_Coordinators, Dim_Topics, Dim_Sites (dimensions)
   - Summary_Monthly, Summary_Coordinators (for quick visuals)

4. Create relationships (if not auto-detected):
   - ECN_FactTable[ECNCoordinator] -> Dim_Coordinators[ECNCoordinator]
   - ECN_FactTable[ECN Topic] -> Dim_Topics[ECN Topic]

5. Set scheduled refresh:
   - Re-run this script daily to update the Excel file
   - Power BI will refresh from the updated Excel file


OPTION B: Direct Database Connection (Real-time)
-------------------------------------------------
1. Open Power BI Desktop
2. Get Data > SQL Server
   - Server: {DB_CONFIG['server']}
   - Database: {DB_CONFIG['database']}

3. Advanced options > SQL statement:
   {query}

4. Transform Data (Power Query Editor):
   - Add same filters and calculations as above
   - Filter out ECN Type 9
   - Calculate ProcCT(days) and time dimensions

5. Publish to Power BI Service
6. Set up scheduled refresh (daily, hourly, etc.)


RECOMMENDED VISUALS:
-------------------
1. KPI Cards:
   - Total ECNs
   - Median Processing Time
   - 90th Percentile

2. Line Chart:
   - X-axis: SubmitMonthName
   - Y-axis: Median_ProcCT
   - Title: "Processing Time Trend"

3. Bar Chart:
   - X-axis: CoordinatorName (top 10)
   - Y-axis: ECN_Count
   - Title: "Top Coordinators by Volume"

4. Column Chart:
   - X-axis: ECN Topic
   - Y-axis: Count of RequestNum
   - Title: "ECNs by Topic"

5. Scatter Plot:
   - X-axis: ECN_Count
   - Y-axis: Median_ProcCT
   - Legend: CoordinatorName
   - Title: "Coordinator Performance"
''')

    print(f'\nNext Steps:')
    print(f'1. Open Power BI Desktop (download from https://powerbi.microsoft.com)')
    print(f'2. Follow Option A or B above to connect to your data')
    print(f'3. Create your dashboard with the recommended visuals')
    print(f'4. Publish to Power BI Service')
    print(f'5. Embed in SharePoint page')

except Exception as e:
    print(f'\nERROR: {e}')
    import traceback
    traceback.print_exc()
