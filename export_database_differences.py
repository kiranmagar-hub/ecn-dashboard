"""
Export database differences for manual comparison
"""
import pandas as pd
import os
import glob
from datetime import datetime

print('='*80)
print('EXPORTING DATABASE vs EXCEL DIFFERENCES')
print('='*80)

# Load reference Excel
print('\n1. Loading Reference Excel...')
df_excel = pd.read_excel(r'C:\Users\kmagar\Downloads\07022026ECN.xlsx', sheet_name='Document_TB11')
df_excel['SubmitDate'] = pd.to_datetime(df_excel['SubmitDate'], errors='coerce')

# Filter to date range
df_excel_filtered = df_excel[
    (df_excel['SubmitDate'] >= '2025-07-02') &
    (df_excel['SubmitDate'] <= '2026-07-02')
].copy()

print(f'   Reference Excel total: {len(df_excel_filtered):,} rows')

# Parse ECN Topic
df_excel_filtered['ECN Topic'] = df_excel_filtered['EcnTopic'].str.split('~').str[0].str.strip().str.upper()
df_excel_filtered.loc[~df_excel_filtered['EcnTopic'].str.contains('~', na=False), 'ECN Topic'] = \
    df_excel_filtered.loc[~df_excel_filtered['EcnTopic'].str.contains('~', na=False), 'EcnTopic'].str.strip().str.upper()

# Filter out Type 9
df_excel_filtered = df_excel_filtered[
    ~df_excel_filtered['ECN Topic'].str.contains(r'^[\(\{\[]?\s*(9|10)[\)\}\][\s\-]|^(9|10)$', case=False, na=False, regex=True)
].copy()

print(f'   After Type 9 filter: {len(df_excel_filtered):,} rows')

# Filter to Closed only
df_excel_closed = df_excel_filtered[df_excel_filtered['State'] == 'Closed'].copy()
print(f'   Closed ECNs: {len(df_excel_closed):,} rows')

# Load database export - find most recent that's not locked
print('\n2. Loading Database Export...')
folders = glob.glob(r'C:\Users\kmagar\OneDrive - Analog Devices, Inc\Documents\ECNMETRICS07022026\ECN_Metrics_*')
folders.sort(key=os.path.getmtime, reverse=True)

df_db = None
db_file_path = None

for folder in folders:
    try:
        db_file = os.path.join(folder, 'source_data.xlsx')
        if os.path.exists(db_file):
            df_db = pd.read_excel(db_file, sheet_name='Document_TB11')
            db_file_path = db_file
            print(f'   Using: {os.path.basename(folder)}')
            break
    except PermissionError:
        continue

if df_db is None:
    print('ERROR: Could not read any database export (files may be locked)')
    print('Please close any open Excel files and try again.')
    exit(1)

df_db['SubmitDate'] = pd.to_datetime(df_db['SubmitDate'], errors='coerce')
print(f'   Database export total: {len(df_db):,} rows')

# Parse ECN Topic
df_db['ECN Topic'] = df_db['EcnTopic'].str.split('~').str[0].str.strip().str.upper()
df_db.loc[~df_db['EcnTopic'].str.contains('~', na=False), 'ECN Topic'] = \
    df_db.loc[~df_db['EcnTopic'].str.contains('~', na=False), 'EcnTopic'].str.strip().str.upper()

# Filter out Type 9
df_db_filtered = df_db[
    ~df_db['ECN Topic'].str.contains(r'^[\(\{\[]?\s*(9|10)[\)\}\][\s\-]|^(9|10)$', case=False, na=False, regex=True)
].copy()

print(f'   After Type 9 filter: {len(df_db_filtered):,} rows')

# Filter to Closed only
df_db_closed = df_db_filtered[df_db_filtered['State'] == 'Closed'].copy()
print(f'   Closed ECNs: {len(df_db_closed):,} rows')

# Find differences
print('\n3. Finding differences...')

excel_requests = set(df_excel_closed['RequestNum'])
db_requests = set(df_db_closed['RequestNum'])

only_in_excel = excel_requests - db_requests
only_in_db = db_requests - excel_requests
in_both = excel_requests & db_requests

print(f'   In both: {len(in_both):,} ECNs')
print(f'   Only in Excel: {len(only_in_excel):,} ECNs')
print(f'   Only in Database: {len(only_in_db):,} ECNs')

# Export differences
output_file = 'Database_vs_Excel_Comparison.xlsx'
print(f'\n4. Exporting to {output_file}...')

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

    # Summary sheet
    summary_data = {
        'Metric': [
            'Date Range',
            'Total Records (Excel)',
            'Total Records (Database)',
            'Closed ECNs (Excel)',
            'Closed ECNs (Database)',
            'In Both',
            'Only in Excel',
            'Only in Database',
            '',
            'Excel 50th Percentile',
            'Excel 75th Percentile',
            'Excel 90th Percentile',
            'Database 50th Percentile',
            'Database 75th Percentile',
            'Database 90th Percentile'
        ],
        'Value': [
            'Jul 2, 2025 - Jul 2, 2026',
            len(df_excel_filtered),
            len(df_db_filtered),
            len(df_excel_closed),
            len(df_db_closed),
            len(in_both),
            len(only_in_excel),
            len(only_in_db),
            '',
            f"{df_excel_closed['ProcCT(days)'].quantile(0.50):.2f} days",
            f"{df_excel_closed['ProcCT(days)'].quantile(0.75):.2f} days",
            f"{df_excel_closed['ProcCT(days)'].quantile(0.90):.2f} days",
            f"{df_db_closed['ProcCT(days)'].quantile(0.50):.2f} days",
            f"{df_db_closed['ProcCT(days)'].quantile(0.75):.2f} days",
            f"{df_db_closed['ProcCT(days)'].quantile(0.90):.2f} days"
        ]
    }
    pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

    # Only in Database (the 937 extra records)
    if len(only_in_db) > 0:
        df_only_db = df_db_closed[df_db_closed['RequestNum'].isin(only_in_db)].copy()
        df_only_db = df_only_db.sort_values('SubmitDate')
        df_only_db.to_excel(writer, sheet_name='Only in Database', index=False)
        print(f'   - Only in Database: {len(df_only_db):,} rows')

    # Only in Excel
    if len(only_in_excel) > 0:
        df_only_excel = df_excel_closed[df_excel_closed['RequestNum'].isin(only_in_excel)].copy()
        df_only_excel = df_only_excel.sort_values('SubmitDate')
        df_only_excel.to_excel(writer, sheet_name='Only in Excel', index=False)
        print(f'   - Only in Excel: {len(df_only_excel):,} rows')

    # Full database export (for reference)
    df_db_closed_sorted = df_db_closed.sort_values('SubmitDate')
    if len(df_db_closed_sorted) > 10000:
        print(f'   - Database Closed (first 10,000 rows only)')
        df_db_closed_sorted.head(10000).to_excel(writer, sheet_name='Database Closed (sample)', index=False)
    else:
        df_db_closed_sorted.to_excel(writer, sheet_name='Database Closed (all)', index=False)
        print(f'   - Database Closed: {len(df_db_closed_sorted):,} rows')

    # Full Excel export (for reference)
    df_excel_closed_sorted = df_excel_closed.sort_values('SubmitDate')
    if len(df_excel_closed_sorted) > 10000:
        print(f'   - Excel Closed (first 10,000 rows only)')
        df_excel_closed_sorted.head(10000).to_excel(writer, sheet_name='Excel Closed (sample)', index=False)
    else:
        df_excel_closed_sorted.to_excel(writer, sheet_name='Excel Closed (all)', index=False)
        print(f'   - Excel Closed: {len(df_excel_closed_sorted):,} rows')

print(f'\n' + '='*80)
print('EXPORT COMPLETE!')
print('='*80)
print(f'\nFile saved: {output_file}')
print(f'\nSheets created:')
print(f'  1. Summary - Overall comparison statistics')
print(f'  2. Only in Database - {len(only_in_db):,} ECNs not in Excel')
print(f'  3. Only in Excel - {len(only_in_excel):,} ECNs not in Database')
print(f'  4. Database Closed (all) - Full database export')
print(f'  5. Excel Closed (all) - Full Excel export')
print(f'\nNext steps:')
print(f'  1. Open {output_file}')
print(f'  2. Review "Only in Database" sheet to see the 937 extra records')
print(f'  3. Check if these records should be included or excluded')
print(f'  4. Identify the pattern (date range, state, or other criteria)')
