#!/usr/bin/env python3
"""
Simple converter: read all CSV files in data/ and write them as sheets in an XLSX workbook.
Produces can_repos_real.xlsx in the repository root.
"""
import pandas as pd
import glob
import os

csv_files = sorted(glob.glob('data/*.csv'))
if not csv_files:
    print('No CSV files found in data/. Nothing to do.')
    exit(0)

with pd.ExcelWriter('can_repos_real.xlsx', engine='openpyxl') as writer:
    for csv in csv_files:
        try:
            df = pd.read_csv(csv)
        except Exception as e:
            print(f'Failed to read {csv}: {e}')
            continue
        # sheet name: filename without extension, max 31 chars
        sheet = os.path.splitext(os.path.basename(csv))[0]
        sheet = sheet.replace('_', ' ')
        sheet = sheet[:31]
        df.to_excel(writer, sheet_name=sheet, index=False)

print('Wrote can_repos_real.xlsx')
