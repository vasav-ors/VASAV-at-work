import os
import pandas as pd
import time

# This script scans the subfolders in the specified parent directory for results_MinPenetration CSV files.
# It extracts SLS and ULS UR from the **results_MinPenetration block and saves a summary CSV in the parent folder.

# User input: parent directory containing subfolders
parent_dir = r'k:\dozr\HOW03\GEO\04_OptiMon Runs\20251017_Lateral_pile_stability_ Installation\variations\05_Hs7p3_loads_it1\monopiles'
# User input: Hs value to extract
hs_value = 'Hs7_3'  # e.g. 'Hs2_5', 'Hs6_4', 'Hs7_3'

results = []

# Process all subfolders (remove the limit)
subfolders = [f for f in os.listdir(parent_dir) if os.path.isdir(os.path.join(parent_dir, f))]
for subfolder in subfolders:
    start_time = time.time()
    subfolder_path = os.path.join(parent_dir, subfolder)
    print(f"Processing {subfolder}...")
    # Find the CSV file starting with 'results_MinPenetration'
    csv_files = [f for f in os.listdir(subfolder_path) if f.startswith('results_MinPenetration') and f.endswith('.csv')]
    print(f"  Found CSV files: {csv_files}")
    if not csv_files:
        print(f"  No results_MinPenetration CSV found in {subfolder}")
        continue
    csv_path = os.path.join(subfolder_path, csv_files[0])
    # Robust block parsing for pdtable-like format
    with open(csv_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]
    # Find block start
    block_start = None
    for i, line in enumerate(lines):
        if line.startswith('**results_MinPenetration;'):
            block_start = i
            break
    if block_start is None:
        print(f"  Table marker not found in {csv_path}")
        continue
    idx = block_start + 1
    # Skip position line
    while idx < len(lines) and not lines[idx]:
        idx += 1
    idx += 1
    # Find header line
    while idx < len(lines) and not lines[idx]:
        idx += 1
    header_line = lines[idx]
    col_names = [c for c in header_line.split(';') if c]
    idx += 1
    # Skip units line
    while idx < len(lines) and not lines[idx]:
        idx += 1
    idx += 1
    # Collect data lines until next block or end
    data_lines = []
    while idx < len(lines):
        if lines[idx].startswith('**'):
            break
        if lines[idx] and not lines[idx].startswith(';'):
            row = [c for c in lines[idx].split(';') if c]
            if len(row) == len(col_names):
                data_lines.append(row)
        idx += 1
    if not data_lines:
        print(f"  No data found in block for {subfolder}")
        continue
    df = pd.DataFrame(data_lines, columns=col_names)
    print(f"  DataFrame for {subfolder}:\n{df}")
    # SLS row: SeaState == hs_value, ULS row: SeaState == hs_value + '_ULS'
    sls_row = df[df['SeaState'] == hs_value]
    uls_row = df[df['SeaState'] == hs_value + '_ULS']
    if not sls_row.empty and not uls_row.empty:
        hs_val = sls_row.iloc[0]['SeaState']
        ur_sls_val = sls_row.iloc[0]['UR_Soil_SLS_Lateral_SRCCyclic']
        ur_uls_val = uls_row.iloc[0]['UR_Soil_ULS_Lateral']
        results.append({
            'Position': subfolder,
            'Hs': hs_val,
            'UR_SLS': ur_sls_val,
            'UR_ULS': ur_uls_val
        })
        print(f"  Extracted for {subfolder}")
    else:
        print(f"  SLS or ULS row not found for {subfolder}")
    elapsed = time.time() - start_time
    print(f"  Done {subfolder} in {elapsed:.2f} seconds.")

# Create DataFrame and save
results_df = pd.DataFrame(results)
output_csv = os.path.join(parent_dir, 'min_penetration_summary.csv')
results_df.to_csv(output_csv, index=False)
print(f'Saved {output_csv}')
