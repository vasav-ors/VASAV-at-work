# minimum_penetration_plot.py
# This script reads MP position data from an Excel file and generates three separate interactive Plotly maps:
# 1. Minimum penetration (minL) for selected Hs value
# 2. Utilisation ratio for SLS (UR_sls) for selected Hs value
# 3. Utilisation ratio for ULS (UR_uls) for selected Hs value
# Each plot is saved as an HTML file and opened automatically.

# User inputs
excel_file = r'k:\dozr\HOW03\GEO\04_OptiMon Runs\20251017_Lateral_pile_stability_ Installation\post-processing\HOW03_minL_load_iter1.xlsm'  # Excel file name
sheet_name = 'Summary 05_combined'         # Sheet name
hs_value = '7.3'                          # Hs value to plot: '2_5', '6_4', or '7_3'

import pandas as pd
import plotly.graph_objects as go
import os
import webbrowser

# Path to the Excel file
excel_path = os.path.join(os.path.dirname(__file__), excel_file)

# Read the Excel file, no header, skip rows above 16 (so rows 0-15 are skipped)
df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, skiprows=16)

# Combine rows 0, 1, 2 (Excel rows 17, 18, 19) into headers
combined_headers = df_raw.iloc[0:3].astype(str).agg(' '.join).str.replace('nan ', '', regex=False)
combined_headers = [str(h) for h in combined_headers]

# Create df with actual data, skipping the header rows
df = df_raw.iloc[3:].copy()
df.columns = pd.Index(combined_headers)
df.reset_index(drop=True, inplace=True)

# Find columns for Easting (E_) and Northing (N_)
easting_cols = [col for col in df.columns if 'E_' in col]
northing_cols = [col for col in df.columns if 'N_' in col]

if not easting_cols or not northing_cols:
    print('Column names:', df.columns)
    raise ValueError('No E_ or N_ columns found in the sheet.')

E_idx = df.columns.get_loc(easting_cols[0])
N_idx = df.columns.get_loc(northing_cols[0])
name_idx = 0

# Convert E and N to numeric, drop rows with NaN
E = pd.to_numeric(df.iloc[:, E_idx], errors='coerce')
N = pd.to_numeric(df.iloc[:, N_idx], errors='coerce')
position_names = df.iloc[:, name_idx]

valid_mask = (~E.isna()) & (~N.isna())
E = E[valid_mask]
N = N[valid_mask]
position_names = position_names[valid_mask]

# Accept hs_value as e.g. '6.4' and convert to '6_4' for column search
hs_col_key = hs_value.replace('.', '_')

# Select columns and colorbar title based on hs_value
sls_col = [col for col in df.columns if f'Lmin_Hs{hs_col_key}' in col and '_ULS' not in col]
uls_col = [col for col in df.columns if f'Lmin_Hs{hs_col_key}_ULS' in col]
if not sls_col or not uls_col:
    raise ValueError(f"Columns for Hs={hs_value} not found in the sheet.")
sls_col = sls_col[int(0)]
uls_col = uls_col[int(0)]
colorbar_title = f'minL; Hs={hs_value}m'

# Convert SLS and ULS columns to numeric
SLS = pd.to_numeric(df[sls_col], errors='coerce')
ULS = pd.to_numeric(df[uls_col], errors='coerce')

# Use the maximum of SLS and ULS for color coding
color_vals = pd.concat([SLS, ULS], axis=1).max(axis=1)

# Filter valid rows (E, N, SLS, ULS must all be present)
valid_mask = (~E.isna()) & (~N.isna()) & (~SLS.isna()) & (~ULS.isna())
E = E[valid_mask]
N = N[valid_mask]
position_names = position_names[valid_mask]
SLS = SLS[valid_mask]
ULS = ULS[valid_mask]
color_vals = color_vals[valid_mask]

# Build custom labels for bottom (position name, bigger)
position_labels = [f"<span style='font-size:12px'><b>{str(name)}</b></span>" for name in position_names]

# Build custom hover text for each position
hover_labels = [
    f"{name}<br>SLS: {sls:.2f} m<br>ULS: {uls:.2f} m"
    for name, sls, uls in zip(position_names, SLS, ULS)
]

# Plotly scatter plot with color coding
fig = go.Figure()
fig.add_trace(go.Scatter(
    x=E,
    y=N,
    mode='markers+text',
    marker=dict(
        size=18,  # Increased marker size
        color=color_vals,
        colorscale='Turbo',  # More distinct color scale
        colorbar=dict(title=colorbar_title),
        showscale=True
    ),
    text=position_labels,
    textposition='bottom center',  # All names below each marker
    textfont=dict(size=12),  # Make position name smaller
    hovertext=hover_labels,
    hoverinfo='text',
    showlegend=False
))

min_easting = E.min()
max_easting = E.max()
range_buffer = max(10, int(0.01 * (max_easting - min_easting)))
xaxis_min = min_easting - range_buffer
xaxis_max = max_easting + range_buffer

fig.update_layout(
    xaxis_title='<b>Easting</b>',
    yaxis_title='<b>Northing</b>',
    title=f'MP Minimum penetration for Hs={hs_value}m',
    template='plotly_white',
    xaxis_range=[xaxis_min, xaxis_max],
    margin=dict(l=40, r=40, t=60, b=40),
    paper_bgcolor='white',
    plot_bgcolor='white',
    shapes=[dict(
        type='rect',
        xref='paper', yref='paper',
        x0=0, y0=0, x1=1, y1=1,
        line=dict(color='black', width=2),
        fillcolor='rgba(0,0,0,0)'
    )]
)

# Save HTML in the same folder as the Excel file
html_path = os.path.join(os.path.dirname(excel_file), f"minimum_penetration_plot_Hs{hs_value}.html")
fig.write_html(html_path, include_plotlyjs='cdn')
print(f"Saved: {html_path}")
webbrowser.open(html_path)

# --- Additional plot: UR_sls map for selected Hs ---
# Extract header rows for column matching
hs_row = df_raw.iloc[0].astype(str)
# Find the correct column: row 0 must be Hs_{hs_search}, combined_headers must contain UR_sls
hs_search = hs_value.replace('.', '_')
hs_col_indices = [i for i, val in enumerate(hs_row) if val.strip() == f'Hs_{hs_search}']
ur_sls_col_indices = [int(i) for i in hs_col_indices if 'UR_sls' in combined_headers[int(i)]]
if not ur_sls_col_indices:
    print('DEBUG: hs_row:', list(hs_row))
    print('DEBUG: combined_headers:', [combined_headers[int(i)] for i in hs_col_indices])
    raise ValueError(f'No column found with Hs_{hs_search} in row 17 and UR_sls in header')
ur_sls_col_idx = int(ur_sls_col_indices[0])
ur_sls_col_name = combined_headers[ur_sls_col_idx]

# Extract UR_sls data for each MP position
UR_sls = pd.to_numeric(df.iloc[:, ur_sls_col_idx], errors='coerce')

# Filter valid rows (E, N, UR_sls must all be present)
valid_mask_ur = (~E.isna()) & (~N.isna()) & (~UR_sls.isna())
E_ur = E[valid_mask_ur]
N_ur = N[valid_mask_ur]
position_names_ur = position_names[valid_mask_ur]
UR_sls = UR_sls[valid_mask_ur]

# Build custom labels and hover text
position_labels_ur = [f"<span style='font-size:12px'><b>{str(name)}</b></span>" for name in position_names_ur]
hover_labels_ur = [
    f"{name}<br>UR_sls: {ur:.2f}" for name, ur in zip(position_names_ur, UR_sls)
]

# Plotly scatter plot for UR_sls
fig_ur = go.Figure()
fig_ur.add_trace(go.Scatter(
    x=E_ur,
    y=N_ur,
    mode='markers+text',
    marker=dict(
        size=18,
        color=UR_sls,
        colorscale='Viridis',
        colorbar=dict(title=f'UR_SLS; Hs={hs_value}m'),
        showscale=True,
        cmin=0,
        cmax=1
    ),
    text=position_labels_ur,
    textposition='bottom center',
    textfont=dict(size=12),
    hovertext=hover_labels_ur,
    hoverinfo='text',
    showlegend=False
))
fig_ur.update_layout(
    xaxis_title='<b>Easting</b>',
    yaxis_title='<b>Northing</b>',
    title=f'MP Utilisation ratio for SLS for Hs={hs_value}m',
    template='plotly_white',
    xaxis_range=[xaxis_min, xaxis_max],
    margin=dict(l=40, r=40, t=60, b=40),
    paper_bgcolor='white',
    plot_bgcolor='white',
    shapes=[dict(
        type='rect',
        xref='paper', yref='paper',
        x0=0, y0=0, x1=1, y1=1,
        line=dict(color='black', width=2),
        fillcolor='rgba(0,0,0,0)'
    )]
)

# Save HTML for UR_sls plot
html_path_ur = os.path.join(os.path.dirname(excel_file), f"UR_sls_map_Hs{hs_value}.html")
fig_ur.write_html(html_path_ur, include_plotlyjs='cdn')
print(f"Saved: {html_path_ur}")
webbrowser.open(html_path_ur)

# --- Additional plot: UR_uls map for selected Hs ---
# Find the correct column: row 0 must be Hs_{hs_search}, combined_headers must contain UR_uls
ur_uls_col_indices = [int(i) for i in hs_col_indices if 'UR_uls' in combined_headers[int(i)]]
if not ur_uls_col_indices:
    print('DEBUG: hs_row:', list(hs_row))
    print('DEBUG: combined_headers:', [combined_headers[int(i)] for i in hs_col_indices])
    raise ValueError(f'No column found with Hs_{hs_search} in row 17 and UR_uls in header')
ur_uls_col_idx = int(ur_uls_col_indices[0])
ur_uls_col_name = combined_headers[ur_uls_col_idx]

# Extract UR_uls data for each MP position
UR_uls = pd.to_numeric(df.iloc[:, ur_uls_col_idx], errors='coerce')

# Filter valid rows (E, N, UR_uls must all be present)
valid_mask_uls = (~E.isna()) & (~N.isna()) & (~UR_uls.isna())
E_uls = E[valid_mask_uls]
N_uls = N[valid_mask_uls]
position_names_uls = position_names[valid_mask_uls]
UR_uls = UR_uls[valid_mask_uls]

# Build custom labels and hover text
position_labels_uls = [f"<span style='font-size:12px'><b>{str(name)}</b></span>" for name in position_names_uls]
hover_labels_uls = [
    f"{name}<br>UR_uls: {ur:.2f}" for name, ur in zip(position_names_uls, UR_uls)
]

# Plotly scatter plot for UR_uls
fig_uls = go.Figure()
fig_uls.add_trace(go.Scatter(
    x=E_uls,
    y=N_uls,
    mode='markers+text',
    marker=dict(
        size=18,
        color=UR_uls,
        colorscale='Viridis',
        colorbar=dict(title=f'UR_ULS; Hs={hs_value}m'),
        showscale=True,
        cmin=0,
        cmax=1
    ),
    text=position_labels_uls,
    textposition='bottom center',
    textfont=dict(size=12),
    hovertext=hover_labels_uls,
    hoverinfo='text',
    showlegend=False
))
fig_uls.update_layout(
    xaxis_title='<b>Easting</b>',
    yaxis_title='<b>Northing</b>',
    title=f'MP Utilisation ratio for ULS for Hs={hs_value}m',
    template='plotly_white',
    xaxis_range=[xaxis_min, xaxis_max],
    margin=dict(l=40, r=40, t=60, b=40),
    paper_bgcolor='white',
    plot_bgcolor='white',
    shapes=[dict(
        type='rect',
        xref='paper', yref='paper',
        x0=0, y0=0, x1=1, y1=1,
        line=dict(color='black', width=2),
        fillcolor='rgba(0,0,0,0)'
    )]
)

# Save HTML for UR_uls plot
html_path_uls = os.path.join(os.path.dirname(excel_file), f"UR_uls_map_Hs{hs_value}.html")
fig_uls.write_html(html_path_uls, include_plotlyjs='cdn')
print(f"Saved: {html_path_uls}")
webbrowser.open(html_path_uls)

# --- Remove combined subplot code ---
# Only save and open the individual HTML files for each plot
# (minimum penetration and UR_sls)
