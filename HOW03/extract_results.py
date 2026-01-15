"""
Extract and plot pile driving analysis results from multiple positions.

This script:
1. Reads results_PileDrivingAnalysis CSV files from position folders (A01, A02, etc.)
2. Parses multiple tables within each CSV file
3. Allows user to select which SRD methods and soil bounds to plot
4. Creates plots with consistent colors per method and line styles per soil bound
"""

import pandas as pd
from pathlib import Path
import plotly.graph_objects as go
import re
from typing import Dict, List, Tuple, Optional
from plotly.subplots import make_subplots
import os

# --- USER CONFIGURABLE CONSTANTS ---
# Set hard driving and refusal blowcount rates here (units: bl/25cm)
HARD_DRIVING_BLOWCOUNT = 75  # e.g. 80 bl/25cm
REFUSAL_BLOWCOUNT = 250      # e.g. 180 bl/25cm
INTERNAL_LIFTING_TOOL = 100 # weight of internal lifting tool in t
HAMMER_WEIGHT = 739 # weight of hammer in t

# USER DEFINED CONSTANT FOR MONOPILE WEIGHTS FILE
MONOPILE_WEIGHTS_FILE = r"C:/Users/vasav/PyCharmProjects/VASAV-at-work/HOW03/summary-01_Primary_Steel_Design_Verification_25yr.xls"


# -----------------------------------


def parse_results_csv(file_path: Path) -> Dict[str, pd.DataFrame]:
    """
    Parse a results CSV file that contains multiple tables separated by empty rows.
    Each table starts with ** and the table name.

    Returns:
        Dictionary with table names as keys and DataFrames as values
    """
    tables = {}

    # Read the entire file
    with open(str(file_path), 'r', encoding='utf-8') as f:
        content = f.read()

    # Split by double asterisks to find table headers
    sections = re.split(r'\*\*', content)

    for section in sections[1:]:  # Skip first empty section
        lines = section.strip().split('\n')
        if not lines:
            continue

        # First line is the table name (remove trailing semicolon)
        table_name = lines[0].strip().rstrip(';')

        # Find where the actual data starts (after position line and header)
        # Typically: table_name, position, column_headers, unit_row, data...
        if len(lines) < 4:
            continue

        # Position is on line 1
        position = lines[1].strip()

        # Column headers on line 2
        headers = lines[2].strip().split(';')

        # Skip unit row (line 3) and read data starting from line 4
        data_lines = []
        for line in lines[4:]:
            line = line.strip()
            if not line:  # Stop at empty line (next table)
                break
            data_lines.append(line)

        if not data_lines:
            continue

        # Parse data into DataFrame
        data_rows = [line.split(';') for line in data_lines]
        df = pd.DataFrame(data_rows, columns=headers)

        # Convert numeric columns
        for col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col])
            except (ValueError, TypeError):
                pass  # Keep as string if conversion fails

        tables[table_name] = df

    return tables


def get_position_folders(root_dir: Path) -> List[Tuple[str, Path]]:
    """
    Get all position folders (A01, A02, etc.) from the root directory.

    Returns:
        List of tuples (position_name, folder_path)
    """
    positions = []

    if not root_dir.exists():
        print(f"Error: Directory does not exist: {root_dir}")
        return positions

    for item in sorted(root_dir.iterdir()):
        if item.is_dir() and re.match(r'^[A-Z]\d+$', item.name):
            positions.append((item.name, item))

    return positions


def extract_method_and_bound(table_name: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Extract SRD method and soil bound from table name.

    Example: 'results_PileDrivingAnalysis_HOW03_lb_MD_4400S' -> ('MD', 'lb')
    Example: 'results_PileDrivingAnalysis_HOW03_lb_Maynard_4400S' -> ('MY', 'lb')
    """
    # Pattern: ..._{lb|be|ub}_{MD|MY|AH|Maynard}_...
    match = re.search(r'_(lb|be|ub)_(MD|MY|AH|Maynard)_', table_name)
    if match:
        method = match.group(2)
        bound = match.group(1)

        # Map "Maynard" to "MY" for consistency
        if method == "Maynard":
            method = "MY"

        return method, bound
    return None, None


def get_available_methods_and_bounds(tables: Dict[str, pd.DataFrame]) -> Tuple[List[str], List[str]]:
    """
    Extract all unique SRD methods and soil bounds from table names.
    """
    methods = set()
    bounds = set()

    for table_name in tables.keys():
        method, bound = extract_method_and_bound(table_name)
        if method and bound:
            methods.add(method)
            bounds.add(bound)

    return sorted(list(methods)), sorted(list(bounds))


def get_available_methods_and_bounds_from_summary(summary_table: pd.DataFrame) -> Tuple[List[str], List[str], Dict[str, Dict]]:
    """
    Extract all unique mapped SRD methods and soil bounds from the summary table.
    Returns:
        - List of mapped methods (AH, MD, MY)
        - List of soil bounds (lb, be, ub)
        - Dict mapping (method, bound) to analysis_ID and original method
    """
    method_map = {
        "Alm and Hamre": "AH",
        "MonoDrive": "MD",
        "Maynard": "MY"
    }
    methods = set()
    bounds = set()
    method_bound_to_id = {}
    for idx, row in summary_table.iterrows():
        orig_method = str(row.get("method", "")).strip()
        mapped_method = method_map.get(orig_method, None)
        soil_case = str(row.get("SoilCase", "")).strip().lower()
        analysis_id = str(row.get("analysis_ID", "")).strip()
        if mapped_method and soil_case:
            methods.add(mapped_method)
            bounds.add(soil_case)
            method_bound_to_id[(mapped_method, soil_case)] = {
                "analysis_ID": analysis_id,
                "orig_method": orig_method
            }
    return sorted(list(methods)), sorted(list(bounds)), method_bound_to_id


def get_monopile_weights(excel_path: Path, positions: list) -> dict:
    """
    Reads the Excel file and returns a dict mapping position name to monopile weight for the given positions.
    Combines rows 7 and 8 for headers, starts data from row 10.
    Uses exact column names 'Position' and 'MP mass'.
    """
    df_raw = pd.read_excel(excel_path, sheet_name=0, header=None)
    header_rows = df_raw.iloc[6:8].astype(str)
    combined_headers = header_rows.agg(' '.join).str.replace('nan ', '', regex=False).str.strip()
    df_raw.columns = combined_headers
    df = df_raw.iloc[9:].reset_index(drop=True)
    position_col = next((col for col in df.columns if col.strip().lower() == 'position'), None)
    weight_col = next((col for col in df.columns if col.strip().lower() == 'mp mass'), None)
    if position_col is None or weight_col is None:
        print(f"ERROR: Could not find required columns for position or MP mass.")
        print(f"Available columns: {df.columns}")
        return {}
    weights = {}
    for _, row in df.iterrows():
        pos = str(row[position_col]).strip()
        if pos in positions:
            try:
                weight = float(row[weight_col])
            except Exception:
                weight = None
            weights[pos] = weight
    return weights


def get_position_info(position_tables, position):
    """
    Extracts hammer name, hammer weight, and target blowcount rate for a given position.
    Returns a dict with keys: 'hammer_name', 'hammer_weight', 'target_blowcount_rate'.
    """
    # Hammer weight always from user constant
    info = {'hammer_name': None, 'hammer_weight': HAMMER_WEIGHT, 'target_blowcount_rate': None}
    tables = position_tables.get(position, {})
    # Hammer name from first table (usually 'inputs')
    for table in tables.values():
        if 'Hammer_name' in table.columns:
            val = table['Hammer_name'].iloc[0]
            if pd.notna(val):
                info['hammer_name'] = str(val)
            break
    # Target blowcount rate from summary table column 'Target_Blowcount_Rate'
    summary_table = tables.get('results_PileDrivingAnalysis_Summary')
    if summary_table is not None and 'Target_Blowcount_Rate' in summary_table.columns:
        val = summary_table['Target_Blowcount_Rate'].iloc[0]
        if pd.notna(val):
            info['target_blowcount_rate'] = float(val)
    return info


def plot_rut_vs_depth(tables: Dict[str, pd.DataFrame], position: str,
                      selected_methods: List[str], selected_bounds: List[str],
                      output_dir: Path = None,
                      monopile_weights: dict = None,
                      position_info: dict = None):
    """
    Plot Rut vs Depth for selected methods and bounds using Plotly (interactive).
    Each method gets a unique color, each bound gets a unique line style.
    Info panel in row 2, column 4 includes hammer info and target blowcount rate.
    """
    # Extract target depth from summary table
    target_depth = None
    summary_table = tables.get('results_PileDrivingAnalysis_Summary')
    if summary_table is not None and 'targetdepth' in summary_table.columns:
        try:
            target_depth = pd.to_numeric(summary_table['targetdepth'].iloc[0])
            print(f"Target depth: {target_depth} m")
        except (ValueError, IndexError):
            print("Warning: Could not extract target depth from summary table")

    # Define colors for methods
    method_colors = {
        'MD': '#1f77b4',  # blue
        'MY': '#d62728',  # red
        'AH': '#2ca02c'   # green
    }

    # Define line dash patterns for bounds (Plotly format)
    bound_dashes = {
        'be': 'solid',      # best estimate: solid line
        'lb': 'dash',       # lower bound: dashed line
        'ub': 'dot'         # upper bound: dotted line
    }

    # Create subplots: 2 rows, 4 columns (info panel in row 1, col 4)
    fig = make_subplots(
        rows=2, cols=4,
        shared_yaxes=False,
        horizontal_spacing=0.06,
        vertical_spacing=0.18,  # Increased from 0.04 to 0.18 for more separation
        subplot_titles=(None, None, None, None, None, None, None, None),
        specs=[[{}, {}, {}, {}], [{}, {}, {}, {}]]
    )
    plot_count = 0

    # Plot each combination
    for table_name, df in tables.items():
        method, bound = extract_method_and_bound(table_name)

        if method not in selected_methods or bound not in selected_bounds:
            continue

        # Check if required columns exist
        if 'Depth' not in df.columns or 'Rut' not in df.columns or 'Blowcount_rate' not in df.columns or 'Input_Energy' not in df.columns:
            continue

        # Convert to numeric if needed
        depth = pd.to_numeric(df['Depth'], errors='coerce')
        rut = pd.to_numeric(df['Rut'], errors='coerce')
        blowcount_rate = pd.to_numeric(df['Blowcount_rate'], errors='coerce') / 4.0  # Convert bl/m to bl/25cm
        input_energy = pd.to_numeric(df['Input_Energy'], errors='coerce')
        # Use Cumulative_input_energy column if present, else fallback to cumsum of Input_Energy
        cumulative_input_energy = pd.to_numeric(df['Cumulative_input_Energy'], errors='coerce') / 1_000_000


        # Remove NaN values
        mask = ~(depth.isna() | rut.isna() | blowcount_rate.isna() | input_energy.isna() | cumulative_input_energy.isna())
        depth = depth[mask]
        rut = rut[mask]
        blowcount_rate = blowcount_rate[mask]
        input_energy = input_energy[mask]
        cumulative_input_energy = cumulative_input_energy[mask]

        if len(depth) == 0:
            continue

        color = method_colors.get(method, '#000000')
        dash = bound_dashes.get(bound, 'solid')
        label = f'{method}_{bound}'

        # Plot in row 1, column 1: SRD [MN] vs Depth [mbsb]
        fig.add_trace(
            go.Scatter(
                x=rut,
                y=depth,
                mode='lines',
                name=label,
                line=dict(color=color, width=2, dash=dash),
                legendgroup=label,
                hovertemplate='<b>%{fullData.name}</b><br>SRD: %{x:.2f} MN<br>Depth: %{y:.2f} m<br><extra></extra>'
            ),
            row=1, col=1
        )
        # Plot in row 1, column 2: Blowcount Rate [bl/25cm] vs Depth [mbsb]
        fig.add_trace(
            go.Scatter(
                x=blowcount_rate,
                y=depth,
                mode='lines',
                name=label,
                line=dict(color=color, width=2, dash=dash),
                legendgroup=label,
                showlegend=False,
                hovertemplate='<b>%{fullData.name}</b><br>Blowcount Rate: %{x:.2f} bl/25cm<br>Depth: %{y:.2f} m<br><extra></extra>'
            ),
            row=1, col=2
        )
        # Plot in row 1, column 3: Input_Energy [kJ/blow] vs Depth [mbsb]
        fig.add_trace(
            go.Scatter(
                x=input_energy,
                y=depth,
                mode='lines',
                name=label,
                line=dict(color=color, width=2, dash=dash),
                legendgroup=label,
                showlegend=False,
                hovertemplate='<b>%{fullData.name}</b><br>Input Energy: %{x:.2f} kJ/blow<br>Depth: %{y:.2f} m<br><extra></extra>'
            ),
            row=1, col=3
        )
        # New plot: Total Energy [GJ] vs Depth [mbsb] in row 2, col 1
        fig.add_trace(
            go.Scatter(
                x=cumulative_input_energy,
                y=depth,
                mode='lines',
                name=label,
                line=dict(color=color, width=2, dash=dash),
                legendgroup=label,
                showlegend=False,
                hovertemplate='<b>%{fullData.name}</b><br>Total Energy: %{x:.2f} GJ<br>Depth: %{y:.2f} m<br><extra></extra>'
            ),
            row=2, col=1
        )
        plot_count += 1

    if plot_count == 0:
        print("Warning: No data found to plot!")
        return

    # Add target depth horizontal line if available (to all subplots)
    if target_depth is not None:
        fig.add_hline(
            y=target_depth,
            line_dash="dash",
            line_color="black",
            line_width=2,
            annotation_text=f"Target Depth: {target_depth} m",
            annotation_position="top left",
            annotation=dict(font_size=11, font_color="black"),
            row=1, col=1
        )
        fig.add_hline(
            y=target_depth,
            line_dash="dash",
            line_color="black",
            line_width=2,
            showlegend=False,
            row=1, col=2
        )
        fig.add_hline(
            y=target_depth,
            line_dash="dash",
            line_color="black",
            line_width=2,
            showlegend=False,
            row=1, col=3
        )
        fig.add_hline(
            y=target_depth,
            line_dash="dash",
            line_color="black",
            line_width=2,
            showlegend=False,
            row=2, col=1
        )

    # Use hardcoded values instead of user input
    hard_driving_val = HARD_DRIVING_BLOWCOUNT
    refusal_val = REFUSAL_BLOWCOUNT

    # Add vertical lines to second subplot if values are provided
    if hard_driving_val is not None:
        fig.add_vline(
            x=hard_driving_val,
            line_dash="dash",
            line_color="orange",
            line_width=3,
            annotation_text="hard driving",
            annotation_position="top left",
            annotation=dict(font_size=12, font_color="orange", textangle=-90),
            row=1, col=2
        )
    if refusal_val is not None:
        fig.add_vline(
            x=refusal_val,
            line_dash="dash",
            line_color="red",
            line_width=3,
            annotation_text="refusal",
            annotation_position="top left",
            annotation=dict(font_size=12, font_color="red", textangle=-90),
            row=1, col=2
        )

    # Info panel in row 1, column 4 as annotation for visibility
    mp_weight = None
    if monopile_weights is not None:
        mp_weight = monopile_weights.get(position, None)
    info_lines = [f"<b>Position:</b> {position}"]
    if mp_weight is not None:
        info_lines.append(f"<b>Monopile weight:</b> {mp_weight:.1f} t")
    else:
        info_lines.append(f"<b>Monopile weight:</b> N/A")
    if position_info:
        if position_info.get('hammer_name'):
            info_lines.append(f"<b>Hammer:</b> {position_info['hammer_name']}")
        if position_info.get('hammer_weight') is not None:
            info_lines.append(f"<b>Hammer weight:</b> {position_info['hammer_weight']:.1f} t")
        if position_info.get('target_blowcount_rate') is not None:
            info_lines.append(f"<b>Target blowcount rate:</b> {position_info['target_blowcount_rate'] / 4.0:.2f} bl/25cm")
    info_text = '<br>'.join(info_lines)
    fig.add_annotation(
        text=info_text,
        xref='x3', yref='y3',
        x=8.0, y=1.0,  # Adjust x as needed for "slightly right"
        xanchor='right', yanchor='top',
        showarrow=False,
        font=dict(size=16),
        align='left',
        row=1, col=4
    )
    fig.update_xaxes(visible=False, row=1, col=4)
    fig.update_yaxes(visible=False, row=1, col=4)

    # Update layout for all subplotsA
    fig.update_layout(
        title={
            'text': f'Position {position}',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 18, 'family': 'Arial, sans-serif'}
        },
        xaxis=dict(
            title='SRD [MN]',
            gridcolor='lightgray',
            showgrid=True,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True
        ),
        xaxis2=dict(
            title='Blowcount Rate [bl/25cm]',
            gridcolor='lightgray',
            showgrid=True,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True,
            range=[0, 250]
        ),
        xaxis3=dict(
            title='Input Energy [kJ/blow]',
            gridcolor='lightgray',
            showgrid=True,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True
        ),
        xaxis4=dict(
            title='Total Energy [GJ]',
            gridcolor=None,
            showgrid=False,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True
        ),
        yaxis=dict(
            title='Depth [mbsb]',
            autorange='reversed',
            rangemode='tozero',
            gridcolor='lightgray',
            showgrid=True,
            dtick=5,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True,
            showticklabels=True
        ),
        yaxis2=dict(
            title='Depth [mbsb]',
            autorange='reversed',
            rangemode='tozero',
            gridcolor='lightgray',
            showgrid=True,
            dtick=5,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True,
            showticklabels=True
        ),
        yaxis3=dict(
            title='Depth [mbsb]',
            autorange='reversed',
            rangemode='tozero',
            gridcolor='lightgray',
            showgrid=True,
            dtick=5,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True,
            showticklabels=True
        ),
        yaxis4=dict(
            title='Depth [mbsb]',
            autorange='reversed',
            rangemode='tozero',
            gridcolor=None,
            showgrid=False,
            dtick=5,
            linecolor='#5a5a5a',
            linewidth=2,
            mirror=True,
            showticklabels=True
        ),
        hovermode='closest',
        template='plotly_white',
        legend=dict(
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.02,
            bgcolor="rgba(255, 255, 255, 0.8)",
            bordercolor="Black",
            borderwidth=1
        ),
        width=2200,
        height=1200,
        font=dict(size=12),
        plot_bgcolor='white',
        paper_bgcolor='white'
    )

    # Update layout for row 2, col 1 (Total Energy plot)
    fig.update_xaxes(
        title='Total Energy [GJ]',
        gridcolor='lightgray',
        showgrid=True,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=1
    )
    fig.update_yaxes(
        title='Depth [mbsb]',
        autorange='reversed',
        rangemode='tozero',
        gridcolor='#5a5a5a',
        showgrid=True,
        dtick=5,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=1
    )

    # Save or show
    if output_dir:
        output_path = output_dir / f'rut_vs_depth_{position}.html'
        fig.write_html(str(output_path))
        print(f"Plot saved to: {output_path}")

    # Show interactive plot
    fig.show()
    print(f"✓ Interactive plot displayed with {plot_count} traces per subplot")

def main():
    """Main execution function"""

    # Set the root directory
    root_dir = Path(r"K:\DOZR\HOW03\GEO\05_Driveability\20241108_Cleaned_Const_blow\variations\const_blow_0.25m_intrvl_200bl_limit\monopiles")

    # For testing locally, allow fallback to current directory
    if not root_dir.exists():
        print(f"Warning: Directory not found: {root_dir}")
        print("Trying local HOW03 directory for testing...")
        root_dir = Path(__file__).parent

    # Get position folders
    positions = get_position_folders(root_dir)

    if not positions:
        print("No position folders found!")
        return

    print(f"Found {len(positions)} positions: {[p[0] for p in positions]}")

    # Gather all available methods and bounds from all selected positions
    all_methods = set()
    all_bounds = set()
    position_tables = {}
    for p_name, p_path in positions:
        csv_file = p_path / f"results_PileDrivingAnalysis-{p_name}.csv"
        if csv_file.exists():
            tables = parse_results_csv(csv_file)
            position_tables[p_name] = tables
            methods, bounds = get_available_methods_and_bounds(tables)
            all_methods.update(methods)
            all_bounds.update(bounds)

    methods = sorted(list(all_methods))
    bounds = sorted(list(all_bounds))

    # Restore user selection for positions
    print("\n" + "="*60)
    print("SELECT POSITION(S) TO PLOT")
    print("="*60)
    print(f"\nAvailable positions: {', '.join([p[0] for p in positions])}")
    selected_positions_input = input("Enter positions to plot (comma-separated, e.g., A01,A02 or 'all'): ").strip()

    if selected_positions_input.lower() == 'all':
        selected_positions = [p[0] for p in positions]
    else:
        selected_positions = [p.strip().upper() for p in selected_positions_input.split(',') if p.strip()]

    # User selection for methods and bounds
    print("\n" + "="*60)
    print("SELECT SRD METHODS AND SOIL BOUNDS")
    print("="*60)
    print(f"\nAvailable SRD methods: {', '.join(methods)}")
    selected_methods_input = input("Enter methods to plot (comma-separated, e.g., MD,AH or 'all'): ").strip()

    if selected_methods_input.lower() == 'all':
        selected_methods = methods
    else:
        selected_methods = [m.strip().upper() for m in selected_methods_input.split(',') if m.strip()]

    print(f"Selected methods: {selected_methods}")

    print(f"\nAvailable soil bounds: {', '.join(bounds)}")
    selected_bounds_input = input("Enter bounds to plot (comma-separated, e.g., lb,ub or 'all'): ").strip()

    if selected_bounds_input.lower() == 'all':
        selected_bounds = bounds
    else:
        selected_bounds = [b.strip().lower() for b in selected_bounds_input.split(',') if b.strip()]

    print(f"Selected bounds: {selected_bounds}")

    # Create output directory if it doesn't exist
    output_dir = root_dir / "plots"
    output_dir.mkdir(exist_ok=True)

    # Read monopile weights for selected positions
    excel_path = Path(MONOPILE_WEIGHTS_FILE)
    if not excel_path.exists():
        print(f"Warning: Monopile weight Excel file not found: {excel_path}")
        monopile_weights = None
    else:
        monopile_weights = get_monopile_weights(excel_path, selected_positions)

    # Plot for selected positions
    for position in selected_positions:
        if position in position_tables:
            tables = position_tables[position]
            position_info = get_position_info(position_tables, position)
            plot_rut_vs_depth(tables, position, selected_methods, selected_bounds, output_dir, monopile_weights, position_info)
        else:
            print(f"Warning: No data found for position {position}")

    print("\n✓ Processing complete. Check plots in the 'plots' directory.")


if __name__ == "__main__":
    main()
