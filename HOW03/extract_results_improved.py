"""
Extract and plot pile driving analysis results from multiple positions.

This script:
1. Reads results_PileDrivingAnalysis CSV files from position folders (A01, A02, etc.)
2. Parses multiple tables within each CSV file
3. Allows user to select which SRD methods and soil bounds to plot
4. Creates plots with consistent colors per method and line styles per soil bound

Features:
- Robust error handling for reading monopile weights from CSV/Excel files
- Flexible column name matching for various input formats
- Professional data tables for position information and weights
- Interactive Plotly visualizations with customizable thresholds
"""

import pandas as pd
from pathlib import Path
import plotly.graph_objects as go
import re
from typing import Dict, List, Tuple, Optional
from plotly.subplots import make_subplots

# --- USER CONFIGURABLE CONSTANTS ---
# Set hard driving and refusal blowcount rates here (units: bl/25cm)
HARD_DRIVING_BLOWCOUNT = 75  # e.g. 80 bl/25cm
REFUSAL_BLOWCOUNT = 250      # e.g. 180 bl/25cm
INTERNAL_LIFTING_TOOL = 100 # weight of internal lifting tool in t
HAMMER_WEIGHT = 736 # weight of hammer in t
ADDITIONAL_WEIGHT = 20 # any additional weight in MP like flange and pins for secondary attachments int

# USER DEFINED CONSTANT FOR MONOPILE WEIGHTS FILE
MONOPILE_WEIGHTS_FILE = r"C:/Users/vasav/PyCharmProjects/VASAV-at-work/HOW03/summary-01_Primary_Steel_Design_Verification_25yr.xls"

# USER DEFINED CONSTANT FOR MONOPILE ROOT DIRECTORY
MONOPILE_ROOT_DIR = Path(r"K:/DOZR/HOW03/GEO/05_Driveability/20241108_Cleaned_Const_blow/variations/const_blow_0.25m_intrvl_200bl_limit/monopiles")

# USER DEFINED CONSTANT FOR OUTPUT DIRECTORY WHERE PLOTS ARE SAVED
PLOTS_OUTPUT_DIR = MONOPILE_ROOT_DIR / "plots"


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

    IMPROVED VERSION with better error handling and more flexible column matching.
    """
    try:
        df_raw = pd.read_excel(excel_path, sheet_name=0, header=None)

        # Combine header rows (rows 7 and 8, 0-indexed as 6 and 7)
        header_rows = df_raw.iloc[6:8].astype(str)
        combined_headers = header_rows.agg(' '.join).str.replace('nan ', '', regex=False).str.strip()
        df_raw.columns = combined_headers

        # Data starts at row 10 (0-indexed as 9)
        df = df_raw.iloc[9:].reset_index(drop=True)

        # Normalize column names for flexible matching
        col_mapping = {col: col.strip().lower() for col in df.columns}

        # Find position and weight columns (case-insensitive and flexible matching)
        position_col = next((col for col, norm in col_mapping.items() if norm == 'position'), None)
        weight_col = next((col for col, norm in col_mapping.items() if 'mp mass' in norm or 'mp_mass' in norm), None)

        if position_col is None or weight_col is None:
            print(f"ERROR: Could not find required columns for position or MP mass.")
            print(f"Available columns: {list(df.columns)}")
            return {}

        # Build weights dictionary
        weights = {}
        for _, row in df.iterrows():
            pos = str(row[position_col]).strip()
            if pos in positions:
                try:
                    weight = float(row[weight_col])
                except (ValueError, TypeError):
                    weight = None
                weights[pos] = weight

        return weights

    except Exception as e:
        print(f"ERROR reading monopile weights from {excel_path}: {e}")
        return {}


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


def plot_driveability_results(tables: Dict[str, pd.DataFrame], position: str,
                      selected_methods: List[str], selected_bounds: List[str],
                      output_dir: Path = None,
                      monopile_weights: dict = None,
                      position_info: dict = None):
    """
    Plot Rut vs Depth for selected methods and bounds using Plotly (interactive).
    Each method gets a unique color, each bound gets a unique line style.

    IMPROVED VERSION: Info panel now uses go.Table for professional appearance.
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

    # Create subplots: 2 rows, 4 columns (info panel in row 1, col 4; weight table in row 2, col 4)
    # IMPORTANT: Specify 'table' type for row 1, col 4 and row 2, col 4 to support go.Table
    # Optimized for A3 paper (landscape): 420mm x 297mm
    # row_heights: [0.5, 0.5] gives equal height to both rows
    # column_widths: [0.23, 0.23, 0.23, 0.31] gives slightly more space to table column
    fig = make_subplots(
        rows=2, cols=4,
        shared_yaxes=False,
        row_heights=[0.5, 0.5],  # Equal height for both rows - optimized for A3 printing
        column_widths=[0.23, 0.23, 0.23, 0.31],  # Equal width for plots, slightly wider for tables
        horizontal_spacing=0.06,
        vertical_spacing=0.12,  # Adequate spacing between rows for A3 layout
        subplot_titles=(None, None, None, None, None, None, None, None),
        specs=[[{'type': 'xy'}, {'type': 'xy'}, {'type': 'xy'}, {'type': 'table'}],
               [{'type': 'xy'}, {'type': 'xy'}, {'type': 'xy'}, {'type': 'table'}]]
    )
    plot_count = 0

    methods_to_plot = []
    bounds_to_plot = []
    ruts_to_plot = []
    depths_to_plot = []
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
        blowcount_rate_per_m = pd.to_numeric(df['Blowcount_rate'], errors='coerce')  # Keep original bl/m for cumulative calculation
        input_energy = pd.to_numeric(df['Input_Energy'], errors='coerce')
        # Use Cumulative_input_energy column if present, else fallback to cumsum of Input_Energy
        cumulative_input_energy = pd.to_numeric(df['Cumulative_input_Energy'], errors='coerce') / 1_000_000

        # Remove NaN values
        mask = ~(depth.isna() | rut.isna() | blowcount_rate.isna() | blowcount_rate_per_m.isna() | input_energy.isna() | cumulative_input_energy.isna())
        depth = depth[mask]
        rut = rut[mask]
        blowcount_rate = blowcount_rate[mask]
        blowcount_rate_per_m = blowcount_rate_per_m[mask]
        input_energy = input_energy[mask]
        cumulative_input_energy = cumulative_input_energy[mask]

        if len(depth) == 0:
            continue

        # Calculate cumulative blows
        # For each depth interval, calculate blows = blowcount_rate * depth_increment
        # Then accumulate
        depth_increments = depth.diff().fillna(depth.iloc[0] if len(depth) > 0 else 0)
        blows_per_interval = blowcount_rate_per_m * depth_increments
        cumulative_blows = blows_per_interval.cumsum()

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
        # New plot: Cumulative Blows vs Depth [mbsb] in row 2, col 2
        fig.add_trace(
            go.Scatter(
                x=cumulative_blows,
                y=depth,
                mode='lines',
                name=label,
                line=dict(color=color, width=2, dash=dash),
                legendgroup=label,
                showlegend=False,
                hovertemplate='<b>%{fullData.name}</b><br>Cumulative Blows: %{x:.0f}<br>Depth: %{y:.2f} m<br><extra></extra>'
            ),
            row=2, col=2
        )
        # New plot: SRD [kN] vs Depth [mbsb] in row 2, col 3 (logarithmic scale)
        rut_kn = rut * 1000  # Convert MN to kN
        fig.add_trace(
            go.Scatter(
                x=rut_kn,
                y=depth,
                mode='lines',
                name=label,
                line=dict(color=color, width=2, dash=dash),
                legendgroup=label,
                showlegend=False,
                hovertemplate='<b>%{fullData.name}</b><br>SRD: %{x:.0f} kN<br>Depth: %{y:.2f} m<br><extra></extra>'
            ),
            row=2, col=3
        )
        # Collect for intersection calculation
        methods_to_plot.append(method)
        bounds_to_plot.append(bound)
        ruts_to_plot.append(rut)
        depths_to_plot.append(depth)
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
        fig.add_hline(
            y=target_depth,
            line_dash="dash",
            line_color="black",
            line_width=2,
            showlegend=False,
            row=2, col=2
        )
        fig.add_hline(
            y=target_depth,
            line_dash="dash",
            line_color="black",
            line_width=2,
            showlegend=False,
            row=2, col=3
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

    # Add vertical lines to row 2, col 3 (SRD in kN log scale) for weight references
    mp_weight = None
    if monopile_weights is not None:
        mp_weight = monopile_weights.get(position, None)

    if mp_weight is not None:
        # Calculate nominal monopile weight (MP + additional weight)
        nominal_mp_weight = mp_weight + ADDITIONAL_WEIGHT

        # Add line for MP + additional weight + lifting tool weight (dashed black line)
        mp_lift_tool_total_weight_kn = (nominal_mp_weight + INTERNAL_LIFTING_TOOL) * 9.81
        fig.add_vline(
            x=mp_lift_tool_total_weight_kn,
            line_dash="dash",
            line_color="black",
            line_width=2,
            annotation_text="MP+add.+lift tool",
            annotation_position="top right",
            annotation=dict(font_size=11, font_color="black", textangle=-90),
            row=2, col=3
        )

        # Add line for MP + additional weight + hammer (solid black line)
        mp_hammer_total_weight_kn = (nominal_mp_weight + HAMMER_WEIGHT) * 9.81
        fig.add_vline(
            x=mp_hammer_total_weight_kn,
            line_dash="solid",
            line_color="black",
            line_width=2,
            annotation_text="MP+add.+hammer",
            annotation_position="top right",
            annotation=dict(font_size=11, font_color="black", textangle=-90),
            row=2, col=3
        )

    # --- Find intersection depths for SWP MP + ILT for each method/bound ---
    intersection_depths = {}  # (method, bound) -> depth string
    if mp_weight is not None:
        mp_lift_tool_total_weight_kn = (nominal_mp_weight + INTERNAL_LIFTING_TOOL) * 9.81
        for method, bound, rut, depth in zip(methods_to_plot, bounds_to_plot, ruts_to_plot, depths_to_plot):
            rut_kn = rut * 1000
            for i in range(1, len(rut_kn)):
                if (rut_kn[i-1] < mp_lift_tool_total_weight_kn <= rut_kn[i]) or (rut_kn[i-1] > mp_lift_tool_total_weight_kn >= rut_kn[i]):
                    d1, d2 = depth[i-1], depth[i]
                    r1, r2 = rut_kn[i-1], rut_kn[i]
                    if r2 != r1:
                        depth_cross = d1 + (mp_lift_tool_total_weight_kn - r1) * (d2 - d1) / (r2 - r1)
                    else:
                        depth_cross = d1
                    intersection_depths[(method, bound)] = f'{depth_cross:.2f}'
                    break
    # --- Find intersection depths for SWP MP + ILT for each bound ---
    swp_mp_ilt_depths = {'LB': '', 'BE': '', 'UB': ''}
    swp_mp_ilt_depths_numeric = {'LB': None, 'BE': None, 'UB': None}  # Store numeric values for later use
    if mp_weight is not None:
        mp_lift_tool_total_weight_kn = (nominal_mp_weight + INTERNAL_LIFTING_TOOL) * 9.81
        # For each plotted line in row=2, col=3, find first crossing
        for method, bound, rut, depth in zip(methods_to_plot, bounds_to_plot, ruts_to_plot, depths_to_plot):
            rut_kn = rut * 1000
            # Find first crossing
            for i in range(1, len(rut_kn)):
                if (rut_kn[i-1] < mp_lift_tool_total_weight_kn <= rut_kn[i]) or (rut_kn[i-1] > mp_lift_tool_total_weight_kn >= rut_kn[i]):
                    # Linear interpolation for depth
                    d1, d2 = depth[i-1], depth[i]
                    r1, r2 = rut_kn[i-1], rut_kn[i]
                    if r2 != r1:
                        depth_cross = d1 + (mp_lift_tool_total_weight_kn - r1) * (d2 - d1) / (r2 - r1)
                    else:
                        depth_cross = d1
                    if bound == 'lb':
                        swp_mp_ilt_depths['LB'] = f'{depth_cross:.2f}'
                        swp_mp_ilt_depths_numeric['LB'] = depth_cross
                    elif bound == 'be':
                        swp_mp_ilt_depths['BE'] = f'{depth_cross:.2f}'
                        swp_mp_ilt_depths_numeric['BE'] = depth_cross
                    elif bound == 'ub':
                        swp_mp_ilt_depths['UB'] = f'{depth_cross:.2f}'
                        swp_mp_ilt_depths_numeric['UB'] = depth_cross
                    break

    # --- Find depth where SRD becomes less than nominal MP + lifting tool (below SWP MP+ILT depth) ---
    pile_run_at_hammer_placement = {'LB': '', 'BE': '', 'UB': ''}
    if mp_weight is not None:
        mp_lift_tool_total_weight_kn = (nominal_mp_weight + INTERNAL_LIFTING_TOOL) * 9.81
        # For each bound, find where SRD drops below mp_lift_tool_total_weight_kn
        for method, bound, rut, depth in zip(methods_to_plot, bounds_to_plot, ruts_to_plot, depths_to_plot):
            rut_kn = rut * 1000
            bound_key = bound.upper()

            # Get the SWP MP+ILT depth for this bound
            swp_depth = swp_mp_ilt_depths_numeric.get(bound_key, None)

            if swp_depth is not None:
                # Look for first depth below swp_depth where SRD < mp_lift_tool_total_weight_kn
                found_risk = False
                for i in range(1, len(rut_kn)):
                    current_depth = depth[i]
                    # Only consider depths below the SWP MP+ILT depth
                    if current_depth > swp_depth:
                        # Check if SRD drops below the lifting tool threshold
                        if rut_kn[i] < mp_lift_tool_total_weight_kn:
                            # Linear interpolation to find exact crossing depth
                            if i > 0 and rut_kn[i-1] >= mp_lift_tool_total_weight_kn:
                                d1, d2 = depth[i-1], depth[i]
                                r1, r2 = rut_kn[i-1], rut_kn[i]
                                if r2 != r1:
                                    depth_cross = d1 + (mp_lift_tool_total_weight_kn - r1) * (d2 - d1) / (r2 - r1)
                                else:
                                    depth_cross = current_depth
                                pile_run_at_hammer_placement[bound_key] = f'{depth_cross:.2f}'
                            else:
                                pile_run_at_hammer_placement[bound_key] = f'{current_depth:.2f}'
                            found_risk = True
                            break

                # If no risk found, report "No risk"
                if not found_risk:
                    pile_run_at_hammer_placement[bound_key] = 'No risk'
    # --- Find depth for SWP MP + Hammer for each bound ---
    swp_mp_hammer_depths = {'LB': '', 'BE': '', 'UB': ''}
    swp_mp_hammer_depths_numeric = {'LB': None, 'BE': None, 'UB': None}
    if mp_weight is not None:
        mp_hammer_total_weight_kn = (nominal_mp_weight + HAMMER_WEIGHT) * 9.81
        for method, bound, rut, depth in zip(methods_to_plot, bounds_to_plot, ruts_to_plot, depths_to_plot):
            rut_kn = rut * 1000
            bound_key = bound.upper()
            # Get numeric values of first two rows
            depth1 = swp_mp_ilt_depths_numeric.get(bound_key, None)
            depth2 = pile_run_at_hammer_placement.get(bound_key, None)
            try:
                depth2_num = float(depth2) if depth2 and depth2 != 'No risk' else None
            except Exception:
                depth2_num = None
            # Compute max depth
            start_depth = None
            if depth1 is not None and depth2_num is not None:
                start_depth = max(depth1, depth2_num)
            elif depth1 is not None:
                start_depth = depth1
            elif depth2_num is not None:
                start_depth = depth2_num
            # Find first crossing after start_depth
            if start_depth is not None:
                for i in range(1, len(rut_kn)):
                    if depth[i] > start_depth:
                        # Only upward crossing: SRD rises above MP+Hammer total weight
                        if rut_kn[i-1] < mp_hammer_total_weight_kn <= rut_kn[i]:
                            d1, d2 = depth[i-1], depth[i]
                            r1, r2 = rut_kn[i-1], rut_kn[i]
                            if r2 != r1:
                                depth_cross = d1 + (mp_hammer_total_weight_kn - r1) * (d2 - d1) / (r2 - r1)
                            else:
                                depth_cross = d1
                            swp_mp_hammer_depths[bound_key] = f'{depth_cross:.2f}'
                            swp_mp_hammer_depths_numeric[bound_key] = depth_cross
                            break
    # --- Find depth for pile run risk top (initiation) for each bound ---
    pile_run_risk_top = {'LB': '', 'BE': '', 'UB': ''}
    if mp_weight is not None:
        mp_hammer_total_weight_kn = (nominal_mp_weight + HAMMER_WEIGHT) * 9.81
        for method, bound, rut, depth in zip(methods_to_plot, bounds_to_plot, ruts_to_plot, depths_to_plot):
            rut_kn = rut * 1000
            bound_key = bound.upper()
            swp_mp_hammer_depth = swp_mp_hammer_depths_numeric.get(bound_key, None)
            if swp_mp_hammer_depth is not None:
                found_risk = False
                for i in range(1, len(rut_kn)):
                    current_depth = depth[i]
                    # Only consider depths below the SWP MP+Hammer depth
                    if current_depth > swp_mp_hammer_depth:
                        if rut_kn[i] < mp_hammer_total_weight_kn:
                            # Linear interpolation to find exact crossing depth
                            if i > 0 and rut_kn[i-1] >= mp_hammer_total_weight_kn:
                                d1, d2 = depth[i-1], depth[i]
                                r1, r2 = rut_kn[i-1], rut_kn[i]
                                if r2 != r1:
                                    depth_cross = d1 + (mp_hammer_total_weight_kn - r1) * (d2 - d1) / (r2 - r1)
                                else:
                                    depth_cross = current_depth
                            else:
                                depth_cross = current_depth
                            pile_run_risk_top[bound_key] = f'{depth_cross:.2f}'
                            found_risk = True
                            break
                if not found_risk:
                    pile_run_risk_top[bound_key] = 'No risk'
    # --- Build dynamic table columns and values ---
    table_headers = []
    table_columns = []
    for method in selected_methods:
        for bound in selected_bounds:
            table_headers.append(f'<b>{method} {bound.upper()}</b>')
            table_columns.append([intersection_depths.get((method, bound), '')] + ['']*4)
    # Table rows
    table_row_labels = ['<b>SWP MP + ILT</b>', '<b>Pile run at hammer placement</b>', '<b>SWP MP + Hammer</b>', '<b>Pile run risk top</b>', '<b>Pile run risk bottom</b>']
    # --- Create table ---
    weight_table = go.Table(
        header=dict(
            values=['<b></b>', '<b>LB</b>', '<b>BE</b>', '<b>UB</b>'],
            fill_color='lightgray',
            align='center',
            font=dict(size=13, color='black', family='Arial')
        ),
        cells=dict(
            values=[
                ['<b>SWP MP + ILT</b>', '<b>Pile run at hammer placement</b>', '<b>SWP MP + Hammer</b>', '<b>Pile run risk top</b>', '<b>Pile run risk bottom</b>'],
                [swp_mp_ilt_depths['LB'], pile_run_at_hammer_placement['LB'], swp_mp_hammer_depths['LB'], pile_run_risk_top['LB'], ''],  # LB column
                [swp_mp_ilt_depths['BE'], pile_run_at_hammer_placement['BE'], swp_mp_hammer_depths['BE'], pile_run_risk_top['BE'], ''],  # BE column
                [swp_mp_ilt_depths['UB'], pile_run_at_hammer_placement['UB'], swp_mp_hammer_depths['UB'], pile_run_risk_top['UB'], '']   # UB column
            ],
            fill_color='white',
            align=['left', 'center', 'center', 'center'],
            font=dict(size=12, color='black', family='Arial'),
            height=30
        )
    )

    fig.add_trace(weight_table, row=2, col=4)
    fig.update_xaxes(visible=False, row=1, col=4)
    fig.update_yaxes(visible=False, row=1, col=4)
    fig.update_xaxes(visible=False, row=2, col=4)
    fig.update_yaxes(visible=False, row=2, col=4)
    # ===============================================================================

    # --- Info panel table in row 1, col 4 ---
    info_table = go.Table(
        header=dict(
            values=['<b>Info</b>', '<b>Value</b>'],
            fill_color='lightgray',
            align='center',
            font=dict(size=13, color='black', family='Arial')
        ),
        cells=dict(
            values=[
                ['Position', 'Monopile Weight [t]', 'Hammer', 'Hammer Weight [t]', 'Target Blowcount Rate [bl/25cm]'],
                [
                    position,
                    f"{mp_weight:.1f}" if mp_weight is not None else '',
                    position_info.get('hammer_name', ''),
                    f"{position_info.get('hammer_weight', '')}",
                    f"{position_info.get('target_blowcount_rate', '')}"
                ]
            ],
            fill_color='white',
            align=['left', 'center'],
            font=dict(size=12, color='black', family='Arial'),
            height=30
        )
    )
    fig.add_trace(info_table, row=1, col=4)

    # Update layout for all subplots
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
            yanchor="bottom",
            y=0.02,
            xanchor="right",
            x=0.98,
            bgcolor="rgba(255, 255, 255, 0.9)",
            bordercolor="Black",
            borderwidth=1
        ),
        width=1587,   # A3 landscape width (420mm) at 96 DPI
        height=1123,  # A3 landscape height (297mm) at 96 DPI
        font=dict(size=11),  # Slightly smaller font for better fit on A3
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
        gridcolor='lightgray',
        showgrid=True,
        dtick=5,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=1
    )

    # Update layout for row 2, col 2 (Cumulative Blows plot)
    fig.update_xaxes(
        title='Cumulative Blows',
        gridcolor='lightgray',
        showgrid=True,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=2
    )
    fig.update_yaxes(
        title='Depth [mbsb]',
        autorange='reversed',
        rangemode='tozero',
        gridcolor='lightgray',
        showgrid=True,
        dtick=5,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=2
    )

    # Update layout for row 2, col 3 (SRD in kN with log scale)
    fig.update_xaxes(
        title='SRD [kN]',
        type='log',
        range=[4, 5.176],  # log10(10000) = 4, log10(150000) ≈ 5.176
        tickmode='array',
        tickvals=[10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000, 100000, 150000],
        ticktext=['10k', '', '', '', '', '', '', '', '', '100k', '150k'],
        gridcolor='lightgray',
        showgrid=True,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=3
    )
    fig.update_yaxes(
        title='Depth [mbsb]',
        autorange='reversed',
        rangemode='tozero',
        gridcolor='lightgray',
        showgrid=True,
        dtick=5,
        linecolor='#5a5a5a',
        linewidth=2,
        mirror=True,
        showticklabels=True,
        row=2, col=3
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
    root_dir = MONOPILE_ROOT_DIR

    # Fallback to local directory if network path is unavailable
    if not root_dir.exists():
        print(f"Warning: Network directory not found: {root_dir}")
        print("Using local HOW03 directory instead...")
        root_dir = Path(__file__).parent

    # Get position folders
    positions = get_position_folders(root_dir)

    if not positions:
        print("No position folders found!")
        return

    print(f"Found {len(positions)} positions: {[p[0] for p in positions]}")

    # Ask for position selection FIRST (before parsing any files)
    print("\n" + "="*60)
    print("SELECT POSITION(S) TO PLOT")
    print("="*60)
    print(f"\nAvailable positions: {', '.join([p[0] for p in positions])}")
    selected_positions_input = input("Enter positions to plot (comma-separated, e.g., A01,A02 or 'all'): ").strip()

    if selected_positions_input.lower() == 'all':
        selected_positions = [p[0] for p in positions]
    else:
        selected_positions = [p.strip().upper() for p in selected_positions_input.split(',') if p.strip()]

    # Now only parse CSV files for SELECTED positions (much faster!)
    print(f"\nParsing data for {len(selected_positions)} selected position(s)...")
    all_methods = set()
    all_bounds = set()
    position_tables = {}
    for p_name, p_path in positions:
        if p_name not in selected_positions:
            continue  # Skip positions not selected
        csv_file = p_path / f"results_PileDrivingAnalysis-{p_name}.csv"
        if csv_file.exists():
            print(f"  Reading {p_name}...")
            tables = parse_results_csv(csv_file)
            position_tables[p_name] = tables
            methods, bounds = get_available_methods_and_bounds(tables)
            all_methods.update(methods)
            all_bounds.update(bounds)

    methods = sorted(list(all_methods))
    bounds = sorted(list(all_bounds))


    # User selection for methods and bounds
    print("\n" + "="*60)
    print("SELECT SRD METHODS AND SOIL BOUNDS")
    print("="*60)
    print(f"\nAvailable SRD methods: {', '.join(methods)}")
    selected_methods_input = input("Enter methods to plot (comma-separated, e.g., MD,AH or 'all'): ").strip()

    if selected_methods_input.lower() == 'all':
        selected_methods = methods
    else:
        selected_methods = [m.strip() for m in selected_methods_input.split(',') if m.strip()]

    print(f"Selected methods: {selected_methods}")

    print(f"\nAvailable soil bounds: {', '.join(bounds)}")
    selected_bounds_input = input("Enter bounds to plot (comma-separated, e.g., lb,ub or 'all'): ").strip()

    if selected_bounds_input.lower() == 'all':
        selected_bounds = bounds
    else:
        selected_bounds = [b.strip().lower() for b in selected_bounds_input.split(',') if b.strip()]

    print(f"Selected bounds: {selected_bounds}")

    # --- MAIN LOOP: Process each selected position ---
    for position in selected_positions:
        print(f"\nProcessing position: {position}")

        # Get corresponding tables
        tables = position_tables.get(position, {})

        # Plot Rut vs Depth
        plot_driveability_results(
            tables=tables,
            position=position,
            selected_methods=selected_methods,
            selected_bounds=selected_bounds,
            output_dir=PLOTS_OUTPUT_DIR,
            monopile_weights=get_monopile_weights(MONOPILE_WEIGHTS_FILE, [position]),
            position_info=get_position_info(tables, position)
        )


if __name__ == "__main__":
    main()
