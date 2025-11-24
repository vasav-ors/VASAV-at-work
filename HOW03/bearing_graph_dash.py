"""
Interactive bearing graph using Dash - TRULY STATEFUL filtering.

This uses Dash callbacks which maintain state across all dropdown selections.
When you select A01 and then select 'be', it shows ONLY A01 + be (not all positions).

Usage:
    python HOW03/bearing_graph_dash.py --root "K:\\...\\monopiles"

Then open http://127.0.0.1:8050 in your browser.
"""
import argparse
import re
from pathlib import Path
import csv
import pandas as pd
import dash
from dash import dcc, html, Output, Input, callback
import plotly.graph_objects as go


def find_result_files(root: Path):
    """Return list of result csv files under root in subfolders like A01, A02..."""
    files = []
    for entry in root.iterdir():
        if entry.is_dir():
            for f in entry.glob('results_PileDrivingAnalysis-*.csv'):
                files.append((entry.name, f))
    return sorted(files)


def parse_results_csv(path: Path, site_name: str):
    """Parse results CSV with semicolon separators."""
    tables = []
    text = path.read_text(encoding='utf-8', errors='replace')
    lines = text.splitlines()
    n = len(lines)
    i = 0
    while i < n:
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith('**'):
            analysis_name = line.lstrip('*').strip().rstrip(';')
            j = i + 1
            while j < n and not lines[j].strip():
                j += 1
            if j >= n:
                break
            position = lines[j].strip()
            j += 1
            while j < n and not lines[j].strip():
                j += 1
            if j >= n:
                break
            header_line = lines[j].strip()
            j += 1
            units_line = ''
            if j < n:
                units_line = lines[j].strip()
            j += 1
            data_rows = []
            while j < n:
                cur = lines[j]
                if cur.strip() == '':
                    break
                if cur.strip().startswith('**'):
                    break
                data_rows.append(cur)
                j += 1

            try:
                header = next(csv.reader([header_line], delimiter=';'))
            except Exception:
                header = [header_line]
            header = [h.strip() for h in header]

            try:
                units = next(csv.reader([units_line], delimiter=';')) if units_line else []
                units = [u.strip() for u in units]
            except Exception:
                units = []

            parsed_rows = []
            for r in data_rows:
                try:
                    vals = next(csv.reader([r], delimiter=';'))
                except Exception:
                    vals = [r]
                if len(vals) < len(header):
                    vals = vals + [''] * (len(header) - len(vals))
                elif len(vals) > len(header):
                    vals = vals[:len(header)]
                parsed_rows.append(dict(zip(header, vals)))

            if not parsed_rows:
                i = j
                continue

            df_table = pd.DataFrame(parsed_rows)
            depth_col = blow_col = rut_col = None
            for c in df_table.columns:
                low = str(c).lower()
                if low.strip() == 'depth' or low.startswith('depth'):
                    depth_col = c
                if 'blow' in low:
                    blow_col = c
                if low.strip() == 'rut' or low.startswith('rut'):
                    rut_col = c
            if depth_col is None:
                depth_col = df_table.columns[0]
            if blow_col is None:
                for c in df_table.columns:
                    if 'blowcount' in str(c).lower():
                        blow_col = c
            if rut_col is None:
                for c in df_table.columns:
                    if 'rut' in str(c).lower():
                        rut_col = c

            df_small = pd.DataFrame()
            try:
                df_small['depth'] = pd.to_numeric(df_table[depth_col].astype(str).str.replace(',', '.'), errors='coerce')
            except Exception:
                df_small['depth'] = pd.to_numeric(df_table.iloc[:, 0].astype(str).str.replace(',', '.'), errors='coerce')
            if blow_col is not None:
                df_small['blowcount'] = pd.to_numeric(df_table[blow_col].astype(str).str.replace(',', '.'), errors='coerce')
            else:
                df_small['blowcount'] = pd.NA
            if rut_col is not None:
                df_small['Rut'] = pd.to_numeric(df_table[rut_col].astype(str).str.replace(',', '.'), errors='coerce')
            else:
                df_small['Rut'] = pd.NA

            df_small['site'] = site_name
            df_small['position'] = position
            df_small['analysis_name'] = analysis_name

            df_small = df_small.dropna(subset=['depth'])
            if not df_small.empty:
                tables.append(df_small)

            i = j
        else:
            i += 1
    if not tables:
        return pd.DataFrame()
    return pd.concat(tables, ignore_index=True, sort=False)


def tidy_dataframe(df: pd.DataFrame):
    """Tidy the dataframe."""
    df['depth'] = pd.to_numeric(df['depth'], errors='coerce')
    df['blowcount'] = pd.to_numeric(df['blowcount'], errors='coerce')
    df['Rut'] = pd.to_numeric(df['Rut'], errors='coerce')

    def parse_analysis(name: str):
        name = str(name).lower()
        soil = None
        method = None
        if '_lb_' in name:
            soil = 'lb'
        elif '_be_' in name:
            soil = 'be'
        elif '_ub_' in name:
            soil = 'ub'
        if '_md_' in name or '-md-' in name:
            method = 'MD'
        elif '_ah_' in name:
            method = 'AH'
        elif 'maynard' in name or '_my_' in name:
            method = 'MY'
        elif 'mono' in name or 'monodrive' in name:
            method = 'MD'
        return soil, method

    parsed = df['analysis_name'].apply(parse_analysis)
    df['soil'] = parsed.apply(lambda x: x[0])
    df['method'] = parsed.apply(lambda x: x[1])
    df = df.dropna(subset=['Rut', 'blowcount'])
    return df


def make_pale_color(color_name):
    """Convert color name to pale version."""
    pale_map = {
        'blue': 'rgba(100, 150, 255, 0.7)',
        'red': 'rgba(255, 100, 100, 0.7)',
        'green': 'rgba(100, 200, 100, 0.7)',
    }
    return pale_map.get(color_name, 'rgba(150, 150, 150, 0.7)')


def load_and_prepare_data(root_path):
    """Load all data and prepare it for the app."""
    files = find_result_files(root_path)
    parsed_all = []
    for site, path in files:
        try:
            df = parse_results_csv(path, site)
            if not df.empty:
                parsed_all.append(df)
        except Exception:
            pass

    if not parsed_all:
        return None

    full = pd.concat(parsed_all, ignore_index=True, sort=False)
    tidy = tidy_dataframe(full)
    return tidy


# Global data storage
app_data = {
    'df': None,
    'positions': [],
    'soils': [],
    'methods': []
}


def run_dash_app(root: Path, host='127.0.0.1', port=8050):
    """Run the Dash app."""
    global app_data

    print('Loading data...')
    df = load_and_prepare_data(root)
    if df is None or df.empty:
        print('No data loaded!')
        return

    app_data['df'] = df
    app_data['positions'] = sorted(df['position'].unique())
    app_data['soils'] = sorted(df['soil'].dropna().unique())
    app_data['methods'] = sorted(df['method'].dropna().unique())

    soil_styles = {'lb': 'solid', 'be': 'dot', 'ub': 'dash'}
    method_colors = {'MD': 'blue', 'AH': 'red', 'MY': 'green'}

    app = dash.Dash(__name__)

    app.layout = html.Div([
        html.H2('Bearing Graph: Interactive Analysis Tool'),
        html.Div([
            html.Div([
                html.Label('Position (MP):', style={'fontWeight': 'bold'}),
                dcc.Dropdown(
                    id='position-dropdown',
                    options=[{'label': 'All Positions', 'value': 'all'}] +
                            [{'label': pos, 'value': pos} for pos in app_data['positions']],
                    value='all',
                    style={'width': '100%'}
                )
            ], style={'width': '30%', 'display': 'inline-block', 'marginRight': '3%'}),

            html.Div([
                html.Label('Soil Bounds:', style={'fontWeight': 'bold'}),
                dcc.Dropdown(
                    id='soil-dropdown',
                    options=[{'label': 'All', 'value': 'all'}] +
                            [{'label': soil, 'value': soil} for soil in app_data['soils']] +
                            [{'label': f"{app_data['soils'][i]}+{app_data['soils'][j]}",
                              'value': f"{app_data['soils'][i]}+{app_data['soils'][j]}"}
                             for i in range(len(app_data['soils']))
                             for j in range(i+1, len(app_data['soils']))],
                    value='all',
                    style={'width': '100%'}
                )
            ], style={'width': '30%', 'display': 'inline-block', 'marginRight': '3%'}),

            html.Div([
                html.Label('Method:', style={'fontWeight': 'bold'}),
                dcc.Dropdown(
                    id='method-dropdown',
                    options=[{'label': 'All', 'value': 'all'}] +
                            [{'label': method, 'value': method} for method in app_data['methods']] +
                            [{'label': f"{app_data['methods'][i]}+{app_data['methods'][j]}",
                              'value': f"{app_data['methods'][i]}+{app_data['methods'][j]}"}
                             for i in range(len(app_data['methods']))
                             for j in range(i+1, len(app_data['methods']))],
                    value='all',
                    style={'width': '100%'}
                )
            ], style={'width': '30%', 'display': 'inline-block'})
        ], style={'marginBottom': '20px', 'display': 'flex', 'gap': '10px'}),

        dcc.Graph(id='bearing-graph'),

        html.Div(id='info-text', style={'marginTop': '20px', 'color': 'gray', 'fontSize': '12px'})
    ], style={'padding': '20px', 'fontFamily': 'Arial, sans-serif'})

    @callback(
        Output('bearing-graph', 'figure'),
        Output('info-text', 'children'),
        Input('position-dropdown', 'value'),
        Input('soil-dropdown', 'value'),
        Input('method-dropdown', 'value')
    )
    def update_graph(selected_pos, selected_soil, selected_method):
        df = app_data['df']

        # Parse selections
        positions = app_data['positions'] if selected_pos == 'all' else [selected_pos]

        if selected_soil == 'all':
            soils = app_data['soils']
        else:
            soils = selected_soil.split('+')

        if selected_method == 'all':
            methods = app_data['methods']
        else:
            methods = selected_method.split('+')

        # Filter dataframe
        dff = df[(df['position'].isin(positions)) &
                 (df['soil'].isin(soils)) &
                 (df['method'].isin(methods))]

        fig = go.Figure()

        # Group by position, soil, method and create traces
        groups = dff.groupby(['position', 'soil', 'method'])
        trace_count = 0

        for (pos, soil, method), group in groups:
            gp = group.groupby('depth', as_index=False).agg({'blowcount': 'mean', 'Rut': 'mean'})
            if len(gp) == 0:
                continue

            x = gp['blowcount'] / 4.0
            y = gp['Rut']
            name = f"{pos} | {soil} | {method}"

            color = method_colors.get(method, 'black')
            dash = soil_styles.get(soil, 'solid')
            pale_color = make_pale_color(color)

            fig.add_trace(go.Scatter(
                x=x, y=y, mode='lines', name=name,
                line=dict(color=pale_color, dash=dash, width=2),
                hovertemplate='%{x:.2f} bl/25cm<br>SRD: %{y:.3f} MN<br>' + name + '<extra></extra>'
            ))
            trace_count += 1

        title_parts = []
        if selected_pos != 'all':
            title_parts.append(f"Position: {selected_pos}")
        if selected_soil != 'all':
            title_parts.append(f"Soil: {selected_soil}")
        if selected_method != 'all':
            title_parts.append(f"Method: {selected_method}")

        title = ' | '.join(title_parts) if title_parts else 'All Positions | All Soils | All Methods'

        fig.update_layout(
            title=f'Bearing Graph: blowcount/4 vs SRD ({title})',
            xaxis_title='blowcount; [bl/25cm]',
            yaxis_title='SRD; [MN]',
            height=700,
            hovermode='closest'
        )

        info_text = f'Displaying {trace_count} trace(s) | Total rows: {len(dff)}'

        return fig, info_text

    print(f'Starting Dash app on http://{host}:{port}')
    print('Press Ctrl+C to stop the server')
    app.run(host=host, port=port, debug=False)


def main():
    parser = argparse.ArgumentParser(description='Interactive bearing graph (Dash version)')
    parser.add_argument('--root', type=str, default=None, help='Root folder containing position subfolders')
    args = parser.parse_args()

    # If no --root argument provided, ask user interactively
    if args.root is None:
        print("\n" + "="*70)
        print("BEARING GRAPH - INTERACTIVE PILE DRIVING ANALYSIS TOOL")
        print("="*70)
        print("\nEnter the path to the monopiles folder:")
        print("Example: K:\\dozr\\HOW03\\GEO\\05_Driveability\\20240909_Final Design for Certification\\variations\\const_en_MENCK_original_cans\\monopiles")
        print()
        root_input = input("Enter folder path: ").strip()

        if not root_input:
            print("[ERROR] No path provided. Exiting.")
            return

        root = Path(root_input)
    else:
        root = Path(args.root)

    if not root.exists():
        print(f'[ERROR] Root not found: {root}')
        return

    print(f'\n[OK] Using folder: {root}')
    run_dash_app(root)


if __name__ == '__main__':
    main()

