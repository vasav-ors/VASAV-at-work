import pandas as pd
import plotly.graph_objects as go

# it reads an input soil profile file and plots the soil profiles
# targets primarily SBJ focusing only on sand or clay and first few meters of soil profile


def extract_tables_from_excel(file_path, position_row_values):
    """
    Extracts tables from an Excel file where tables are marked by '**' in the first column
    and filters them based on the table name and position_row values.

    Parameters:
    file_path (str): Path to the Excel file.
    position_row_values (pd.Series): Series of position row values to filter tables.

    Returns:
    list: A list of dictionaries, each containing metadata and data for a filtered table.
    """

    # Read the Excel file (single sheet)
    data = pd.read_excel(file_path, header=None)
    column_a = data[0]  # Assuming column A is the first column (index 0)

    # Find the start of tables
    start_indices = column_a[column_a.str.startswith('**', na=False)].index
    tables = []

    for start in start_indices:
        # Extract the four metadata rows
        table_name = data.iloc[start, 0]  # Table name
        position_row = data.iloc[start + 1, 0]  # Position labels
        header_row = data.iloc[start + 2, :]  # Headers

        # Determine the last column index for the current table
        if table_name == '**soil':
            last_col_index = len(header_row) + 1
            for i, value in enumerate(header_row):
                if pd.isna(value):
                    last_col_index = i
                    break
            header_row = data.iloc[start + 2, :last_col_index]
            units_row = data.iloc[start + 3, :last_col_index]  # Units

            # Check if the position_row matches the specified values
            if any(value in position_row for value in position_row_values):
                # Print statement to inform which table and position is being extracted
                print(f"Extracting table: {table_name}, Position: {position_row}")

                # Find the end of the table (next empty row after data rows)
                data_start = start + 4  # Data starts after the four metadata rows
                empty_row = data.iloc[data_start:, :].index[data.iloc[data_start:, 0].isnull()].min()
                data_end = empty_row if pd.notna(empty_row) else len(data)

                # Extract the table data
                table_data = data.iloc[data_start:data_end, :last_col_index].reset_index(drop=True)

                # Create a MultiIndex for the columns using header_row and units_row
                multi_index = pd.MultiIndex.from_arrays([header_row, units_row])
                table_data.columns = multi_index

                # Find the z_top column
                z_top_column = table_data.columns.get_level_values(0).tolist().index('z_top')
                z_top = table_data.iloc[:, z_top_column]

                # Calculate the thickness values
                thickness = -z_top.diff().fillna(0).shift(-1).fillna(0)

                # Add the new thickness column with header ('thickness', 'm')
                table_data[('thickness', 'm')] = thickness.values

                # Combine metadata and data into a dictionary
                tables.append({
                    'table_name': table_name,
                    'position': position_row,
                    'data': table_data
                })

        elif table_name == '**geoUnits':
            last_col_index = header_row.first_valid_index()
            for i, value in enumerate(header_row):
                if pd.isna(value):
                    last_col_index = i
                    break
            header_row = data.iloc[start + 2, :last_col_index]
            units_row = data.iloc[start + 3, :last_col_index]  # Units

            # Print statement to inform which table and position is being extracted
            print(f"Extracting table: {table_name}, Position: {position_row}")

            # Find the end of the table (next empty row after data rows)
            data_start = start + 4  # Data starts after the four metadata rows
            empty_row = data.iloc[data_start:, :].index[data.iloc[data_start:, 0].isnull()].min()
            data_end = empty_row if pd.notna(empty_row) else len(data)

            # Extract the table data
            table_data = data.iloc[data_start:data_end, :last_col_index].reset_index(drop=True)

            # Create a MultiIndex for the columns using header_row and units_row
            multi_index = pd.MultiIndex.from_arrays([header_row, units_row])
            table_data.columns = multi_index

            # Combine metadata and data into a dictionary
            tables.append({
                'table_name': table_name,
                'position': position_row,
                'data': table_data
            })

    return tables


def create_interactive_plot(tables, min_elevation=None):
    """
    Creates an interactive stacked column plot from the extracted tables.

    Parameters:
    tables (list): A list of dictionaries, each containing metadata and data for a filtered table.
    min_elevation (float, optional): The minimum elevation to display in the plot. Defaults to None.

    Returns:
    None: Displays the interactive plot.
    """


    try:
        fig = go.Figure()

        # Print statement to inform that plotting has started
        print(f"Plotting has started")

        # Extract the geoUnits table
        geo_units_table = next((table['data'] for table in tables if table['table_name'] == '**geoUnits'), None)
        if geo_units_table is None:
            raise ValueError("The geoUnits table is missing.")

        # Print the columns to debug
        #print("Columns in geo_units_table:", geo_units_table.columns)

        # Ensure the required columns are present
        required_columns = [('geoUnit', 'text'), ('type', 'text'), ('plot_R', '-'), ('plot_G', '-'), ('plot_B', '-')]
        for col in required_columns:
            if col not in geo_units_table.columns:
                raise ValueError(f"Required column {col} is missing in the geoUnits table.")

        # Debugging: Print the first few rows of the geo_units_table
        #print("First few rows of geo_units_table:\n", geo_units_table.head())

        # Create a color map based on the RGB values in the geoUnits table
        geo_units_table = geo_units_table.set_index(('geoUnit', 'text'))
        geo_units_types = geo_units_table[('type', 'text')].to_dict()
        type_color_map = {
            unit: f'rgba({geo_units_table.loc[unit, ("plot_R", "-")]}, '
                  f'{geo_units_table.loc[unit, ("plot_G", "-")]}, '
                  f'{geo_units_table.loc[unit, ("plot_B", "-")]}, 0.8)'
            for unit in geo_units_table.index
        }

        geoUnit_traces = {unit: [] for unit in geo_units_types.keys()}

        # Filter the tables to include only those with table_name '**soil'
        soil_tables = [table for table in tables if table['table_name'] == '**soil']

        for table in soil_tables:
            position = table['position']
            data = table['data']

            if ('z_top', 'm') not in data.columns or ('thickness', 'm') not in data.columns or ('geoUnit', 'text') not in data.columns:
                raise ValueError("Required columns 'z_top', 'thickness', or 'geoUnit' are missing in the data.")

            z_top = data[('z_top', 'm')]
            thickness = data[('thickness', 'm')]
            geoUnit = data[('geoUnit', 'text')]

            for unit in geoUnit.unique():
                if unit == "EOP":
                    continue  # Skip plotting for 'EOP'

                unit_mask = geoUnit == unit
                unit_type = geo_units_types.get(unit, 'default')
                color = type_color_map.get(unit, 'rgba(0, 0, 0, 0.8)')  # Default color if type not found

                fig.add_trace(go.Bar(
                    x=[position] * unit_mask.sum(),
                    y=thickness[unit_mask],
                    base=z_top[unit_mask] - thickness[unit_mask],
                    name=unit,
                    marker_color=color,
                    hoverinfo='x+y',
                    text=[unit] * unit_mask.sum(),
                    textposition='auto',
                    showlegend=True  # Ensure legend is shown for all geoUnits
                ))
                geoUnit_traces[unit].append(len(fig.data) - 1)

        # Ensure that clicking on the legend affects all traces of the same geoUnit
        for unit, trace_indices in geoUnit_traces.items():
            for trace_index in trace_indices:
                fig.data[trace_index].legendgroup = unit
                fig.data[trace_index].showlegend = (
                    trace_index == trace_indices[0])  # Show legend only for the first trace of each geoUnit

        # Update layout with min_elevation if provided
        layout_update = {
            'title': 'Interactive Stacked Column Plot of Thickness at Different Positions',
            'xaxis_title': 'Position',
            'yaxis_title': 'Elevation [m]',
            'barmode': 'stack',
            'hovermode': 'closest'
        }
        if min_elevation is not None:
            layout_update['yaxis'] = {'range': [min_elevation, 0]}

        fig.update_layout(**layout_update)

        fig.show()

    except Exception as e:
        print(f"An error occurred: {e}")


def simplify_profiles(tables, min_elevation, mixed_replacement=None):
    """
    Processes the extracted tables to introduce a new column 'type/text' in **soil tables
    and associates the type with each geoUnit.

    Parameters:
    tables (list): A list of dictionaries, each containing metadata and data for a filtered table.
    min_elevation (float): The minimum elevation to consider for simplifying profiles.
    mixed_replacement (str, optional): The type to replace 'MIXED' with. Defaults to None.


    Returns:
    list: A list of dictionaries with simplified_sprofiles.
    """
    try:

        # Print statement to inform that plotting has started
        print(f"Simplifying profiles has started")

        # Extract the geoUnits table
        geo_units_table = next((table['data'] for table in tables if table['table_name'] == '**geoUnits'), None)
        if geo_units_table is None:
            raise ValueError("The geoUnits table is missing.")

        # Create a dictionary to map geoUnit to type
        geo_unit_type_map = geo_units_table.set_index(('geoUnit', 'text'))[('type', 'text')].to_dict()

        simplified_profiles = []

        # Process each **soil table
        for table in tables:
            if table['table_name'] == '**soil':
                data = table['data']

                # Ensure the required column is present
                if ('geoUnit', 'text') not in data.columns:
                    raise ValueError("Required column 'geoUnit' is missing in the **soil table.")

                # Introduce the new column 'type/text'
                data[('type', 'text')] = data[('geoUnit', 'text')].map(geo_unit_type_map)

                # Optionally replace 'MIXED' types
                if mixed_replacement:
                    data.loc[data[('type', 'text')] == 'MIXED', ('type', 'text')] = mixed_replacement

                # Simplify profiles down to min_elevation
                simplified_profile = []
                for _, row in data.iterrows():
                    z_top = row[('z_top', 'm')]
                    thickness = row[('thickness', 'm')]
                    geo_type = row[('type', 'text')]

                    if z_top <= min_elevation:
                        break

                    if z_top - thickness < min_elevation:
                        thickness = z_top - min_elevation

                    simplified_profile.append({
                        'z_top': z_top,
                        'type': geo_type
                    })

                # Merge consecutive rows with the same type
                merged_profile = []
                current_type = None

                for layer in simplified_profile:
                    if layer['type'] != current_type:
                        merged_profile.append({
                            'z_top': layer['z_top'],
                            'type': layer['type']
                        })
                        current_type = layer['type']

                # Calculate thickness for each layer
                for i in range(len(merged_profile)):
                    if i < len(merged_profile) - 1:
                        merged_profile[i]['thickness'] = merged_profile[i]['z_top'] - merged_profile[i + 1]['z_top']
                    else:
                        merged_profile[i]['thickness'] = merged_profile[i]['z_top'] - min_elevation

                simplified_profiles.append({
                    'position': table['position'],
                    'profile': pd.DataFrame(merged_profile)
                })

        # Print the simplified profiles once at the end
        # for profile in simplified_profiles:
        #     print(f"Simplified profile for position {profile['position']}:")
        #     print(profile['profile'])

        # Create a list of dictionaries with position names and simplified profiles
        simplified_profiles_list = [{'position': profile['position'], 'profile': profile['profile']} for profile in simplified_profiles]

        return simplified_profiles_list

    except Exception as e:
        print(f"An error occurred: {e}")
        return []


def identify_unique_profiles(simplified_profiles_list):
    """
    Identifies unique types across all simplified profiles.

    Parameters:
    simplified_profiles_list (list): A list of dictionaries, each containing a position name and a simplified profile DataFrame.

    Returns:
    dict: A dictionary where keys are unique sequences of types and values are lists of positions that have each unique profile.
    """
    unique_profiles = {}
    total_positions = len(simplified_profiles_list)

    for profile in simplified_profiles_list:
        position = profile['position']
        profile_df = profile['profile']
        # Convert the sequence of types to a tuple
        type_sequence = tuple(profile_df['type'])

        if type_sequence not in unique_profiles:
            unique_profiles[type_sequence] = []
        unique_profiles[type_sequence].append(position)

    # Sort profiles by the number of positions in descending order
    sorted_profiles = sorted(unique_profiles.items(), key=lambda item: len(item[1]), reverse=True)

    # Print the unique profiles with positions and the number of positions
    for profile, positions in sorted_profiles:
        print(f"Profile: {profile}")
        print(f"Number of Positions: {len(positions)} out of {total_positions}")
        print(f"Positions: {', '.join(positions)}\n")

    return unique_profiles


# Example usage
file_path = "input_soilProfiles_BAFO_batch0.xlsx"  # Replace with your Excel file path
position_row_values = pd.Series([
    'L286_041' #'HOW04_002', 'HOW04_003', 'HOW04_004', 'HOW04_005', 'HOW04_006', 'HOW04_010', 'HOW04_011',
    #'HOW04_013', 'HOW04_014', 'HOW04_015', 'HOW04_016', 'HOW04_017', 'HOW04_022', 'HOW04_024',
    #'HOW04_028', 'HOW04_043', 'HOW04_045', 'HOW04_047', 'HOW04_050', 'HOW04_059', 'HOW04_060',
    #'HOW04_065', 'HOW04_067', 'HOW04_069', 'HOW04_072', 'HOW04_075', 'HOW04_077', 'HOW04_080',
    #'HOW04_081', 'HOW04_089', 'HOW04_093', 'HOW04_096', 'HOW04_099', 'HOW04_104', 'HOW04_105',
    #'HOW04_118', 'HOW04_119', 'HOW04_123', 'HOW04_125', 'HOW04_127', 'HOW04_135', 'HOW04_137',
    #'HOW04_146', 'HOW04_147', 'HOW04_158', 'HOW04_166', 'HOW04_168'
])  # Your specified position row values
min_elevation = -12
mixed_replacement = "SAND"

# Step 1: Extract tables from the Excel file
tables = extract_tables_from_excel(file_path, position_row_values)

# Step 2: Process the extracted tables to introduce the 'type/text' column in **soil tables
simplified_profiles_list = simplify_profiles(tables, min_elevation, mixed_replacement)

# Step 3: Create an interactive plot from the processed tables
create_interactive_plot(tables, min_elevation)

# Step 4: identifies the uniques sequences in the top xxm of profiles
unique_profiles = identify_unique_profiles(simplified_profiles_list)


