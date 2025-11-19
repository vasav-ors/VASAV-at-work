import pandas as pd
from pathlib import Path
import numpy as np
import plotly.express as px
import plotly.graph_objects as go



path_xl = Path(r"k:\dozr\HOW04\WTG_Jacket\07_GEO\12_Results discussion\MP_SBJ_47_CPT_BAFO\MP_ SBJ_47_CPT_BAFO_risk.xlsx")

df = pd.read_excel(path_xl, sheet_name="Sheet1")



# Combine row 5, 6, 7 to headers
df.columns = df.iloc[5:9].astype(str).agg(' '.join).str.replace('nan ', '', regex=False)

#only take data from row 8 and down
df = df.iloc[9:].reset_index(drop=True)

# turn numeric columns to numeric
for col in df.columns:
    try:
        df[col] = pd.to_numeric(df[col])
    except:
        pass

df['CPT coverage %'] = df['CPT coverage %'] * 100
df['CPT coverage SBJ %'] = df['CPT coverage SBJ %'] * 100

#df_MP = df[df['Foundation Type -'] == "MP"]
df_MP_SBJ = df[df['Foundation Type -'].isin(["MP", "SBJ"])]


# Create a dictionary to track occurrences of each column name
column_count = {}

# Create a list to store new column names
new_columns = []

# Iterate over each column name
for col in df_MP_SBJ.columns:
    if col in column_count:
        # Increment the count for duplicate columns
        column_count[col] += 1
        # Append the new name with a unique suffix
        new_columns.append(f"{col}_{column_count[col]}")
    else:
        # Initialize count for new columns
        column_count[col] = 1
        # Append the original name
        new_columns.append(col)

# Assign the new column names to the DataFrame
df_MP_SBJ.columns = new_columns

# Verify the changes
print("Columns after renaming:", df_MP_SBJ.columns)



# Plot CPT coverage
custom_color_scale = [
    (0.0, 'white'),  # 0% mapped to grey
    (0.01, '#F8696B'),  # Slightly above 0% starts the color scale
    (0.5, '#FFEB84'),  # 50% mapped to yellow
    (0.99, '#63BE7B'),   # 100% mapped to green
    (1.0, 'green')   # 100% mapped to green
]

######
fig = px.scatter(df_MP_SBJ, x='Easting', y='Northing', color='CPT coverage %', range_color=[0, 100],
                 color_continuous_scale=custom_color_scale,
                 text=df_MP_SBJ.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1))  # Corrected line
fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)))
fig.update_layout(template="plotly_white", plot_bgcolor='lightgrey')
fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)),
                        textposition='top center',
                        textfont=dict(size=8))  # Adjusted text position
fig.write_html('CPT coverage MP.html', include_plotlyjs='cdn')

#######


# Fill missing values in 'CPT depth m' with 0
df_MP_SBJ['CPT depth m'] = df_MP_SBJ['CPT depth m'].fillna(0)
# Normalize 'CPT depth m' for color scaling
#df_CPT_interpr['CPT depth m normalized'] = df_CPT_interpr['CPT depth m'] / df_CPT_interpr['CPT depth m'].max()

# Plot CPT depth with conditional annotations and adjusted text position
fig_depth = px.scatter(df_MP_SBJ, x='Easting', y='Northing', color='CPT depth m',
                       color_continuous_scale=custom_color_scale, range_color=[0, df_MP_SBJ['CPT depth m'].max()],
                       #text=df_MP.apply(lambda row: f"{row['CPT depth m']}m" if row['CPT depth m'] > 0 else "", axis=1))
                       text=df_MP_SBJ.apply(lambda row: f"{row['Position [-]'][-3:]}: {row['CPT depth m']}m" if row['CPT depth m'] > 0 else "", axis=1))

fig_depth.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)),
                        textposition='top center',
                        textfont=dict(size=8))  # Adjusted text position
fig_depth.update_layout(template="plotly_white", plot_bgcolor='lightgrey')
fig_depth.write_html('CPT_depth.html', include_plotlyjs='cdn')

########

fig = px.scatter(df_MP_SBJ, x='Easting', y='Northing', color='CPT coverage SBJ %', range_color=[50, 200],
                 color_continuous_scale=custom_color_scale,
                 text=df_MP_SBJ.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1))  # Corrected line
fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)))
fig.update_layout(template="plotly_white", plot_bgcolor='lightgrey')
fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)),
                        textposition='top center',
                        textfont=dict(size=8))  # Adjusted text position
fig.write_html('CPT coverage SBJ.html', include_plotlyjs='cdn')

########
# Define a color mapping for risk levels for SBJ
color_map = {
    "High risk": "red",
    "Medium risk": "orange",
    "Low risk": "green"
}

# Map risk levels to colors, leaving NaN values unmapped
df_MP_SBJ['Risk Color'] = df_MP_SBJ['Risk [-]'].map(color_map)

# Create the scatter plot
fig_risk = go.Figure()

# Add scatter trace for each risk level
for risk_level, color in color_map.items():
    df_risk_level = df_MP_SBJ[df_MP_SBJ['Risk [-]'] == risk_level]
    fig_risk.add_trace(go.Scatter(
        x=df_risk_level['Easting'],
        y=df_risk_level['Northing'],
        mode='markers+text',
        marker=dict(
            color=color,
            size=10,
            line=dict(color='grey', width=0.5)
        ),
        text=df_risk_level.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1),
        textposition='top center',
        hoverinfo='text',
        name=risk_level  # Add legend entry
    ))

# Add scatter trace for NaN values (transparent)
df_nan = df_MP_SBJ[df_MP_SBJ['Risk [-]'].isna()]
fig_risk.add_trace(go.Scatter(
    x=df_nan['Easting'],
    y=df_nan['Northing'],
    mode='markers',
    marker=dict(
        color='rgba(0,0,0,0)',  # Transparent color for NaN
        size=10,
        line=dict(color='grey', width=0.5)
    ),
    text=df_nan.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1),
    textposition='top center',
    hoverinfo='text',
    showlegend=False  # No legend entry for NaN
))

# Update layout
fig_risk.update_layout(
    xaxis_title="Easting",
    yaxis_title="Northing",
    template="plotly_white",
    plot_bgcolor='lightgrey',
    legend_title_text="SBJ Installation risk"
)

# Save the plot as an HTML file
fig_risk.write_html('Risk_Levels_SBJ_Installation.html', include_plotlyjs='cdn')


########
# Define a color mapping for risk levels for MP design
color_map = {
    "High risk": "red",
    "Medium risk": "orange",
    "Low risk": "green"
}

# Map risk levels to colors, leaving NaN values unmapped
df_MP_SBJ['Risk Color'] = df_MP_SBJ['Design Risk MP [-]'].map(color_map)

# Create the scatter plot
fig_risk = go.Figure()

# Add scatter trace for each risk level
for risk_level, color in color_map.items():
    df_risk_level = df_MP_SBJ[df_MP_SBJ['Design Risk MP [-]'] == risk_level]
    fig_risk.add_trace(go.Scatter(
        x=df_risk_level['Easting'],
        y=df_risk_level['Northing'],
        mode='markers+text',
        marker=dict(
            color=color,
            size=10,
            line=dict(color='grey', width=0.5)
        ),
        text=df_risk_level.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1),
        textposition='top center',
        hoverinfo='text',
        name=risk_level  # Add legend entry
    ))

# Add scatter trace for NaN values (transparent)
df_nan = df_MP_SBJ[df_MP_SBJ['Design Risk MP [-]'].isna()]
fig_risk.add_trace(go.Scatter(
    x=df_nan['Easting'],
    y=df_nan['Northing'],
    mode='markers',
    marker=dict(
        color='rgba(0,0,0,0)',  # Transparent color for NaN
        size=10,
        line=dict(color='grey', width=0.5)
    ),
    text=df_nan.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1),
    textposition='top center',
    hoverinfo='text',
    showlegend=False  # No legend entry for NaN
))

# Update layout
fig_risk.update_layout(
    xaxis_title="Easting",
    yaxis_title="Northing",
    template="plotly_white",
    plot_bgcolor='lightgrey',
    legend_title_text="MP Design risk"
)

# Save the plot as an HTML file
fig_risk.write_html('Risk_Levels_MP_Design.html', include_plotlyjs='cdn')


########
# Define a color mapping for risk levels for MP driveability
color_map = {
    "High risk": "red",
    "Medium risk": "orange",
    "Low risk": "green"
}

# Map risk levels to colors, leaving NaN values unmapped
df_MP_SBJ['Risk Color'] = df_MP_SBJ['Driveability Risk MP [-]'].map(color_map)

# Create the scatter plot
fig_risk = go.Figure()

# Add scatter trace for each risk level
for risk_level, color in color_map.items():
    df_risk_level = df_MP_SBJ[df_MP_SBJ['Driveability Risk MP [-]'] == risk_level]
    fig_risk.add_trace(go.Scatter(
        x=df_risk_level['Easting'],
        y=df_risk_level['Northing'],
        mode='markers+text',
        marker=dict(
            color=color,
            size=10,
            line=dict(color='grey', width=0.5)
        ),
        text=df_risk_level.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1),
        textposition='top center',
        hoverinfo='text',
        name=risk_level  # Add legend entry
    ))

# Add scatter trace for NaN values (transparent)
df_nan = df_MP_SBJ[df_MP_SBJ['Driveability Risk MP [-]'].isna()]
fig_risk.add_trace(go.Scatter(
    x=df_nan['Easting'],
    y=df_nan['Northing'],
    mode='markers',
    marker=dict(
        color='rgba(0,0,0,0)',  # Transparent color for NaN
        size=10,
        line=dict(color='grey', width=0.5)
    ),
    text=df_nan.apply(lambda row: f"{row['Position [-]'][-3:]}", axis=1),
    textposition='top center',
    hoverinfo='text',
    showlegend=False  # No legend entry for NaN
))

# Update layout
fig_risk.update_layout(
    xaxis_title="Easting",
    yaxis_title="Northing",
    template="plotly_white",
    plot_bgcolor='lightgrey',
    legend_title_text="MP Driveability risk"
)

# Save the plot as an HTML file
fig_risk.write_html('Risk_Levels_MP_Driveability.html', include_plotlyjs='cdn')
# # Plot refusals
# df_CPT = df_MP[df_MP['CPT coverage %'] > 0]
# colors = {True: 'red', False: 'green'}
# lgnd = {True: 'Refusal', False: 'No refusal'}
#
# # Global parameters
# fig = go.Figure()
# for refusal, df_refusal in df_MP.groupby('Reference Driv Refusal'):
#     fig.add_trace(go.Scatter(x=df_refusal['Easting'], y=df_refusal['Northing'], mode='markers', marker=dict(symbol='circle', color=colors[refusal]), showlegend=False))
#
# fig.add_trace(go.Scatter(x=df_MP['Easting'], y=df_MP['Northing'], mode='markers', marker=dict(symbol='circle', color='rgba(255, 255, 255, 0.8)'), showlegend=False))
#
# for refusal, df_refusal in df_CPT.groupby('Reference Driv Refusal'):
#     fig.add_trace(go.Scatter(x=df_refusal['Easting'], y=df_refusal['Northing'], mode='markers', marker=dict(symbol='circle', color=colors[refusal]), name=lgnd[refusal]))
#
# fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)))
# fig.update_layout(template="plotly_white")
# fig.write_html('Refusal no CPT.html', include_plotlyjs='cdn')
#
# # CPTs
# fig = go.Figure()
# for refusal, df_refusal in df_MP.groupby('CPT Results Driv Refusal'):
#     fig.add_trace(go.Scatter(x=df_refusal['Easting'], y=df_refusal['Northing'], mode='markers', marker=dict(symbol='circle', color=colors[refusal]), showlegend=False))
#
# fig.add_trace(go.Scatter(x=df_MP['Easting'], y=df_MP['Northing'], mode='markers', marker=dict(symbol='circle', color='rgba(255, 255, 255, 0.8)'), showlegend=False))
#
# for refusal, df_refusal in df_CPT.groupby('CPT Results Driv Refusal'):
#     fig.add_trace(go.Scatter(x=df_refusal['Easting'], y=df_refusal['Northing'], mode='markers', marker=dict(symbol='circle', color=colors[refusal]), name=lgnd[refusal]))
#
# fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)))
# fig.update_layout(template="plotly_white")
# fig.write_html('Refusal w CPT.html', include_plotlyjs='cdn')
#
# # Plot bars
# df_CPT_refusal = df_CPT[(df_CPT['Reference Driv Refusal'] == True)|(df_CPT['CPT Results Driv Refusal'] == True)]
#
# fig = go.Figure()
# fig.add_trace(go.Bar(x=df_CPT_refusal["Position [-]"], y=-df_CPT_refusal["Reference Driv targetdepth m"], marker_color='steelblue', marker_pattern_shape='/', offsetgroup='Global', showlegend=False))
# fig.add_trace(go.Bar(x=df_CPT_refusal["Position [-]"], y=-df_CPT_refusal["Reference Driv Maxdepth m"], marker_color='steelblue', offsetgroup='Global', name='No CPT'))
#
# fig.add_trace(go.Bar(x=df_CPT_refusal["Position [-]"], y=-df_CPT_refusal["CPT Results Driv targetdepth m"], marker_color='skyblue', marker_pattern_shape='/', offsetgroup='CPT', showlegend=False))
# fig.add_trace(go.Bar(x=df_CPT_refusal["Position [-]"], y=-df_CPT_refusal["CPT Results Driv Maxdepth m"], marker_color='skyblue', offsetgroup='CPT', name='CPT'))
#
# fig.add_trace(go.Bar(x=[np.nan], y=[np.nan], marker_color='white', marker_pattern_shape='/', name='Refusal'))
# fig.update_yaxes(title_text="Depth [m]")
#
# fig.update_layout(template="plotly_white")
# fig.write_html('Refusal bars.html', include_plotlyjs='cdn')