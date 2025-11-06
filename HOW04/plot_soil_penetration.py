from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

path_fou_type = Path(r'k:\dozr\HOW04\WTG_MP\07_GEO\02_Driveability\20250130 L286\WorkingFolder\L286_FoundationTypeWithMonopileRisk - BE.xlsx')
path = r"K:\dozr\HOW04\WTG_MP\05_PS\20250514 BAFO DELIVERABLES\02 All positions\01 Without PICASO imitation\summary\data_summary\summary-root.xls"

sheets = pd.read_excel(path, sheet_name=["Summary ..", "GEO stratigraphy"])
fou_type_data = pd.read_excel(path_fou_type)

#vasav comment
# manipulate Summary ..
sheet = sheets["Summary .."]
headers1 = sheet.iloc[5,:].fillna('').values
headers2 = sheet.iloc[6,:].values
headers = [f'{h2}' if h1 == '' else f'{h1}: {h2}' for h1, h2 in zip(headers1, headers2)]
data_end = sheet.iloc[7:, 0].isna().idxmax()
df_summary = sheet.iloc[8:data_end,:]
df_summary.columns = headers
df_summary = df_summary.reset_index(drop=True)
df_summary['Fou type'] = fou_type_data['Foundation Type']

# Manipulate the data to a df
sheet = sheets['GEO stratigraphy']
sheet.columns = sheet.iloc[0]
sheet = sheet.drop(index=[0, 1]).reset_index(drop=True)
data_end = sheet.iloc[7:, 0].isna().idxmax()

df_strat = sheet.iloc[:data_end, :]
for col in df_strat.columns:
    try:
        print(col)
        df_strat[col] = pd.to_numeric(df_strat[col])
    except:
        pass

# Add a column called TCL equal to the column TCL-1-C
df_summary['TCL penetration [m]'] = df_strat['TCL-1-C'] + df_strat['TCL-2-C'] + df_strat['TCL-2-Mudstone']
df_summary['Speeton penetration [m]'] = df_strat['SPE-C']

df_MP = df_summary[df_summary['Fou type'] == 'MP']

custom_color_scale = [
    (0.0, 'white'),  # 0% mapped to white
    (0.01, '#63BE7B'),  # Slightly above 0% starts the color scale
    (0.5, '#FFEB84'),  # 50% mapped to yellow
    (1.0, '#F8696B'),   # 100% mapped to green
]

fig = px.scatter(df_MP, x='Pos: E_WGS84', y='Pos: N_WGS84', color='TCL penetration [m]', color_continuous_scale=custom_color_scale)
fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)))
fig.update_layout(template="plotly_white")
fig.write_html(r"Map TCL penetration.html", include_plotlyjs="cdn")

fig = px.scatter(df_MP, x='Pos: E_WGS84', y='Pos: N_WGS84', color='Speeton penetration [m]', color_continuous_scale=custom_color_scale)
fig.update_traces(marker=dict(size=10, line=dict(color='grey', width=0.5)))
fig.update_layout(template="plotly_white")
fig.write_html(r"Map Speeton penetration.html", include_plotlyjs="cdn")

