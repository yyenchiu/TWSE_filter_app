import numpy as np
import pandas as pd

import dash
from dash import dash_table
from dash import dcc
from dash import html


app = dash.Dash(__name__)
server = app.server

twse = pd.read_csv("上市一月.csv", header=0)
tpex = pd.read_csv("上櫃一月.csv", header=0)
df = pd.concat([twse, tpex], axis=0).set_axis(["代碼", "名稱", "產業", "月營收", 
                                  "年累計營收", "%MoM", "%YoY", "%YoY累計", "盈利率", "BPS", "P/B", 
                                  "EPS(12個月)", "EPS預測", "P/E", "P/E預測"], axis=1)

app.layout = html.Div([
    html.H1("上市櫃排行（2024年一月）", style={"text-align": "center"}),
    dash_table.DataTable(df.to_dict("records"),
                         [{"name": i, "id": i} for i in df.columns],
                         filter_action="native",
                         filter_options={"placeholder_text": "Filter column..."},
                        )        
])


if __name__ == '__main__':
    app.run_server(debug=False)
