import numpy as np
import pandas as pd
import datetime

import yfinance as yf

import dash
from dash import dash_table
from dash import dcc
from dash import html
from dash import Input, Output, callback, State

import base64
import io


twse = pd.read_csv("TWSE_list.csv", header=0)
tpex = pd.read_csv("TPEX_list.csv", header=0)

app = dash.Dash(__name__)
server = app.server

if "daily_info" not in locals():
    twse_symbols = twse["Symbol"].to_list()
    tickers = [str(symbol)+".TW" for symbol in twse_symbols]
    combined = " ".join(tickers)
    twse = twse.set_index("Symbol")
    if "twse_get_info" not in locals():
        twse_get_info = yf.Tickers(combined)
    for symbol in twse_symbols:
        stock = twse_get_info.tickers[str(symbol)+".TW"].info
        for attribute in ["currentPrice", "grossMargins", "bookValue", "priceToBook", 
                          "trailingEps", "forwardEps", "trailingPE", "forwardPE"]:
            if attribute in stock.keys():
                twse.loc[symbol, attribute] = round(float(stock[attribute]), 2)
            else:
                twse.loc[symbol, attribute] = np.nan

    tpex_symbols = tpex["Symbol"].to_list()
    tickers = [str(symbol)+".TWO" for symbol in tpex_symbols]
    combined = " ".join(tickers)
    tpex = tpex.set_index("Symbol")
    if "tpex_get_info" not in locals():
        tpex_get_info = yf.Tickers(combined)
    for symbol in tpex_symbols:
        stock = tpex_get_info.tickers[str(symbol)+".TWO"].info
        for attribute in ["currentPrice", "grossMargins", "bookValue", "priceToBook", 
                          "trailingEps", "forwardEps","trailingPE", "forwardPE"]:
            if attribute in stock.keys():
                tpex.loc[symbol, attribute] = round(float(stock[attribute]), 2)
            else:
                tpex.loc[symbol, attribute] = np.nan

    daily_info = pd.concat([twse, tpex], axis=0)

app.layout = html.Div([
    html.H1("上市櫃排行", style={"text-align": "center"}),
    html.A("上市報表", href="https://www.twse.com.tw/zh/trading/statistics/index04.html"),
    html.Div(),
    html.A("上櫃報表", href="https://www.tpex.org.tw/web/regular_emerging/statistics/sales_revenue/list.php?l=zh-tw"),
    dcc.Upload(
        id='upload_data',
        children=html.Div([
            '上傳檔案'
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        multiple=True
    ),
    html.Div(id='output_df') 
])

def parse_contents(contents, filename):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    try:
        if "xls" in filename.lower():
            if "O" in filename:
                # Assume that the user uploaded an excel file
                df = pd.read_excel(io.BytesIO(decoded), header=8)
                df.drop(["Unnamed: 1", "本年度\r\n2024", "上年度\r\n2023", "Unnamed: 7",
                "Unnamed: 8"], axis=1, inplace=True)
                df = df.set_axis(["Symbol", "Last_Month", "Cur_Month", "Accum", 
                                                  "Prev_Year_Month", "Prev_Year_Accum"], axis=1)
                df.fillna("NAN", inplace=True)
                df.drop(index=0, inplace=True)
                for i in df.index:
                    df.loc[i, "Name"] = str(df.loc[i, "Symbol"])[4:].strip()
                    df.loc[i, "Symbol"] = str(df.loc[i, "Symbol"])[:4].strip()
                for i in df.index:
                    if df["Symbol"][i].isnumeric() == False:
                        df.drop(index=i, inplace=True)
                    elif len(df["Symbol"][i])<4:
                        df.drop(index=i, inplace=True)
                df["Symbol"] = df["Symbol"].astype(int)
                df.replace(0, np.nan, inplace=True)
                df["%MoM"] = round((df["Cur_Month"]-
                                    df["Last_Month"])/df["Last_Month"]*100, 2)
                df["%YoY"] = round((df["Cur_Month"]-
                                    df["Prev_Year_Month"])/df["Prev_Year_Month"]*100, 2)
                df["%YoY_Accum"] = round((df["Accum"]-
                                          df["Prev_Year_Accum"])/df["Prev_Year_Accum"]*100, 2)
                df.drop(["Last_Month", "Prev_Year_Month", "Prev_Year_Accum"], axis=1, inplace=True)
            else:
                df = pd.read_excel(io.BytesIO(decoded), header=9)
                df.rename(columns={'                ': "ticker", '        ': ""}, inplace=True)
                df.drop(df.iloc[:, 6:].columns.to_list(), axis=1, inplace=True)
                for i in df.index:
                    df.loc[i, "name"] = str(df.loc[i, "ticker"])[4:].strip()
                    df.loc[i, "ticker"] = str(df.loc[i, "ticker"])[:4].strip()
                for i in df.index:
                    if df["ticker"][i].isnumeric() == False:
                        df.drop(index=i, inplace=True)
                    elif len(df["ticker"][i])<4:
                        df.drop(index=i, inplace=True)
                df["ticker"] = df["ticker"].astype(int)
                df = df.set_axis(["Symbol", "Last_Month", "Cur_Month", "Accum", 
                                      "Prev_Year_Month", "Prev_Year_Accum", "Name"], axis=1)
                df.replace(0, np.nan, inplace=True)
                df["%MoM"] = round((df["Cur_Month"]-
                                    df["Last_Month"])/df["Last_Month"]*100, 2)
                df["%YoY"] = round((df["Cur_Month"]-
                                    df["Prev_Year_Month"])/df["Prev_Year_Month"]*100, 2)
                df["%YoY_Accum"] = round((df["Accum"]-
                                          df["Prev_Year_Accum"])/df["Prev_Year_Accum"]*100, 2)
                df.drop(["Last_Month", "Prev_Year_Month", "Prev_Year_Accum"], axis=1, inplace=True)
    except Exception as e:
        print(e)
        return html.Div([
            '上傳發生錯誤'
        ])
    
    return df


@callback(Output('output_df', 'children'),
              Input('upload_data', 'contents'),
              State('upload_data', 'filename'))
def update_output(dfs, names):
    if dfs is not None:
        contents = [parse_contents(c, n) for c, n in
            zip(dfs, names)]
        if len(contents)==1:
            final_df = contents[0].merge(daily_info, left_on="Symbol", right_on="Symbol", how="left")
            final_df["Symbol"] = final_df["Symbol"].astype("str") + " " + final_df["Name"]
            final_df.drop(["Name"], axis=1, inplace=True)
            final_df = final_df.set_axis(["名稱", "月營收", "年累計營收", "%MoM", "%YoY", 
                                          "%YoY累計", "股價", "毛利率", "BPS", "P/B", "EPS(12個月)", 
                                          "EPS預測", "P/E", "P/E預測"], axis=1)
        else:
            list_of_df = [parse_contents(c, n) for c, n in zip(dfs, names)]
            contents = pd.concat([list_of_df[0], list_of_df[1]])
            final_df = contents.merge(daily_info, left_on="Symbol", right_on="Symbol", how="left")
            final_df["Symbol"] = final_df["Symbol"].astype("str") + " " + final_df["Name"]
            final_df.drop(["Name"], axis=1, inplace=True)
            final_df = final_df.set_axis(["名稱", "月營收", "年累計營收", "%MoM", "%YoY", 
                                          "%YoY累計", "股價", "毛利率", "BPS", "P/B", "EPS(12個月)", 
                                          "EPS預測", "P/E", "P/E預測"], axis=1)
        return dash_table.DataTable(
                final_df.to_dict("records"),
                [{'name': i, 'id': i} for i in final_df.columns],
                filter_action="native",
                sort_action='native',
                filter_options={"placeholder_text": "Filter column..."},
            ),


if __name__ == '__main__':
    app.run_server(debug=False)
