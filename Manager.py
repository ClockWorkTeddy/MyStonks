import pygsheets
import yfinance as yf
import finviz as fv

#+
def CellName(column, row_num):
    name = column + str(row_num)
    return name
#+
def GetCell(column, row_num, wks):
    adress = CellName(column, row_num)
    value = wks.cell(adress).value
    return value
#+
def SetCell(column, row_num, value, wks):
    adress = CellName(column, row_num)
    wks.cell(adress).value = value
#+
def GetTicker(row_num, wks):
    ticker = GetCell('B', row_num, wks)
    return ticker
#+
def GetYahooApi(ticker):
    res = []
    company = yf.Ticker(ticker)
    rec = company.info['recommendationMean']
    price = company.info['targetMeanPrice']
    res.append(rec)
    res.append(price)
    return res
#+
def GetFinvizApi(ticker):
    res = []
    data = fv.get_stock(ticker)
    peg = data['PEG']
    eps = data['EPS past 5Y']
    res.append(peg)
    res.append(eps)    
    return res

def Main():
    gc = pygsheets.authorize()
    sh = gc.open('Daily')
    wks = sh.worksheet('title', 'NewStonks')
    row_index = input()
    rows_qnt = input()
    limit = row_index + rows_qnt

    while (True):
        ticker = GetTicker(row_index, wks)
        if ticker != '' and row_index < limit:
            fin_viz_data = GetFinvizApi(ticker)
            SetCell('F', row_index, fin_viz_data[0], wks)
            SetCell('G', row_index, fin_viz_data[1], wks)
            yahoo_data = GetYahooApi(ticker)
            SetCell('K', row_index, yahoo_data[0], wks)
            SetCell('J', row_index, yahoo_data[1], wks)
            print(ticker + str(row_index))
            row_index += 1
        else:
            break

Main()
print("hui")

