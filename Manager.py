import pygsheets
import yfinance as yf
import finviz as fv
#+
def CellName(column, row_num):
    name = column + str(row_num)
    return name
#+
def GetCell(column, row_num):
    adress = CellName(column, row_num)
    value = wks.cell(adress).value
    return value
#+
def SetCell(column, row_num, value):
    adress = CellName(column, row_num)
    wks.cell(adress).value = value
#+
def GetTicker(row_num):
    ticker = GetCell('B', row_num)
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

gc = pygsheets.authorize()
sh = gc.open('Daily')
wks = sh.worksheet('title', 'NewStonks')
row_index = 2

while (True):
    ticker = GetTicker(row_index)
    if ticker != '':
        fin_viz_data = GetFinvizApi(ticker)
        SetCell('F', row_index, fin_viz_data[0])
        SetCell('G', row_index, fin_viz_data[1])
        yahoo_data = GetYahooApi(ticker)
        SetCell('K', row_index, yahoo_data[0])
        SetCell('J', row_index, yahoo_data[1])
        print(ticker + str(row_index))
        row_index += 1
    else:
        break

print("hui")

