from pickle import FALSE, TRUE
import pygsheets
import yfinance as yf
import finviz as fv
import datetime as dt

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
    res.append(company.info['recommendationMean'])
    res.append(company.info['targetMeanPrice'])
    
    return res
#+
def GetFinvizApi(ticker):
    res = []
    data = fv.get_stock(ticker)
    res.append(data['PEG'])
    res.append(data['EPS past 5Y'])    
    
    return res

def GetTimeStamp():
    date_time = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    return date_time

def OverLimit(row_index, rows_qnt, limit):
    if rows_qnt == 0 :
        return TRUE
    else:
        if row_index >= limit :
            return FALSE
        else:
            return TRUE

def GetWorkSheet() :
    gc = pygsheets.authorize()
    sh = gc.open('Daily')
    return sh.worksheet('title', 'Stonks')

def FinVizProcessing(row_index, wks, ticker) :
    fin_viz_data = GetFinvizApi(ticker)
    SetCell('G', row_index, fin_viz_data[0], wks)
    SetCell('H', row_index, fin_viz_data[1], wks)

def YahooFinProcessing(row_index, wks, ticker) :
    yahoo_data = GetYahooApi(ticker)
    SetCell('L', row_index, yahoo_data[0], wks)
    SetCell('K', row_index, yahoo_data[1], wks)

def FillCells(row_index, date_time, wks, ticker) :
    SetCell('C', row_index, date_time, wks)
    FinVizProcessing(row_index, wks, ticker)
    YahooFinProcessing(row_index, wks, ticker)
    
    print(ticker + str(row_index))

def Main():
    wks = GetWorkSheet()
    row_index = int(input()) + 1
    rows_qnt = int(input())
    limit = row_index + rows_qnt

    while (True):
        ticker = GetTicker(row_index, wks)
        date_time = GetTimeStamp()

        if ticker != '' and OverLimit(row_index, rows_qnt, limit) != FALSE:
            FillCells(row_index, date_time, wks, ticker)
            row_index += 1
        else:
            break

Main()

print("hui")

