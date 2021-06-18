import pygsheets
import numpy as np
import requests
from bs4 import BeautifulSoup


def CellName(column, row_num):
    name = column + str(row_num)
    return name

def GetCell(column, row_num):
    adress = CellName(column, row_num)
    value = wks.cell(adress).value
    return value

def SetCell(column, row_num, value):
    adress = CellName(column, row_num)
    wks.cell(adress).value = value

def GetTicker(row_num):
    ticker = GetCell('B', row_num)
    return ticker

def GetSite(page_url, ticker):
    url = page_url + ticker
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:45.0) Gecko/20100101 Firefox/45.0'}
    r = requests.get(url, headers = headers)
    with open('test.html', 'w', encoding="utf-8") as output_file:
        output_file.write(r.text)

    return r.text

def GetFinViz(ticker):
    url = "https://finviz.com/quote.ashx?t="
    html_text = GetSite(url, ticker)
    str = []
    st = ''
    soup = BeautifulSoup(html_text, 'html.parser')
    
    table = soup.find_all('td', class_ = 'snapshot-td2')
    #with open('test.html', 'w', encoding="utf-8") as output_file:
        #output_file.write(st)

    str.append(table[13].text)
    str.append(table[38].text)
    print(str)


gc = pygsheets.authorize()

# Open spreadsheet
sh = gc.open('sampleTitle')
wks = sh.sheet1
start_row_index = 2
# Read Tickers from cells in 1sr column do-while
row_index = start_row_index

while (True):
    ticker = GetTicker(row_index)
    if ticker != '':
        GetFinViz(ticker)
        #SetCell('E', row_index, "=GOOGLEFINANCE(\""+ticker+"\",\"pe\")")
        row_index += 1
    else:
        break

print("hui")

