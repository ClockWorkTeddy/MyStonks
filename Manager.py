import pygsheets
import numpy as np
import requests
from selenium import webdriver
from bs4 import BeautifulSoup
import time

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
    r = requests.get(url, timeout=30, headers = headers)
    return r.text

def GetFinViz(ticker):
    str = []
    html_text = GetSite("https://finviz.com/quote.ashx?t=", ticker)
    soup = BeautifulSoup(html_text, 'html.parser')
    table = soup.find_all('td', class_ = 'snapshot-td2')
    str.append(table[13].text)
    str.append(table[38].text)
    return str

def GetHtml(wd):
    time_out = 0.5
    wd.get("https://finance.yahoo.com/quote/" + ticker)
    time.sleep(time_out)
    wd.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(time_out)
    wd.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    html_text = wd.execute_script('return document.body.innerHTML;')
    return html_text

def GetYahoo(ticker, wd):
    res = []
    html_text = GetHtml(wd)
    soup = BeautifulSoup(html_text)
    coeffs = soup.find_all('div', class_ = 'Bdbw(2px) Bdbs(s) Bdbc($seperatorColor) H(1em) Pos(r) Mt(30px) Mx(10%)')
    coeff = coeffs[0].contents[5].contents[0]
    res.append(coeff)
    price = soup.find_all('div', class_ = 'Pos(r) T(5px) Miw(100px) Fz(s) Fw(500) D(ib) C($primaryColor)Ta(c) Translate3d($half3dTranslate)')
    data = price[0].contents[4].contents[0]
    res.append(data)
    return res

gc = pygsheets.authorize()
sh = gc.open('Daily')
wks = sh.worksheet('title', 'NewStonks')
row_index = 2
wd = webdriver.Chrome()

while (True):
    ticker = GetTicker(row_index)
    if ticker != '':
        fin_viz_data = GetFinViz(ticker)
        SetCell('F', row_index, fin_viz_data[0])
        SetCell('G', row_index, fin_viz_data[1])
        yahoo_data = GetYahoo(ticker, wd)
        SetCell('K', row_index, yahoo_data[0])
        SetCell('J', row_index, yahoo_data[1])
        row_index += 1
    else:
        break

wd.close()

print("hui")

