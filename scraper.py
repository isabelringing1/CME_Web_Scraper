import xlwings as xw
import requests
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from datetime import date, time, timedelta
import time
import sys

REG_DATE = str(xw.Range("G1").value[1:-1])
EURO_DATES = [xw.Range("G2").value[1:-1], xw.Range("H2").value[1:-1], xw.Range("I2").value[1:-1]]
PATH = str(xw.Range("J2").value)

class update_error(Exception):
    pass

def get_data():
    full_update = True
    stem = 'https://www.cmegroup.com/trading/interest-rates/'
    items = ['us-treasury/2-year-us-treasury-note_quotes_settlements_futures.html', 
            'us-treasury/3-year-us-treasury-note_quotes_settlements_futures.html', 
            'us-treasury/5-year-us-treasury-note_quotes_settlements_futures.html', 
            'us-treasury/10-year-us-treasury-note_quotes_settlements_futures.html', 
            'us-treasury/ultra-t-bond_quotes_settlements_futures.html',
            'us-treasury/30-year-us-treasury-bond_quotes_settlements_futures.html']
    today = date.today().strftime("%m/%d/%Y")
    #today = (date.today() - timedelta(days=1)).strftime("%m/%d/%Y")
    settles = {} 
    driver = webdriver.Chrome(executable_path=PATH)
    while len(items):
        incomplete = []
        for item in items:
            driver.get(stem + item) #combine into url
            html = driver.page_source
            try:
                soup = BeautifulSoup(html, 'html.parser')
                #fetches latest date
                latest = soup.find_all('option', selected=True)[1]['value']
                if latest != today:
                    raise update_error()
                table = soup.find('table', attrs={'id':'settlementsFuturesProductTable'})
                table_body = table.find('tbody') 
                row = table_body.find_all('tr')[0]
                for tr in table_body.find_all('tr'):
                    if tr.find_all('th')[0].text == REG_DATE:
                        row = tr
                        break
                settle_col = row.find_all('td')[5]
                title = soup.find('h1', ).text
                settles[re.sub('[\n+]', '', title)] = re.sub("'", ".", settle_col.text)
            except update_error:
                #print(item + " has not been updated yet today. Try again in a few minutes.")
                title = soup.find('h1', ).text
                settles[re.sub('[\n+]', '', title)] = "This value has not been updated today. "
                full_update = False
            except:
                print("Can't get " + item + ", trying again...")
                incomplete.append(item)
        items = incomplete

    #Get Eurodollar settles
    euro = 'stir/eurodollar_quotes_settlements_futures.html'
    EURO_DATES_COPY = EURO_DATES[:]
    while len(EURO_DATES_COPY):
        driver.get(stem + euro)
        html = driver.page_source
        try:
            soup = BeautifulSoup(html, 'html.parser')
            latest = soup.find_all('option', selected=True)[1]['value']
            if latest != today:
                raise update_error()
            table = soup.find('table', attrs={'id':'settlementsFuturesProductTable'})
            table_body = table.find('tbody') 
            title = soup.find('h1', ).text
            for tr in table_body.find_all('tr'):
                if tr.find_all('th')[0].text in EURO_DATES:
                    settle_col = tr.find_all('td')[5]
                    settles[re.sub('[\n+]', '', title) + " " + tr.find_all('th')[0].text] = settle_col.decode_contents()
                    EURO_DATES_COPY.remove(tr.find_all('th')[0].text)
        except update_error: 
            settles['Eurodollar FuturesSettlements ' + EURO_DATES[0]] = "This value has not been updated today. "
            settles['Eurodollar FuturesSettlements ' + EURO_DATES[1]] = "This value has not been updated today. "
            settles['Eurodollar FuturesSettlements ' + EURO_DATES[2]] = "This value has not been updated today. "
            EURO_DATES_COPY = []
            full_update = False
        except:
            print("Can't get " + euro + ", trying again...")

    print(settles)
    insert_row(settles, full_update)

def insert_row(settles, full_update):
    #row: GE GE GE TU FV ZN ZB
    row = [
        date.today().strftime("%m/%d/%Y"),
        settles['Eurodollar FuturesSettlements ' + EURO_DATES[0]],
        settles['Eurodollar FuturesSettlements ' + EURO_DATES[1]],
        settles['Eurodollar FuturesSettlements ' + EURO_DATES[2]],
        settles['2-Year T-Note FuturesSettlements'],
        settles['5-Year T-Note FuturesSettlements'],
        settles['10-Year T-Note FuturesSettlements'],
        settles['U.S. Treasury Bond FuturesSettlements']
    ]
    last = xw.Range('A1:OO1').end('down')
    if str(last.offset(0, 3).value) == "This value has not been updated today. ":
        next_empty = last
        last = last.offset(-1)
    else:
        next_empty = last.offset(1)
        next_empty_row = xw.Range(next_empty, next_empty.offset(0, 300))
        next_empty_row.api.Insert()
        next_empty = next_empty.offset(-1)
    xw.Range(last, next_empty).formula = last.formula
    xw.Range(next_empty.offset(0, 2), next_empty.offset(0, 9)).value = row
    if full_update:
        for i in range (10, 300):
            xw.Range(last.offset(0,i), next_empty.offset(0,i)).formula = last.offset(0,i).formula

def main(): 
    get_data()
