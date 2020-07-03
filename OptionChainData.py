# coding: utf-8
from __future__ import unicode_literals

import requests
import json
import math

url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"



import uno

from apso_utils import msgbox

def roundup(x):
    return int(math.ceil(x / 100.0)) * 100

def currentSheet():
    context = uno.getComponentContext()
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)  
    document = desktop.getCurrentComponent()  
    sheet = document.Sheets[0]
    return sheet


def run():
    '''
    context = uno.getComponentContext()  
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)  
    document = desktop.getCurrentComponent()  
    sheet = document.Sheets[0]
    sheetname = sheet.getName()
    '''
    sheet = currentSheet()
    payload = {}
    headers = {
    'authority': 'www.nseindia.com',
    'pragma': 'no-cache',
    'cache-control': 'no-cache',
    'user-agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36',
    'accept': '*/*',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-mode': 'cors',
    'sec-fetch-dest': 'empty',
    'referer': 'https://www.nseindia.com/get-quotes/derivatives?symbol=NIFTY&identifier=OPTIDXNIFTY09-07-2020CE10600.00',
    'accept-language': 'en-US,en;q=0.9',
    #'cookie': cookies
    }
    
    response = requests.request("GET", url, headers=headers, data = payload)
    res = json.loads(response.text)
    niftyPrice = res['records']['underlyingValue']
    niftyPrice = roundup(niftyPrice)
    strikePrices = res['records']['strikePrices']
    position = strikePrices.index(niftyPrice)
    strikePrices = strikePrices[position-5:position+5]
    data = res['records']['data']
    expiry = '09-Jul-2020'
    finalData = {}
    for s in strikePrices:
        finalData[s] = {}
    for d in data:
        sp = d['strikePrice']
        exp = d['expiryDate']
        if exp == expiry and sp in strikePrices:
            finalData[sp]['PE'] = d['PE']
            finalData[sp]['CE'] = d['CE']
    
    #for f in finalData:
    sheet['A1'].setString("Call Price")
    sheet['B1'].setString("Strike Price")
    sheet['C1'].setString("Put Price")
    i = 2
    for f in finalData:
        sheet['A'+str(i)].setString(finalData[f]['CE']['lastPrice'])
        sheet['B'+str(i)].setString(finalData[f]['CE']['strikePrice'])
        sheet['C'+str(i)].setString(finalData[f]['PE']['lastPrice'])
        i = i+1
