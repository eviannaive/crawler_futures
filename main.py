from itertools import product

import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd

data_list = []

def crawl(date):
    response = requests.get('https://www.taifex.com.tw/cht/3/futContractsDate?queryDate={}%2F{}%2F{}'.format(date.year,date.month,date.day))
    if response.status_code == requests.codes.ok:
        soup = BeautifulSoup(response.text, "html.parser")
        day_mark = f"{date.year}/{date.month}/{date.day}"

    else:
        print("connection error")
        return

    try:
        table = soup.find("table", class_="table_f")
        body = table.find("tbody")
        rows = body.find_all("tr")
        for d in rows:
            tds = d.find_all("td")
            cells = [td.text.strip() for td in tds]
            if len(cells) == 15:
                product = cells[1]
                data = [day_mark] + cells[1:]
            else:
                data = [day_mark] + [product] + cells
            if(cells[0] == '期貨小計'):
                break
            data_list.append(data)
        print(day_mark+" 資料轉出...")
    except AttributeError:
        print(day_mark," no data")

def day_loop():
    date = datetime.today()
    while True:
        # print(date)
        crawl(date)
        date = date - timedelta(days=1)
        if date < datetime.today() - timedelta(days=3):
            break
    build_file()

def build_file():
    print("轉出excel...",data_list)
    headers = ["日期","商品","身份別","交易多方口數","交易多方金額","交易空方口數","交易空方金額","交易口數淨額","交易契約淨額","未平倉多方口數","未平倉多方金額","未平倉空方口數","未平倉空方金額","未平倉口數","未平倉淨額"]
    pf = pd.DataFrame(data_list, columns=headers)
    pf.to_excel("futures.xlsx", index=False, engine="openpyxl")



day_loop()



