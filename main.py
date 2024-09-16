from itertools import product

import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd

data_list = []

def crawl(date,contracts):
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
        current_row = rows[2]
        td = current_row.find_all("td")
        amount = int(td[-2].text.replace(",",""))
        num = int(contracts)
        data = [day_mark] + [amount]
        if num > 0 and amount >= num:
            print(data)
            data_list.append(data)
        if num < 0 and amount <= num:
            data_list.append(data)
            print(data)

    except AttributeError:
        pass
        # print(day_mark," no data")

def main():
    date = datetime.today()
    contracts = input("請輸入判斷口數")
    data_days = input("抓取天數")
    while True:
        # print(date)
        crawl(date,contracts)
        date = date - timedelta(days=1)
        if date < datetime.today() - timedelta(days=int(data_days)):
            break
    print(f"資料共 {len(data_list)} 筆")
    save = input("轉出excel?(y/n)")
    if save == "y":
        build_file()
    again = input("重新查詢?(y/n)")
    if again == "y":
        main()

def build_file():
    print("轉出excel...")
    # headers = ["日期","商品","身份別","交易多方口數","交易多方金額","交易空方口數","交易空方金額","交易口數淨額","交易契約淨額","未平倉多方口數","未平倉多方金額","未平倉空方口數","未平倉空方金額","未平倉口數","未平倉淨額"]
    pf = pd.DataFrame(data_list, columns=["日期","未平倉"])
    pf.to_excel("futures.xlsx", index=False, engine="openpyxl")


if __name__ == "__main__":
    main()



