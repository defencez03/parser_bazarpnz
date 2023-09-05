import requests
import os
import pandas as pd
from openpyxl import Workbook
from bs4 import BeautifulSoup


def request(URL):
    r = requests.get(URL)
    if r.status_code != 200:
        print("Error page")
        return 0
    else:
        return r


def countPage():
    r = request(url)
    soup = BeautifulSoup(r.text.encode("1251"), features="html.parser")
    str_count = soup.find("form", attrs={"name": "topage"})
    for i in str_count.find_all("a"):
        if i.text == "следующей":
            return int(i_prev.text)
        i_prev = i


def createExcel():
    wb = Workbook()
    wb.save("new.xlsx")


def parse(url):
    count = 0
    price = []
    text = []
    count_page = countPage()
    for i in range(count_page):
        r = request(url)
        print(f"Parse page №{i+1}...")
        soup = BeautifulSoup(r.text.encode("1251"), features="html.parser")
        for item in soup.find_all("tr", class_="norm"):
            price.append(item.find("td", class_="price").get_text())
            text.append(item.find("td", class_="text").get_text())
        url = url.replace(str(count), str(count + 45))
        count += 45
    print("Complete")
    data_set = {
        "text": text,
        "price": price
    }
    return data_set


def writeToFile(data_set):
    df = pd.DataFrame(data_set)
    createExcel()
    with pd.ExcelWriter('new.xlsx') as writer:
        df.to_excel(writer, index=False)
    print("Please, check your file...")


if __name__ == "__main__":
    countUrl = 0
    countDs = 0
    while True:
        print("Парсер сайта bazarpnz:\n"
              "1. Ввести URL\n"
              "2. Парсинг сайта\n"
              "3. Сохранение в Excel-файл\n"
              "4. Выход")
        var = input("> ")
        if var == "1":
            countUrl = 1
            url = str(input("Введите URL: "))
        elif var == "2":
            if countUrl == 0: continue
            countDs = 1
            ds = parse(url)
            input()
        elif var == "3":
            if countDs == 0 or countUrl == 0: continue
            writeToFile(ds)
            input()
        elif var == "4":
            os.system('cls')
            break
        os.system('cls')