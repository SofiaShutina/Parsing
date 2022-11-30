import requests 
from bs4 import BeautifulSoup
import re
import json
import time
from selenium import webdriver
import openpyxl



chrome_driver = 'C:/Tools/ChromeDriver/chromedriver.exe' 
chrome_options = webdriver.ChromeOptions()
driver = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)
item_info =[]
for item in range(1,6,1):
    driver.get(f"https://www.kinopoisk.ru/lists/movies/top250/?page={item}")
    time.sleep(30)
    html= driver.page_source
    with open(f"data/{item}.html","w",encoding='utf-8') as file:
        file.write(html)
    time.sleep(40)
    headers= {
    "Accept": "*/*",
    "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36"
    }

    with open(f"data/{item}.html",encoding='utf-8') as file:
        src=file.read()
    soup=BeautifulSoup(src,"lxml")
    hrefs=[]
    all_films=soup.find_all("div",class_="styles_root__ti07r")


    ch=","
    hrefs=[]
    
    count=0
    info=[]
    for item in all_films:
        
        item_russian_title=item.find("div",class_="base-movie-main-info_mainInfo__ZL_u3").text
     
        item_mark=item.find("div",class_="styles_rating__ni2L0 styles_root___s7Tg styles_rootMd__ZvdRj styles_rootPositive__PIwO2").text
        
        try:
                item_native_title=item.find("span",class_="desktop-list-main-info_secondaryTitle__ighTt").text
        except Exception:
                item_native_title ="Нет"
        item_href="https://www.kinopoisk.ru/"+item.find("a",class_="base-movie-main-info_link__YwtP1").get("href")
     
        hrefs.append(item_href)
        item_years=item.find("span",class_="desktop-list-main-info_secondaryText__M_aus").text.strip()
        years=item_years[:6].replace(",","")
     
        item_country=item.find("span",class_="desktop-list-main-info_truncatedText__IMQRP").text
        
        first_word = re.findall(r'\w+', item_country)[0]
        
        second_word = re.findall(r'\w+', item_country)[1]
        
        item_info.append(
                {
                    "Russian_title": item_russian_title,
                    "Native_title":item_native_title,
                    "Year": years,
                    "Mark":item_mark,
                    "Genre":second_word,
                    "Country":first_word
                }
            )

with open(f"index.json","a",encoding='utf-8') as file:
    json.dump(item_info,file,indent=4,ensure_ascii=False)
with open('index.json',encoding='utf-8') as file:
    data = json.load(file)
book=openpyxl.Workbook()

sheet = book.active

sheet['A1'] ='Русское название'
sheet['B1'] ='Оригинальное название'
sheet['C1'] ='Год'
sheet['D1'] ='Рейтинг'
sheet['E1'] ='Жанр'
sheet['F1'] ='Страна производства'

row=2 
for item in data:
    sheet[row][0].value = item['Russian_title']
    sheet[row][1].value = item['Native_title']
    sheet[row][2].value = item['Year']
    sheet[row][3].value = item['Mark']
    sheet[row][4].value = item['Genre']
    sheet[row][5].value = item['Country']
    row+=1

book.save("my_book.xlsx")
book.close()

print("#"*20,"DONE","#"*20)