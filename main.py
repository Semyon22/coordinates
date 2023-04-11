import openpyxl
import numpy
from  selenium import  webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import lxml
from fake_useragent import UserAgent
import time
import requests
import json

#импорт данных с экселя



def read_matrix_from_excel(file_name):
    """
        Данная подпрограмма считывает матрицу  из формата xlsx
        Матрица может быть любой размерности
        На вход подается имя файла или путь до него
        Возражается словарь , где ключом является адрес значением - кортеж , содержащий url
        адресса в яндекс картах и кортеж содержащий координаты объекта
    """
    "to_do:Доделать чтобы нормально считывался файл"
    data_list= {}
    book=openpyxl.open(file_name,read_only=True)
    sheet=book.active
    rows = sheet.max_row
    columns = sheet.max_column
    column=3
    for row in range(2,70):#обработка данных из excel
        data_list[sheet[row][column].value]=('url',(0,0))
    return data_list
    book.close()

data_list = read_matrix_from_excel("Лавина1_Карасукский_район_Тест.xlsx")

for key, value in data_list.items():

  print("{0}: {1}".format(key,value))

def flt_from_str(string):
    """
    :param string:строка должна передаваться в формате '51.660781, 39.200296'
    :return: tuple (x_coord,y_coord)
    """
    str_x_coord=''
    str_y_coord=''
    for i in range(0,len(string)):
        if string[i]!=',':
            str_x_coord+=string[i]
        else:
            str_y_coord=string[i+2:]
            coordinates=(float(str_x_coord),float(str_y_coord))
            return coordinates

def get_adr_url(data_list):
    """
    :param data_list:Функция принимает на вход словарь ,где ключом является адрес ,значением - кортеж , содержащий url
    адреса в яндекс картах и кортеж содержащий координаты объекта.
    :return:data_list В ходе работы функций заполняется url адреса объекта в яндекс картах , координаты объекта
    """
    url = "https://yandex.ru/maps/65/novosibirsk/?ll=82.920430%2C55.030199&z=12"
    driver = webdriver.Chrome(executable_path="C:\\Users\\User\\PycharmProjects\\Coordinates\\hromdriver\\chromedriver.exe")
    try:
        driver.get(url=url)
        time.sleep(2)
        input_adr = driver.find_element(By.CLASS_NAME, 'input__control')
        for key in data_list:
            input_adr = driver.find_element(By.CLASS_NAME, 'input__control')
            input_adr.send_keys(key)
            input_adr.send_keys(Keys.ENTER)
            time.sleep(5)
            soup = BeautifulSoup(driver.page_source,'lxml')
            soup=soup.find(class_="toponym-card-title-view__coords-badge")
            coordinates=flt_from_str(soup.text)
            input_adr.clear()
            time.sleep(3)
            driver.refresh()
            data_list[key]=(driver.current_url,coordinates)
    except Exception as ex:
        print(ex)
    finally:
        return data_list
        driver.close()
        driver.quit()
get_adr_url(data_list=data_list)
for key, value in data_list.items():

  print("{0}: {1}".format(key,value))
# json_object = json.dumps(data_list, indent = 4,encodings='UTF-8')
# print(json_object)