mport openpyxl
import numpy
import pandas
from  selenium import  webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import lxml
# from fake_useragent import UserAgent
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
    for row in range(492,593):#обработка данных из excel
        data_list[sheet[row][column].value]=('url',(0,0))
    return data_list
    book.close()

data_list = read_matrix_from_excel("Новосибирский р-н_на заполнение.xlsx")

for key, value in data_list.items():

  print("{0}: {1}",value)

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
    driver = webdriver.Chrome(executable_path="C:\\Users\\helpdesk\\PycharmProjects\\pythonProject\\hromdriver\\hromdriver.exe")
    try:
        driver.get(url=url)
        time.sleep(2)
        input_adr = driver.find_element(By.CLASS_NAME, 'input__control')
        for key in data_list:
            input_adr = driver.find_element(By.CLASS_NAME, 'input__control')
            input_adr.send_keys(key)
            input_adr.send_keys(Keys.ENTER)
            time.sleep(3)
            try:
                soup = BeautifulSoup(driver.page_source,'lxml')
                soup=soup.find(class_="toponym-card-title-view__coords-badge")
                coordinates=flt_from_str(soup.text)
            except Exception as ex:
                print(ex)


                continue
                # driver.get(url=url)
            finally:
                input_adr.clear()
                time.sleep(1)
                driver.refresh()

                data_list[key]=(driver.current_url,coordinates)
    except Exception as ex:
        print(ex)
    finally:
        return data_list
        driver.close()
        driver.quit()
data_list=get_adr_url(data_list=data_list)
for key, value in data_list.items():

  print("{0}: {1}".format(key,value))
def outPut(data_list):
    wb = openpyxl.load_workbook('output.xlsx')  # Открываем тестовый Excel файл
    wb.create_sheet('Sheet1')  # Создаем лист с названием "Sheet1"
    worksheet = wb['Sheet1']  # Делаем его активным
    i=1
    for key, value in data_list.items():

        worksheet[f'A{i}']=key #A указанную ячейку на активном листе пишем все, что в кавычках
        url,coord=value
        x,y=coord
        x=str(x)
        y=str(y)
        coord=x+', '+y
        worksheet[f'B{i}'] =url
        worksheet[f'C{i}']=coord

        i=i+1
    wb.save('output.xlsx') #Сохраняем измененный файл
outPut(data_list)
