import win32com.client
from selenium import webdriver
import os

driver = webdriver.Chrome('./chromedriver')

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

_path_ = os.getcwd()+r'\data.xls'
wb = excel.Workbooks.Open(_path_)
ws = wb.ActiveSheet

for i in range(2,550):
    driver.get('https://search.naver.com/search.naver?where=nexearch&sm=tab_org&qvt=0&query='+ ws.Cells(i,7).Value)

    try :
        elem = driver.find_element_by_id('no-matched-address-list')
    except Exception as e:
        elem = driver.find_element_by_id('unique')

    arrAdress = elem.text.split()

    ws.Cells(i,2).Value = arrAdress[2]

    if arrAdress[3][-1] == "리" :
        ws.Cells(i,3).Value = arrAdress[3]
        arrNum = arrAdress[4].split("-")
    else :
        arrNum = arrAdress[3].split("-")

    ws.Cells(i,5).Value = arrNum[0]
    
    if len(arrNum) >= 2 :
        ws.Cells(i,6).Value = arrNum[1]

wb.SaveAs('output.xls')
excel.Quit()
driver.quit()
