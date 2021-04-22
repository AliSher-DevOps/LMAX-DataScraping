from selenium import webdriver
from csv import writer
import pandas as pd
import xlwings as xw
import time
from openpyxl import Workbook
workbook = Workbook()
sheet = workbook.active
driver=webdriver.Chrome("chromedriver.exe")
driver.get("https://lmax-resprime.mtp-cdn.com/resprime2.html?ld4=1&depth=4&cdn=1&bodycss=ts")
driver2=webdriver.Chrome("chromedriver.exe")
driver2.get("https://lmax-resprime.mtp-cdn.com/resprime2.html?ny4=1&depth=4&cdn=1&show=10&bodycss=ts")
time.sleep(10)
try:
     wb = xw.Book('DataFile3.xlsx')
     sht1 = wb.sheets['Sheet']
except:
    print("file not found")
time.sleep(15)
try:
    while(True):  
        currency = []
        c2 = []
        c3 = []
        bid = []
        ask = []
        c5 = []
        currencyY = []
        c2Y = []
        c3Y = []
        bidY = []
        askY = []
        c5Y = []
        a=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[1]').text
        v = a.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        b=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[2]').text
        v = b.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        c=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[3]').text
        v = c.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        d=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[4]').text
        v = d.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        e=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[5]').text
        v = e.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        f=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[6]').text
        v = f.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        g=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[7]').text
        v = g.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        h=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[8]').text
        v = h.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        i=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[9]').text
        v = i.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        j=driver.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[10]').text
        v = j.split()
        currency.append(v[0])
        c2.append(v[1])
        c3.append(v[2])
        bid.append(v[3])
        ask.append(v[4])
        c5.append(v[5])
        a2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[1]').text
        v = a2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5]) 
        b2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[2]').text
        v = b2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])    
        c22=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[3]').text
        v = c22.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5]) 
        d2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[4]').text
        v = d2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])   
        e2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[5]').text
        v = e2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])    
        f2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[6]').text
        v = f2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])     
        g2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[7]').text
        v = g2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])   
        h2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[8]').text
        v = h2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])     
        i2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[9]').text
        v = i2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])   
        j2=driver2.find_element_by_xpath('/html/body/div/div/div[4]/div[2]/div/table/tbody/tr[10]').text
        v = j2.split()
        currencyY.append(v[0])
        c2Y.append(v[1])
        c3Y.append(v[2])
        bidY.append(v[3])
        askY.append(v[4])
        c5Y.append(v[5])



       

        sheet["A1"] = "London_Currency"
        sheet["B1"] = "London_data"
        sheet["C1"] = "London_data"
        sheet["D1"] = "London_BID"
        sheet["E1"] = "London_ASK"
        sheet["F1"] = "London_data"

        sheet["H1"] = "New_York_Currency"
        sheet["I1"] = "New_York_data"
        sheet["J1"] = "New_York_data"
        sheet["K1"] = "New_York_BID"
        sheet["L1"] = "New_York_ASK"
        sheet["M1"] = "New_York_data"

        sht1.range("A2").value = currency[0]
        sht1.range("B2").value = c2[0]
        sht1.range("C2").value = c2[0]
        sht1.range("D2").value = bid[0]
        sht1.range("E2").value = ask[0]
        sht1.range("F2").value = c5[0]

        sht1.range("A3").value = currency[1]
        sht1.range("B3").value = c2[1]
        sht1.range("C3").value = c2[1]
        sht1.range("D3").value = bid[1]
        sht1.range("E3").value = ask[1]
        sht1.range("F3").value = c5[1]

        sht1.range("A4").value = currency[2]
        sht1.range("B4").value = c2[2]
        sht1.range("C4").value = c2[2]
        sht1.range("D4").value = bid[2]
        sht1.range("E4").value = ask[2]
        sht1.range("F4").value = c5[2]

        sht1.range("A5").value = currency[3]
        sht1.range("B5").value = c2[3]
        sht1.range("C5").value = c2[3]
        sht1.range("D5").value = bid[3]
        sht1.range("E5").value = ask[3]
        sht1.range("F5").value = c5[3]

        sht1.range("A6").value = currency[4]
        sht1.range("B6").value = c2[4]
        sht1.range("C6").value = c2[4]
        sht1.range("D6").value = bid[4]
        sht1.range("E6").value = ask[4]
        sht1.range("F6").value = c5[4]

        sht1.range("A7").value = currency[5]
        sht1.range("B7").value = c2[5]
        sht1.range("C7").value = c2[5]
        sht1.range("D7").value = bid[5]
        sht1.range("E7").value = ask[5]
        sht1.range("F7").value = c5[5]

        sht1.range("A8").value = currency[6]
        sht1.range("B8").value = c2[6]
        sht1.range("C8").value = c2[6]
        sht1.range("D8").value = bid[6]
        sht1.range("E8").value = ask[6]
        sht1.range("F8").value = c5[6]

        sht1.range("A9").value = currency[7]
        sht1.range("B9").value = c2[7]
        sht1.range("C9").value = c2[7]
        sht1.range("D9").value = bid[7]
        sht1.range("E9").value = ask[7]
        sht1.range("F9").value = c5[7]

        sht1.range("A10").value = currency[8]
        sht1.range("B10").value = c2[8]
        sht1.range("C10").value = c2[8]
        sht1.range("D10").value = bid[8]
        sht1.range("E10").value = ask[8]
        sht1.range("F10").value = c5[8]

        sht1.range("A11").value = currency[9]
        sht1.range("B11").value = c2[9]
        sht1.range("C11").value = c2[9]
        sht1.range("D11").value = bid[9]
        sht1.range("E11").value = ask[9]
        sht1.range("F11").value= c5[9]

        sht1.range("H2").value = currencyY[0]
        sht1.range("I2").value = c2Y[0]
        sht1.range("J2").value = c3Y[0]
        sht1.range("k2").value = bidY[0]
        sht1.range("L2").value = askY[0]
        sht1.range("M2").value = c5Y[0]

        sht1.range("H3").value = currencyY[1]
        sht1.range("I3").value = c2Y[1]
        sht1.range("J3").value= c3Y[1]
        sht1.range("k3").value = bidY[1]
        sht1.range("L3").value = askY[1]
        sht1.range("M3").value = c5Y[1]

        sht1.range("H4").value = currencyY[2]
        sht1.range("I4").value = c2Y[2]
        sht1.range("J4").value = c3Y[2]
        sht1.range("k4").value = bidY[2]
        sht1.range("L4").value = askY[2]
        sht1.range("M4").value = c5Y[2]

        sht1.range("H5").value = currencyY[3]
        sht1.range("I5").value = c2Y[3]
        sht1.range("J5").value = c3Y[3]
        sht1.range("k5").value = bidY[3]
        sht1.range("L5").value = askY[3]
        sht1.range("M5").value = c5Y[3]

        sht1.range("H6").value = currencyY[4]
        sht1.range("I6").value = c2Y[4]
        sht1.range("J6").value = c3Y[4]
        sht1.range("k6").value = bidY[4]
        sht1.range("L6").value = askY[4]
        sht1.range("M6").value = c5Y[4]

        sht1.range("H7").value = currencyY[5]
        sht1.range("I7").value = c2Y[5]
        sht1.range("J7").value = c3Y[5]
        sht1.range("k7").value = bidY[5]
        sht1.range("L7").value = askY[5]
        sht1.range("M7").value = c5Y[5]

        sht1.range("H8").value = currencyY[6]
        sht1.range("I8").value = c2Y[6]
        sht1.range("J8").value = c3Y[6]
        sht1.range("k8").value = bidY[6]
        sht1.range("L8").value = askY[6]
        sht1.range("M8").value = c5Y[6]

        sht1.range("H9").value = currencyY[7]
        sht1.range("I9").value = c2Y[7]
        sht1.range("J9").value = c3Y[7]
        sht1.range("k9").value = bidY[7]
        sht1.range("L9").value = askY[7]
        sht1.range("M9").value = c5Y[7]


        sht1.range("H10").value = currencyY[8]
        sht1.range("I10").value = c2Y[8]
        sht1.range("J10").value = c3Y[8]
        sht1.range("k10").value = bidY[8]
        sht1.range("L10").value = askY[8]
        sht1.range("M10").value = c5Y[8]


        sht1.range("H11").value = currencyY[9]
        sht1.range("I11").value = c2Y[9]
        sht1.range("J11").value = c3Y[9]
        sht1.range("k11").value = bidY[9]
        sht1.range("L11").value = askY[9]
        sht1.range("M11").value = c5Y[9]
        wb.save()

except:
    print("end")
  
