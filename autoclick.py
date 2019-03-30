import pyautogui,time,sys,xlrd
def robcli(code):
    pyautogui.click(63, 263)
    time.sleep(0.2)
    pyautogui.click(910, 366)
    pyautogui.typewrite(code)
    time.sleep(0.2)
    pyautogui.click(1300, 360)
    time.sleep(1)
    pyautogui.click(1703, 600)
    time.sleep(5)
    pyautogui.click(1158, 252)
    time.sleep(0.2)
    pyautogui.click(964, 380)
    time.sleep(2)
    pyautogui.click(56, 269)
    time.sleep(2)
xlsdata = xlrd.open_workbook("new.xlsx")
sheet = xlsdata.sheet_by_index(0)
nrowsdata = sheet.col_values(0)
for i in nrowsdata:
    robcli(i)
    print("{}同步成功".format(i))

