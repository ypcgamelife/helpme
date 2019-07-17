import pyautogui
import time
import xlrd
import os
import os.path
import pyperclip

mypath=os.getcwd()
file = 'd:/helpme.xls'
wb = xlrd.open_workbook(filename=file)#打开文件
#print(wb.sheet_names())#获取所有表格名字
sheet1 = wb.sheet_by_index(0)#通过索引获取表格
bzrow=sheet1.nrows
#print(bzrow)
# bz = sheet1.cell(1, 0).value
# print(bz)
# pyautogui.hotkey('winleft', 'r')
# bz = sheet1.cell(2, 0).value
# print(bz)
# pyautogui.hotkey('winleft', 'r')
# bz = sheet1.cell(3, 0).value
# print(bz)
bzs=1
while bzs < bzrow:
    #print('bzs=',str(bzs))
    sheet1=wb.sheet_by_index(0)
    bz = sheet1.cell(bzs, 0).value
    x1 = sheet1.cell(bzs, 1).value
    y1 = sheet1.cell(bzs, 2).value
    mousebutton = sheet1.cell(bzs, 3).value
    content = sheet1.cell(bzs, 4).value
    key = sheet1.cell(bzs, 5).value
    hotkey1 = sheet1.cell(bzs, 6).value
    hotkey2 = sheet1.cell(bzs, 7).value
    hotkey3 = sheet1.cell(bzs, 8).value
    sleeptime = int(sheet1.cell(bzs, 9).value)
    # print(bz)
    # print(x1)
    # print(y1)
    # print(mousebutton)
    # print(content)
    # print(key)
    # print(hotkey1,hotkey2,hotkey3)
    # print(sleeptime)
    if bz<=0:
        print('结束')
        break
    pyautogui.moveTo(x1, y1, duration=0.25)
    if mousebutton>'0':
       if mousebutton=='doubleClick':
           pyautogui.doubleClick()
       elif mousebutton=='left' or mousebutton=='wright' or mousebutton=='middle':
            pyautogui.click(button=mousebutton)
       else:
           #滚动
           if mousebutton.find('+'):
               scrollvalue=int(mousebutton[7:])
           else:
               scrollvalue=0-int(mousebutton[7:])
           pyautogui.scroll(scrollvalue)
    time.sleep(1)
    #输入文字，不支持中文，只能复制，粘贴
    if content>'0':
        pyautogui.hotkey('ctrlleft', 'a')
        pyautogui.press('delete')
        pyperclip.copy(content)
        pyautogui.hotkey('ctrlleft','v')
        #pyautogui.typewrite(oauser, 0.25)
        time.sleep(1)
    #按键
    if key>'0':
        pyautogui.press(key)
        time.sleep(1)
    #热键
    if hotkey1>'0':
        if hotkey3>'0':
            pyautogui.hotkey(hotkey1, hotkey2,hotkey3)
        else:
            pyautogui.hotkey(hotkey1,hotkey2)
    bzs+=1
    time.sleep(sleeptime)

