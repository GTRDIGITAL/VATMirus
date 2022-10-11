import pyautogui
from win32com.client import DispatchEx
from time import sleep

excel = DispatchEx('Excel.Application')
wbP=excel.Workbooks.Open("C:/Users/Bogdan.Constantinesc/Documents/ba cainele.xlsx")
excel.Visible=True
wbP.Worksheets("Training").Select()
sleep(5)



pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'c')
pyautogui.press("alt")
pyautogui.press("e")
pyautogui.press("s")
pyautogui.press("v")
pyautogui.press("enter")
pyautogui.hotkey('ctrl','s')
pyautogui.hotkey('alt','f4')