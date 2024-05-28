# Bibliotecas.
import pyautogui as p
import rpa as r
import os
import pyperclip as pc
import pygetwindow as gw
import pandas as pan

os.startfile(r'C:\Program Files\Google\Chrome\Application\chrome.exe')
p.sleep(2)

r.init(chrome_browser=False, visual_automation=True)

r.close()
