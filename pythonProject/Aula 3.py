import rpa as r
import pyautogui as p
import pandas as pd
import os
import pygetwindow as pw

r.init(visual_automation=True, chrome_browser=True)

r.url('https://br.advfn.com/monitor/')
p.sleep(4)
janela = pw.getActiveWindow()
janela.maximize()
p.sleep(2)
# r.click('xpath=//*[@id="afnmainbodid"]/div[1]/div/div[5]/a[1]')