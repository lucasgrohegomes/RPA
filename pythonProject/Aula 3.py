import pyperclip
import rpa as r
import pyautogui as p
import pandas as pd
import os
import pygetwindow as pw
import selenium
import pyperclip as pc
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys

r.init(visual_automation=True, chrome_browser=True)

r.url('https://br.advfn.com/p.php?pid=front&logout=1')
p.sleep(20)
janela = pw.getActiveWindow()
janela.maximize()
p.sleep(3)
r.click('/html/body/div/div[2]/div[2]/div/form/div[1]/input')
p.sleep(2)
p.hotkey('ctrl','a')
p.sleep(1)
p.hotkey('backspace')
p.sleep(20)
r.type('//*[@id="afnmainbodid"]/div/div[2]/div[2]/div/form/div[1]/input', 'lucasgomes2')
r.type('//*[@id="password-input"]', 'Lucasgg10*[enter]')
p.sleep(3)
r.close()