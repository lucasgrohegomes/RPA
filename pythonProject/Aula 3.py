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
r.type('//*[@id="password-input"]', 'Lucasgg10*[enter]')
p.sleep(4)
p.click(x=1495, y=536)
p.click(x=1462, y=640)
p.sleep(4)

r.table('//*[@id="monitorApp_monGrid_grid_table"]', 'cotacoes.csv')
dados = pd.read_csv('cotacoes.csv')
dados.to_csv(r'bolsa.csv', index=False, mode='a', header=True)
csv_xlsx = pd.read_csv(r'bolsa.csv')
csv_xlsx.to_excel(r'bolsa.xlsx')
p.sleep(4)
r.close()

os.startfile(r'C:\Users\lucas.gomes\Downloads\PESSOAL\RPA\pythonProject\bolsa.xlsx')