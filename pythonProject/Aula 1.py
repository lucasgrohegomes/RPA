# Bibliotecas.
import pyautogui as p
import rpa as r
import os
import pyperclip
import pygetwindow as gw


# Abre o Notepad.
os.startfile(r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Notepad.lnk')
p.sleep(5)


# Liga o robô.
r.init(chrome_browser = False, visual_automation = True)


# Detecta se o Notepad está aberto.
if r.present('notepadAberto.png')== True:
    print('Notepad está aberto.')
else:
    r.wait(r.present('notepadAberto.png'))
    print('Aguardando Notepad...')


# Maximiza a tela do Notepad.
janela = gw.getActiveWindow()
janela.maximize()


# Copia e cola um texto no Notepad.
pyperclip.copy('Robô funcionando!')
p.sleep(1)
p.hotkey('ctrl','v')


# Desliga o Robô.
r.close()


# option = p.confirm('Olá, por favor, escolha o programa desejado:', buttons=['FILIAIS-POR-REGIONAL', 'WORD', 'NOTEPAD'])
#
# if option == 'EXCEL':
#     print('Você escolheu Excel')
#     os.startfile(r'C:\Users\lucas.gomes\Desktop\Filiais_Regionais.xlsx')
