# Bibliotecas.
import pyautogui as p
import rpa as r
import os
import pyperclip
import pygetwindow as gw
#
#
# # Abre o Notepad.
# os.startfile(r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Notepad.lnk')
# p.sleep(5)
#
#
# # Liga o robô.
# r.init(chrome_browser = False, visual_automation = True)
#
#
# # Detecta se o Notepad está aberto.
# if r.present('notepadAberto.png')== True:
#     print('Notepad está aberto.')
# else:
#     r.wait(r.present('notepadAberto.png'))
#     print('Aguardando Notepad...')
#
#
# # Maximiza a tela do Notepad.
# janela = gw.getActiveWindow()
# janela.maximize()
#
#
# # Copia e cola um texto no Notepad.
# pyperclip.copy('Robô funcionando!')
# p.sleep(1)
# p.hotkey('ctrl','v')
#
#
# # Desliga o Robô.
# r.close()


option = p.confirm('Olá, por favor, escolha o programa desejado:', buttons=['EXCEL', 'WORD', 'NOTEPAD'])

if option == 'EXCEL':
    print('Você escolheu o Excel.')
    os.startfile(r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk')
    p.sleep(7.5)
    janela = gw.getWindowsWithTitle('Excel')
    janela.maximize()

    r.init(chrome_browser=False, visual_automation=True)

    if r.present('excelAberto.png'):

        p.hotkey('enter')

    else:
        r.wait(r.present('excelAberto.png'))
        print('Aguardando Excel.')
    r.close()

elif option == 'WORD':
    print('Você escolheu o Word.')
    os.startfile(r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk')
    p.sleep(7.5)
    janela = gw.getWindowsWithTitle('Word')
    janela.maximize()

elif option == 'NOTEPAD':
    print('Você escolheu o NotePad')
    os.startfile(r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Notepad.lnk')
    p.sleep(7.5)
    