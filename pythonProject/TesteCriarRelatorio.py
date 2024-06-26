import pyautogui as p
from selenium.webdriver import Chrome
from openpyxl import Workbook
import warnings
import pyperclip
import rpa as r
import pandas as pd
from openpyxl import load_workbook
import os
import pygetwindow as gw
import selenium
import pyperclip as pc
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys

option = p.confirm('Olá, selecione a entidade do relatório selecionado:', buttons=['IEL','FIESC','SENAI','SESI'])

if option == 'IEL':
    os.startfile(r'\\sistemafiesc.org\NETLOGON\Atalhos\RemoteAPP\BENNER\Produção\CORPORATIVO_OFICIAL_NS06.rdp')
    r.init(visual_automation=True, chrome_browser=False)
    p.sleep(6)

    # O robô vai esperar que o app do Benner seja aberto, ele vai verificar se existe os campos de usuário e senha para preencher e depois ele irá clicar no botão ok.

    if r.present('bennerAberto.png')==True:
        if r.present('usuarioBenner.png')==True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl','v')
            if r.present('senhaBenner.png')==True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
    else:
        r.wait(r.present('bennerAberto.png'))
        if r.present('usuarioBenner.png')==True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl','v')
            if r.present('senhaBenner.png')==True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')

    if r.present('okBotao.png')==True:
        r.click('okBotao.png')
    else:
        r.wait(r.present('okBotao.png'))
        r.click('okBotao.png')
    p.sleep(15)

    janela = gw.getActiveWindow()
    janela.maximize()

    # Após isso, ele vai verificar qual entidade está sendo visualizada.
    if r.present('abaIEL.png')==True:
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    else:
        r.click('casaEntidade.png')
        r.click('diminuirLista.png')
        r.dclick('selecionarIEL.png')
        r.dclick('botaoTF.png')
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    p.sleep(2)
    r.dclick('grupoRelatorios.png')
    r.dclick('inventariosIEL.png')
    r.dclick('relatorios.png')
    r.click('bensPorLocalizacao.png')
    r.click('imprimirRelatorio.png')
    r.click('selecionarAT_ITENS.png')
    r.click('monitor.png')
    r.click('filtroIEL.png')
    r.click('filtrar.png')
    p.sleep(5)
    r.click('salvarRelatorio.png')
    r.click('esteComputador.png')
    r.dclick('unidadeC.png')
    r.dclick('pastaUsuarios.png')
    r.dclick('pastaLucasGomes.png')
    r.dclick('areaDeTrabalho.png')
    r.dclick('conferenciasIEL.png')
    nomearquivo = p.prompt('Qual nome você dará ao arquivo?')
    pyperclip.copy(str(nomearquivo))
    r.click('nomeDoArquivo.png')
    p.hotkey('ctrl', 'v')
    r.click('tipoDeArquivo.png')
    r.click('arquivoTipoExcel.png')
    r.click('salvarXLSX.png')
    r.close()

elif option == 'FIESC':
    os.startfile(r'\\sistemafiesc.org\NETLOGON\Atalhos\RemoteAPP\BENNER\Produção\CORPORATIVO_OFICIAL_NS06.rdp')
    r.init(visual_automation=True, chrome_browser=False)
    p.sleep(6)

    # O robô vai esperar que o app do Benner seja aberto, ele vai verificar se existe os campos de usuário e senha para preencher e depois ele irá clicar no botão ok.

    if r.present('bennerAberto.png') == True:
        if r.present('usuarioBenner.png') == True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
    else:
        r.wait(r.present('bennerAberto.png'))
        if r.present('usuarioBenner.png') == True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')

    if r.present('okBotao.png') == True:
        r.click('okBotao.png')
    else:
        r.wait(r.present('okBotao.png'))
        r.click('okBotao.png')
    p.sleep(15)

    janela = gw.getActiveWindow()
    janela.maximize()

    # Após isso, ele vai verificar qual entidade está sendo visualizada.
    if r.present('abaFIESC.png') == True:
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    else:
        r.click('casaEntidade.png')
        r.click('diminuirLista.png')
        r.dclick('selecionarFIESC.png')
        r.dclick('botaoTF.png')
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    p.sleep(2)
    r.dclick('grupoRelatorios.png')
    r.dclick('inventariosFIESC.png')
    r.dclick('relatorios.png')
    r.click('bensPorLocalizacao.png')
    r.click('imprimirRelatorio.png')
    r.click('selecionarAT_ITENS.png')
    r.click('monitor.png')
    r.click('filtrar.png')
    p.sleep(5)
    r.click('salvarRelatorio.png')
    r.click('esteComputador.png')
    r.dclick('unidadeC.png')
    r.dclick('pastaUsuarios.png')
    r.dclick('pastaLucasGomes.png')
    r.dclick('areaDeTrabalho.png')
    r.dclick('conferenciasFIESC.png')
    nomearquivo = p.prompt('Qual nome você dará ao arquivo?')
    pyperclip.copy(str(nomearquivo))
    r.click('nomeDoArquivo.png')
    p.hotkey('ctrl', 'v')
    r.click('tipoDeArquivo.png')
    r.click('arquivoTipoExcel.png')
    r.click('salvarXLSX.png')
    r.close()
elif option == 'SENAI':
    os.startfile(r'\\sistemafiesc.org\NETLOGON\Atalhos\RemoteAPP\BENNER\Produção\CORPORATIVO_OFICIAL_NS06.rdp')
    r.init(visual_automation=True, chrome_browser=False)
    p.sleep(6)

    # O robô vai esperar que o app do Benner seja aberto, ele vai verificar se existe os campos de usuário e senha para preencher e depois ele irá clicar no botão ok.

    if r.present('bennerAberto.png') == True:
        if r.present('usuarioBenner.png') == True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
    else:
        r.wait(r.present('bennerAberto.png'))
        if r.present('usuarioBenner.png') == True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')

    if r.present('okBotao.png') == True:
        r.click('okBotao.png')
    else:
        r.wait(r.present('okBotao.png'))
        r.click('okBotao.png')
    p.sleep(15)

    janela = gw.getActiveWindow()
    janela.maximize()

    # Após isso, ele vai verificar qual entidade está sendo visualizada.
    if r.present('aba.png') == True:
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    else:
        r.click('casaEntidade.png')
        r.click('diminuirLista.png')
        r.dclick('selecionarSENAI.png')
        r.dclick('botaoTF.png')
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    p.sleep(2)
    r.dclick('grupoRelatorios.png')
    r.dclick('inventariosSENAI.png')
    r.dclick('relatorios.png')
    r.click('bensPorLocalizacao.png')
    r.click('imprimirRelatorio.png')
    r.click('selecionarAT_ITENS.png')
    r.click('monitor.png')
    r.click('filtroSENAI.png')
    r.click('filtrar.png')
    p.sleep(5)
    r.click('salvarRelatorio.png')
    r.click('esteComputador.png')
    r.dclick('unidadeC.png')
    r.dclick('pastaUsuarios.png')
    r.dclick('pastaLucasGomes.png')
    r.dclick('areaDeTrabalho.png')
    r.dclick('conferenciasSENAI.png')
    nomearquivo = p.prompt('Qual nome você dará ao arquivo?')
    pyperclip.copy(str(nomearquivo))
    r.click('nomeDoArquivo.png')
    p.hotkey('ctrl', 'v')
    r.click('tipoDeArquivo.png')
    r.click('arquivoTipoExcel.png')
    r.click('salvarXLSX.png')
    r.close()
elif option == 'SESI':
    os.startfile(r'\\sistemafiesc.org\NETLOGON\Atalhos\RemoteAPP\BENNER\Produção\CORPORATIVO_OFICIAL_NS06.rdp')
    r.init(visual_automation=True, chrome_browser=False)
    p.sleep(6)

    # O robô vai esperar que o app do Benner seja aberto, ele vai verificar se existe os campos de usuário e senha para preencher e depois ele irá clicar no botão ok.

    if r.present('bennerAberto.png') == True:
        if r.present('usuarioBenner.png') == True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
    else:
        r.wait(r.present('bennerAberto.png'))
        if r.present('usuarioBenner.png') == True:
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
        else:
            r.wait(r.present('usuarioBenner.png'))
            r.click('usuarioBenner.png')
            pyperclip.copy('lucas.gomes')
            p.sleep(1)
            p.hotkey('ctrl', 'v')
            if r.present('senhaBenner.png') == True:
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')
            else:
                r.wait(r.present('senhaBenner.png'))
                r.click('senhaBenner.png')
                pyperclip.copy('Lucasgg17*')
                p.sleep(1)
                p.hotkey('ctrl', 'v')

    if r.present('okBotao.png') == True:
        r.click('okBotao.png')
    else:
        r.wait(r.present('okBotao.png'))
        r.click('okBotao.png')
    p.sleep(15)

    janela = gw.getActiveWindow()
    janela.maximize()

    # Após isso, ele vai verificar qual entidade está sendo visualizada.
    if r.present('aba.png') == True:
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    else:
        r.click('casaEntidade.png')
        r.click('diminuirLista.png')
        r.dclick('selecionarSESI.png')
        r.dclick('botaoTF.png')
        if r.present('ativoDesselecionado.png') == True:
            r.click('ativoDesselecionado.png')
    p.sleep(2)
    r.dclick('grupoRelatorios.png')
    r.dclick('inventariosSESI.png')
    r.dclick('relatorios.png')
    r.click('bensPorLocalizacao.png')
    r.click('imprimirRelatorio.png')
    r.click('selecionarAT_ITENS.png')
    r.click('monitor.png')
    r.click('filtroSESI.png')
    r.click('filtrar.png')
    p.sleep(5)
    r.click('salvarRelatorio.png')
    r.click('esteComputador.png')
    r.dclick('unidadeC.png')
    r.dclick('pastaUsuarios.png')
    r.dclick('pastaLucasGomes.png')
    r.dclick('areaDeTrabalho.png')
    r.dclick('conferenciasSESI.png')
    nomearquivo = p.prompt('Qual nome você dará ao arquivo?')
    pyperclip.copy(str(nomearquivo))
    r.click('nomeDoArquivo.png')
    p.hotkey('ctrl', 'v')
    r.click('tipoDeArquivo.png')
    r.click('arquivoTipoExcel.png')
    r.click('salvarXLSX.png')
    r.close()
else:
    p.alert('Processo cancelado.')


