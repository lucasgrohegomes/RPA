import pyautogui as p
import pandas as pd
from selenium.webdriver import Chrome
import warnings

# IMPORTANDO OS DADOS

with warnings.catch_warnings(record=True):
    warnings.simplefilter('always')
    dados = pd.read_excel(r'C:\Users\lucas.gomes\Desktop\dadosformulario.xlsx', sheet_name='dados')

# TRANSFORMAR EM DATA FRAME

df = pd.DataFrame(dados, columns=['NOME', 'EMAIL', 'TELEFONE', 'SEXO', 'PROFISSAO'])

contagem = len(df)
linha = 1

for row in df.itertuples():
    print(linha, '/', contagem)
    p.sleep(3)
    navegador = Chrome()
    navegador.get('https://pt.surveymonkey.com/r/Y9Y6FFR')
    navegador.maximize_window()
    p.sleep(3)

    # NOME
    navegador.find_element('xpath', '//*[@id="683928983"]').send_keys(row[1])
    p.sleep(1)

    # EMAIL
    navegador.find_element('xpath','//*[@id="683932318"]').send_keys(row[2])
    p.sleep(1)

    # TELEFONE
    navegador.find_element('xpath', '//*[@id="683930688"]').send_keys(row[3])
    p.sleep(1)

    # SEXO
    if row[4] == 'Feminino':
        navegador.find_element('xpath','//*[@id="683931881_4497366119_label"]/span[1]').click()
    else:
        navegador.find_element('xpath','//*[@id="683931881_4497366118_label"]/span[1]').click()
    p.sleep(1)

    # SOBRE
    navegador.find_element('xpath','//*[@id="683932969"]').send_keys(row[5])
    p.sleep(1)

    # ENVIAR DADOS
    navegador.find_element('xpath','//*[@id="patas"]/main/article/section/form/div[2]/button').click()
    p.sleep(2)

    print('Indo para a próxima linha')
    navegador.close()
    linha = linha + 1

print('Acabaram as linhas. Processo concluído com sucesso.')
