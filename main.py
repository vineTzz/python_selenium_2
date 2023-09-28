from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import pandas as pd

# Criar Navegador e Importar Base de Dados #
navegador = webdriver.Chrome()
tabela = pd.read_excel(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + selenium\consulta_juridica\Processos.xlsx')

# abrir a lista de cidades #
for linha in tabela.index:

    navegador.get(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + selenium\consulta_juridica\index.html')

    menu = navegador.find_element(By.CLASS_NAME, 'dropdown-menu')
    ActionChains(navegador).move_to_element(menu).perform()
    cidade = tabela.loc[linha, 'Cidade']
    navegador.find_element(By.PARTIAL_LINK_TEXT, cidade).click()

# mudar para a nova aba
    aba_index = navegador.window_handles[0]
    indice = linha + 1
    aba_formulario = navegador.window_handles[indice]
    navegador.switch_to.window(aba_formulario)

# Preencher formulario #
    navegador.find_element(By.ID, 'nome').send_keys(tabela.loc[linha, 'Nome'])
    navegador.find_element(By.ID, 'advogado').send_keys(tabela.loc[linha, 'Advogado'])
    navegador.find_element(By.ID, 'numero').send_keys(tabela.loc[linha, 'Processo'])
    navegador.find_element(By.XPATH, '/html/body/div/form/div/button').click()

# Confirmar Alerta de pesquisa #
    alerta = navegador.switch_to.alert
    alerta.accept()

# Esperar o Resultado da Pesquisa e Agir de Acordo com o Resultado #
    while True:
        try:
            alerta = navegador.switch_to.alert
            break
        except:
            sleep(1)
    texto = alerta.text
    if 'Processo encontrado com sucesso' in texto:
        alerta.accept()
        tabela.loc[linha, 'Status'] = 'Encontrado'
    else:
        alerta.accept()
        tabela.loc[linha, 'Status'] = 'Não Encotrado'


# Criar um arquivo com o Resultado #
    tabela.to_excel(r'C:\Users\alxia\OneDrive\Área de Trabalho\github\python + selenium\consulta_juridica\consultas_juridicas.xlsx')
