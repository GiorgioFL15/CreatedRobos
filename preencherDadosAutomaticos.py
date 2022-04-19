#pip install selenium
#pip install xlsxwriter
#pip install openpyxl
#instalar driver referente ao seu google chrome

from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import os 

nomeCaminhoArquivo = "/home/fontes-1064/workspace/criandorobos/ListaDeDolar"

planilha_aberta = load_workbook(filename=nomeCaminhoArquivo)

sheet_selecionada = planilha_aberta['Dados']

for linha in range(2, len(sheet_selecionada['A']) + 1):

    nome = sheet_selecionada['A%s' % linha].value
    email = sheet_selecionada['B%s' % linha].value
    telefone = sheet_selecionada['C%s' % linha].value
    sexo = sheet_selecionada['D%s' % linha].value
    sobre = sheet_selecionada['E%s' % linha].value

    tempoEspera.sleep(2)

    navegadorFormulario = opcoesSelenium.Chrome()
    navegadorFormulario.get("https://pt.surveymonkey.com/r/Y9Y6FFR")

    #Aguardar para o computador processar as informações
    tempoEspera.sleep(6)

    #Preenche Nome
    navegadorFormulario.find_element(By.NAME, "683928983").send_keys(nome)

    #Preenche Email
    navegadorFormulario.find_element(By.NAME, "683932318").send_keys(email)

    #Preenche Telefone
    navegadorFormulario.find_element(By.NAME, "683930688").send_keys(telefone)

    #Preenche Sobre
    navegadorFormulario.find_element(By.NAME, "683932969").send_keys(sobre)

    if sexo == "Masculino":

        #Preenche Radio Button Feminino
        navegadorFormulario.find_element(By.ID,"683931881_4497366118_label").click()

    else:
        #Preenche Radio Button Feminino
        navegadorFormulario.find_element(By.ID,"683931881_4497366119_label").click()

    #Aguardar para o computador processar as informações
    tempoEspera.sleep(6)

    #Clica para enviar as informações
    navegadorFormulario.find_element(By.XPATH,'//*[@id="patas"]/main/article/section/form/div[2]/button').click()
