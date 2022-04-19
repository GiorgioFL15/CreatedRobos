#pip install selenium
#pip install xlsxwriter
#pip install openpyxl
#instalar driver referente ao seu google chrome

from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui as tempoEspera

from openpyxl import load_workbook
import subprocess, sys
#import os

#Pegamos o caminho do arquivo no computador
nome_arquivo_tabela = "/home/fontes-1064/workspace/criandorobos/ListaDeDolar/testesRobos.xlsx"
planilhaDadosTabela = load_workbook(nome_arquivo_tabela)

#Selecionamos a sheet de dados
sheet_selecionada = planilhaDadosTabela['Dados']


#Abre o navegador da web Google Chrome
navegador = opcoesSelenium.Chrome()
navegador.get("https://rpachallengeocr.azurewebsites.net/")

linha = 1 

i = 1
while i < 4:
    #Selecionamos a sheet de dados
    sheet_dados = planilhaDadosTabela['Dados']

    # Copia o Xpath da tabela 
    elementoTabela = navegador.find_element(By.XPATH, '//*[@id="tableSandbox"]')

    #Pega linhas e colunas
    linhas = elementoTabela.find_elements(By.TAG_NAME, "tr")
    colunas = elementoTabela.find_elements(By.TAG_NAME, "td")

    for linhaAtual in linhas:

        linha = linha + 1
        #Pegamos a ultima linha +1 
        linha = len(sheet_dados['A']) + 1

        #Demos o nome da coluna + o numero da linha
        colunaA = "A" + str(linha)
        colunaB = "B" + str(linha)
        colunaC = "C" + str(linha)

        #Pegamos o texto da linha
        texto = linhaAtual.text
        #Separamos com o split todas as palavras com critério do espaço entre texto
        texto2 = texto.split(" ")

        #Imprimimos os dados da tabela no Excel
        sheet_dados[colunaA] = texto2[0]
        sheet_dados[colunaB] = texto2[1]
        sheet_dados[colunaC] = texto2[2]

    i += 1 
    #Aguarda 2 segundos para o computador ou site processar as informações
    tempoEspera.sleep(2)
    #Encontra o XPATH do botão Next e clica
    navegador.find_element(By.XPATH, '//*[@id="tableSandbox_next"]').click()
    #Aguarda 2 segundos para o computador ou site processar as informações
    tempoEspera.sleep(2)
else:

    print("Pronto!")

#salva o arquivo com as alterações
    planilhaDadosTabela.save(filename=nome_arquivo_tabela)

def open_file(filename):    
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])
