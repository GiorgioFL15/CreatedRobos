#Mercado livre
#pip install selenium
#pip install xlsxwriter
#pip install openpyxl
#instalar driver referente ao seu google chrome

from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui as tempoEspera
import subprocess, sys
import os
from openpyxl import load_workbook

#Pegamos o caminho do arquivo no computador
nome_arquivo_tabela = "/home/fontes-1064/workspace/criandorobos/ListaDeDolar/testePesquisaInformacoesRobos.xlsx"
planilhaDadosTabela = load_workbook(nome_arquivo_tabela)

#Selecionamos a sheet de dados
sheet_selecionada = planilhaDadosTabela['Dados']

#Abre o navegador da Web do Google
navegador = opcoesSelenium.Chrome()
navegador.get("https://www.mercadolivre.com.br/")

#Procura o campo Name, digita a palavra que queremos procurar e cola
navegador.find_element(By.NAME, 'as_word').send_keys('carteira')

tempoEspera.sleep(2)

#Procura o botão com o Xpath e clica no botão para pesquisar
navegador.find_element(By.XPATH, '/html/body/header/div/form/button').click()

tempoEspera.sleep(3)

dadosProduto = navegador.find_elements(By.CLASS_NAME, 'ui-search-layout__item')

linha = 2 
urlProduto = ""

for informacoes in dadosProduto:

    sheet_dados = planilhaDadosTabela['Dados']

    nomeProduto = informacoes.find_element(By.CLASS_NAME, 'ui-search-item__title').text
    precoProduto = informacoes.find_element(By.CLASS_NAME, 'price-tag-fraction').text
    try:
        centavosProduto = informacoes.find_element(By.CLASS_NAME, 'price-tag-cents').text
    except:
        centavosProduto = "0"
    
    urlProduto = informacoes.find_element(By.TAG_NAME, 'a').get_attribute('href')

    #print(nomeProduto + "-" + precoProduto + "," + centavosProduto + "-" + urlProduto)

    #Pegamos a ultima linha +1 
    linha = len(sheet_dados['A']) + 1

    #Demos o nome da coluna + o numero da linha
    colunaA = "A" + str(linha)
    colunaB = "B" + str(linha)
    colunaC = "C" + str(linha)

    #Imprimimos os dados da tabela no Excel
    sheet_dados['A1'] = "Produto"
    sheet_dados['B1'] = "Preço"
    sheet_dados['C1'] = "Imagem"

    precoTexto = precoProduto + "," + centavosProduto

    precoSemPonto = precoTexto.replace('.', '')
    precoSemPonto2 = precoSemPonto.replace(',', '.')

    precoSemPonto2 =  float(precoSemPonto2)

    #Imprimimos os dados da tabela no Excel
    sheet_dados[colunaA] = nomeProduto
    sheet_dados[colunaB] = precoProduto + "-" + centavosProduto
    sheet_dados[colunaC] = urlProduto

#salva o arquivo com as alterações
planilhaDadosTabela.save(filename=nome_arquivo_tabela)

def open_file(filename):    
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])
