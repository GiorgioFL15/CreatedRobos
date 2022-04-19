from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoEspera

from selenium.webdriver.common.by import By

#Abre o navegador do Google Chrome e abre o site
navegador = opcoesSelenium.Chrome()
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

#Imprime o CEP no campo do CEP
navegador.find_element(By.NAME, "endereco").send_keys("05892387")

#Aguarda 2 segundos
tempoEspera.sleep(2)

#Clica no botão de Pesquisar
navegador.find_element(By.NAME, "btn_pesquisar").click()

#Aguarda 6 segundos
tempoEspera.sleep(6)

#Pega os dados da Rua no site
rua = navegador.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text
print("Rua: ", rua)

#Pega os dados do bairro no site
bairro = navegador.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text
print("Bairro: ", bairro)

#Pega os dados da Cidade no site
cidade = navegador.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text
print("Cidade: ", cidade)

#Pega os dados do CEP no site
cep = navegador.find_element(By.XPATH,'//*[@id="resultado-DNEC"]/tbody/tr/td[4]').text
print("CEP: ", cep)

from openpyxl import load_workbook
import os 

nome_arquivo = "/home/fontes-1064/workspace/criandorobos/ListaDeDolar/PesquisarEndereco.xlsx"
planilhaDadosEndereco = load_workbook(nome_arquivo)

sheet_selecionada = planilhaDadosEndereco['Dados']

linha = len(sheet_selecionada['A']) + 1

colunaA = "A" + str(linha)
colunaB = "B" + str(linha)
colunaC = "C" + str(linha)
colunaD = "D" + str(linha)

sheet_selecionada[colunaA] = rua
sheet_selecionada[colunaB] = bairro
sheet_selecionada[colunaC] = cidade
sheet_selecionada[colunaD] = cep

planilhaDadosEndereco.save(filename=nome_arquivo)
os.startfile(nome_arquivo)

#pip install selenium
#instalar o google chrome driver de acordo com a versão do seu browser
#pip install xlsxwriter
#utilizar o excel em Inglês
