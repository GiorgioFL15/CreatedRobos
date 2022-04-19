#Importando as bibliotecas do selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#Importando as bibliotecas do pyautogui
import pyautogui as tempoEspera
import pyautogui as teclasTeclado
#By para trabalhar com os computadores mais recentes
from selenium.webdriver.common.by import By
#Importando a biblioteca openpyxl para trabalhar com o Excel
from openpyxl import load_workbook
import subprocess, sys

#Pegamos o caminho do arquivo + nome do arquivo que está no computador
nome_arquivo_contatos = "/home/fontes-1064/workspace/criandorobos/ListaDeDolar/WhatsApp.xlsx"
planilhaDadosContato = load_workbook(nome_arquivo_contatos)

#Selecionamos a sheet de Dados
sheet_selecionada = planilhaDadosContato['Dados']

#Emulando o navegador do Chrome
navegadorChrome = webdriver.Chrome()

#Passando e abrindo a pagina da web que devemos abrir
navegadorChrome.get('https://web.whatsapp.com/')

while len(navegadorChrome.find_elements(By.ID, 'side')) < 1: 
    
    tempoEspera.sleep(3)

for linha in range(2, len(sheet_selecionada['A']) + 1):

    nomeContato = sheet_selecionada['A%s' % linha].value
    mensagemContato = sheet_selecionada['B%s' % linha].value


    navegadorChrome.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/label/div/div[2]').send_keys(nomeContato)

    tempoEspera.sleep(2)

    teclasTeclado.press('enter')

    navegadorChrome.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(mensagemContato)

    tempoEspera.sleep(1)

    teclasTeclado.press('enter')

    tempoEspera.sleep(2)