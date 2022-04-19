from selenium import webdriver as opcoes_selenium_aula
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoPausaComputador
import xlsxwriter
import os 
from selenium.webdriver.common.by import By

meuNavegador = opcoes_selenium_aula.Chrome()
meuNavegador.get("https://www.google.com.br/")

tempoPausaComputador.sleep(2)

meuNavegador.find_element(By.NAME, "q").send_keys("Dolar hoje")

tempoPausaComputador.sleep(2)

meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(2)

valorDolarPeloGoogle = meuNavegador.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

tempoPausaComputador.sleep(2)

meuNavegador.find_element(By.NAME, "q").send_keys("")

tempoPausaComputador.sleep(2)

tempoPausaComputador.press('tab')

tempoPausaComputador.sleep(2)

tempoPausaComputador.press('enter')

meuNavegador.find_element(By.NAME, "q").send_keys("Euro hoje")

tempoPausaComputador.sleep(2)

meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(2)

valorEuroPeloGoogle = meuNavegador.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

nomeCaminhoArquivo = '/home/fontes-1064/workspace/criandorobos/ListaDeDolar/Excel com Euro e Dolar.xlsx'
planilhaCriada = xlsxwriter.Workbook(nomeCaminhoArquivo)
sheet1 = planilhaCriada.add_worksheet()

tempoPausaComputador.sleep(2)

sheet1.write("A1", "Dolar")
sheet1.write("B1", "Euro")
sheet1.write("A2", valorDolarPeloGoogle)
sheet1.write("B2", valorEuroPeloGoogle)

tempoPausaComputador.sleep(2)

planilhaCriada.close()

os.startfile(nomeCaminhoArquivo)

print("Dolar e Euro extraido com sucesso")