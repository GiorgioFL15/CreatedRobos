#!pip install xlsxwriter
from openpyxl import Workbook
import xlsxwriter as opcoesDoXlsxWriter
import os

nomeCaminhoArquivo = "/home/fontes-1064/workspace/criandorobos/ListaDeDolar/PintaFundoEFonte.xlsx"
workbook = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)
sheetPadrao = workbook.add_worksheet("Dados")

corFundo = workbook.add_format({'fg_color' : 'yellow'})

corFonte = workbook.add_format()
corFonte.set_font_color('blue')

sheetPadrao.write("A1", "Nome", corFundo)
sheetPadrao.write("B1", "Idade", corFundo)
sheetPadrao.write("A2", "Amanda", corFonte)
sheetPadrao.write("B2", 21, corFonte)
sheetPadrao.write("A3", "Allan", corFonte)
sheetPadrao.write("B3", 28, corFonte)

workbook.close()

os.startfile(nomeCaminhoArquivo)