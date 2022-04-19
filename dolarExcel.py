import xlsxwriter
import os 

nomeCaminhoArquivo = '/home/fontes-1064/workspace/criandorobos/ListaDeDolar/Excel com Euro e Dolar.xlsx'
planilhaCriada = xlsxwriter.Workbook(nomeCaminhoArquivo)

sheet1 = planilhaCriada.add_worksheet()

sheet1.write("A1", "Nome")
sheet1.write("B1", "Idade")
sheet1.write("A2", "Amanda")
sheet1.write("B2", 28)
sheet1.write("A3", "Roberto")
sheet1.write("B3", 25)

planilhaCriada.close()

os.startfile(nomeCaminhoArquivo)