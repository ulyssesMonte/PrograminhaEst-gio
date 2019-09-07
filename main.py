from openpyxl import * #biblioteca para ler
import xlsxwriter #biblioteca para escrever
ph1 = load_workbook('PREVISÕES HORÁRIO 1.xlsx')  #abrindo planilha inicial
ph2 = xlsxwriter.Workbook('PREVISÕES PROBLEMÁTICAS.xlsx') #criando planilha com previsões erradas
pps = ph2.add_worksheet() 
phs = ph1['Plan1']
x=0 #variável auxiliar
for line in phs:  #percorrer planilha e copiar tuplas em que não houve previsão para um ônibus que chegou
  if (line[2].value)==None:
    pps.write(x,0,line[0].value)
    pps.write(x,1,line[1].value)
    pps.write(x,2,line[2].value)
    pps.write(x,3,line[3].value)
    pps.write(x,4,line[4].value)
    pps.write(x,5,line[5].value)
    pps.write(x,6,"não houve previsão")
    x+=1
  if(line[5].value)==None: #copiar tuplas em que houve previsão mas não chegou ônibus
    pps.write(x,0,line[0].value)
    pps.write(x,1,line[1].value)
    pps.write(x,2,line[2].value)
    pps.write(x,3,line[3].value)
    pps.write(x,4,line[4].value)
    pps.write(x,5,line[5].value)
    pps.write(x,6,"ônibus não chegou")

print(x)
ph2.close()
