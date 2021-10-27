import pyautogui as p
import os
import datetime
from dotenv import load_dotenv
from main import abrir_rm, muda_modulo
import pandas as pd

load_dotenv()

dados = pd.read_excel('rateios.xlsx', sheet_name='Planilha1')

for i in range(0, 1):
    senha = os.environ.get('senha_rm')
    usuario = 'vinicius.oliveira'
    tipo = 'backoffice'
    modulo = 'nucleus'
    rateio = str(dados['ID'][i])
    coligada = str(dados['COLIGADA'][i])
    mes = str(dados['MES'][i]) if int(dados['MES'][i]) >= 10 else '0'+str(dados['MES'][i])
    ano = str(dados['ANO'][i])
    valor = str(dados['VALOR'][i])
    aditivo = str(dados['ADITIVO'][i])
    acordo = str(dados['ACORDO'][i])

print(rateio, coligada, mes, ano, valor)

# p.sleep(3)
# print(p.position())
# p.sleep(10)
# # Abrir o RM
# abrir_rm(usuario, senha)
# p.sleep(15)

# ir para o modulo
muda_modulo(tipo, modulo)
p.sleep(2)

# Selecionar menu "customização"
p.moveTo(x=811, y=41, duration=0.2)
p.click()
print("customização - indo para sistema de rateio")
p.sleep(0.2)
# Selecionar sistema de rateio
p.moveTo(x=443, y=87, duration=0.2)
p.click()
print("sistema de rateio aberto")
p.sleep(0.2)
# Selecionar filtro "Todos Global"
p.moveTo(x=518, y=196, duration=0.2)
p.click()
# Executar filtro
p.moveTo(x=769, y=600, duration=0.2)
p.click()
p.sleep(0.2)
# Selecionar caixa de pesquisa
p.moveTo(x=239, y=184, duration=0.2)
p.click()
p.sleep(0.2)
# pesquisar rateio
p.write(rateio)
p.sleep(0.2)
# Ocultar não filtrados
p.moveTo(x=476, y=214, duration=0.2)
p.click()
p.sleep(0.2)
# Selecionar e abrir rateio
p.moveTo(x=212, y=329, duration=0.2)
p.doubleClick()
p.sleep(2)
# Abrir parametros do rateio
p.moveTo(x=710, y=201, duration=0.2)
p.click()
p.sleep(0.2)
# Incluir novo rateio
p.moveTo(x=333, y=234, duration=0.2)
p.click()
p.sleep(0.2)
# preencher dados do rateio (coligada)
p.write(coligada)
p.press('tab')
p.sleep(0.2)
# prencher dados do rateio (mes)
p.moveTo(x=340, y=294, duration=0.2)
p.click()
p.write(mes)
p.sleep(0.2)
# prencher dados do rateio (ano)
p.press('tab')
p.sleep(0.2)
p.write(ano)
p.press('tab')
p.sleep(0.2)
# prencher dados do rateio (valor)
p.write(valor)
p.press('tab')
p.sleep(0.2)
# prencher dados do rateio (aditivo)
p.write(aditivo)
p.press('tab')
p.sleep(0.2)
# prencher dados do rateio (acordo)
p.write(acordo)
p.sleep(1)
# Salvar dados(clicar em OK)
p.moveTo(x=825, y=608, duration=0.2)
p.click()
p.sleep(1)
# Salvar rateio (clicar em OK)
p.click()
p.sleep(1)
# Abrir planilha .NET (ir para "gestão")
p.moveTo(x=890, y=42, duration=0.2)
p.click()
p.sleep(2)
# Selecionar "planilhas"
p.moveTo(x=148, y=87, duration=0.2)
p.click()
p.sleep(1)
# Selecionar o filtro "Planilhas Net"
p.moveTo(x=503, y=176, duration=0.2)
p.click()
# Executar filtro
p.moveTo(x=767, y=604, duration=0.2)
p.click()
p.sleep(2)
# Selecionar caixa de pesquisa
p.moveTo(x=239, y=184, duration=0.2)
p.click()
p.sleep(1)
# pesquisar planilha
p.write('368')
p.sleep(1)
# Ocultar não filtrados
p.moveTo(x=476, y=214, duration=0.2)
p.click()
p.sleep(1)
# Selecionar e abrir rateio
p.moveTo(x=139, y=276, duration=0.2)
p.doubleClick()
p.sleep(2)
# Selecionar filtro planilha "mes"
p.moveTo(x=129, y=238, duration=0.2)
p.click()
p.sleep(0.2)
p.write(mes)
# preencher ano do rateio
p.press('down')
p.sleep(0.2)
p.write(ano)
p.sleep(0.2)
# preencher id do rateio
p.press('down')
p.sleep(0.2)
p.write(rateio)
p.sleep(0.2)
p.press('down')
p.sleep(0.2)
# Executar o rateio
p.moveTo(x=801, y=102, duration=0.2)
p.click()
p.sleep(0.2)
p.moveTo(x=834, y=289)
p.click()
p.moveTo(x=764, y=457, duration=0.2)
p.click()
p.sleep(2)

#voltar para o indice de planilhas
p.moveTo(x=244, y=154, duration=0.2)
p.click()
p.sleep(1)
# pesquisar planilha
p.moveTo(x=250, y=212, duration=0.2)
p.click()
p.write('367')
p.sleep(1)
# Selecionar e abrir rateio
p.moveTo(x=139, y=276, duration=0.2)
p.doubleClick()
p.sleep(2)

# Selecionar filtro planilha "mes"
p.moveTo(x=155, y=258, duration=0.2)
p.click()
p.sleep(0.2)
p.write(mes)
# preencher ano do rateio
p.press('down')
p.sleep(0.2)
p.write(ano)
p.sleep(0.2)
# preencher id do rateio
p.press('down')
p.sleep(0.2)
p.write(rateio)
p.sleep(0.2)
p.press('down')
p.sleep(0.2)
# Executar o rateio
p.moveTo(x=801, y=102, duration=0.2)
p.click()
p.sleep(0.2)
p.moveTo(x=834, y=289)
p.click()
p.moveTo(x=764, y=457, duration=0.2)
p.click()
p.sleep(2)

# Imprimir rateio
p.moveTo(x=58, y=111, duration=0.3)
p.click()
p.sleep(1)
p.moveTo(x=665, y=420, duration=0.3)
p.click()
p.sleep(2)
p.moveTo(x=263, y=405, duration=0.3)
p.click()
