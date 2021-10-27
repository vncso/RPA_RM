from main import muda_modulo, abrir_rm, planilha_net, \
    executar_planilha, rateio_transportes
import pyautogui as p
import pandas as pd

dados = pd.read_excel('dados.xlsx', sheet_name='Planilha1')

valor = str(dados['VALOR'].tolist()[0].replace('.', ','))
mes = str(dados['MES'].tolist()[0])
ano = str(dados['ANO'].tolist()[0])
coligada = str(dados['COLIGADA'].tolist()[0])
tipo = str(dados['TIPO'].tolist()[0])

p.sleep(2)
p.moveTo(x=157, y=750)
p.click()
p.sleep(2)
p.write(str(dados['VALOR'].tolist()[0]).replace('.', ','))
# planilha = 247
# modulo = 'labore'
# tipo = 'rh'
#
# p.sleep(5)
#
# muda_modulo(tipo, modulo)
#
# planilha_net(planilha)
#
# rateio_transportes(8, 2021, '3', 1, '38457,69')