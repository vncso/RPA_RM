import pyautogui as p
import os
import datetime
from dotenv import load_dotenv
from main import planilha_net, muda_modulo, rateio_ti, \
    imprimir_planilha, seleciona_area_impressao, gera_rateio_definitivo
import pandas as pd

p.sleep(7)

dados = pd.read_excel('dados.xlsx', sheet_name='TOTVS')
inicial = 1
muda_modulo('backoffice', 'nucleus')

# Selecionar menu "customização"
try:
    customizacao = p.locateOnScreen('imgs/rateios/customizacao.png', confidence=0.9)
    if customizacao is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=811, y=41, duration=0.2)
    p.click()
    p.sleep(0.2)
else:
    p.moveTo(p.center(customizacao), duration=0.2)
    p.click()
    p.sleep(0.2)

print("customização - indo para sistema de rateio")
p.sleep(0.2)
# Selecionar sistema de rateio
try:
    sis_rateio = p.locateOnScreen('imgs/rateios/sistema_rateio.png', confidence=0.9)
    if sis_rateio is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=443, y=87, duration=0.2)
    p.click()
    p.sleep(0.3)
else:
    p.moveTo(p.center(sis_rateio), duration=0.2)
    p.click()
    p.sleep(0.3)
print("sistema de rateio aberto")
p.sleep(0.2)
# Selecionar filtro "Todos Global"
p.press('down')
p.press('enter')
"""
# p.moveTo(x=518, y=196, duration=0.2)
# p.click()
# Executar filtro
p.moveTo(x=769, y=600, duration=0.2)
p.click()
p.sleep(0.2)
"""
p.sleep(1.5)

# Selecionar caixa de pesquisa
try:
    busca = p.locateOnScreen('imgs/procurar.png', confidence=0.9)
    if busca is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=239, y=184, duration=0.2)
    p.click()
    p.sleep(0.4)
else:
    p.moveTo(p.center(busca), duration=0.2)
    p.click()
    p.sleep(0.4)

for i in range(0, 4):
    rateio = str(dados['ID'][i])
    coligada = str(dados['COLIGADA'][i])
    mes = str(dados['MES'][i]) if int(dados['MES'][i]) >= 10 else '0' + str(dados['MES'][i])
    ano = str(dados['ANO'][i])
    valor = str(dados['VALOR'][i])
    aditivo = str(dados['ADITIVO'][i])
    acordo = str(dados['ACORDO'][i])

    # pesquisar rateio
    p.moveTo(x=308, y=213, duration=0.2)
    p.doubleClick()
    p.write(rateio)
    p.sleep(0.4)
    # Ocultar não filtrados
    if inicial == 1:
        try:
            ocultar = p.locateOnScreen('imgs/ocultar_regs.png', confidence=0.9)
            if ocultar is None:
                raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
        except:
            p.moveTo(x=710, y=201, duration=0.2)
            p.click()
            p.sleep(0.4)
        else:
            p.moveTo(p.center(ocultar), duration=0.2)
            p.click()
            p.sleep(0.4)
    else:
        pass
    # Selecionar e abrir rateio
    p.moveTo(x=212, y=329, duration=0.2)
    p.doubleClick()
    p.sleep(0.5)
    # Abrir parametros do rateio
    try:
        params_rateio = p.locateOnScreen('imgs/rateios/parametros_rateio.png', confidence=0.9)
        if params_rateio is None:
            raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
    except:
        p.moveTo(x=706, y=204, duration=0.2)
        p.click()
        p.sleep(0.2)
    else:
        p.moveTo(p.center(params_rateio), duration=0.2)
        p.click()
        p.sleep(0.3)
    # Incluir novo rateio
    try:
        novo = p.locateOnScreen('imgs/rateios/novo_param.png', confidence=0.9)
        if novo is None:
            raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
    except:
        p.moveTo(x=333, y=234, duration=0.2)
        p.click()
        p.sleep(0.2)
    else:
        p.moveTo(p.center(novo), duration=0.2)
        p.click()
        p.sleep(0.3)
    # preencher dados do rateio (coligada)
    p.write(coligada)
    p.press('tab')
    p.sleep(0.2)
    # prencher dados do rateio (mes)
    p.press('tab', presses=2)
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
    inicial = 0


planilha_net(368)

for i in range(0, 4):
    rateio = str(dados['ID'][i])
    coligada = str(dados['COLIGADA'][i])
    mes = str(dados['MES'][i]) if int(dados['MES'][i]) >= 10 else '0' + str(dados['MES'][i])
    ano = str(dados['ANO'][i])

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

# voltar para o indice de planilhas
try:
    indice = p.locateOnScreen('imgs/planilhas/aba_planilhas.png', confidence=0.9)
    if indice is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
        print('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=333, y=234, duration=0.2)
    p.click()
    p.sleep(0.2)
else:
    p.moveTo(p.center(indice), duration=0.2)
    p.click()
    p.sleep(0.3)
# p.moveTo(x=244, y=154, duration=0.2)
# p.click()
# p.sleep(1)

# pesquisar planilha
p.moveTo(x=250, y=212, duration=0.2)
p.click()
p.write('367')
p.sleep(1)
# Selecionar e abrir rateio
p.moveTo(x=139, y=276, duration=0.2)
p.doubleClick()
p.sleep(2)

for i in range(0, 4):
    rateio = str(dados['ID'][i])
    coligada = str(dados['COLIGADA'][i])
    mes = str(dados['MES'][i]) if int(dados['MES'][i]) >= 10 else '0' + str(dados['MES'][i])
    ano = str(dados['ANO'][i])
    rateio_ti(mes, ano, rateio)

    # Imprimir rateio
    seleciona_area_impressao()
    p.sleep(0.3)
    print('imprimindo ;)')
    imprimir_planilha()

    scroll_up = p.locateOnScreen('imgs/planilhas/scroll_up.png', confidence=0.9)
    p.moveTo(p.center(scroll_up), duration=0.2)
    p.click(clicks=85)
    p.sleep(0.3)
    p.press('down')

    definitivo = input('deseja gerar o rateio em definitivo? (s/n)')
    if definitivo.lower() == 's':
        gera_rateio_definitivo(['24', '25', '26', '27'], mes, ano)

    """
    p.moveTo(x=58, y=111, duration=0.3)
    p.click()
    p.sleep(1)
    p.moveTo(x=665, y=420, duration=0.3)
    p.click()
    p.sleep(2)
    p.moveTo(x=263, y=405, duration=0.3)
    p.click()
    """