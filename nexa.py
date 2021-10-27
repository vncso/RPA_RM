import pyautogui as p
import pandas as pd
from main import imprimir_planilha, seleciona_area_impressao, gera_pdf_rateio, muda_modulo, envia_email

dados = pd.read_excel('dados.xlsx', sheet_name='NEXA')

mes = str(dados['MES'][0]) if int(dados['MES'][0]) >= 10 else '0' + str(dados['MES'][0])
ano = str(dados['ANO'][0])
inicio = 1

print(mes, ano)

p.sleep(7)
muda_modulo('backoffice', 'nucleus')
p.sleep(1)
# Abrir planilha .NET (ir para "gestão")
gestao = p.locateOnScreen('imgs/gestao.png', confidence=0.9)
p.moveTo(p.center(gestao), duration=0.2)
p.click()
print("Abrindo 'Gestão' em Nucleus")
p.sleep(3)
# Selecionar "planilhas"
planilhas = p.locateOnScreen('imgs/planilhas/planilha_net.png', confidence=0.9)
p.moveTo(p.center(planilhas), duration=0.2)
p.click()
p.sleep(3)
print("Abrindo Planilhas .NET")
p.sleep(1)
# Selecionar o filtro "Planilhas Net"
p.moveTo(x=503, y=176, duration=0.2)
p.click()
# Executar filtro
p.moveTo(x=767, y=604, duration=0.2)
p.click()
print("Planilhas .NET abertas")
p.sleep(3)
for rateio in range(11, 12):
    # Selecionar caixa de pesquisa
    p.moveTo(x=239, y=184, duration=0.2)
    p.click()
    print(f"pesquisando pela planilha 'grava rateio'...")
    p.sleep(1)
    # pesquisar planilha
    p.moveTo(x=249, y=214, duration=0.2)
    p.click()
    p.write('368')
    p.sleep(1)
    # Ocultar não filtrados
    if inicio == 1:
        p.moveTo(x=476, y=214, duration=0.2)
        p.click()
        p.sleep(1)
    # Selecionar e abrir rateio
    p.moveTo(x=139, y=276, duration=0.2)
    p.doubleClick()
    print("Abrindo planilha 'grava rateio'")
    p.sleep(3)
    # Selecionar filtro planilha "mes"
    p.moveTo(x=129, y=238, duration=0.2)
    p.click()
    p.sleep(0.2)
    print(f"Preenchendo dados do rateio de id '{rateio}'...")
    p.write(mes)
    # preencher ano do rateio
    p.press('down')
    p.sleep(0.2)
    p.write(ano)
    p.sleep(0.2)
    # preencher id do rateio
    p.press('down')
    p.sleep(0.2)
    p.write(str(rateio))
    p.sleep(0.2)
    p.press('down')
    p.sleep(0.2)
    # Executar o rateio
    p.moveTo(x=801, y=102, duration=0.2)
    p.click()
    print(f"Iniciando execução do rateio de id {rateio} para gravar dados")
    p.sleep(0.2)
    p.moveTo(x=834, y=289)
    p.click()
    p.moveTo(x=764, y=457, duration=0.2)
    p.click()
    p.sleep(5)

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

    # voltar para o indice de planilhas
    # p.moveTo(x=244, y=154, duration=0.2)
    # p.click()
    # print("Retornando ao menu de planilhas .NET")
    # p.sleep(1)
    # pesquisar planilha
    p.moveTo(x=250, y=212, duration=0.2)
    p.click()
    print("Pesquisando planilha 'sistema rateio'...")
    p.write('367')
    p.sleep(1)
    # Selecionar e abrir rateio
    p.moveTo(x=139, y=276, duration=0.2)
    p.doubleClick()
    print("Abrindo planilha 'sistema rateio'...")
    p.sleep(5)

    # Selecionar filtro planilha "mes"
    p.moveTo(x=155, y=258, duration=0.5)
    p.click()
    print(f"preenchendo dados do rateio de id '{rateio}'")
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
    p.write(str(rateio))
    p.sleep(0.2)
    p.press('down')
    p.sleep(0.2)
    # Executar o rateio
    p.moveTo(x=801, y=102, duration=0.2)
    p.click()
    print(f"gerando rateio de id '{rateio}'")
    p.sleep(0.2)
    p.moveTo(x=834, y=289)
    p.click()
    p.moveTo(x=764, y=457, duration=0.2)
    p.click()
    p.sleep(5)
    # if rateio > 12:

    seleciona_area_impressao()

    imprimir_planilha(pdf='sim')

    gera_pdf_rateio()

    envia_email(11, mes, ano)
        # Imprimir rateio
        # p.moveTo(x=58, y=111, duration=0.2)
        # p.click()
        # print(f"Imprimindo rateio de id '{rateio}' na impressora padrão..")
        # p.sleep(1)
        # p.moveTo(x=665, y=420, duration=0.2)
        # p.click()
        # p.sleep(2)
        # p.moveTo(x=395, y=192, duration=0.2)
        # p.click()
        # p.sleep(0.2)
        # p.moveTo(x=128, y=122, duration=0.2)
        # p.click()
        # p.sleep(0.2)
        # # rateio impresso em orientação paisagem
        # p.moveTo(x=217, y=222, duration=0.2)
        # p.click()
        # p.moveTo(x=533, y=688, duration=0.2)
        # p.click()
        # p.moveTo(x=263, y=405, duration=0.2)
        # p.sleep(0.2)
        # p.click()
        # p.sleep(1)

    # Fechar planilhas de rateio
    print("fechando planilhas em aberto..")
    # p.moveTo(x=672, y=155, duration=0.2)
    # p.click()
    # p.sleep(0.2)
    # p.moveTo(x=849, y=439, duration=0.2)
    # p.click()
    # p.sleep(0.2)
    # p.moveTo(x=486, y=155, duration=0.2)
    # p.click()
    # p.sleep(0.2)
    # p.moveTo(x=849, y=439, duration=0.2)
    # p.click()

    inicio = 0