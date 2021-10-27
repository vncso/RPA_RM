import pandas as pd
import pyautogui as p
from main import cadastra_user_portal, muda_modulo

usuarios_dados = pd.read_excel('integracao.xlsx', sheet_name='RELACAO INTEGRACAO', skiprows=5)

muda_modulo('global')
p.sleep(2)

# Abrir aba "Segurança" no RM:
try:
    seguranca = p.locateOnScreen('imgs/muda_modulo/globais/seguranca.png', confidence=0.9)
    if seguranca is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
        print('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=184, y=39, duration=0.2)
    p.click()
    p.sleep(1)
else:
    p.moveTo(p.center(seguranca), duration=0.2)
    p.click()
    p.sleep(1)

# Abrir "Usuários" no RM:
try:
    usuarios = p.locateOnScreen('imgs/muda_modulo/globais/usuarios.png', confidence=0.9)
    if usuarios is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
        print('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=184, y=84, duration=0.2)
    p.click()
    p.sleep(1)
else:
    p.moveTo(p.center(usuarios), duration=0.2)
    p.click()
    p.sleep(1)

# Selecionar o filtro "TODOS":
try:
    filtro_todos = p.locateOnScreen('imgs/muda_modulo/globais/filtro_todos.png', confidence=0.9)
    if filtro_todos is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
        print('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=493, y=213, duration=0.2)
    p.doubleClick()
    p.sleep(2)
else:
    p.moveTo(p.center(filtro_todos), duration=0.2)
    p.doubleClick()
    p.sleep(2)

# Mudar para o sistema de Folha de pagamento:
sistema = p.locateOnScreen('imgs/cria_user/muda_sistema.png', confidence=0.8)
p.moveTo(p.center(sistema), duration=0.2)
p.click()
p.sleep(0.5)

# Selecionar RH:
try:
    rh = p.locateOnScreen('imgs/cria_user/sistema_rh.png', confidence=0.9)
    if rh is None:
        raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
        print('Imagem não encontrada, tentando localizar por coordenada')
except:
    p.moveTo(x=12, y=206, duration=0.2)
    p.click()
    p.sleep(0.5)
else:
    p.moveTo(p.center(rh), duration=0.2)
    p.click()
    p.sleep(0.5)

for i in range(0, 6):  # in usuarios:
    nome = str(usuarios_dados['Nome'][i])
    cpf = str(usuarios_dados['CPF'][i])
    coligada = str(usuarios_dados['Col'][i])
    chapa = str(usuarios_dados['Chapa'][i])
    cadastra_user_portal(nome, cpf, coligada, chapa)
