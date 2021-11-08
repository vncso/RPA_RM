import pyautogui as p
import PyPDF2
from os.path import exists
rateios = [
    {
        'id': 11,
        'empresa': 'Nexa'
    },
    {
        'id': 12,
        'empresa': 'Totalware'
    },
    {
        'id': 13,
        'empresa': 'Vivo movel'
    },
    {
        'id': 14,
        'empresa': 'Vivo Office 365 - Business Essentials'
    },
    {
        'id': 15,
        'empresa': 'Vivo Office 365 - Business Premium'
    },
    {
        'id': 16,
        'empresa': 'Vivo Office 365 - Enterprise E3'
    },
    {
        'id': 17,
        'empresa': 'Vivo Office 365 - Sec Productive Enterprise E5'
    },
    {
        'id': 18,
        'empresa': 'Vivo Office 365 - Project Online Professional'
    },
    {
        'id': 19,
        'empresa': 'Vivo Office 365 - Visio Online Plan 2'
    },
    {
        'id': 20,
        'empresa': 'Vivo Office 365 - Azure Information Protection Premium 1'
    },
    {
        'id': 21,
        'empresa': 'Oracle Licença Anual - Licenciamento'
    },
    {
        'id': 22,
        'empresa': 'Oracle Licença Anual - Suporte'
    },
    {
        'id': 23,
        'empresa': 'Sou Cloud - Azure'
    },
    {
        'id': 24,
        'empresa': 'TOTVS - 373445'
    },
    {
        'id': 25,
        'empresa': 'TOTVS - 351140'
    },
    {
        'id': 26,
        'empresa': 'TOTVS - 384130'
    },
    {
        'id': 27,
        'empresa': 'TOTVS - 353463'
    },
    {
        'id': 28,
        'empresa': 'Global Rastreadores - Baldin'
    },
    {
        'id': 29,
        'empresa': 'Global Rastreadores - Agrícola'
    },
    {
        'id': 30,
        'empresa': 'VIVO Tech - 21528943'
    },
    {
        'id': 31,
        'empresa': 'VIVO Tech - 21509058'
    },
    {
        'id': 32,
        'empresa': 'VIVO Tech - 21536610'
    },
    {
        'id': 33,
        'empresa': 'VIVO Tech - 21516003'
    },
    {
        'id': 34,
        'empresa': 'VIVO Tech - 21555025'
    },
    {
        'id': 35,
        'empresa': 'VIVO Tech - 21575414'
    },
    {
        'id': 36,
        'empresa': 'VIVO Tech - 21580417'
    },
    {
        'id': 44,
        'empresa': 'VIVO Tech - 21583161'
    },
    {
        'id': 37,
        'empresa': 'Markanti - AGR'
    },
    {
        'id': 38,
        'empresa': 'Markanti - Baldin'
    },
    {
        'id': 39,
        'empresa': 'Rádios Mendonça - Baldin'
    },
    {
        'id': 40,
        'empresa': 'Rádios Mendonça - AGR'
    },
    {
        'id': 41,
        'empresa': 'Rádios Mendonça - Acordo Baldin'
    },
    {
        'id': 42,
        'empresa': 'Rádios Mendonça - Acordo AGR'
    },
    {
        'id': 43,
        'empresa': 'VIVO Fixo'
    },
    {
        'id': 45,
        'empresa': 'Oracle Cloud - Trimestral'
    }
]
vivo = [
    {
        'rateio': '14',
        'coligada': '1',
        'valor': '472,00',
        'aditivo': '00',
        'acordo': '00'
    },
    {
        'rateio': '15',
        'coligada': '1',
        'valor': '6887,50',
        'aditivo': '00',
        'acordo': '00'
    },
    {
        'rateio': '16',
        'coligada': '1',
        'valor': '1580,00',
        'aditivo': '00',
        'acordo': '00'
    },
    {
        'rateio': '17',
        'coligada': '1',
        'valor': '283,99',
        'aditivo': '00',
        'acordo': '00'
    },
    {
        'rateio': '18',
        'coligada': '1',
        'valor': '329,00',
        'aditivo': '00',
        'acordo': '00'
    },
    {
        'rateio': '19',
        'coligada': '1',
        'valor': '82,50',
        'aditivo': '00',
        'acordo': '00'
    },
    {
        'rateio': '20',
        'coligada': '1',
        'valor': '109,70',
        'aditivo': '00',
        'acordo': '00'
    },
]


def abrir_rm(usuario, senha):
    """
        Função "abrir_rm" recebe como parametros o usuário e a senha para acesso
        ao ERP.

        É uma função sem retorno, realizando apenas a tarefa de acessar o sistema.

        :param usuario: string, usuário que irá acessar o RM
        :param senha: string, senha do usuário para acesso ao RM

    """
    # Abrir o RM
    p.moveTo(x=116, y=629, duration=0.5)
    p.sleep(2)

    p.doubleClick()
    p.sleep(7)

    # digitar usuário:
    p.moveTo(x=634, y=237, duration=0.5)
    p.click()
    p.press('backspace', presses=30)
    p.write(usuario)

    # digitar senha:
    p.moveTo(x=527, y=288, duration=0.5)
    p.click()
    p.write(senha)
    p.sleep(1)

    p.moveTo(x=458, y=407)
    p.sleep(1)

    p.click()
    for s in range(15):
        print(s)
        p.sleep(1)


def muda_modulo(tipo, modulo='rm'):
    """

    Função para alternar entre os módulos do sistema (RH/Backoffice/Global).

    :param tipo: string, pode possuir os valores "rh", "backoffice" ou "global". Módulo a ser acessado
    :param modulo: string, nome do módulo que será acessado. No caso do tipo "global" o módulo não precisa ser informado
    :return: sem retorno
    """
    muda_icone = p.locateOnScreen('imgs/muda_modulo.png', confidence=0.9)
    p.moveTo(p.center(muda_icone), duration=0.2)
    p.click()
    print("menu para mudança de módulo")
    p.sleep(0.5)
    if tipo == 'backoffice':
        # Selecionar o módulo Backoffice
        backoffice = p.locateOnScreen('imgs/muda_modulo/backoffice/backoffice.png', confidence=0.9)
        p.moveTo(p.center(backoffice), duration=0.2)
        print("selecionando Backoffice")
        p.sleep(0.5)
        if modulo == 'nucleus':
            # Selecionar o módulo Nucleus
            nucleus = p.locateOnScreen('imgs/muda_modulo/backoffice/nucleus.png', confidence=0.9)
            p.moveTo(p.center(nucleus), duration=0.2)
            p.click()
            print("Nucleus selecionado")
            p.sleep(7)
        elif modulo == 'bonum':
            # Selecionar o módulo Nucleus
            nucleus = p.locateOnScreen('imgs/muda_modulo/backoffice/bonum.png', confidence=0.9)
            p.moveTo(p.center(nucleus), duration=0.2)
            p.click()
            print("Bonum selecionado")
            p.sleep(7)
        elif modulo == 'saldus':
            # Selecionar o módulo Nucleus
            nucleus = p.locateOnScreen('imgs/muda_modulo/backoffice/saldus.png', confidence=0.9)
            p.moveTo(p.center(nucleus), duration=0.2)
            p.click()
            print("Saldus selecionado")
            p.sleep(7)
        elif modulo == 'fluxus':
            # Selecionar o módulo Nucleus
            nucleus = p.locateOnScreen('imgs/muda_modulo/backoffice/fluxus.png', confidence=0.9)
            p.moveTo(p.center(nucleus), duration=0.2)
            p.click()
            print("Fluxus selecionado")
            p.sleep(7)
        else:
            print(f'O módulo {modulo} não existe ou foi digitado incorretamente')
            print('utilize "nucleus","fluxus", "saldus" ou "bonum"')
            exit()
    elif tipo == 'rh':
        rh = p.locateOnScreen('imgs/muda_modulo/rh/rh.png', confidence=0.9)
        p.moveTo(p.center(rh), duration=0.2)
        p.click()
        p.sleep(0.5)
        if modulo == 'labore':
            # Selecionar o módulo Nucleus
            labore = p.locateOnScreen('imgs/muda_modulo/rh/labore.png', confidence=0.9)
            p.moveTo(p.center(labore), duration=0.2)
            p.click()
            print("Labore selecionado")
            p.sleep(4)
        elif modulo == 'vitae':
            # Selecionar o módulo Nucleus
            vitae = p.locateOnScreen('imgs/muda_modulo/rh/vitae.png', confidence=0.9)
            p.moveTo(p.center(vitae), duration=0.2)
            p.click()
            print("Vitae selecionado")
            p.sleep(4)
        elif modulo == 'chronus':
            # Selecionar o módulo Nucleus
            chronus = p.locateOnScreen('imgs/muda_modulo/rh/chronus.png', confidence=0.9)
            p.moveTo(p.center(chronus), duration=0.2)
            p.click()
            print("Chronus selecionado")
            p.sleep(7)
        elif modulo == 'sst':
            # Selecionar o módulo Nucleus
            sst = p.locateOnScreen('imgs/muda_modulo/rh/sst.png', confidence=0.9)
            p.moveTo(p.center(sst), duration=0.2)
            p.click()
            print("SST selecionado")
            p.sleep(7)
        else:
            print(f'O módulo {modulo} não existe ou foi digitado incorretamente')
            print('utilize "labore","vitae", "sst" ou "chronus"')
            exit()
    elif tipo == 'global':
        servicos_globais = p.locateOnScreen('imgs/muda_modulo/globais/servicos_globais.png', confidence=0.9)
        p.moveTo(p.center(servicos_globais), duration=0.2)
        p.click()
        p.sleep(0.5)
    else:
        print(f'O módulo {tipo} não existe ou foi digitado incorretamente')
        print('utilize "rh" ou "backoffice"')
        exit()


def planilha_net(id):
    """
    Abre uma planilha .Net no RM para o usuário

    :param id: (str) identificador numérico da planilha que será executada
    :return: none
    """
    # Ir para a aba "gestão"
    gestao = p.locateOnScreen('imgs/gestao.png', confidence=0.9)
    p.moveTo(p.center(gestao), duration=0.2)
    p.click()
    p.sleep(0.5)
    # Abrir Planilhas .Net
    planilhas = p.locateOnScreen('imgs/planilhas/planilha_net.png', confidence=0.9)
    p.moveTo(p.center(planilhas), duration=0.2)
    p.click()
    p.sleep(0.5)
    # Selecionar o filtro "Planilhas Net"
    p.moveTo(x=503, y=176, duration=0.5)
    # filtro_planilhas = p.locateOnScreen('imgs/filtro_planilhas.png', confidence=0.7)
    # p.moveTo(p.center(filtro_planilhas), duration=0.2)
    p.click()
    p.sleep(0.5)
    # Executar o filtro selecionado
    executar = p.locateOnScreen('imgs/executar.png', confidence=0.9)
    p.moveTo(p.center(executar), duration=0.2)
    p.click()
    p.sleep(0.5)
    # Buscar pela planilha informada @param: id
    try:
        procurar = p.locateOnScreen('imgs/procurar.png', confidence=0.9)
        if procurar is None:
            raise TypeError('"ícone não localizado, utilizando a movimentação por coordenadas"')
    except:
        print("ícone não localizado, utilizando a movimentação por coordenadas")
        p.moveTo(x=239, y=184, duration=0.2)
        p.click()
        p.sleep(1)
    else:
        p.moveTo(p.center(procurar), duration=0.2)
        p.click()
        p.sleep(1)

    #buscar planilha pelo ID
    input_procurar = p.locateOnScreen('imgs/input_procurar.png', confidence=0.9)
    p.moveTo(p.center(input_procurar), duration=0.2)
    p.click()
    p.write(str(id))

    # Ocultar não filtrados
    try:
        params_rateio = p.locateOnScreen('imgs/rateios/parametros_rateio.png', confidence=0.9)
        if params_rateio is None:
            raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
            print(TypeError)
    except:
        p.moveTo(x=476, y=214, duration=0.2)
        p.click()
        p.sleep(0.2)
    else:
        p.moveTo(p.center(params_rateio), duration=0.2)
        p.click()
        p.sleep(0.3)

    # Selecionar e abrir planilha pesquisada
    p.moveTo(x=139, y=276, duration=0.2)
    p.doubleClick()
    p.sleep(2)
    # ok = p.locateOnScreen('imgs/planilhas/ok.png', confidence=0.9)
    # p.moveTo(p.center(ok), duration=0.2)
    # p.click()
    # p.sleep(2)


def executar_planilha():
    executar = p.locateOnScreen('imgs/planilhas/visao_dados.png', confidence=0.9)
    p.moveTo(p.center(executar), duration=0.5)
    p.click()
    p.moveTo(x=834, y=289)
    p.click()
    p.sleep(0.1)
    p.moveTo(x=764, y=457, duration=0.2)
    p.click()


def seleciona_area_impressao():
    # Seleciona área de impressão
    header_rateio = p.locateOnScreen('imgs/rateios/header_rateios.png', confidence=0.9)
    p.moveTo(p.center(header_rateio))
    p.click()
    scroll_down = p.locateOnScreen('imgs/planilhas/scroll_down.png', confidence=0.9)
    p.moveTo(p.center(scroll_down))
    p.click(clicks=65)
    p.keyDown('shift')
    p.moveTo(x=447, y=623)
    p.click()
    p.keyUp('shift')
    p.sleep(0.3)


def imprimir_planilha(orientacao='portrait', pdf='nao'):
    """
    Função para realizar a impressão de um rateio gerado no sistema;

    :param orientacao: str - opcional. Esse parâmetro recebe uma string para definir a orientação da página que
    será impressa. "landscape" irá enviar uma impressão no modo "paisagem" (folha deitada). Qualquer outro valor
    irá enviar uma impressão em modo "retrato" (folha em pé).

    :return: null (realiza a impressão)
    """
    if pdf == 'nao':
        imprimir = p.locateOnScreen('imgs/planilhas/imprimir.png', confidence=0.9)
        p.moveTo(p.center(imprimir), duration=0.2)
        p.click()
        p.sleep(0.5)
        if orientacao != 'landscape':
            imprimir = p.locateOnScreen('imgs/planilhas/imprimir_selecao.png', confidence=0.9)
            p.moveTo(p.center(imprimir), duration=0.2)
            p.click()
        imprimir = p.locateOnScreen('imgs/planilhas/ok_imprimir.png', confidence=0.9)
        p.moveTo(p.center(imprimir), duration=0.2)
        p.click()
        p.sleep(3)
        # rateio impresso em orientação paisagem
        if orientacao == 'landscape':
            preferencias = p.locateOnScreen('imgs/planilhas/print_preferences.png', confidence=0.9)
            p.moveTo(p.center(preferencias), duration=0.2)
            p.click()
            p.sleep(0.3)
            basic = p.locateOnScreen('imgs/planilhas/preferences_basic.png', confidence=0.9)
            p.moveTo(p.center(basic), duration=0.2)
            p.click()
            p.sleep(0.3)
            landscape = p.locateOnScreen('imgs/planilhas/landscape.png', confidence=0.9)
            p.moveTo(p.center(landscape), duration=0.2)
            p.click()
            p.sleep(0.3)
            p.press('enter')
            p.sleep(0.3)
        imprimir = p.locateOnScreen('imgs/planilhas/print.png', confidence=0.9)
        p.moveTo(p.center(imprimir), duration=0.2)
        p.sleep(1)
        p.click()
        p.sleep(1)
    else:
        imprimir = p.locateOnScreen('imgs/planilhas/imprimir.png', confidence=0.9)
        p.moveTo(p.center(imprimir), duration=0.2)
        p.click()
        p.sleep(0.5)
        imprimir = p.locateOnScreen('imgs/planilhas/imprimir_selecao.png', confidence=0.9)
        p.moveTo(p.center(imprimir), duration=0.2)
        p.click()
        imprimir = p.locateOnScreen('imgs/planilhas/ok_imprimir.png', confidence=0.9)
        p.moveTo(p.center(imprimir), duration=0.2)
        p.click()
        p.sleep(3)
        p.press('c')
        p.press('enter')
        save_as_pdf = p.locateOnScreen('imgs/rateios/save_as_pdf.png', confidence=0.9)
        while save_as_pdf is None:
            p.sleep(1)
            print('Aguardando Wizard para salvar o arquivo...')
            save_as_pdf = p.locateOnScreen('imgs/rateios/save_as_pdf.png', confidence=0.9)
        print('Salvando o rateio como PDF')
        p.sleep(0.3)
        p.press('tab', presses=5)
        dir_rateios_pdf = p.locateOnScreen('imgs/rateios/dir_rateios_pdf.png', confidence=0.8)
        if dir_rateios_pdf is None:
            p.write(r'RPA_rateio\rateio_gerado.pdf')
        else:
            p.write(r'rateio_gerado.pdf')
        p.press('tab', presses=2)
        p.sleep(0.3)
        p.press('enter')
        # imprimir = p.locateOnScreen('imgs/planilhas/print.png', confidence=0.9)
        # p.moveTo(p.center(imprimir), duration=0.2)
        # p.sleep(15)
        # p.click()


def gera_pdf_rateio():
    """
    Essa função tem o objetivo de gerar o arquivo PDF que era digitalizado pelo setor para ser armazenado
    no disco N:/. Ela recebe 3 arquivos PDF (NF, Boleto e Rateio) que são unificados em um único arquivo
    e salvos no diretório especificado.

    Arquivos:
        NF - Nota Fiscal recebida pelo setor (normalmente por e-mail)
        Boleto - Boleto referente a NF
        rateio - Rateio gerado pelo processo de RPA e salvo no disco N:/

        Todos os arquivos devem estar no diretório especificado antes do início da execução
        do robo.

    :return: Novo arquivo PDF com todas as páginas do rateio.
    """
    # Abre os arquivos que devem ser unificados
    nf1 = open('nf.pdf', 'rb') if exists('nf.pdf') else None
    nf2 = open('nf2.pdf', 'rb') if exists('nf2.pdf') else None
    nf3 = open('nf3.pdf', 'rb') if exists('nf3.pdf') else None
    nf4 = open('nf4.pdf', 'rb') if exists('nf4.pdf') else None
    nf5 = open('nf5.pdf', 'rb') if exists('nf5.pdf') else None
    nf6 = open('nf6.pdf', 'rb') if exists('nf6.pdf') else None
    nf7 = open('nf7.pdf', 'rb') if exists('nf7.pdf') else None
    nf8 = open('nf8.pdf', 'rb') if exists('nf8.pdf') else None

    boleto1 = open('boleto.pdf', 'rb') if exists('boleto.pdf') else None
    boleto2 = open('boleto2.pdf', 'rb') if exists('boleto2.pdf') else None
    boleto3 = open('boleto3.pdf', 'rb') if exists('boleto3.pdf') else None
    boleto4 = open('boleto4.pdf', 'rb') if exists('boleto4.pdf') else None
    boleto5 = open('boleto5.pdf', 'rb') if exists('boleto5.pdf') else None
    boleto6 = open('boleto6.pdf', 'rb') if exists('boleto6.pdf') else None
    boleto7 = open('boleto7.pdf', 'rb') if exists('boleto7.pdf') else None
    boleto8 = open('boleto8.pdf', 'rb') if exists('boleto8.pdf') else None

    rateio1 = open('rateio.pdf', 'rb') if exists('rateio.pdf') else None
    rateio2 = open('rateio2.pdf', 'rb') if exists('rateio2.pdf') else None
    rateio3 = open('rateio3.pdf', 'rb') if exists('rateio3.pdf') else None
    rateio4 = open('rateio4.pdf', 'rb') if exists('rateio4.pdf') else None
    rateio5 = open('rateio5.pdf', 'rb') if exists('rateio5.pdf') else None
    rateio6 = open('rateio6.pdf', 'rb') if exists('rateio6.pdf') else None
    rateio7 = open('rateio7.pdf', 'rb') if exists('rateio7.pdf') else None
    rateio8 = open('rateio8.pdf', 'rb') if exists('rateio8.pdf') else None

    # Lê os arquivos que foram abertos
    nf1reader = PyPDF2.PdfFileReader(nf1) if nf1 is not None else None
    nf2reader = PyPDF2.PdfFileReader(nf2) if nf2 is not None else None
    nf3reader = PyPDF2.PdfFileReader(nf3) if nf3 is not None else None
    nf4reader = PyPDF2.PdfFileReader(nf4) if nf4 is not None else None
    nf5reader = PyPDF2.PdfFileReader(nf5) if nf5 is not None else None
    nf6reader = PyPDF2.PdfFileReader(nf6) if nf6 is not None else None
    nf7reader = PyPDF2.PdfFileReader(nf7) if nf7 is not None else None
    nf8reader = PyPDF2.PdfFileReader(nf8) if nf8 is not None else None

    boleto1reader = PyPDF2.PdfFileReader(boleto1) if boleto1 is not None else None
    boleto2reader = PyPDF2.PdfFileReader(boleto2) if boleto2 is not None else None
    boleto3reader = PyPDF2.PdfFileReader(boleto3) if boleto3 is not None else None
    boleto4reader = PyPDF2.PdfFileReader(boleto4) if boleto4 is not None else None
    boleto5reader = PyPDF2.PdfFileReader(boleto5) if boleto5 is not None else None
    boleto6reader = PyPDF2.PdfFileReader(boleto6) if boleto6 is not None else None
    boleto7reader = PyPDF2.PdfFileReader(boleto7) if boleto7 is not None else None
    boleto8reader = PyPDF2.PdfFileReader(boleto8) if boleto8 is not None else None

    rateio1reader = PyPDF2.PdfFileReader(rateio1) if rateio1 is not None else None
    rateio2reader = PyPDF2.PdfFileReader(rateio2) if rateio2 is not None else None
    rateio3reader = PyPDF2.PdfFileReader(rateio3) if rateio3 is not None else None
    rateio4reader = PyPDF2.PdfFileReader(rateio4) if rateio4 is not None else None
    rateio5reader = PyPDF2.PdfFileReader(rateio5) if rateio5 is not None else None
    rateio6reader = PyPDF2.PdfFileReader(rateio6) if rateio6 is not None else None
    rateio7reader = PyPDF2.PdfFileReader(rateio7) if rateio7 is not None else None
    rateio8reader = PyPDF2.PdfFileReader(rateio8) if rateio8 is not None else None

    arquivos = [
        {
            'nf': nf1reader,
            'boleto': boleto1reader,
            'rateio': rateio1reader
        },
        {
            'nf': nf2reader,
            'boleto': boleto2reader,
            'rateio': rateio2reader
        },
        {
            'nf': nf3reader,
            'boleto': boleto3reader,
            'rateio': rateio3reader
        },
        {
            'nf': nf4reader,
            'boleto': boleto4reader,
            'rateio': rateio4reader
        },
        {
            'nf': nf5reader,
            'boleto': boleto5reader,
            'rateio': rateio5reader
        },
        {
            'nf': nf6reader,
            'boleto': boleto6reader,
            'rateio': rateio6reader
        },
        {
            'nf': nf7reader,
            'boleto': boleto7reader,
            'rateio': rateio7reader
        },
        {
            'nf': nf8reader,
            'boleto': boleto8reader,
            'rateio': rateio8reader
        },
    ]

    # Cria um novo Objeto "PdfFileWriter" que representa um arquivo em branco
    pdfWriter = PyPDF2.PdfFileWriter()

    # Passa pelas páginas dos arquivos para gerar o conteúdo.
    for arquivo in arquivos:
        if arquivo['nf'] is not None:
            for pageNum in range(arquivo['nf'].numPages):
                pageObj = arquivo['nf'].getPage(pageNum)
                pdfWriter.addPage(pageObj)

        if arquivo['boleto'] is not None:
            for pageNum in range(arquivo['boleto'].numPages):
                pageObj = arquivo['boleto'].getPage(pageNum)
                pdfWriter.addPage(pageObj)

        if arquivo['rateio'] is not None:
            for pageNum in range(arquivo['rateio'].numPages):
                pageObj = arquivo['rateio'].getPage(pageNum)
                pdfWriter.addPage(pageObj)

    # Salva o novo arquivo gerado.
    pdfOutputFile = open('MergedFiles.pdf', 'wb')
    pdfWriter.write(pdfOutputFile)

    # Fecha os arquivos - Criados no processo e Abertos no início
    pdfOutputFile.close()
    if nf1 is not None:
        nf1.close()
    if nf2 is not None:
        nf2.close()
    if nf3 is not None:
        nf3.close()
    if nf4 is not None:
        nf4.close()
    if nf5 is not None:
        nf5.close()
    if nf6 is not None:
        nf6.close()
    if nf7 is not None:
        nf7.close()
    if nf8 is not None:
        nf8.close()

    if boleto1 is not None:
        boleto1.close()
    if boleto2 is not None:
        boleto2.close()
    if boleto3 is not None:
        boleto3.close()
    if boleto4 is not None:
        boleto4.close()
    if boleto5 is not None:
        boleto5.close()
    if boleto6 is not None:
        boleto6.close()
    if boleto7 is not None:
        boleto7.close()
    if boleto8 is not None:
        boleto8.close()

    if rateio1 is not None:
        rateio1.close()
    if rateio2 is not None:
        rateio2.close()
    if rateio3 is not None:
        rateio3.close()
    if rateio4 is not None:
        rateio4.close()
    if rateio5 is not None:
        rateio5.close()
    if rateio6 is not None:
        rateio6.close()
    if rateio7 is not None:
        rateio7.close()
    if rateio8 is not None:
        rateio8.close()


def rateio_transportes(mes, ano, coligada, tipo, valor):
    """
    Função geradora do rateio dos transportes mensais. Desenvolvida a pedido do Maganha e do Nivaldo.
    Essa função realiza todas as operações do rateio que antes era feito manualmente pelo Maganha.
    A origem dos dados é uma planilha de carga que é lida pelo robo;

    :param mes: [str] - mes de referencia do rateio
    :param ano: [str] - ano de referencia do rateio
    :param coligada: [str] - coligada para qual o rateio está sendo realizado
    :param tipo: [srt] - tipo do rateio pode assumir 2 valores (1 e 2) para ACN transportes ou Viação Pirassununga,
                    respectivamente
    :param valor: [str] - valor total da NF
    :return: None
    """
    aba = p.locateOnScreen('imgs/rateio_transporte/aba_rateio.png', confidence=0.9)
    p.moveTo(p.center(aba), duration=0.2)
    p.click()
    p.sleep(0.5)
    # ir para celula B1
    filtro_mes = p.locateOnScreen('imgs/rateio_transporte/mes.png', confidence=0.9)
    p.moveTo(p.center(filtro_mes), duration=0.2)
    p.click()
    # Preencher dados do rateio:
    p.write(str(mes))
    p.press('down')
    p.write(str(ano))
    p.press('down')
    p.write(str(coligada))
    p.press('right', presses=2)
    p.press('up', presses=2 )
    p.write(str(tipo))
    p.press('down')
    p.write(str(valor))
    p.press('right')
    executar_planilha()
    # Selecionar area de impressão
    a1 = p.locateOnScreen('imgs/rateio_transporte/mes2.png', confidence=0.9) # célula A1
    p.moveTo(p.center(a1))
    p.click()
    p.sleep(0.5)
    scroll_down = p.locateOnScreen('imgs/planilhas/scroll_down.png', confidence=0.9)
    p.moveTo(p.center(scroll_down))
    if coligada == '%%' or coligada == '%':
        p.click(clicks=57)
    else:
        p.click(clicks=27)
    p.sleep(0.5)
    ult_cel = p.locateOnScreen('imgs/rateio_transporte/ultima_celula.png', confidence=0.9)
    p.moveTo(p.center(scroll_down))
    p.keyDown('shift')
    p.moveTo(x=507, y=507, duration=0.2)
    p.click()
    p.keyUp('shift')
    p.sleep(15)

    imprimir_planilha()


def cadastra_user_portal(nome, cpf, coligada, chapa):

    # Selecionar "incluir novo:
    try:
        rh = p.locateOnScreen('imgs/cria_user/folha_de_pagamento.png', confidence=0.9)
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

    # Selecionar "incluir novo:
    try:
        novo = p.locateOnScreen('imgs/cria_user/novo.png', confidence=0.9)
        if novo is None:
            raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
            print('Imagem não encontrada, tentando localizar por coordenada')
    except:
        p.moveTo(x=12, y=206, duration=0.2)
        p.click()
        p.sleep(0.5)
    else:
        p.moveTo(p.center(novo), duration=0.2)
        p.click()
        p.sleep(0.5)

    # preencher dados do usuário
    p.sleep(0.5)
    p.write(f'{coligada}-{chapa}')
    p.press('tab')
    p.write(nome)
    p.press('tab', presses=3)
    p.write('def')
    p.press('tab')
    # Selecionar mudar senha
    try:
        muda_senha = p.locateOnScreen('imgs/cria_user/muda_senha.png', confidence=0.9)
        if muda_senha is None:
            raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
            print('Imagem não encontrada, tentando localizar por coordenada')
    except:
        p.moveTo(x=467, y=443, duration=0.2)
        p.click()
        p.sleep(0.5)
    else:
        p.moveTo(p.center(muda_senha), duration=0.2)
        p.click()
        p.sleep(0.5)
    p.write(cpf)
    p.sleep(0.2)
    p.press('tab')
    p.write(cpf)
    p.sleep(0.2)
    p.press('tab')
    p.sleep(0.2)
    p.press('enter')
    p.sleep(0.3)
    # Ir para a aba "Segurança" para habilitar o perfil
    try:
        seguranca = p.locateOnScreen('imgs/cria_user/seguranca_perfis.png', confidence=0.9)
        if seguranca is None:
            raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
            print('Imagem não encontrada, tentando localizar por coordenada')
    except:
        p.moveTo(x=470, y=220, duration=0.2)
        p.click()
        p.sleep(0.5)
    else:
        p.moveTo(p.center(seguranca), duration=0.2)
        p.click()
        p.sleep(0.5)

    if coligada == '1':
        # Ir para a coligada 1
        try:
            col = p.locateOnScreen('imgs/cria_user/col_1.png', confidence=0.9)
            if col is None:
                raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
                print('Imagem não encontrada, tentando localizar por coordenada')
        except:
            p.moveTo(x=356, y=274, duration=0.2)
            p.click()
            p.sleep(0.5)
        else:
            p.moveTo(p.center(col), duration=0.2)
            p.click()
            p.sleep(0.5)
    elif coligada == '3':
        # Ir para a coligada 3
        try:
            col = p.locateOnScreen('imgs/cria_user/col_3.png', confidence=0.9)
            if col is None:
                raise TypeError('Imagem não encontrada, tentando localizar por coordenada')
                print('Imagem não encontrada, tentando localizar por coordenada')
        except:
            p.moveTo(x=356, y=274, duration=0.2)
            p.click()
            p.sleep(0.5)
        else:
            p.moveTo(p.center(col), duration=0.2)
            p.click()
            p.sleep(0.5)
    p.press('tab')
    p.write('portal_rh')
    p.sleep(0.2)
    p.moveTo(x=861, y=312, duration=0.1)
    p.click()
    p.sleep(0.2)
    p.press('tab', presses=2)
    # p.press('enter')
    p.sleep(1.3)
    p.sleep(10)


def rateio_ti(mes, ano, rateio):
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


def envia_email(id_rateio, mes, ano):
    try:
        outlook = p.locateOnScreen('imgs/email/outlook.png', confidence=0.9)
        if outlook is None:
            raise TypeError('imagem não encontrada, tentando imagem com e-mail')
    except:
        outlook = p.locateOnScreen('imgs/email/outlook2.png', confidence=0.9)
        p.moveTo(p.center(outlook), duration=0.2)
        p.rightClick()
        p.sleep(0.2)
    else:
        p.moveTo(p.center(outlook), duration=0.2)
        p.rightClick()
        p.sleep(0.2)
    p.sleep(0.6)
    nova_msg = p.locateOnScreen('imgs/email/novo_email.png', confidence=0.8)
    p.moveTo(p.center(nova_msg), duration=0.2)
    p.click()
    p.sleep(1)

    assunto = p.locateOnScreen('imgs/email/assunto.png', confidence=0.9)
    p.moveTo(p.center(assunto), duration=0.2)
    p.click()
    p.sleep(0.3)

    p.write('guilherme.baccan@baldin-bioenergia.com.br')
    p.sleep(0.2)
    p.press('tab', presses=3)
    nome = ''
    for rateio in rateios:
        if rateio['id'] == id_rateio:
            nome = rateio['empresa']
    p.write(f'[RATEIO] - {nome} - {mes}/{ano}')
    p.sleep(0.2)

    p.press('tab')
    p.write('Segue em anexo rateio gerado')
    p.sleep(1)

    anexo = p.locateOnScreen('imgs/email/anexo.png', confidence=0.9)
    p.moveTo(p.center(anexo), duration=0.2)
    p.click()
    p.sleep(0.3)

    neste_pc = p.locateOnScreen('imgs/email/neste_pc.png', confidence=0.9)
    p.moveTo(p.center(neste_pc), duration=0.2)
    p.click()
    p.sleep(0.5)

    dir_path = p.locateOnScreen('imgs/email/dir_path.png', confidence=0.9)
    p.moveTo(p.center(dir_path), duration=0.2)
    p.click()
    p.write(r'N:\TI\RPA_rateio')
    p.sleep(0.5)
    p.press('enter')
    p.sleep(0.5)
    p.press('tab', presses=6)
    p.sleep(0.5)
    p.write('rateio.pdf')
    p.sleep(0.5)
    p.press('enter')
    p.sleep(0.5)
    enviar = p.locateOnScreen('imgs/email/enviar.png', confidence=0.9)
    p.moveTo(p.center(enviar), duration=0.5)
    # p.click()


def gera_rateio_definitivo(rateios, mes, ano):
    planilha_net(368)
    p.sleep(1)
    for rateio in rateios:
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
        # alterar para execução em definitivo do rateio
        p.press('right', presses=3)
        p.sleep(0.2)
        p.press('up', presses=3)
        p.sleep(0.2)
        p.write('1')
        p.press('right')
        # Executar o rateio
        # p.moveTo(x=801, y=102, duration=0.2)
        # p.click()
        # print(f"Iniciando execução do rateio de id {rateio} para gravar dados")
        # p.sleep(0.2)
        # p.moveTo(x=834, y=289)
        # p.click()
        # p.moveTo(x=764, y=457, duration=0.2)
        # p.click()
        # p.sleep(5)
