from __future__ import print_function

from threading import Thread
import os
import os.path
import pyperclip as pc
import numpy as np
import sys
from datetime import date
import pyautogui
import pywinauto, time
import pandas as pd
import keyboard
from pywinauto.application import Application
from winotify import Notification, audio
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

#API GOOGLE SHEETS

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '#######'
SAMPLE_RANGE_NAME = 'Página1!A2:B'

def startParametrizacao():
    
    task = "Domínio Folha - Versão: "

    app = pywinauto.Application(backend="win32").connect(title_re=task)
    time.sleep(0.5)
    window = app["Domínio Folha - Versão: "]
    window.set_focus()
    window.maximize()

    #window.print_control_identifiers(filename="././resultado.txt") #para ver todos os elementos da domínio
    #window.Properties.TabControlSharing.select("Demonstrativo de INSS Folha e INSS eSocial")

    usuario = os.getlogin()

    wb = pd.read_excel('parametrizacao.xls')

    #solicitando o mês e o ano atual
    data = date.today()
    month, year = (data.month-1, data.year) if data.month != 1 else (12, data.year-1)
    data = data.replace(month=month, year=year)
    data = data.strftime(str("%m") + "/" + str("%Y"))
    

    # notificação de erro
    mensagem = Notification(
                    app_id="Automatização E-social",
                    title="ALERTA!",
                    msg=f"Diferença maior que R$ 0,20 identificada!\nAnotando na planilha.",
                    duration="short"
                    )

    mensagem.set_audio(sound=audio.LoopingAlarm9, loop=True)

    alerta = Notification(
                    app_id="Parametrização E-social",
                    title="ALERTA!",
                    msg=f"O Fechamento não foi realizado.",
                    duration="short"
                    )
    alerta.set_audio(sound=audio.Reminder, loop=True)
    
    for cod in wb["COD"]:
        
        #validação de token e credencial google planilhas
        credentials = None

        if os.path.exists('token.json'):
            credentials = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not credentials or not credentials.valid:
            if credentials and credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                credentials = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(credentials.to_json())
        
        def startDominio():
        #1 - abrir o perfil da empresa (por código)
            time.sleep(1)
            pywinauto.keyboard.send_keys('{F8}')
            time.sleep(3)
            pywinauto.keyboard.send_keys(str(cod))
            pywinauto.keyboard.send_keys('{ENTER}')
            pywinauto.keyboard.send_keys('{ESC}')
            time.sleep(8)

            #2 - ir em relatório -> e-Social -> Eventos Periódicos:

            relatorios = pyautogui.locateOnScreen('imagens/relatorios.png', region=(0,0, 1366,768))
            pyautogui.click(relatorios)
            time.sleep(0.5)
            esocial = pyautogui.locateOnScreen('imagens/esocial.png', region=(0,0, 1366,768))
            pyautogui.click(esocial)
            time.sleep(0.5)
            eventos = pyautogui.locateOnScreen('imagens/eventos.png', region=(0,0, 1366,768))
            pyautogui.click(eventos)
            time.sleep(0.9)
                    
                # 	1 - deixar o mês vigente
            window.child_window(title="Competência:", class_name="Button").click_input(coords=(100,10, 1366,768))
            
            pc.copy(data)
            pyautogui.hotkey('ctrl','v')
            
                #  	2 - marcar "Remuneração" e "Pagamentos"
            window.child_window(title="S-1200 - Remuneração", class_name="Button").click_input(coords=(5,7, 1366,768))
            window.child_window(title="S-1210 - Pagamentos", class_name="Button").click_input(coords=(5,7, 1366,768))

                # 	3 - clicar em "Enviar"
            window.child_window(title="En&viar", class_name="Button").click_input(coords=(0,0, 1366,768))
            time.sleep(120)

            aviso_rubricas = pyautogui.locateOnScreen('imagens/aviso_rubricas.PNG', region=(0,0, 1366,768))
            
            atencao = pyautogui.locateOnScreen('imagens/atencao.png', region=(0,0, 1366,768))
            atencao2 = pyautogui.locateOnScreen('imagens/atencao2.png', region=(0,0, 1366,768))

            if aviso_rubricas:
                pyautogui.click(aviso_rubricas)
                pywinauto.keyboard.send_keys('{ENTER}')
                time.sleep(8)
                window.child_window(title="En&viar", class_name="Button").click_input(coords=(0,0, 1366,768))
                time.sleep(120)

            if atencao2:
                pyautogui.click(atencao)
                pywinauto.keyboard.send_keys('{ENTER}')
                time.sleep(0.5)
                    
            if atencao:
                pyautogui.click(atencao)
                pywinauto.keyboard.send_keys('{ENTER}')
                time.sleep(0.5)

            enviado_sucesso = pyautogui.locateOnScreen('imagens/enviado_sucesso.png', region=(0,0, 1366,768))
            
            if enviado_sucesso:
                pyautogui.click(enviado_sucesso)
                pywinauto.keyboard.send_keys('{ENTER}')
            else:
                enviado_sucesso == False   
            
            #3 - acompanhar o envio no "Painel de Pendências"
            time.sleep(15)
            window.child_window(title="&Painel Pendências", class_name="Button").click_input(coords=(0,0, 1366,768))
            time.sleep(1)
            em_processamento = pyautogui.locateOnScreen('imagens/em_processamento.png', region=(0,0, 1366,768))
            pyautogui.click(em_processamento)

                # 	1 - caso o a coluna "Evento" estiver escrito "Rubricas", o procedimento de envio deve ser feito duas vezes (pois posteriormente será enviado os "Pagamentos" e "Remuneração")

            rubrica = pyautogui.locateOnScreen('imagens/Rubrica.png', region=(0,0, 1366,768))

            if rubrica:
                while pyautogui.locateOnScreen('imagens/Rubrica.png', region=(0,0, 1366,768)): # se Rubrica existir, faça:
                    # 	2 - clicar em "Atualizar" para verificar se já foram enviadas
                    window.child_window(title="&Atualizar", class_name="Button").double_click_input(coords=(0,0, 1366,768))
                    time.sleep(1.5)
                else: # quando não existir mais, faça:
                    # 	3 - quando não tiver mais nada na aba "Em Processamento", podemos dar continuidade (em caso de erro, o mesmo deve ser excluído e enviado de novo até ser enviado)
                    window.child_window(title="Fe&char",class_name="Button").click_input(coords=(0,0, 1366,768)) #procurar o botao fechar
                    window.child_window(title="En&viar", class_name="Button").click_input(coords=(0,0, 1366,768))
                    time.sleep(4.5)
                    pyautogui.click(enviado_sucesso)
                    pywinauto.keyboard.send_keys('{ENTER}')
                    window.child_window(title="&Painel Pendências", class_name="Button").click_input(coords=(0,0, 1366,768))
                    time.sleep(1)
                    pyautogui.click(em_processamento)
                    
                    #  4 - Para verificar se os eventos foram validados, ir no "Painel de Pendências" -> Em Processamento e aguardar ser enviado. Depois ir em Validados -> Periódicos e filtrar pela competência
            fechar = pyautogui.locateOnScreen('imagens/fechar.png', region=(0,0, 1366,768))

            while pyautogui.locateOnScreen('imagens/S-1200.png', region=(0,0, 1366,768)) or pyautogui.locateOnScreen('imagens/S-1210.png', region=(0,0, 1366,768)):
                window.child_window(title="&Atualizar", class_name="Button").double_click_input(coords=(0,0, 1366,768))
                time.sleep(1)
            else:
                pyautogui.click(fechar)
                time.sleep(1)

            pywinauto.keyboard.send_keys('{ESC}')
            time.sleep(0.5)

            # 3 - Depois de enviado, ir em "Relatórios" -> "Demonstrativo de INSS Folha..."
                # 	1 - Ok
                # 	2 - Checar se a coluna "Situação" está com o status de "Enviado"
                # 	3 - Checar se as colunas "Base de Cálculo" e "Valor INSS" da tabela "Valor INSS Sistema" coincidem com os valores da tabela "Valor INSS eSocial".
                # 	4 - Esc
                # 	5 - Esc

            pyautogui.click(relatorios)
            time.sleep(0.5)
            esocial = pyautogui.locateOnScreen('imagens/esocial.png', region=(0,0, 1366,768))
            pyautogui.click(esocial)
            time.sleep(0.5)
            demons_inss = pyautogui.locateOnScreen('imagens/demons_inss.png', region=(0,0, 1366,768))
            pyautogui.click(demons_inss)
            time.sleep(1)
            window.child_window(title="Competência:", class_name="Button").click_input(coords=(150,10, 1366,768))
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            window.child_window(title="Competência:", class_name="Button").click_input(coords=(250,10, 1366,768))
            pyautogui.hotkey('ctrl','v')
            window.child_window(title="&OK", class_name="Button").click_input(coords=(0,0, 1366,768))
            time.sleep(15)

            # ------ Salvar o demonstrativo em uma planilha excel em um local padrão, abrir com python e comparar as colunas
            salvar = pyautogui.locateOnScreen('imagens/salvar.png', region=(0,0, 1366,768))
            pyautogui.click(salvar)
            time.sleep(2)
            opcoes = pyautogui.locateOnScreen('imagens/opcoes.png', region=(0,0, 1366,768))
            pyautogui.click(opcoes)
            time.sleep(2)
            planilha = pyautogui.locateOnScreen('imagens/planilha.png', region=(0,0, 1366,768))
            pyautogui.click(planilha)
            time.sleep(1)
            local = pyautogui.locateOnScreen('imagens/local.png', region=(0,0, 1366,768))
            pyautogui.click(local)
            time.sleep(3)
            
            desktop = pyautogui.locateOnScreen('imagens/desktop.png', region=(0,0, 1366,768))    
            desktop2 = pyautogui.locateOnScreen('imagens/desktop2.png', region=(0,0, 1366,768))
            desktop3 = pyautogui.locateOnScreen('imagens/desktop3.png', region=(0,0, 1366,768))        
            if desktop: 
                pyautogui.doubleClick(desktop)
            elif desktop2:
                pyautogui.doubleClick(desktop2)
            elif desktop3:
                pyautogui.doubleClick(desktop3)
            else:
                False

            time.sleep(1)
            campo_nome = pyautogui.locateOnScreen('imagens/campo_nome.png', region=(0,0, 1366,768))
            pyautogui.click(campo_nome)
            time.sleep(1)
            pyautogui.write("Dados INSS")
            time.sleep(1)
            salvar2 = pyautogui.locateOnScreen('imagens/salvar2.png', region=(0,0, 1366,768))
            pyautogui.click(salvar2)
            time.sleep(1)
            salvar3 = pyautogui.locateOnScreen('imagens/salvar3.png', region=(0,0, 1366,768))
            pyautogui.click(salvar3)
            time.sleep(1)

            # ----- Verificação se os dados conferem através da planilha

            data_atual = date.today()
            month, year = (data_atual.month-1, data_atual.year) if data_atual.month != 1 else (12, data_atual.year-1)
            data_atual = data_atual.replace(day=1, month=month, year=year)
            competencia = data_atual.strftime('%d/%m/%Y')
            #print(competencia)

            tabela = pd.read_excel(f"C:\\Users\\{usuario}\\Desktop\\Dados INSS.xls", engine='xlrd') #usar engine='xlrd' para ler arquivos antigos como xls

            # RECUPERANDO O CNPJ
            # cnpj=str(tabela.loc[0, "cp_cgce_emp"])
            # empresa=str(tabela.loc[0, "cp_nome_emp"])
            # para cada empresa, colocar a variavel cnpj em cada linha de uma planilha

            # append_data = pd.DataFrame([{'COD':cod, 'EMPRESA':empresa, 'CNPJ':cnpj}])
            # wb2 = load_workbook('C:\\Users\\Suporte\\Desktop\\Automatizacao e-social\\cnpj.xlsx')
            # aba_ativa = wb2["Planilha1"]
            
            # for linha in dataframe_to_rows(append_data, index=False, header=False):
            #     aba_ativa.append(linha)
            # wb2.save('C:\\Users\\Suporte\\Desktop\\Automatizacao e-social\\cnpj.xlsx')


            #SISTEMA
            sistema_base = tabela.loc[tabela["competencia"]==competencia,"base_inss"]
            sistema_val = tabela.loc[tabela["competencia"]==competencia,"valor_inss"]
            #print(sistema_base)
            #print(sistema_val)

            # ------ ESOCIAL
            esocial_base = tabela.loc[tabela["competencia"]==competencia, "cp_base_inss_esocial"]
            esocial_val = tabela.loc[tabela["competencia"]==competencia, "cp_valor_calculado_inss_esocial"]
            #print(esocial_base)
            #print(esocial_val)

            # ------ COMPARAÇÃO
            base = (np.array_equal(sistema_base, esocial_base)) #compara se os dois arrays são iguais (base calculo sistema, base calculo esocial)
            valor = (np.array_equal(sistema_val, esocial_val)) # valor inss (sistema), valor calculado inss (inss)
            #print(base)
            #print(valor)
            
            if base and valor:
                pyautogui.alert(text=f'Conferência realizada, os valores do INSS da Empresa {cod} coincidem. ', title='Automatização eSocial', button='OK', timeout="15000")
                time.sleep(3)
                pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia
           

                pyautogui.click(relatorios)
                time.sleep(0.5)
                esocial = pyautogui.locateOnScreen('imagens/esocial.png', region=(0,0, 1366,768))
                pyautogui.click(esocial)
                time.sleep(0.5)
                demons_fgts = pyautogui.locateOnScreen('imagens/demons_fgts.png', region=(0,0, 1366,768))
                pyautogui.click(demons_fgts)
                time.sleep(1)
                window.child_window(title="Competência:", class_name="Button").click_input(coords=(150,10, 1366,768))
                #pc.copy("12/2022")
                pyautogui.hotkey('ctrl','v')
                window.child_window(title="Competência:", class_name="Button").click_input(coords=(250,10, 1366,768))
                pyautogui.hotkey('ctrl','v')
                window.child_window(title="&OK", class_name="Button").click_input(coords=(0,0, 1366,768))
                time.sleep(15)

                # ------ Verificação se os dados conferem através da planilha

                salvar = pyautogui.locateOnScreen('imagens/salvar.png', region=(0,0, 1366,768))
                pyautogui.click(salvar)
                time.sleep(2)

                opcoes = pyautogui.locateOnScreen('imagens/opcoes.png', region=(0,0, 1366,768))
                pyautogui.click(opcoes)
                time.sleep(1)

                planilha = pyautogui.locateOnScreen('imagens/planilha.png', region=(0,0, 1366,768))
                pyautogui.click(planilha)

                local = pyautogui.locateOnScreen('imagens/local.png', region=(0,0, 1366,768))
                pyautogui.click(local)
                time.sleep(3)

                desktop = pyautogui.locateOnScreen('imagens/desktop.png', region=(0,0, 1366,768))    
                desktop2 = pyautogui.locateOnScreen('imagens/desktop2.png', region=(0,0, 1366,768))
                desktop3 = pyautogui.locateOnScreen('imagens/desktop3.png', region=(0,0, 1366,768))        
                if desktop: 
                    pyautogui.doubleClick(desktop)
                elif desktop2:
                    pyautogui.doubleClick(desktop2)
                elif desktop3:
                    pyautogui.doubleClick(desktop3)
                else:
                    False
                
                time.sleep(1)

                campo_nome = pyautogui.locateOnScreen('imagens/campo_nome.png', region=(0,0, 1366,768))
                pyautogui.click(campo_nome)
                time.sleep(0.5)

                pyautogui.write("Dados FGTS")
                time.sleep(1)

                salvar2 = pyautogui.locateOnScreen('imagens/salvar2.png', region=(0,0, 1366,768))
                pyautogui.click(salvar2)
                time.sleep(1)

                salvar3 = pyautogui.locateOnScreen('imagens/salvar3.png', region=(0,0, 1366,768))
                pyautogui.click(salvar3)
                time.sleep(1)

                # verificação se os dados conferem através da planilha

                tabela_2 = pd.read_excel(f"C:\\Users\\{usuario}\\Desktop\\Dados FGTS.xls", engine='xlrd') 
                #C:\\Users\\Suporte\\Desktop\\

                # SE A BASE_FGTS_DO_SISTEMA - BASE_DO_ESOCIAL FOR MENOR OU IGUAL A 20 CENTAVOS
                calculo_fgts = (tabela_2.loc[tabela_2["competencia"]==competencia,"base_fgts"]-tabela_2.loc[tabela_2["competencia"]==competencia, "cp_base_fgts_esocial"]) <= 0.20
                resultado_base = calculo_fgts.all()
                # A FUNÇÃO .all() testa se todos os resultados são TRUE

                calculo_valor = (tabela_2.loc[tabela_2["competencia"]==competencia, "valor_fgts"]-tabela_2.loc[tabela_2["competencia"]==competencia,"cp_valor_fgts_esocial"]) <= 0.20
                resultado_valor = calculo_valor.all()

                if resultado_base and resultado_valor:
                    #print("Ok, a diferença é menor que 0.20, pode continuar!!!!")
                    pyautogui.alert(text=f'Conferência realizada com sucesso!', title='Automatização eSocial', button='OK', timeout="15000")
                    time.sleep(3)
                    pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                    pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia
                    # 5 - "Rotina" -> "Eventos Periódicos"
                        
                    pyautogui.click(relatorios)
                    time.sleep(0.5)
                    esocial = pyautogui.locateOnScreen('imagens/esocial.png', region=(0,0, 1366,768))
                    pyautogui.click(esocial)
                    time.sleep(0.5)
                    eventos = pyautogui.locateOnScreen('imagens/eventos.png', region=(0,0, 1366,768))
                    pyautogui.click(eventos)
                    time.sleep(1)
                    window.child_window(title="Competência:", class_name="Button").click_input(coords=(100,10, 1366,768))
                    #pc.copy("12/2022")
                    pc.copy(data)
                    # 	1 - Clicar em "Fechamento dos Eventos Periódicos"
                    s1299 = pyautogui.locateOnScreen('imagens/S-1299.png', region=(0,0, 1366,768))
                    pyautogui.click(s1299)
                    # 	2 - Clicar em "Enviar"
                    window.child_window(title="En&viar", class_name="Button").click(coords=(0,0, 1366,768))
                    time.sleep(15)
                    # 	3 - Ok
                    fechamentomsg = pyautogui.locateOnScreen('imagens/atencao_dctf.png', region=(0,0, 1366,768))

                    if atencao2:
                        pyautogui.click(atencao)
                        pywinauto.keyboard.send_keys('{ENTER}')
                        time.sleep(0.5)

                    if fechamentomsg:
                        pyautogui.click(fechamentomsg)
                        pywinauto.keyboard.send_keys('{ENTER}')

                    enviado_sucesso = pyautogui.locateOnScreen('imagens/enviado_sucesso.png', region=(0,0, 1366,768))
                    
                    if enviado_sucesso:
                        pyautogui.click(enviado_sucesso)
                        pywinauto.keyboard.send_keys('{ENTER}')
                    else:
                        enviado_sucesso == False

                    window.child_window(title="&Painel Pendências", class_name="Button").click_input(coords=(0,0, 1366,768))
                    time.sleep(1)
                    em_processamento = pyautogui.locateOnScreen('imagens/em_processamento.png', region=(0,0, 1366,768))
                    pyautogui.click(em_processamento)

                    # 	4 - Conferir no "Painel de Pendências" o status do processamento clicando em atualizar até sumir.
                    # se o evento ainda estiver sendo processado, identifique a imagem "fechamento" e clique em atualizar até q ela suma
                    while pyautogui.locateOnScreen('imagens/fechamento.png', region=(0,0, 1366,768)):
                        window.child_window(title="&Atualizar", class_name="Button").click_input(coords=(0,0, 1366,768))
                        time.sleep(1)

                    window.child_window(title="&Validados", class_name="Button").click_input(coords=(0,0, 1366,768))
                    time.sleep(3)
                    # 	5 - Conferir na aba "Periódicos" se aparece "Fechamento" na coluna "Evento" (aqui será feito a filtragem do mês/ano do fechamento e o donwload da planilha para pesquisar por FECHAMENTO nela)
                    periodicos = pyautogui.locateOnScreen('imagens/periodicos.png', region=(0,0, 1366,768))
                    pyautogui.click(periodicos)
                    window.child_window(title="&Filtrar", class_name="Button").click_input(coords=(0,0, 1366,768))
                    time.sleep(3)
                    pyautogui.hotkey('ctrl','v')
                    pywinauto.keyboard.send_keys('{TAB}')
                    pyautogui.hotkey('ctrl','v')
                    pywinauto.keyboard.send_keys('{TAB}')
                    pywinauto.keyboard.send_keys('{TAB}')
                    pywinauto.keyboard.send_keys('{ENTER}')
                    
                    window.child_window(title="&Relatório", class_name="Button").click_input(coords=(0,0, 1366,768))
                    time.sleep(15)
                    
                    salvar = pyautogui.locateOnScreen('imagens/salvar.png', region=(0,0, 1366,768))
                    pyautogui.click(salvar)
                    time.sleep(2)
                    opcoes = pyautogui.locateOnScreen('imagens/opcoes.png', region=(0,0, 1366,768))
                    pyautogui.click(opcoes)
                    time.sleep(1)
                    planilha = pyautogui.locateOnScreen('imagens/planilha.png', region=(0,0, 1366,768))
                    pyautogui.click(planilha)
                    time.sleep(1)
                    local = pyautogui.locateOnScreen('imagens/local.png', region=(0,0, 1366,768))
                    pyautogui.click(local)
                    time.sleep(3)
                    
                    desktop = pyautogui.locateOnScreen('imagens/desktop.png', region=(0,0, 1366,768))    
                    desktop2 = pyautogui.locateOnScreen('imagens/desktop2.png', region=(0,0, 1366,768))
                    desktop3 = pyautogui.locateOnScreen('imagens/desktop3.png', region=(0,0, 1366,768))        
                    if desktop: 
                        pyautogui.doubleClick(desktop)
                    elif desktop2:
                        pyautogui.doubleClick(desktop2)
                    elif desktop3:
                        pyautogui.doubleClick(desktop3)
                    else:
                        False

                    time.sleep(1)
                    campo_nome = pyautogui.locateOnScreen('imagens/campo_nome.png', region=(0,0, 1366,768))
                    pyautogui.click(campo_nome)
                    time.sleep(0.5)

                    pyautogui.write("RelatorioValidados")
                    time.sleep(1)

                    salvar2 = pyautogui.locateOnScreen('imagens/salvar2.png', region=(0,0, 1366,768))
                    pyautogui.click(salvar2)
                    time.sleep(1)

                    salvar3 = pyautogui.locateOnScreen('imagens/salvar3.png', region=(0,0, 1366,768))
                    pyautogui.click(salvar3)
                    time.sleep(1)
                    
                    # verificação se os dados conferem através da planilha

                    tabela_validados = pd.read_excel(f"C:\\Users\\{usuario}\\Desktop\\RelatorioValidados.xls", engine='xlrd')
                    
                    # data2 = date.today()
                    # data2 = data2.strftime(str("%d") + "/" + str("%m") + "/" + str("%Y"))

                    # tabela_3 = tabela_validados.loc[tabela_validados["data_conclusao"]==data2,"evento"]
                    #print(tabela_3)
                    
                    for coluna in tabela_validados.columns:
                        fechamento_excel = tabela_validados["evento"]=="S-1299 Fechamento"
                        
                    if fechamento_excel.any():
                        pyautogui.alert(text='Conferência realizada, o Fechamento foi enviado com sucesso!. ', title='Automatização eSocial', button='OK', timeout="10000")
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        
                    else:
                        alerta.show()
                        pyautogui.alert(text='Alerta! O Fechamento não foi realizado.', title='Automatização eSocial', button='OK', timeout="15000")
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')
                    
                    # 6 - Ir em "Controle" -> "Empresas" e copiar o CNPJ da empresa (para inserir no site do ecac)
                    
                else:
                    #print("Valor maior que R$0,20 centavos")
                    # sairDominio()
                    procedimento = "FGTS"
                    
                    # exibindo notificação de ALERTA
                    mensagem.show()
                     
                    try:
                        service = build('sheets', 'v4', credentials=credentials)

                        # Call the Sheets API
                        sheet = service.spreadsheets()

                        valores_adicionar = [
                                [cod, procedimento]
                            ]

                        result = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED", body={'values':valores_adicionar}).execute()

                    except HttpError as err:
                        DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
                        service = build('sheets', 'v4', credentials=credentials, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
                        print(err)
                        

                    time.sleep(3)
                    pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                    pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia
                
            else:
                #print("Erro: OS DADOS NÃO CONFEREM")
                # sairDominio()
                procedimento = "INSS"
                
                # exibindo notificação de ALERTA
                mensagem.show()

                # ABRE O GOOGLE PLANILHAS, DIGITA O CÓDIGO DA EMPRESA E O PROCEDIMENTO QUE TEVE DIFERENÇA (INSS)
                try:
                    service = build('sheets', 'v4', credentials=credentials)

                    # Call the Sheets API
                    sheet = service.spreadsheets()

                    valores_adicionar = [
                            [cod, procedimento]
                        ]

                    result = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED", body={'values':valores_adicionar}).execute()
                    
                except HttpError as err:
                    DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
                    service = build('sheets', 'v4', credentials=credentials, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
                    print(err)

                time.sleep(3)

                pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia

            

            # * PASSOS NAVEGADOR

            # 1 - Acessar o site do Ecac
            # 	1 - Cliclar em "Entrar com gov.br"

            # 2 - Clicar em "Seu certificado digital"

            # 3 - Selecionar "ORGANIZAÇÃO E CONSULTORIA..."
            # 	1 - Clicar em "ok"
            # 	2 - caso o site apresente um erro ao entrar, verificar uma solução

            # 4 - Clicar em "Alterar perfil de acesso"
            # 	1 - No campo "Procurador de pessoa jurídica - CNPJ" colocar o CNPJ da empresa
            # 	2 - Clicar em "Alterar"

            # 5 - Clicar no campo "Localizar Serviço"
            # 	1 - Digitar "DCTFWEB" e clicar na primeira opção

            # 6 - Voltar na domínio
            # 	1 - Clicar em "Processos" -> "Apuração Previdenciária"
            # 	2 - "Apurar"
            # 	3 - "Sim"
            # 	4 - Verificar se o "Saldo a Recolher" é igual ao do Ecac

            # 7 - Se conferir, clicar em "Transmitir" na coluna "Serviços" do Ecac.
            # 	1 - Clicar em "Trasmitir sem efetuar vinculações" (irá começar a fazer o download do arquivo)

            # 8 - Clicar no arquivo baixado

            # 9 - Clicar em "executar"
            # 	1 - clicar em "Ok"

            # 10 - Voltar no Ecac, selecionar a caixa na coluna "Saldo a Pagar" do arquivo e depois clicar em "Guia" para baixar a Guia.

            # 11 - Procurar onde o arquivo foi salvo e extrair.

            # 12 - Renomear o nome do arquivo para "DCTFWEB mês(mm)_ano(aaaa) NOME DA EMPRESA"

            # 13 - Enviar o arquivo no Onvio         
        startDominio()

result = pyautogui.confirm(text="Se você realmente deseja iniciar o programa, clique em Iniciar", title="Automatização eSocial", buttons=['Iniciar', 'Sair'])
if result == "Iniciar":
    startParametrizacao()
    
elif result == "Sair":
    #print('saindo')
    pyautogui.alert(text='Programa não iniciado. ', title='Automatização eSocial', button='OK')
    sys.exit()

if startParametrizacao != True:
    pyautogui.alert(text='Automatização Concluída com Sucesso.', title='Automatização eSocial', button='OK')