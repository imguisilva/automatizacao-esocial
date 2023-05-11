from __future__ import print_function

import pyscreeze
import psutil
import os
import os.path
import pyperclip as pc
import numpy as np
import sys
from datetime import date
import datetime
from time import sleep
import pyautogui
import pywinauto, time
import pandas as pd
from pywinauto.application import Application
from winotify import Notification, audio

#API GOOGLE SHEETS

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1vTKSgXA0iecRLT57M7O3v_0A_V2v8fpcd1luswuY7XU'
SAMPLE_RANGE_NAME = 'Página1!A2:B'

def RunningProcess(processName):
    for proc in psutil.process_iter():
        try:
            if processName.lower() in proc.name().lower():
                return True
        except(psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False;

def pidbyname(processName):
    listOfProcessObjects = []
    for proc in psutil.process_iter():
        try:
            pinfo = proc.as_dict(attrs=['pid', 'name'])

            if processName.lower() in pinfo['name'].lower():
                listOfProcessObjects.append(pinfo)
        except(psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return listOfProcessObjects;

if RunningProcess('contabil'):
    listOfProcessIds = pidbyname('contabil')
    if len (listOfProcessIds) > 0:
        for i in listOfProcessIds:
            processID = i['pid']
            print(processID)

def startParametrizacao():
    
    app = Application().connect(process = i['pid'])
    window = app.top_window()
    window.set_focus()
    window.maximize()

    usuario = os.getlogin()

    wb = pd.read_excel('parametrizacao.xlsx')

    #dia/mês/ano atual
    data_hoje = date.today()
    data_hoje = data_hoje.strftime(str("%d") + "/" + str("%m") + "/" + str("%Y"))

    # notificação de erro
    mensagem = Notification(
                    app_id="Automatização E-social",
                    title="ALERTA!",
                    msg=f"Diferença maior que R$ 0,20 identificada!\nAnotando na planilha.",
                    duration="short"
                    )

    mensagem.set_audio(sound=audio.LoopingAlarm9, loop=True)

    msg = ''
    alerta = Notification(
                    app_id="Parametrização E-social",
                    title="ALERTA!",
                    msg=f"{msg}",
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
            
            # ABRE O GOOGLE PLANILHAS, DIGITA O CÓDIGO DA EMPRESA E O PROCEDIMENTO QUE HOUVE DIFERENÇA/ERRO
            def open_google():
                try:
                    service = build('sheets', 'v4', credentials=credentials)

                    # Call the Sheets API
                    sheet = service.spreadsheets()
                    
                    # Pegando hora atual
                    hora = datetime.datetime.now()
                    hora = hora.strftime("%H:%M")

                    valores_adicionar = [
                            [cod, procedimento, data_hoje, hora]
                        ]

                    result = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME, valueInputOption="USER_ENTERED", body={'values':valores_adicionar}).execute()

                except HttpError as err:
                    DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
                    service = build('sheets', 'v4', credentials=credentials, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
                    print(err)
                    time.sleep(0.5)

            #Mês atual
            data = date.today()
            data = data.strftime(str("%m") + "/" + str("%Y"))

            #Mês anterior
            data2 = date.today()
            month, year = (data2.month-1, data2.year) if data2.month != 1 else (12, data2.year-1)
            data2 = data2.replace(month=month, year=year)
            data2 = data2.strftime(str("%m") + "/" + str("%Y"))

        #1 - abrir o perfil da empresa (por código)
            time.sleep(1)
            pywinauto.keyboard.send_keys('{F8}')
            time.sleep(3)
            pywinauto.keyboard.send_keys(str(cod))

            pywinauto.keyboard.send_keys('{ENTER}')

            app = Application().connect(process = i['pid'])
            window = app.top_window()

            app2 = Application(backend='uia').connect(process = i['pid'])
            window2 = app2.top_window()

            acesso = app.window(title_re='Aviso')

            try:
                app.window(title_re = 'Aviso').wait('visible', timeout=10, retry_interval=1)
            except:
                pass

            if not acesso.exists():

                pywinauto.keyboard.send_keys('{ESC}')
                time.sleep(1)

                #2 - ir em relatório -> e-Social -> Eventos Periódicos:

                # TOOLBAR DO BOTÃO DE RELATÓRIOS
                window2.child_window(title="Relatórios", control_type="MenuItem").click_input(coords=(0,0, 1366,768))
                window2.child_window(title="e-Social", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))
                window2.child_window(title="Eventos periódicos", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))
                        
                    # 	1 - deixar o mês vigente
                window.child_window(title="Competência:", class_name="Button").wait('exists', timeout=15, retry_interval=1).click_input(coords=(100,10, 1366,768))
                pc.copy(data2)
                pywinauto.keyboard.send_keys('^v')

                window.child_window(title="S-1200 - Remuneração", class_name="Button").wait('exists', timeout=15, retry_interval=1).click_input(coords=(100,10, 1366,768))
                window.child_window(title="S-1210 - Pagamentos", class_name="Button").wait('exists', timeout=15, retry_interval=1).click_input(coords=(100,10, 1366,768))

                window.child_window(title="En&viar", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))
            
                window.child_window(title="En&viar").wait('exists', timeout=500, retry_interval=2) #Espera o "Eventos Periodicos ficar disponivel dnv"
                #print("existe o enviar")

                dlg = app.window(title_re='Atenção')

                recisoes = dlg.child_window(title_re='Existem rescisões calculadas que não foram enviadas ao eSocial e se enviadas não será gerado de forma automática o evento de pagamento para o eSocial. Deseja enviar as rescisões agora?', class_name='Static')
                
                #A competência está fechada!
                comp_fechada = dlg.child_window(title_re='A competência informada está fechada no eSocial. Para envio de eventos nesta competência é necessário reabri-la.', class_name='Edit')
                
                #Não existem cálculos nesta competência.
                nao_calculo = dlg.child_window(title_re='Não existem cálculos nesta competência para envio.', class_name='Static')
                nao_pagamentos = dlg.child_window(title_re='Não existem pagamentos nesta competência para envio.', class_name='Static')

                #Eventos periódicos enviados com sucesso!
                enviado_sucesso = dlg.child_window(title_re='Eventos periódicos enviados com sucesso!', class_name='Static')
                
                aviso_rubricas = app.window(title="Avisos Rubricas")
                
                if recisoes.exists():
                    try:
                        time.sleep(1)
                        dlg.child_window(title="&Sim", class_name="Button").click_input(coords=(0,0, 1366,768))
                        dlg.child_window(title_re='Eventos periódicos enviados com sucesso!', class_name='Static').wait('visible', timeout=180)
                    except:
                        pass

                if aviso_rubricas.exists():
                    time.sleep(1)
                    aviso_rubricas.child_window(title="&Sim", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))
                    window.child_window(title="&Painel Pendências", class_name="Button").wait('visible', timeout=30).click_input(coords=(0,0, 1366,768))
                    window.child_window(class_name="PBTabControl32_100").click_input(coords=(550,10, 1366,768))
                    
                    #Definindo o temporizador
                    segundos = 60
                    intervalo = 1

                    while not pyautogui.locateOnScreen('imagens/Rubrica.png', region=(0,0, 1366,768)): # se Rubrica existir, faça:
                        try:
                            print("NÃO EXISTE, CLICANDO NOVAMENTE...")
                            window.child_window(title="&Atualizar", class_name="Button", found_index=0).double_click_input(coords=(0,0, 1366,768))

                            #Temporizador
                            if segundos>0:
                                print(segundos)
                                segundos-=intervalo
                                sleep(intervalo)
                            if segundos == 0:
                                print('Finalizado')
                                break
                            elif segundos < 0:
                                print('Finalizado')
                                break

                        except:
                            #print("AGORA EXISTE, SAINDO...")
                            pass

                    while pyautogui.locateOnScreen('imagens/Rubrica.png', region=(0,0, 1366,768)): # se Rubrica existir, faça:
                        try:
                            #print("TEM RUBRICA")
                            window.child_window(title="&Atualizar", class_name="Button", found_index=0).double_click_input(coords=(0,0, 1366,768))
                            time.sleep(1)
                        except:
                            #print("NÃO TEM MAIS RUBRICA")
                            pass

                    else:
                        # 	3 - quando não tiver mais nada na aba "Em Processamento", podemos dar continuidade (em caso de erro, o mesmo deve ser excluído e enviado de novo até ser enviado)
                        window.child_window(title="Fe&char",class_name="Button").click_input(coords=(0,0, 1366,768)) #procurar o botao fechar
                        window.child_window(title="En&viar", class_name="Button").click_input(coords=(0,0, 1366,768))
                        window.child_window(title="En&viar").wait('exists', timeout=500, retry_interval=2)
                        #print("existe o enviar 2")

                if not aviso_rubricas.exists():
                    if not nao_calculo.exists(): #não existem calculos
                        if not nao_pagamentos.exists():
                            if not comp_fechada.exists(): #A competência está fechada
                                if enviado_sucesso.exists(): #enviado com sucesso
                                    time.sleep(2)
                                    dlg.child_window(title="OK", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))
                                
                                    #3 - acompanhar o envio no "Painel de Pendências"
                                    window.child_window(title="&Painel Pendências", class_name="Button").wait('visible', timeout=30).click_input(coords=(0,0, 1366,768))
                                    time.sleep(1)

                                    invalidados = pyautogui.locateOnScreen('imagens/invalidados.png', region=(0,0, 1366,768))
                                    if not invalidados:

                                        #Clicar na aba "Em Processamento"
                                        window.child_window(class_name="PBTabControl32_100").click_input(coords=(550,10, 1366,768))
                                        window.child_window(title="&Atualizar", class_name="Button", found_index=0).double_click_input(coords=(0,0, 1366,768))
                                        
                                        #Definindo o temporizador
                                        segundos = 60
                                        intervalo = 1

                                        while not pyautogui.locateOnScreen('imagens/S-1200.png', region=(0,0, 1366,768)) or pyautogui.locateOnScreen('imagens/S-1210.png', region=(0,0, 1366,768)):
                                            try:
                                                window.child_window(title="&Atualizar", class_name="Button", found_index=0).double_click_input(coords=(0,0, 1366,768))
                                                print("nao tem pagamentos")

                                                #Temporizador
                                                if segundos>0:
                                                    print(segundos)
                                                    segundos-=intervalo
                                                    sleep(intervalo)
                                                if segundos == 0:
                                                    print('Finalizado')
                                                    break
                                                elif segundos < 0:
                                                    print('Finalizado')
                                                    break

                                            except:
                                                pass

                                        while pyautogui.locateOnScreen('imagens/S-1200.png', region=(0,0, 1366,768)) or pyautogui.locateOnScreen('imagens/S-1210.png', region=(0,0, 1366,768)):
                                            try:
                                                window.child_window(title="&Atualizar", class_name="Button", found_index=0).double_click_input(coords=(0,0, 1366,768))
                                                print("tem pagamentos")
                                                time.sleep(1)
                                            except:
                                                pass
                                            
                                        else:
                                            window.child_window(title="&Fechar", class_name="Button")

                                        pywinauto.keyboard.send_keys('{ESC}')
                                        pywinauto.keyboard.send_keys('{ESC}')
                                        pywinauto.keyboard.send_keys('{ESC}')
                                        time.sleep(0.5)

                                        # 3 - Depois de enviado, ir em "Relatórios" -> "Demonstrativo de INSS Folha..."
                                            # 	1 - Ok
                                            # 	2 - Checar se a coluna "Situação" está com o status de "Enviado"
                                            # 	3 - Checar se as colunas "Base de Cálculo" e "Valor INSS" da tabela "Valor INSS Sistema" coincidem com os valores da tabela "Valor INSS eSocial".
                                            # 	4 - Esc
                                            # 	5 - Esc

                                        # TOOLBAR DO BOTÃO DE RELATÓRIOS
                                        window2.child_window(title="Relatórios", control_type="MenuItem").click_input(coords=(0,0, 1366,768))
                                        window2.child_window(title="e-Social", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))
                                        window2.child_window(title="Demonstrativo de INSS Folha e INSS eSocial", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))
                                        window.child_window(title="Competência:", class_name="Button", found_index=0).wait('visible', timeout=20).click_input(coords=(150,10, 1366,768))
                                        pyautogui.hotkey('ctrl','v')
                                        time.sleep(0.5)
                                        window.child_window(title="Competência:", class_name="Button", found_index=0).click_input(coords=(250,10, 1366,768))
                                        pyautogui.hotkey('ctrl','v')
                                        window.child_window(title="&OK", class_name="Button").click_input(coords=(0,0, 1366,768))

                                        # ------ Salvar o demonstrativo em uma planilha excel em um local padrão, abrir com python e comparar as colunas
                                        window.child_window(title="none", class_name="FNUDO3190", found_index=4).wait('ready', timeout=120).click_input(coords=(0,0, 1366,768))

                                        #Espera 30seg até que a janela 'Salvar o relatório' seja verdadeira
                                        app.window(title_re = 'Salvar o Relatório').wait('visible', timeout=60, retry_interval=1)

                                        salvar_relatorio = app.window(title_re = 'Salvar o Relatório', class_name = 'FNWNS3190')

                                        if salvar_relatorio.exists():
                                            salvar_relatorio.child_window(title="Tipo:", class_name="Button").click_input(coords=(100,5, 1366,768))
                                            pywinauto.keyboard.send_keys('{DOWN}')
                                            pywinauto.keyboard.send_keys('{ENTER}')
                                            salvar_relatorio.child_window(title="...", class_name="Button").click_input(coords=(0,0, 1366,768))

                                        window2.child_window(title="Área de Trabalho", control_type="Button").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))
                                        window2.child_window(title="Nome:", auto_id="1148", control_type="Edit").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))
                                        pyautogui.write("Dados INSS")
                                        window2.child_window(title="Salvar", auto_id="1", control_type="Button").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))

                                        window = app.top_window()
                                        window.child_window(title="&Salvar", class_name="Button").click_input(coords=(0,0, 1366,768))

                                        # ----- Verificação se os dados conferem através da planilha

                                        data_atual = date.today()
                                        month, year = (data_atual.month-1, data_atual.year) if data_atual.month != 1 else (12, data_atual.year-1)
                                        data_atual = data_atual.replace(day=1, month=month, year=year)
                                        competencia = data_atual.strftime('%d/%m/%Y')

                                        tabela = pd.read_excel(f"C:\\Users\\{usuario}\\Desktop\\Dados INSS.xls", engine='xlrd')

                                        #SISTEMA
                                        sistema_base = tabela.loc[tabela["competencia"]==competencia,"base_inss"]
                                        sistema_val = tabela.loc[tabela["competencia"]==competencia,"valor_inss"]


                                        # ------ ESOCIAL
                                        esocial_base = tabela.loc[tabela["competencia"]==competencia, "cp_base_inss_esocial"]
                                        esocial_val = tabela.loc[tabela["competencia"]==competencia, "cp_valor_calculado_inss_esocial"]

                                        # ------ COMPARAÇÃO
                                        base = (np.array_equal(sistema_base, esocial_base)) #compara se os dois arrays são iguais (base calculo sistema, base calculo esocial)
                                        valor = (np.array_equal(sistema_val, esocial_val)) # valor inss (sistema), valor calculado inss (inss)
                                        
                                        if base and valor:
                                            pyautogui.alert(text=f'Conferência realizada, os valores do INSS da Empresa {cod} coincidem. ', title='Automatização eSocial', button='OK', timeout="15000")
                                            window = app.FNWND3190
                                            window.set_focus()
                                            time.sleep(3)
                                            pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                                            pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia
                                    
                                            window2.child_window(title="Relatórios", control_type="MenuItem").click_input(coords=(0,0, 1366,768))
                                            window2.child_window(title="e-Social", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))
                                            window2.child_window(title="Demonstrativo de FGTS Folha e FGTS eSocial", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))

                                            window = app.top_window()
                                            window.child_window(title="Competência:", class_name="Button", found_index=0).wait('visible', timeout=20).click_input(coords=(150,10, 1366,768))
                                            pc.copy(data2)
                                            pyautogui.hotkey('ctrl','v')
                                            time.sleep(0.5)
                                            window.child_window(title="Competência:", class_name="Button", found_index=0).click_input(coords=(250,10, 1366,768))
                                            pyautogui.hotkey('ctrl','v')
                                            window.child_window(title="&OK", class_name="Button").click_input(coords=(0,0, 1366,768))

                                            # ------ Verificar se os dados conferem através da planilha

                                            # ------ Salvar o demonstrativo em uma planilha excel em um local padrão, abrir com python e comparar as colunas
                                            window.child_window(title="none", class_name="FNUDO3190", found_index=4).wait('ready', timeout=120).click_input(coords=(0,0, 1366,768))

                                            #Espera 30seg até que a janela 'Salvar o relatório' seja verdadeira
                                            app.window(title_re = 'Salvar o Relatório').wait('visible', timeout=60, retry_interval=1)

                                            salvar_relatorio = app.window(title_re = 'Salvar o Relatório', class_name = 'FNWNS3190')

                                            if salvar_relatorio.exists():
                                                salvar_relatorio.child_window(title="Tipo:", class_name="Button").click_input(coords=(100,5, 1366,768))
                                                pywinauto.keyboard.send_keys('{DOWN}')
                                                pywinauto.keyboard.send_keys('{ENTER}')
                                                salvar_relatorio.child_window(title="...", class_name="Button").click_input(coords=(0,0, 1366,768))

                                            window2.child_window(title="Área de Trabalho", control_type="Button").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))
                                            window2.child_window(title="Nome:", auto_id="1148", control_type="Edit").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))
                                            pyautogui.write("Dados FGTS")
                                            window2.child_window(title="Salvar", auto_id="1", control_type="Button").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))
                                            
                                            window = app.top_window()
                                            window.child_window(title="&Salvar", class_name="Button").click_input(coords=(0,0, 1366,768))

                                            # verificação se os dados conferem através da planilha

                                            tabela_2 = pd.read_excel(f"C:\\Users\\{usuario}\\Desktop\\Dados FGTS.xls", engine='xlrd') 
                                            #C:\\Users\\Suporte\\Desktop\\

                                            # SE A BASE_FGTS_DO_SISTEMA - BASE_DO_ESOCIAL FOR MENOR OU IGUAL A 20 CENTAVOS
                                            calculo_fgts = (tabela_2.loc[tabela_2["competencia"]==competencia,"base_fgts"]-tabela_2.loc[tabela_2["competencia"]==competencia, "cp_base_fgts_esocial"])*(-1) <= 0.20
                                            resultado_base = calculo_fgts.all()
                                            # A FUNÇÃO .all() testa se todos os resultados são TRUE

                                            calculo_valor = (tabela_2.loc[tabela_2["competencia"]==competencia, "valor_fgts"]-tabela_2.loc[tabela_2["competencia"]==competencia,"cp_valor_fgts_esocial"])*(-1) <= 0.20
                                            resultado_valor = calculo_valor.all()

                                            if resultado_base and resultado_valor:
                                                #print("Ok, a diferença é menor que 0.20, pode continuar!!!!")
                                                pyautogui.alert(text=f'Conferência realizada com sucesso!', title='Automatização eSocial', button='OK', timeout="15000")
                                                window = app.FNWND3190
                                                window.set_focus()
                                                time.sleep(3)
                                                pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                                                pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia
                                                # 5 - "Rotina" -> "Eventos Periódicos"
                                                    
                                                window2.child_window(title="Relatórios", control_type="MenuItem").click_input(coords=(0,0, 1366,768))
                                                window2.child_window(title="e-Social", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))
                                                window2.child_window(title="Eventos periódicos", control_type="MenuItem").wait('exists', timeout=15, retry_interval=1).click_input(coords=(0,0, 1366,768))

                                                window = app.top_window()
                                                window.child_window(title="Competência:", class_name="Button", found_index=0).click_input(coords=(100,10, 1366,768))
                                                pc.copy(data2)
                                                pyautogui.hotkey('ctrl','v')
                                                # 	1 - Clicar em "Fechamento dos Eventos Periódicos"
                                                window.child_window(title="S-1299 - Fechamento dos Eventos Periódicos", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                # 	2 - Clicar em "Enviar"
                                                window.child_window(title="En&viar", class_name="Button").wait('ready', timeout=15).click_input(coords=(0,0, 1366,768))
        
                                                window.child_window(title="En&viar").wait('exists', timeout=500, retry_interval=2)

                                                try:
                                                    atencao_esocial = app.window(title_re="Atenção", class_name="#32770")
                                                    app.window(title_re="Atenção", class_name="#32770")
                                                    if atencao_esocial.exists():
                                                        atencao_esocial.child_window(title="&Sim", class_name="Button").wait('ready', timeout=60).click_input(coords=(0,0, 1366,768))
                                                except:
                                                    pass
                                                
                                                try:
                                                    avisos_esocial = app.window(title_re="Avisos eSocial", class_name="FNWNS3190")
                                                    app.window(title_re="Avisos eSocial", class_name="FNWNS3190")
                                                    if avisos_esocial.exists(): 
                                                        dlg.child_window(title="&Sim", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                except:
                                                    pass

                                                #"OS EVENTOS PERIÓDICOS COM EXCEÇÃO DOS EVENTOS DE REMUNERAÇÃO E PAGAMENTOS DEVERÃO SER ENVIADOS PELA EMPRESA CENTRALIZADORA"
                                                atencao_periodicos = app.window(title_re='Atenção', class_name='FNWNS3190')

                                                if not atencao_periodicos.exists():

                                                    # 	3 - Ok
                                                    fechamentomsg = pyautogui.locateOnScreen('imagens/atencao_dctf.png', region=(0,0, 1366,768))
                                                    if fechamentomsg:
                                                        pyautogui.click(fechamentomsg)
                                                        pywinauto.keyboard.send_keys('{ENTER}')
                                                    
                                                    if enviado_sucesso.exists():
                                                        time.sleep(2)
                                                        dlg.child_window(title="OK", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))

                                                        window.child_window(title="&Painel Pendências", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                        window.child_window(class_name="PBTabControl32_100").click_input(coords=(550,10, 1366,768))

                                                        # 	4 - Conferir no "Painel de Pendências" o status do processamento clicando em atualizar até sumir.
                                                        # se o evento ainda estiver sendo processado, identifique a imagem "fechamento" e clique em atualizar até q ela suma
                                                        while pyautogui.locateOnScreen('imagens/fechamento.png', region=(0,0, 1366,768)):
                                                            try:
                                                                window.child_window(title="&Atualizar", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))
                                                            except:
                                                                pass
                                                            time.sleep(1)

                                                        window.child_window(title="&Validados", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                        # 	5 - Conferir na aba "Periódicos" se aparece "Fechamento" na coluna "Evento" (aqui será feito a filtragem do mês/ano do fechamento e o donwload da planilha para pesquisar por FECHAMENTO nela)
                                                        window.child_window(class_name="PBTabControl32_100", found_index=0).wait('visible', timeout=60).click_input(coords=(150,10, 1366,768))
                                                        window.child_window(title="&Filtrar", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                        
                                                        dlg = app.window(title="Filtro de Eventos", class_name="FNWNS3190")
                                                        app.window(title="Filtro de Eventos", class_name="FNWNS3190").wait('visible', timeout=15)
                                                        time.sleep(1)
                                                        #recuperando o mês atual
                                                        pc.copy(data)
                                                        dlg.child_window(title="Período:", class_name="Button").double_click_input(coords=(70,8, 1366,768))
                                                        pyautogui.hotkey('ctrl','v')
                                                        time.sleep(1)
                                                        dlg.child_window(title="Período:", class_name="Button").double_click_input(coords=(170,8, 1366,768))
                                                        pyautogui.hotkey('ctrl','v')
                                                        dlg.child_window(title="&Ok", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                        
                                                        window.child_window(title="&Relatório", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                        
                                                        # ------ Salvar o demonstrativo em uma planilha excel em um local padrão, abrir com python e comparar as colunas
                                                        window.child_window(title="none", class_name="FNUDO3190", found_index=4).wait('ready', timeout=30).click_input(coords=(0,0, 1366,768))

                                                        #Espera 30seg até que a janela 'Salvar o relatório' seja verdadeira
                                                        app.window(title_re = 'Salvar o Relatório').wait('visible', timeout=30, retry_interval=1)

                                                        salvar_relatorio = app.window(title_re = 'Salvar o Relatório', class_name = 'FNWNS3190')

                                                        if salvar_relatorio.exists():
                                                            salvar_relatorio.child_window(title="Tipo:", class_name="Button").click_input(coords=(100,5, 1366,768))
                                                            pywinauto.keyboard.send_keys('{DOWN}')
                                                            pywinauto.keyboard.send_keys('{ENTER}')
                                                            salvar_relatorio.child_window(title="...", class_name="Button").click_input(coords=(0,0, 1366,768))

                                                        window2.child_window(title="Área de Trabalho", control_type="Button").click_input(coords=(0,0, 1366,768))
                                                        window2.child_window(title="Nome:", auto_id="1148", control_type="Edit").click_input(coords=(0,0, 1366,768))
                                                        pyautogui.write("RelatorioValidados")
                                                        window2.child_window(title="Salvar", auto_id="1", control_type="Button").click_input(coords=(0,0, 1366,768))
                                                        
                                                        window = app.top_window()
                                                        window.child_window(title="&Salvar", class_name="Button").click_input(coords=(0,0, 1366,768))
                                                        
                                                        # verificação se os dados conferem através da planilha

                                                        tabela_validados = pd.read_excel(f"C:\\Users\\{usuario}\\Desktop\\RelatorioValidados.xls", engine='xlrd')
                                                        
                                                        # data2 = date.today()
                                                        # data2 = data2.strftime(str("%d") + "/" + str("%m") + "/" + str("%Y"))

                                                        # tabela_3 = tabela_validados.loc[tabela_validados["data_conclusao"]==data2,"evento"]
                                                        #print(tabela_3)
                                                        
                                                        for coluna in tabela_validados.columns:
                                                            fechamento_excel = tabela_validados["evento"]=="S-1299 Fechamento"
                                                            
                                                        if fechamento_excel.any():
                                                            procedimento = "Fechamento enviado com sucesso!"
                                                            pyautogui.alert(text='Conferência realizada, o Fechamento foi enviado com sucesso!. ', title='Automatização eSocial', button='OK', timeout="10000")
                                                            open_google()
                                                            window = app.top_window()
                                                            window.set_focus()
                                                            
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            
                                                        else:

                                                            procedimento = "Fechamento não enviado"
                                                            
                                                            open_google()

                                                            alerta.show()
                                                            pyautogui.alert(text='Alerta! O Fechamento não foi realizado.', title='Automatização eSocial', button='OK', timeout="15000")                       
                                                            window = app.top_window()
                                                            window.set_focus()

                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')
                                                            pywinauto.keyboard.send_keys('{ESC}')

                                                else:
                                                    procedimento = "Aguardando o envio pela empresa centralizadora"

                                                    atencao_periodicos.child_window(title='Os eventos periódicos com exceção dos eventos de remuneração e pagamentos deverão ser enviados pela empresa centralizadora.', class_name='Edit').click_input(coords=(260,80, 1366,768))
                                                    alerta.show()
                                                    open_google()
                                                    window = app.top_window()
                                                    window.set_focus()
                                                    time.sleep(3)
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
                                                
                                                open_google()
                                                window = app.top_window()
                                                window.set_focus()
                                                time.sleep(3)
                                                pywinauto.keyboard.send_keys('{ESC}') #saindo do painel de inss
                                                pywinauto.keyboard.send_keys('{ESC}') # fechar a tela de competencia
                                    
                                        else:
                                            #print("Erro: OS DADOS NÃO CONFEREM")
                                            # sairDominio()
                                            procedimento = "INSS"
                                            
                                            # exibindo notificação de ALERTA
                                            mensagem.show()

                                            open_google()

                                            time.sleep(3)
                                            window = app.top_window()
                                            window.set_focus()
                                            pywinauto.keyboard.send_keys('{ESC}')
                                            pywinauto.keyboard.send_keys('{ESC}')
                                
                                    else:
                                        #print("Erro: Possui evento invalidado")
                                        # sairDominio()
                                        procedimento = "Evento Invalidado no Painel de Pendências"
                                        
                                        # exibindo notificação de ALERTA
                                        mensagem.show()

                                        open_google()

                                        time.sleep(3)
                                        window = app.top_window()
                                        window.set_focus()
                                        pywinauto.keyboard.send_keys('{ESC}')
                                        pywinauto.keyboard.send_keys('{ESC}')  

                            else:
                                procedimento = "Competencia fechada no eSocial"

                                try:
                                    dlg.child_window(title="&OK", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))
                                except:
                                    pass

                                open_google()
                                
                                alerta.show()

                                time.sleep(3)
                                window = app.top_window()
                                window.set_focus()
                                pywinauto.keyboard.send_keys('{ESC}')
                                pywinauto.keyboard.send_keys('{ESC}')

                        else:
                            procedimento = "Não existem pagamentos nesta competência."

                            try:
                                dlg.child_window(title_re='Não existem pagamentos nesta competência para envio.', class_name='Static').click_input(coords=(230,55, 1366,768))
                            except:
                                pass
                                                
                            open_google()
                            
                            alerta.show()
                            time.sleep(3)
                            window = app.top_window()
                            window.set_focus()
                            pywinauto.keyboard.send_keys('{ESC}')
                            pywinauto.keyboard.send_keys('{ESC}')
                            pywinauto.keyboard.send_keys('{ESC}')

                    else:
                        procedimento = "Não existem cálculos nesta competência."

                        try:
                            dlg.child_window(title_re='Não existem cálculos nesta competência para envio.', class_name='Static').wait('visible', timeout=3).click_input(coords=(230,55, 1366,768))
                        except:
                            pass
                                
                        open_google()
                        
                        alerta.show()
                        time.sleep(3)
                        window = app.top_window()
                        window.set_focus()
                        pywinauto.keyboard.send_keys('{ESC}')
                        pywinauto.keyboard.send_keys('{ESC}')

                else:
                    procedimento = "Rubrica Invalidada."

                    try:
                        dlg = app.window(title='Avisos Rubricas')
                        dlg.child_window(title="&Fechar", class_name="Button", found_index=0).wait('visible', timeout=15).click_input(coords=(0,0, 1366,768))
                    except:
                        pass
                            
                    open_google()
                    
                    alerta.show()
                    time.sleep(3)
                    window = app.top_window()
                    window.set_focus()
                    pywinauto.keyboard.send_keys('{ESC}')
                    pywinauto.keyboard.send_keys('{ESC}')  
                    

            else:
                acesso.child_window(title="OK", class_name="Button", found_index=0).click_input(coords=(0,0, 1366,768))
                pywinauto.keyboard.send_keys('{ESC}') 
                pywinauto.keyboard.send_keys('{ESC}') 
            

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
            wb.drop(0, inplace=True)
            wb.reset_index(drop=True, inplace=True)
            wb.to_excel('parametrizacao.xlsx', index=False)             

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