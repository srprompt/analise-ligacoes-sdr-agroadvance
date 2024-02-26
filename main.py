import requests
import datetime

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import time
import os
import markdown
import openai
import unicodedata
import re
import csv
from bs4 import BeautifulSoup
import pymsteams
from tqdm import tqdm
import json
from dotenv import load_dotenv
from openai import AzureOpenAI
import whisper

load_dotenv()

options = Options()
options.set_preference('browser.download.folderList', 2)  # custom location
options.set_preference('browser.download.manager.showWhenStarting', False)
options.set_preference('browser.download.dir', "C:\\Projetos Agroadvance\\Azure-whisper\\meetime-audios")
options.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/octet-stream')  # MIME do tipo de arquivo que não deve perguntar

browser = webdriver.Firefox(options=options)

##################### VARIÁVEIS GLOBAIS
login = 0
checa_limite = 0
total_segundos = 60
qtd_audios = 0
nome_sdr = ''
nome_lead = ''
email_lead = ''
link_chamada = ''


caminho_origem = "./meetime-audios"

#AZURE CONFIG
client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_KEY"),  
    api_version="2023-12-01-preview",
    azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
)

def acessoGoogleSheets(respostas_csv):
    import os.path
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError

    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = os.getenv("GOOGLE_PLANILHA_ID")
    SAMPLE_RANGE_NAME = 'dbMeetime.csv!A1:N'

    creds = None

    #print(f"Respostas recebidas na funcao de enviar para o Google Sheets: {respostas_csv}")

    # Se o token.json existir, tenta usá-lo
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # Se não houver credenciais válidas, solicita ao usuário que faça login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Salva as credenciais para a próxima execução
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Construir o serviço Google Sheets
    service = build('sheets', 'v4', credentials=creds)

    # Obter a instância da planilha
    sheet = service.spreadsheets()

    # Obter os valores atuais na planilha
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    with open("valores.json", 'w') as file:
        json.dump(values, file)

    # Determinar a última linha ocupada
    last_row = len(values) + 1

    # Adicionar/editar valores no Google Sheets abaixo da última linha ocupada
    valores_adicionar = [respostas_csv]

    range_to_update = f'dbMeetime.csv!A{last_row}:O{last_row + len(valores_adicionar) - 1}'
    
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=range_to_update,
                                    valueInputOption="USER_ENTERED",
                                    body={"values": valores_adicionar}).execute()

    print(f"Google Sheets: {result.get('updatedCells')} células atualizadas.")

def processaTranscricao(data_csv):

    # CARREGA O ARQUIVO DE TRANSCRIÇÃO DO AUDIO
    with open("transcript.txt", "r") as txt_file:
        transcript_text = txt_file.read()   

    messages = [
    {
        'role': 'system',
        'content': "Responda as questões a seguir com base na transcrição, respeitando os limites de resposta\
                    definidos em cada uma das perguntas e o formato HTML estabelecido.\
                    Exemplo de formato desejado:\
                        <h2>Resumo</h2><p>Coloque o resumo aqui</p>\
                        <hr><h2>Perguntas Exploratórias</h2>\
                        <div id='cargo'><h3>Qual o cargo do cliente atualmente?</h3><p>Produtor Rural</p></div>\
                        <div id='desafio'><h3>Qual desafio ou problema o cliente enfrenta atualmente?</h3><p>Resposta</p>\
                        <div id='transicao'><h3>O cliente busca transição de carreira?</h3><p>Sim</p></div>\
                        <div id='capacitacao'><h3>Qual o motivo de busca pela capacitação?</h3><p>Resposta</p></div>\
                        <div id='area'><h3>Qual a área plantada em hectare ou alqueire que o cliente possui?</h3><p>Valor em hectare ou alqueire</p></div>\
                        <div id='tempo'><h3>Quanto tempo de atuação o cliente possui na área?</h3><p>Quantidade em anos</p></div>\
                        <div id='escolaridade'><h3>Qual é o seu nível de escolaridade?</h3><p>Superior</p></div>\
                        <div id='empresa'><h3>Em qual empresa o cliente trabalha?</h3><p>Syngenta</p></div>\
                    1. Resumo: \
                        - Resuma a ligação focando nos pontos mais importantes comentados pelo cliente.\
                    2. Perguntas Exploratórias:\
                        Caso não tenha a resposta apenas coloque `Não mencionado`\
                        - Com o que o cliente trabalha atualmente?(Utilize APENAS as seguintes opções:\
                            Analista, Analista de Dados, Assistente Técnico Agrícola,\
                            Consultor Agrícola, Coordenador, Coordenador Agrícola, Desenvolvedor de Mercado,\
                            Diretor, Engenheiro Agrônomo, Especialista, Estudante, Gerente, Gerente Comercial,\
                            Gerente de Fazenda, Operador de Máquinas, Pesquisa e Desenvolvimento, Produtor Rural,\
                            RTV/RTC, Sucessor, Supervisor, Supervisor Agrícola, Técnico Agrícola, Técnico de Informação ou Outros)\n\
                        - Descreva de forma suscinta os desafios ou problemas que o cliente enfrenta atualmente.\
                            Por exemplo, se o cliente enfrenta dificuldades na gestão, a resposta desejada seria `Dificuldades na gestão`.\
                        - O cliente busca transição de carreira? (Sim, não ou não mencionado)\
                        - Descreva de forma suscinta o motivo de busca pela capacitação?\
                            Por exemplo, se o cliente quer se aperfeiçoar, a resposta desejada seria `Aperfeiçoamento`.\
                        - Qual a área plantada em hectare ou alqueire que o cliente possui? Responda APENAS\
                            se o cliente mencionar a área plantada.\
                        - Quanto tempo de atuação o cliente possui na área do agronegócio? Descreva o valor usando apenas números, em anos.\
                            Por exemplo, se o cliente possui sete anos de experiência, a resposta desejada seria '7 anos'.\
                        - Qual é o nível de escolaridade do cliente? Por favor, UTILIZE APENAS AS OPÇÕES A SEGUIR:\
                            Superior, Técnico, Ensino Médio ou Básico.\
                        - Qual o nome da empresa que o cliente trabalha?\
                    "
    },
    {
        'role': 'user',
        'content': f'Por favor, responda APENAS o que foi solicitado. A seguir segue a transcrição da chamada:\n\n{transcript_text}.'
    },
    ]

    resposta = client.chat.completions.create(
        model="gpt-35t-16k-agroadvance",
        messages=messages,
        temperature=0.4,
        max_tokens=1000
    )

    resposta_texto = resposta.choices[0].message.content

    tokens_usados = resposta.usage.total_tokens
    print(f"Total de Tokens P&R = {tokens_usados}")
    print("--------------------------------------\n")

##################### BEAUTIFUL SOUP
    soup = BeautifulSoup(resposta_texto, 'html.parser')

    # Encontrar a tag h2 com o texto 'Resumo'
    resumo_tag = soup.find('h2', string='Resumo')
    cargo_tag = soup.find(id='cargo')
    desafio_tag = soup.find(id='desafio')
    transicao_tag = soup.find(id='transicao')
    capacitacao_tag = soup.find(id='capacitacao')
    area_tag = soup.find(id='area')
    tempo_tag = soup.find(id='tempo')
    escolaridade_tag = soup.find(id='escolaridade')
    empresa_tag = soup.find(id='empresa')

 
    if resumo_tag:
        p_tag = resumo_tag.find_next('p')
        if p_tag:
            b_resumo = p_tag.get_text(strip=True)  
        else:
            print('Tag <p> não encontrada após <h2>Resumo</h2>.') 
            
    if cargo_tag:
        p_tag2 = cargo_tag.find_next('p')
        if p_tag2:
            b_cargo = p_tag2.get_text(strip=True)
            if b_cargo == 'Analista' or b_cargo == 'Analista de Dados' or b_cargo == 'Assistente Técnico Agrícola' or b_cargo == 'Consultor Agrícola' or b_cargo == 'Coordenador' or b_cargo == 'Coordenador Agrícola' or b_cargo == 'Desenvolvedor de Mercado' or b_cargo == 'Diretor' or b_cargo == 'Especialista' or b_cargo == 'Gerente' or b_cargo == 'Gerente Comercial' or b_cargo == 'Gerente de Fazenda' or b_cargo == 'Operador de Máquinas' or b_cargo == 'Pesquisa e Desenvolvimento' or b_cargo == 'Produtor Rural' or b_cargo == 'RTV/RTC' or b_cargo == 'Sucessor' or b_cargo == 'Supervisor' or b_cargo == 'Técnico Agrícola' or b_cargo == 'Técnico de Informação' or b_cargo == 'Outros':
                b_cargo = b_cargo
            else:
                messages = [
                {
                    'role': 'system',
                    'content': f"O cliente relatou seu cargo como {b_cargo}, e precisamos categorizá-lo em uma das opções padrões a seguir:\
                                Analista, Analista de Dados, Assistente Técnico Agrícola,\
                                Consultor Agrícola, Coordenador, Coordenador Agrícola, Desenvolvedor de Mercado,\
                                Diretor, Engenheiro Agrônomo, Especialista, Estudante, Gerente, Gerente Comercial,\
                                Gerente de Fazenda, Operador de Máquinas, Pesquisa e Desenvolvimento, Produtor Rural,\
                                RTV/RTC, Sucessor, Supervisor, Supervisor Agrícola, Técnico Agrícola, Técnico de Informação.\
                                Se a opção exata não existir, escolha a opção mais próxima ou coloque Outros.\
                                Responda apenas o cargo, sem nenhum acrescimo."
                }
                ]
                cargo = client.chat.completions.create(
                        model="gpt35t-agroadvance",
                        messages=messages,
                        temperature=0.3,
                        max_tokens=100
                    )
                print(f"\nTokens para Cargo: {cargo.usage.total_tokens}")
                b_cargo = cargo.choices[0].message.content       
        else:
            print('Tag <p> não encontrada após <h3>Cargo</h3>.')
            
    if desafio_tag:
        p_tag3 = desafio_tag.find_next('p')
        if p_tag3:
            b_desafio = p_tag3.get_text(strip=True)
            if b_desafio.startswith('O cliente'):
                messages = [
                {
                    'role': 'system',
                    'content': f"O cliente relatou seu desafio como {b_desafio}, porém precisamos\
                        melhorar essa resposta. Por favor, descreva de forma suscinta os desafios\
                        ou problemas que o cliente enfrenta atualmente.\
                        Por exemplo, se o cliente enfrenta dificuldades na gestão, a resposta desejada seria `Dificuldades na gestão`."
                }
                ]
                desafio = client.chat.completions.create(
                        model="gpt35t-agroadvance",
                        messages=messages,
                        temperature=0.3,
                        max_tokens=100
                    )
                print(f"\nTokens para Desafio: {desafio.usage.total_tokens}")
                b_desafio = desafio.choices[0].message.content 
        else:
            print('Tag <p> não encontrada após <h3>Desafio</h3>.')
    
    if transicao_tag:
        p_tag4 = transicao_tag.find_next('p')
        if p_tag4:
            b_transicao = p_tag4.get_text(strip=True)
        else:
            print('Tag <p> não encontrada após <h3>Transicao</h3>.')
            
    if capacitacao_tag:
        p_tag5 = capacitacao_tag.find_next('p')
        if p_tag5:
            b_capacitacao = p_tag5.get_text(strip=True)
            if b_capacitacao.startswith('O cliente'):
                messages = [
                {
                    'role': 'system',
                    'content': f"O cliente relatou seu motivo de capacitação como: {b_capacitacao}, porém precisamos\
                        melhorar essa resposta. Por favor, descreva de forma suscinta o motivo da\
                        necessidade de capacitação atualmente.\
                        Por exemplo, se o cliente quer se aperfeiçoar, a resposta desejada seria `Aperfeiçoamento`."
                }
                ]
                capacitacao = client.chat.completions.create(
                        model="gpt35t-agroadvance",
                        messages=messages,
                        temperature=0.3,
                        max_tokens=100
                    )
                print(f"\nTokens para Capacitação: {capacitacao.usage.total_tokens}")
                b_capacitacao = capacitacao.choices[0].message.content
        else:
            print('Tag <p> não encontrada após <h3>Capacitacao</h3>.')
    
    if area_tag:
        p_tag6 = area_tag.find_next('p')
        if p_tag6:
            b_area = p_tag6.get_text(strip=True)
        else:
            print('Tag <p> não encontrada após <h3>Area</h3>.')
            
    if tempo_tag:
        p_tag7 = tempo_tag.find_next('p')
        if p_tag7:
            b_tempo = p_tag7.get_text(strip=True)
        else:
            print('Tag <p> não encontrada após <h3>Tempo</h3>.')
    
    if escolaridade_tag:
        p_tag8 = escolaridade_tag.find_next('p')
        if p_tag8:
            b_escolaridade = p_tag8.get_text(strip=True)
            if b_escolaridade == 'Superior' or b_escolaridade == 'Técnico' or b_escolaridade == 'Ensino Médio':
                b_escolaridade = b_escolaridade
            else:
                messages = [
                {
                    'role': 'system',
                    'content': f"Sua função é associar o curso do cliente com o nível de escolaridade.\
                                Por exemplo, se o cliente é arquiteto, você deve colocar Superior.\
                                ###\
                                Portanto, você deve associar esse dado: {b_escolaridade}\
                                com o nível de escolaridade e me responder somente se é\
                                Superior, Técnico, Ensino Médio ou Básico. Sem nenhum acrescimo.\
                                Caso não seja possível associar, responda `Não mencionado`.\
                                ###"
                }
                ]

                escolaridade = client.chat.completions.create(
                    model="gpt35t-agroadvance",
                    messages=messages,
                    temperature=0.3,
                    max_tokens=100
                )
                print(f"\nTokens para Escolaridade: {escolaridade.usage.total_tokens}")
                b_escolaridade = escolaridade.choices[0].message.content
                
        else:
            print('Tag <p> não encontrada após <h3>Escolaridade</h3>.')
    
    if empresa_tag:
        p_tag9 = empresa_tag.find_next('p')
        if p_tag9:
            b_empresa = p_tag9.get_text(strip=True)
        else:
            print('Tag <p> não encontrada após <h3>Empresa</h3>.')
######################
    
    bloco_formatado = ""

    with open('transcript.txt', 'r', encoding='utf-8') as file:
        transcript = file.read()
        file.close()

        resumo_formatado = (markdown.markdown(f"# SDR: {nome_sdr}\n\n--------------------\n\nINFORMAÇÕES DO LEAD\n=======================\nNOME: {nome_lead}\n=======================\nE-MAIL: {email_lead}\n=======================\nURL MEETIME:<a href='{link_chamada}'> {link_chamada}</a>\n\n--------------------\n\n"))
        bloco_formatado += resumo_formatado
        respostas_csv = [data_csv, nome_sdr, email_lead, nome_lead]
        try:
            respostas_csv.append(b_cargo)
        except:
            print("\n**Erro ao obter cargo")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_desafio)
        except:
            print("\n**Erro ao obter desafio")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_transicao)
        except:
            print("\n**Erro ao obter transicao")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_capacitacao)
        except:
            print("\n**Erro ao obter capacitacao")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_area)
        except:
            print("\n**Erro ao obter area")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_tempo)
        except:
            print("\n**Erro ao obter tempo")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_escolaridade)
        except:
            print("\n**Erro ao obter escolaridade")
            respostas_csv.append("Não mencionado")
        try:
            respostas_csv.append(b_empresa)
        except:
            print("\n**Erro ao obter empresa")
            respostas_csv.append("Não mencionado")

        respostas_csv.append(link_chamada)
        respostas_csv.append(transcript)
        respostas_csv.append(tokens_usados)
        

        resposta_formatada = (markdown.markdown(f"{resposta_texto}"))
        bloco_formatado += resposta_formatada
    
    with open('dbMeetime.csv', 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(respostas_csv[0:])
        file.close()

    acessoGoogleSheets(respostas_csv)
    print("--------------------------------------\n")
    

    # Salva as perguntas e respostas em um arquivo de texto
    with open("resumo.txt", "w") as arquivo:
        arquivo.write(resumo_formatado)
        
    myTeamsMessage = pymsteams.connectorcard(os.getenv("WEBHOOK_URL_OFICIAL"))
    myTeamsMessage.text(bloco_formatado)
    myTeamsMessage.send()
    print("\nTeams: Mensagem enviada com sucesso!")
    print("--------------------------------------\n")
    
    return b_area, b_capacitacao, b_cargo, b_desafio, b_empresa, b_escolaridade, b_tempo, b_transicao

def dadosMeetime(res_area, res_capacitacao, res_cargo, res_desafio, res_empresa, res_escolaridade, res_tempo, res_transicao):
    try:
        pagina_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,'a[ui-sref="mt.app.prospector.lead.timeline({id:flowLeadDetails.id})"]'))
        )
        pagina_lead.click()
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível clicar na pagina do lead. Erro:", str(e))  
             
    try:
        ### MUDANÇA DE CONTEXTO DE ABA
        handles = browser.window_handles
        nova_aba_handle = handles[-1]
        browser.switch_to.window(nova_aba_handle)
                                    
        editar_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/container/div/div/div/div/div[2]/div[1]/div/div/div[3]/div[1]/div/a"))
        )
        editar_lead.click()                            
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível clicar em editar lead. Erro:", str(e))
    
    try:
        transcricao_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[1]/div/div/span/div/textarea'))
        )
        transcricao_lead.send_keys(formatted_transcript)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher a transcrição do lead. Erro:", str(e))
                    
    try:
        impeditivo_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[2]/div/div/span/div/input'))
        ) 
        impeditivo_lead.send_keys(res_desafio)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher o impeditivo do lead. Erro:", str(e))
                                
    try:
        transicao_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[3]/div/div/span/div/input'))
        )
        transicao_lead.send_keys(res_transicao)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher a transição do lead. Erro:", str(e))
                                
    try:
        capacitacao_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[4]/div/div/span/div/input'))
        )
        capacitacao_lead.send_keys(res_capacitacao)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher o motivo da capacitação do lead. Erro:", str(e))
                                
    try:
        area_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[5]/div/div/span/div/input'))
        )
        area_lead.send_keys(res_area)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher a área plantada do lead. Erro:", str(e))

    try:
        tempo_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[6]/div/div/span/div/input'))
        )
        tempo_lead.send_keys(res_tempo)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher o tempo de atuação do lead. Erro:", str(e))
                                
    try:
        escolaridade_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[7]/div/div/span/div/input'))
        )
        escolaridade_lead.send_keys(res_escolaridade)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher a escolaridade do lead. Erro:", str(e))
                                
    try:
        empresa_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[8]/div/div/span/div/input'))
        )
        empresa_lead.send_keys(res_empresa)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher a empresa do lead. Erro:", str(e))
                                
    try:
        cargo_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/fieldset[3]/div[2]/div[9]/div/div/span/div/input'))
        )
        cargo_lead.send_keys(res_cargo)
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível preencher o cargo do lead. Erro:", str(e))
                                
    try:
        salvar_lead = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/container/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/form/div[2]/div/div/button[2]'))
        )
        #salvar_lead.send_keys(b_cargo)
        salvar_lead.click()
    except (TimeoutException, NoSuchElementException) as e:
        print("Não foi possível clicar em salvar do lead. Erro:", str(e))

def consultaAPIMeetime():
    data_atual = datetime.date.today()
    data_formatada = data_atual.strftime("%Y-%m-%d")
    #data_formatada = '2024-02-05'
    data_csv = str(data_atual)
    print(f"Data de Hoje (formato Meetime API): {data_formatada}")

    #&user_id=32333

    url = f"https://api.meetime.com.br/v2/calls?limit=100&start=0&started_after={data_formatada}&output=MEANINGFUL&status=CONNECTED"

    headers = {
        "accept": "application/json",
        "Authorization": os.getenv("MEETIME_AUTHORIZATION"),
        "Ocp-Apim-Subscription-Key": os.getenv("MEETIME_OCP_SUBSCRIPTION_KEY")
    }

    response = requests.get(url, headers=headers)
    data = response.json()
    
    return data, data_csv
    
##################### LOOP PRINCIPAL

while True:

    # ACESSA API MEETIME
    data, data_csv = consultaAPIMeetime()

    if 'data' in data:
        call_links = [call['call_link'] for call in data['data']]
        print(f"Ligações significativas: {len(call_links)}")
        qtd_audios = len(call_links)
    else:
        print("A resposta não possui a chave 'data'")
    
    links_json = 'links_visitados.json'

    try:
        with open(links_json, 'r') as file:
            links_visitados = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        links_visitados = []

    for link in call_links:
        if link not in links_visitados:
            link_chamada = link
            browser.get(link_chamada)
                
            time.sleep(1)
                
            if login == 0:
                email = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.ID, 'email')))
                #email = browser.find_element(By.ID, 'email')
                email.send_keys(os.getenv("MEETIME_EMAIL") + Keys.ENTER)
                
                senha = WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.ID, 'current-password')))
                #senha = browser.find_element(By.ID, 'current-password')
                senha.send_keys(os.getenv("MEETIME_PASS") + Keys.ENTER)
                login += 1
            
            # Botão três pontos
            botao = WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.XPATH, "//button[@class='btn btn-flat btn-ellipsis']")))
            botao.click()
            
            # Botão de download
            try:
                element = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[starts-with(@href,'https://dialer-audios.s3.amazonaws.com/')]")))
                element.click()
            except (TimeoutException, NoSuchElementException) as e:
                print("Não foi possível fazer o download da ligação", str(e))

            try:
                nome = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.XPATH, "//small[@class='text-muted']")))
                nome_sdr = nome.text
                print(f"SDR: {nome_sdr}")
            except (TimeoutException, NoSuchElementException) as e:
                print("Não foi possível obter o nome do SDR. Erro:", str(e))
                nome_sdr = '' 

            try:
                nomeLead = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.XPATH, "//li[@ng-show='flowLeadDetails.name']")))
                nome_lead = ''
                nome_lead = nomeLead.text
                print(f"Nome do lead: {nome_lead}")
            except (TimeoutException, NoSuchElementException) as e:
                print("Não foi possível obter o nome do Lead. Erro:", str(e))
                nome_lead = ''
            
            try:
                email = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.XPATH, "//li[@ng-show='flowLeadDetails.email']")))
                email_lead = ''
                email_lead = email.text
                print(f"Email do lead: {email_lead}")
            except (TimeoutException, NoSuchElementException) as e:
                print("Não foi possível obter o email do Lead. Erro:", str(e))
                email_lead = ''  
            
            time.sleep(10)
           
            try:
                if os.path.exists(caminho_origem) and os.path.isdir(caminho_origem):
                    # Verifique se a pasta existe e é uma pasta válida
                    for nome_arquivo in os.listdir(caminho_origem):
                        caminho_completo = os.path.join(caminho_origem, nome_arquivo)

                        if os.path.isfile(caminho_completo):
                            # Verifique se o caminho é um arquivo
                            #audio = open(caminho_completo, "rb")
                            
                            
                            '''
                            transcricao = openai.audio.transcriptions.create(
                                            model="whisper-1", 
                                            file=audio,
                                            response_format="text"
                                        )
                            '''
                            model = whisper.load_model("small")
                            resposta = model.transcribe(caminho_completo)
                            
                            transcricao = resposta["text"]
                            
                            formatted_transcript = transcricao

                            # Remove caracteres especiais e normaliza acentuações
                            formatted_transcript = unicodedata.normalize('NFKD', formatted_transcript).encode('ASCII', 'ignore').decode('utf-8')

                            # Remove quebras de linha desnecessárias
                            formatted_transcript = re.sub(r'\n+', '\n', formatted_transcript)

                            # Salva o texto transcrição em um arquivo de texto
                            with open("transcript.txt", "w") as txt_file:
                                txt_file.write(formatted_transcript)

                            print("\nTranscrição formatada e tratada salva no arquivo transcript.txt") 

                            #audio.close()

                            os.remove(caminho_completo)

                            links_visitados.append(link)
                            with open(links_json, 'w') as file:
                                json.dump(links_visitados, file)

                            if checa_limite < 8:
                                res_area, res_capacitacao, res_cargo, res_desafio, res_empresa, res_escolaridade, res_tempo, res_transicao = processaTranscricao(data_csv)
                                checa_limite += 1
                                print(f"\nAudios processados: {checa_limite}")
                                dadosMeetime(res_area, res_capacitacao, res_cargo, res_desafio, res_empresa, res_escolaridade, res_tempo, res_transicao)
                                time.sleep(2)

                            elif checa_limite == 8:
                                with tqdm(total=total_segundos, desc="Limite atingido, aguarde:") as pbar:
                                    for i in range(total_segundos):
                                        pbar.update(1)
                                        time.sleep(1)
                                    res_area, res_capacitacao, res_cargo, res_desafio, res_empresa, res_escolaridade, res_tempo, res_transicao = processaTranscricao(data_csv)
                                    checa_limite = 0  
                        else:
                            print(f"{caminho_completo} não é um arquivo.")
                            os.system('cls')
                            print(f"Ligações restantes: {len(call_links) - len(links_visitados)}")
                else:
                    print(f"A pasta {caminho_origem} não existe.")
            except Exception as e:
                print(f"Erro durante o processamento do link {link}: {e}")

    # Feche o navegador quando terminar
    browser.quit()
    login=0
    print("Aguardando 30 minutos para a próxima execução...")
    time.sleep(1800)
    
