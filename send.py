import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests

# Configurações da Google Sheets API
SPREADSHEET_ID = '1mRNwU-h9EUxYmtG8WNfDLMI8fqWviKGYEERD4k5aegY'  # ID da sua planilha
RANGE_NAME = 'Sheet1!A:G'  # Ajuste o nome da aba e o intervalo conforme necessário

# Configurações da API do Kommo
KOMMO_API_URL = 'https://creditoessencial.kommo.com/api/v4/chats/messages'
KOMMO_TOKEN = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6IjhkYmQxOTlmZTE0ZTIzMDA5ZGFhZmQ5YWNhNWZiMjU4MjU0ZjNkNmIwYTExNTlkY2UyMTI1ZWQ5ZmU3MTY0ZTM2ODVhNDEzODUzMGUxZjllIn0.eyJhdWQiOiIzMWU5ZjgzNy1jOTlhLTQ5YTUtOGFkYS1mNzEwMmJiNmMzYzciLCJqdGkiOiI4ZGJkMTk5ZmUxNGUyMzAwOWRhYWZkOWFjYTVmYjI1ODI1NGYzZDZiMGExMTU5ZGNlMjEyNWVkOWZlNzE2NGUzNjg1YTQxMzg1MzBlMWY5ZSIsImlhdCI6MTcyNjIzNDg2NCwibmJmIjoxNzI2MjM0ODY0LCJleHAiOjE3MjYzMjEyNjQsInN1YiI6IjEwNDY2MDM1IiwiZ3JhbnRfdHlwZSI6IiIsImFjY291bnRfaWQiOjMyMDYxNDU1LCJiYXNlX2RvbWFpbiI6ImtvbW1vLmNvbSIsInZlcnNpb24iOjIsInNjb3BlcyI6WyJwdXNoX25vdGlmaWNhdGlvbnMiLCJmaWxlcyIsImNybSIsImZpbGVzX2RlbGV0ZSIsIm5vdGlmaWNhdGlvbnMiXSwiaGFzaF91dWlkIjoiMmQ1ODUzYjMtZGY2Ni00ZDgyLWExMGYtNDUwNDZjMDM3M2EwIiwiYXBpX2RvbWFpbiI6ImFwaS1nLmtvbW1vLmNvbSJ9.pzq1_nGT9y-pmbKXT5SMDWMmPs4x5GXq55fDVE3nsHjf8TlFt_w2XcUrQu5raezgH3vDYFM3LgMP79o6fLT7PPOKxbg7ilXTfuYmGfLKZgrzvqiwMyENaTIf_pT2x5iKvumc0gxaJyGWMM-U4y8HdQI6gptEf_NZoOrZP-WOpjeTfRJX-i2bX4LVaYKZMd0YBxFArRX9qOeKYL2B8AWXa70_TC_InnHVpK7R26CWgZlh77I2n_uzB9U-fOPgrH1LOsggNvK9hxXqWkk-Izzxdms_xmGROZ-QcWiYtQlZKi35vz_Aj7SkJTdpp-LIDcA0N9FiKnVImThsOpN1-ux_pQ'  # Substitua pelo seu token de API do Kommo

# Função para acessar os dados da planilha
def acessar_planilha():
    # Carrega as credenciais do GitHub Secrets
    google_credentials = json.loads(os.getenv('GOOGLE_CREDENTIALS'))
    
    # Autenticação usando credenciais de serviço a partir do JSON das credenciais
    creds = Credentials.from_service_account_info(google_credentials)
    service = build('sheets', 'v4', credentials=creds)

    # Chama a API do Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    return values

# Função para enviar mensagem via Kommo
def enviar_mensagem_kommo(telefone_cliente, mensagem):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {KOMMO_TOKEN}'
    }
    payload = {
        'contact': {
            'phone': telefone_cliente
        },
        'text': mensagem
    }

    response = requests.post(KOMMO_API_URL, json=payload, headers=headers)

    if response.status_code == 200:
        print(f'Mensagem enviada para {telefone_cliente}: {mensagem}')
    else:
        print(f'Erro ao enviar mensagem para {telefone_cliente}: {response.status_code} - {response.text}')

# Função principal para monitorar mudanças e enviar mensagens
def monitorar_mudancas():
    dados = acessar_planilha()

    # Exemplo de lógica para verificar mudanças e enviar mensagens
    for linha in dados:
        id_cliente, nome_cliente, id_pedido, status_pedido, transportadora, data_atualizacao, telefone_cliente = linha
        
        # Lógica para detectar mudança de status e enviar mensagem
        mensagem = f'Olá {nome_cliente}, o status do seu pedido mudou para: {status_pedido}.'
        enviar_mensagem_kommo(telefone_cliente, mensagem)

if __name__ == '__main__':
    monitorar_mudancas()
