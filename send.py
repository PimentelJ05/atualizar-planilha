import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import json

# Carregar credenciais do ambiente
GOOGLE_CREDENTIALS_PATH = os.getenv('GOOGLE_CREDENTIALS')
KOMMO_API_TOKEN = os.getenv('KOMMO_API_TOKEN')

# Configuração das credenciais do Google Sheets
def obter_credenciais_google():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_PATH, scope)
    client = gspread.authorize(creds)
    return client

# Função para enviar mensagem pelo Kommo
def enviar_mensagem_kommo(numero_cliente, mensagem):
    url = "https://creditoessencial.kommo.com/chats/"
    headers = {
        "Authorization": f"Bearer {KOMMO_API_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {
        "to": numero_cliente,
        "message": mensagem
    }
    response = requests.post(url, json=payload, headers=headers)
    return response.json()

# Dicionário de tradução de status
status_traducao = {
    "posted": "postado",
    "delivered": "entregue",
    "canceled": "cancelado",
    "received": "recebido"
}

# Função para processar a planilha e enviar mensagens
def processar_planilha():
    client = obter_credenciais_google()
    sheet = client.open_by_key("YOUR_SHEET_ID").sheet1  # Abre a planilha pelo ID
    dados = sheet.get_all_records()  # Obtendo todos os dados da planilha
    
    for linha in dados:
        numero_cliente = linha.get('Número do Cliente')
        id_pedido = linha.get('ID do Pedido')
        nome_cliente = linha.get('Nome do Cliente')
        status_pedido = linha.get('Status do Pedido')
        transportadora = linha.get('Transportadora')
        
        if numero_cliente:  # Verifica se o número do cliente está presente
            status_traduzido = status_traducao.get(status_pedido.lower(), status_pedido)  # Traduz o status
            mensagem = (f"Olá {nome_cliente},\n\n"
                        f"Seu pedido com ID {id_pedido} está com o status: {status_traduzido}.\n"
                        f"Transportadora: {transportadora}.\n\n"
                        f"Obrigado por escolher nossos serviços!")
            
            resposta = enviar_mensagem_kommo(numero_cliente, mensagem)
            print(f"Mensagem enviada para {numero_cliente}: {resposta}")

if __name__ == "__main__":
    processar_planilha()
