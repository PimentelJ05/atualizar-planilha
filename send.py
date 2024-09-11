import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials

# Lendo as credenciais do Google Sheets
credentials_json = os.environ.get("GOOGLE_CREDENTIALS")
if not credentials_json:
    raise ValueError("As credenciais do Google não foram encontradas. Verifique a variável de ambiente 'GOOGLE_CREDENTIALS'.")
credentials_info = json.loads(credentials_json)
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credenciais = Credentials.from_service_account_info(credentials_info, scopes=scopes)
kommo_leads_token = os.getenv('KOMMO_LEADS_TOKEN')

# Autenticando no Google Sheets
client = gspread.authorize(credenciais)
planilha_id = '1mRNwU-h9EUxYmtG8WNfDLMI8fqWviKGYEERD4k5aegY'
worksheet = client.open_by_key(planilha_id).sheet1

# Função para buscar lead no Kommo usando o número de telefone
def buscar_lead_kommo(numero_telefone):
    url = "https://creditoessencial.kommo.com/api/v4/leads"
    headers = {
        "Authorization": f"Bearer {kommo_leads_token}",
        "Content-Type": "application/json"
    }
    params = {'query': numero_telefone}
    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        leads = response.json().get('_embedded', {}).get('leads', [])
        if leads:
            return leads[0]  # Retorna o primeiro lead encontrado
    print(f"Erro ao buscar lead: {response.status_code} - {response.text}")
    return None

# Função para atualizar o lead no Kommo
def atualizar_lead_kommo(lead_id, pipeline_id, novo_status_id):
    url = f"https://creditoessencial.kommo.com/api/v4/leads/{lead_id}"
    headers = {
        "Authorization": f"Bearer {kommo_leads_token}",
        "Content-Type": "application/json"
    }
    data = {
        "pipeline_id": pipeline_id,
        "status_id": novo_status_id
    }
    response = requests.patch(url, headers=headers, json=data)
    if response.status_code == 200:
        print(f"Lead {lead_id} atualizado com sucesso para o pipeline {pipeline_id} e status {novo_status_id}.")
    else:
        print(f"Erro ao atualizar lead: {response.status_code} - {response.text}")

# Função para monitorar mudanças na planilha e atualizar leads no Kommo
def monitorar_mudancas_planilha():
    rows = worksheet.get_all_records()
    for row in rows:
        status_atual = row['Status do Pedido']
        telefone_cliente = row['Telefone do Cliente']
        lead = buscar_lead_kommo(telefone_cliente)
        if lead:
            pipeline_id = 9423092  # ID do funil específico
            # Mapeamento de status do pedido para IDs de status no Kommo
            if status_atual == 'posted':
                novo_status_id = 72892352  # Status ID para 'posted'
            elif status_atual == 'delivered':
                novo_status_id = 72892360  # Status ID para 'delivered'
            elif status_atual == 'received':
                novo_status_id = 72892356  # Status ID para 'received'
            else:
                continue  # Ignora se o status não for um dos especificados
            atualizar_lead_kommo(lead['id'], pipeline_id, novo_status_id)

# Chame a função de monitoramento
monitorar_mudancas_planilha()
