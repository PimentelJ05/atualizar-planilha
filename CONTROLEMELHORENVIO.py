import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
from fuzzywuzzy import process
import unicodedata

# Lendo as credenciais do Google Sheets do ambiente
credentials_json = os.environ.get("GOOGLE_CREDENTIALS")

# Verificando se as credenciais foram carregadas corretamente
if not credentials_json:
    raise ValueError("As credenciais do Google não foram encontradas. Verifique a variável de ambiente 'GOOGLE_CREDENTIALS'.")

# Convertendo o JSON das credenciais para um dicionário
try:
    credentials_info = json.loads(credentials_json)
except json.JSONDecodeError as e:
    raise ValueError(f"Erro ao decodificar o JSON das credenciais: {e}")

# Configurando as credenciais usando o dicionário
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
credenciais = Credentials.from_service_account_info(credentials_info, scopes=scopes)

# Autenticando no Google Sheets
client = gspread.authorize(credenciais)

# ID da planilha fornecida
planilha_id = '1mRNwU-h9EUxYmtG8WNfDLMI8fqWviKGYEERD4k5aegY'

# Acessar a planilha usando o ID fornecido
spreadsheet = client.open_by_key(planilha_id)
worksheet = spreadsheet.sheet1

# Variáveis de tokens do Melhor Envio e Kommo
melhor_envio_access_token = os.getenv('MELHOR_ENVIO_ACCESS_TOKEN').strip()
melhor_envio_client_id = os.getenv('MELHOR_ENVIO_CLIENT_ID')
melhor_envio_client_secret = os.getenv('MELHOR_ENVIO_CLIENT_SECRET')
kommo_access_token = os.getenv('KOMMO_ACCESS_TOKEN').strip()

# Função para normalizar nomes
def normalizar_nome(nome):
    nome = nome.strip().lower()
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')  # Remove acentuação
    nome = nome.replace('\u200b', '')  # Remove caracteres invisíveis como zero-width space
    return nome

# Função para obter os nomes e IDs dos clientes do Kommo
def obter_nomes_ids_clientes():
    api_url = 'https://creditoessencial.kommo.com/api/v4/contacts'
    headers = {
        'Authorization': f'Bearer {kommo_access_token}'
    }

    response = requests.get(api_url, headers=headers)

    if response.status_code == 200:
        contacts = response.json()['_embedded']['contacts']
        # Normaliza os nomes dos clientes para melhor correspondência
        clientes = {normalizar_nome(contact['name']): contact['id'] for contact in contacts}
        print(f"Clientes obtidos do Kommo: {len(clientes)}")
        return clientes
    else:
        print(f"Erro ao obter contatos: {response.status_code} - {response.text}")
    return {}

# Função para obter todos os pedidos do Melhor Envio
def obter_todos_pedidos():
    base_url = "https://www.melhorenvio.com.br/api/v2/me/orders"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f'Bearer {melhor_envio_access_token}',
        "User-Agent": "Controle Melhor Envio (julia.pimentel@creditoessencial.com.br)"
    }
    todos_pedidos = []
    pagina = 1

    while True:
        response = requests.get(f"{base_url}?page={pagina}", headers=headers)
        if response.status_code == 204:
            print(f"Página {pagina}: Sem conteúdo (204).")
            break
        elif response.status_code != 200:
            print(f"Erro ao fazer a requisição: {response.status_code}")
            print("Conteúdo da resposta:", response.text)
            break

        dados = response.json()
        pedidos = dados.get('data', [])
        if not pedidos:
            break
        todos_pedidos.extend(pedidos)
        pagina += 1

    print(f"Pedidos obtidos do Melhor Envio: {len(todos_pedidos)}")
    return todos_pedidos

# Função para encontrar o nome mais próximo usando fuzzy matching
def encontrar_nome_semelhante(nome_cliente_pedido, clientes):
    nomes_kommo = list(clientes.keys())
    nome_cliente_normalizado = normalizar_nome(nome_cliente_pedido)
    nome_correspondente, pontuacao = process.extractOne(nome_cliente_normalizado, nomes_kommo)
    
    print(f"Comparando: '{nome_cliente_normalizado}' com '{nome_correspondente}', Pontuação: {pontuacao}")
    
    if pontuacao > 90:  # Mantendo o limiar de 90 para alta precisão
        return nome_correspondente
    return None

# Função para atualizar a planilha com os dados dos clientes e pedidos
def atualizar_planilha_google_sheets(pedidos, clientes, worksheet):
    lista_pedidos = []
    ids_adicionados = set()

    for pedido in pedidos:
        nome_cliente_pedido = pedido.get('to', {}).get('name', 'N/A')
        nome_correspondente = encontrar_nome_semelhante(nome_cliente_pedido, clientes)
        id_cliente = clientes.get(nome_correspondente, 'CLIENT ID NOT FOUND')
        id_pedido = pedido.get('id', 'N/A')

        if id_pedido not in ids_adicionados:
            lista_pedidos.append([
                id_cliente,
                nome_correspondente or nome_cliente_pedido,
                id_pedido,
                pedido.get('status', 'N/A'),
                pedido.get('service', {}).get('company', {}).get('name', 'N/A'),
                pedido.get('updated_at', 'N/A'),
                pedido.get('to', {}).get('phone', 'N/A')
            ])
            ids_adicionados.add(id_pedido)

    worksheet.clear()
    worksheet.append_row(["ID do Cliente", "Nome do Cliente", "ID do Pedido", "Status do Pedido", "Transportadora",
                          "Data de Atualização", "Telefone do Cliente"])
    worksheet.append_rows(lista_pedidos, value_input_option='USER_ENTERED')

# Executando as funções para obter dados e atualizar a planilha
clientes = obter_nomes_ids_clientes()
pedidos = obter_todos_pedidos()

if pedidos and clientes:
    atualizar_planilha_google_sheets(pedidos, clientes, worksheet)
    print(f"Planilha '{spreadsheet.title}' atualizada com sucesso! ID: {planilha_id}")
else:
    print("Nenhum pedido ou cliente encontrado para salvar.")
