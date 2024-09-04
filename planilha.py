import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
from fuzzywuzzy import process

# Lendo as credenciais do secret do GitHub Actions
credentials_json = os.environ.get("GOOGLE_CREDENTIALS")

# Convertendo o JSON das credenciais para um dicionário
credentials_info = json.loads(credentials_json)

# Configurando as credenciais usando o dicionário
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
credenciais = Credentials.from_service_account_info(credentials_info, scopes=scopes)

# Autenticando no Google Sheets
client = gspread.authorize(credenciais)

# Acessar a planilha do Google Sheets (ID da sua planilha)
planilha_id = '1-_lGGdU1gH_IkehJrUPydfQ_KxW3J_R9rAS-5dLyGFA'
spreadsheet = client.open_by_key(planilha_id)
worksheet = spreadsheet.sheet1

# Função para obter os nomes e IDs dos clientes do Kommo
def obter_nomes_ids_clientes():
    api_url = 'https://creditoessencial.kommo.com/api/v4/contacts'
    headers = {
        'Authorization': 'Bearer SEU_TOKEN_KOMMO'  # Substitua pelo seu token do Kommo
    }
    
    response = requests.get(api_url, headers=headers)

    if response.status_code == 200:
        contacts = response.json()['_embedded']['contacts']
        # Criar um dicionário com Nome como chave e ID como valor
        clientes = {contact['name']: contact['id'] for contact in contacts}
        print("Clientes obtidos do Kommo:", clientes)  # Verificar os clientes
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
        "Authorization": "Bearer SEU_TOKEN_MELHOR_ENVIO",  # Substitua pelo seu token do Melhor Envio
        "User-Agent": "Planilha Crédito Essencial (julia.pimentel@creditoessencial.com.br)"
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

    print("Pedidos obtidos do Melhor Envio:", pedidos)  # Verificar os pedidos
    return todos_pedidos

# Função para encontrar o nome mais próximo usando fuzzy matching
def encontrar_nome_semelhante(nome_cliente_pedido, clientes):
    nomes_kommo = list(clientes.keys())
    nome_correspondente, pontuacao = process.extractOne(nome_cliente_pedido, nomes_kommo)
    # Considera uma correspondência válida se a pontuação for alta o suficiente (ex: > 80)
    if pontuacao > 80:
        return nome_correspondente
    return None

# Função para atualizar a planilha com os dados dos clientes e pedidos
def atualizar_planilha_google_sheets(pedidos, clientes, worksheet):
    lista_pedidos = []
    for pedido in pedidos:
        nome_cliente_pedido = pedido.get('to', {}).get('name', 'N/A')  # Nome do cliente no pedido
        nome_correspondente = encontrar_nome_semelhante(nome_cliente_pedido, clientes)  # Encontrar o nome correspondente
        if nome_correspondente:
            id_cliente = clientes[nome_correspondente]
            lista_pedidos.append([
                id_cliente,                                # ID do Cliente do Kommo
                nome_correspondente,                      # Nome do Cliente
                pedido.get('id', 'N/A'),                   # ID do Pedido
                pedido.get('status', 'N/A'),               # Status do Pedido
                pedido.get('service', {}).get('company', {}).get('name', 'N/A'),  # Transportadora
                pedido.get('updated_at', 'N/A'),           # Data de Atualização
                pedido.get('to', {}).get('phone', 'N/A')   # Telefone do Cliente
            ])
        else:
            print(f"Nome {nome_cliente_pedido} não encontrado ou não suficientemente próximo no Kommo.")

    # Adicionar cabeçalhos e limpar dados existentes
    worksheet.clear()  
    worksheet.append_row(["ID do Cliente", "Nome do Cliente", "ID do Pedido", "Status do Pedido", "Transportadora", "Data de Atualização", "Telefone do Cliente"])
    worksheet.append_rows(lista_pedidos, value_input_option='USER_ENTERED')  # Adicionar dados

# Executando as funções para obter dados e atualizar a planilha
clientes = obter_nomes_ids_clientes()
pedidos = obter_todos_pedidos()

if pedidos and clientes:
    atualizar_planilha_google_sheets(pedidos, clientes, worksheet)
    print("Planilha atualizada com sucesso!")
else:
    print("Nenhum pedido ou cliente encontrado para salvar.")

