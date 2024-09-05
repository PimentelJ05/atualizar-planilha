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
        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6ImM0YWU0ZDkzNjkxYjI1NzA3MmZlNzQwMjNjMWUzYjdhMzg3NjY1MWRhM2UyNGUwYjFkYTBkYzQyN2Q1NDZmNjYyNWZmYTdlNGVmZTk3NzUzIn0.eyJhdWQiOiIyYmFjMWY2OC1jMjA0LTQ3MmEtYWRmZS0wNjMwYTI1OTJjZDQiLCJqdGkiOiJjNGFlNGQ5MzY5MWIyNTcwNzJmZTc0MDIzYzFlM2I3YTM4NzY2NTFkYTNlMjRlMGIxZGEwZGM0MjdkNTQ2ZjY2MjVmZmE3ZTRlZmU5Nzc1MyIsImlhdCI6MTcyNTI5MzExNSwibmJmIjoxNzI1MjkzMTE1LCJleHAiOjE3OTg3NjE2MDAsInN1YiI6IjEwNDY2MDM1IiwiZ3JhbnRfdHlwZSI6IiIsImFjY291bnRfaWQiOjMyMDYxNDU1LCJiYXNlX2RvbWFpbiI6ImtvbW1vLmNvbSIsInZlcnNpb24iOjIsInNjb3BlcyI6WyJjcm0iLCJmaWxlcyIsImZpbGVzX2RlbGV0ZSIsIm5vdGlmaWNhdGlvbnMiLCJwdXNoX25vdGlmaWNhdGlvbnMiXSwiaGFzaF91dWlkIjoiZTRmZWYxMDgtODViZi00ZmY0LTllOWYtMGRmZDAwNWYzNWJlIiwiYXBpX2RvbWFpbiI6ImFwaS1nLmtvbW1vLmNvbSJ9.rW4oG9BwkoTSdLi6DLFlCiL0wE8tMPN5dBnNAZtnQYGjdOe1kSKx4fU2s3Tm8vrHF0aI7_1YlubA85ty4uhsh4x_1IGC9593zmOKN-Z2nkK0qSaX0ANQwTNST5XjuhF03FcLEpnqJSb-bjPW-U15vg2SIwR0qezbrPuJMKtjFdiGNwWQW3Jjx2VogZzRQuuRXA30VT8bdDtzySnSQnG0NIb8wGie9QYsZPcYT3c4HQVlPHL8sr9OPhNujTi7YTpiCDnrwDQvO4JBt0CstD78X_Snf4bGQfSOUa8KoAX9DkrHz8-LDkhGc6O1Rwq92iZk6nANI34a8SVyz2oVXwntTw'
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
        "Authorization": "Bearer  eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxNTg5NyIsImp0aSI6IjcyOGFjN2U2MDJmNDM1ZjE4YWJmY2ZjOWI2OTU0Nzk5MDgwMzI2OTE5YmMyNzQ2ZjMxYmZmMGY1OWYzZGE0M2Q1NjMyYWRkMzRkZThiNGU2IiwiaWF0IjoxNzI1MDI3ODAxLjA5NzQ3NiwibmJmIjoxNzI1MDI3ODAxLjA5NzQ3OCwiZXhwIjoxNzI3NjE5ODAxLjA1NTUzMywic3ViIjoiOWM3MzFiZjUtMDk4Yi00MWIzLWJjNzMtYmNjMWI5ZDcxYWQwIiwic2NvcGVzIjpbImNhcnQtcmVhZCIsImNhcnQtd3JpdGUiLCJjb21wYW5pZXMtcmVhZCIsImNvbXBhbmllcy13cml0ZSIsImNvdXBvbnMtcmVhZCIsImNvdXBvbnMtd3JpdGUiLCJub3RpZmljYXRpb25zLXJlYWQiLCJvcmRlcnMtcmVhZCIsInByb2R1Y3RzLXJlYWQiLCJwcm9kdWN0cy13cml0ZSIsInB1cmNoYXNlcy1yZWFkIiwic2hpcHBpbmctY2FsY3VsYXRlIiwic2hpcHBpbmctY2FuY2VsIiwic2hpcHBpbmctY2hlY2tvdXQiLCJzaGlwcGluZy1jb21wYW5pZXMiLCJzaGlwcGluZy1nZW5lcmF0ZSIsInNoaXBwaW5nLXByZXZpZXciLCJzaGlwcGluZy1wcmludCIsInNoaXBwaW5nLXNoYXJlIiwic2hpcHBpbmctdHJhY2tpbmciLCJlY29tbWVyY2Utc2hpcHBpbmciLCJ0cmFuc2FjdGlvbnMtcmVhZCIsInVzZXJzLXJlYWQiLCJ1c2Vycy13cml0ZSJdfQ.2uIiaMclQhDERAaQ_qSIShaS1j1UcM-RHrL23H21xk9v0u7M3kvvc4EKQi7U5cnUzi_-2IkJsnr4J7iF04j7ZmxZfmySvH6ttbnFDI6csThrZSw7NanibXIbRnrrAZ6_LmbSk61dj-M12KbPS7OaaMHYYsut04MjowixYlRPDlVHo7bglajHN_fRxTjlGPKIIltevnwtdXkHiIVvEhxavYQvo5PQTs-w6zMmqCAllYWxCJQAyQdtGIJaDoc-zpQEr2XWKIoy4aFmwA_W7fKDXayiPni-K_3J0hvr2wn6p1V-REHqJJ9Wg_RzswU6tL7trArFOghaR05TEYTq1fcLpz99ZjPyX7ES3VfGGXcNwmTzuo-0IWQOmman9_hTjL_YXO_uH-UL9C1p8KQRKVT7qQ2hG1JvzLvuZoOPIO4WKbFso1M4eKyool2_kKjnyKeUDDR1MHfCu_Elt9g_Oo-psHeK9ecubPIq1vjt_gNqKGNjIZidXpc3zmSjNNcQ7_ihVtXkOm0rZ5nw9PC4oFGic_Dj31KjaXawHgttp7tNZpV-LL8P3Zh2u7A6Oggayeaz4WtDTtzgJe1Le9lpGjhpED2xpTHlLmSiFmf0DQ1lqQ5O74aYQg4zY-yhMrgKcjxdHBfsZlixFbOD_O-iZ_roZqdw2Nk1r5Fwow0VIU4tkBM",
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
    if pontuacao > 85:
        return nome_correspondente
    return None

# Função para atualizar a planilha com os dados dos clientes e pedidos, evitando duplicatas
def atualizar_planilha_google_sheets(pedidos, clientes, worksheet):
    lista_pedidos = []
    ids_adicionados = set()  # Conjunto para rastrear IDs de pedidos adicionados
    
    for pedido in pedidos:
        nome_cliente_pedido = pedido.get('to', {}).get('name', 'N/A')  # Nome do cliente no pedido
        nome_correspondente = encontrar_nome_semelhante(nome_cliente_pedido, clientes)  # Encontrar o nome correspondente
        id_cliente = clientes.get(nome_correspondente, '') if nome_correspondente else 'CLIENT ID NOT FOUND'
        id_pedido = pedido.get('id', 'N/A')

        # Verificar se o ID do pedido já foi adicionado
        if id_pedido not in ids_adicionados:
            lista_pedidos.append([
                id_cliente,                                # ID do Cliente do Kommo ou 'CLIENT ID NOT FOUND'
                nome_correspondente or nome_cliente_pedido,  # Nome do Cliente ou o nome do pedido
                id_pedido,                                # ID do Pedido
                pedido.get('status', 'N/A'),              # Status do Pedido
                pedido.get('service', {}).get('company', {}).get('name', 'N/A'),  # Transportadora
                pedido.get('updated_at', 'N/A'),          # Data de Atualização
                pedido.get('to', {}).get('phone', 'N/A')  # Telefone do Cliente
            ])
            ids_adicionados.add(id_pedido)  # Adicionar o ID do pedido ao conjunto

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
