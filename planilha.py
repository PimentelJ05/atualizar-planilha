import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
from fuzzywuzzy import process

# Defina os tokens iniciais
access_token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxNTg5NyIsImp0aSI6ImIwNTU3ZjViYWJiZmQ4Y2JlMTM1MGYzMTBiODk5OTQ5OTk1ODgxYzIzNzhiMjcwMGI4ZTA2MTVmOWUyNDRjYTJiODc0YmM1OGM0YTc0Mzk5IiwiaWF0IjoxNzI1NTQzMjA4Ljk1NjgsIm5iZiI6MTcyNTU0MzIwOC45NTY4MDIsImV4cCI6MTcyODEzNTIwOC44ODM0NjQsInN1YiI6IjljNzMxYmY1LTA5OGItNDFiMy1iYzczLWJjYzFiOWQ3MWFkMCIsInNjb3BlcyI6WyJjYXJ0LXJlYWQiLCJjYXJ0LXdyaXRlIl19.sDfnwkohO72MNIn5vd-2mDVl50ftG3SNxVlchOmgmCkiHD69PUkkNbiX702ZzvcJdRDoSzUWgdUGwHesISAUOwqXhDfB3RU7XEA1gIFtbTosNlgHksalbCU7Aev7FO0mdyahCAjzMo8nydGHqcsb2cuM02zC3I2O1esAi_ZmSoqUI37AyAyiwlpb46LDRobgi88qzoowiQb5FQ93yQjfI55XTtAXJ2WSa23r9LHSxROp8VU7jlVJH-tLtBYNHrfAK5YrNYl7j--ZzKch_4atQBKmLtDWNzchKmEzBu_IyUdBdee-nYnN6PXzwJgitjLXRS1aoSvAytAwyHZcllNys9qsKf8MgL7yqnY5rEBn3IHC9A6Cj6NQiyBm8x4J0K-H6lnmabLyQ0ZG6zWoYopT9oihYNUBueoYhczqxuwx4h0SGFpLGGWZtq_yGOJbOu5_2ujDG9NjHdKnuBEzJbCAEMLr1VMnwZzKBBq7KF6mfpa7njWrC6Shsj_Env_Z5o2WIctbuBuy6WBDJmd8VgR0lRdi8UXcicD-s9SPyToYDnvmTBmqdJOS7alj2cdX3viPw42dSJ3I1eRL19MXxdNkpm-JzEmFNU5uNJi5trxqG9FRhNN-bthRixhB6EhwDX9MHYawTX14Q2j4NZNeCu2QUDnUGn5dotG5I5BjiV8wQGE'
refresh_token = 'def50200cc347aa16bbf2326ff0a2b659e52abe511a1bc1d2bb773a9f8363fbdf46e9a1fc2744364fe8869365765829ff04caa4eac5c837302c1dc25a42ea0e7fd3bfd1ea02f9f32fbb4ab80a7cffbb38e1ce8ae1dd61b66ae4396abda49c41c495f8ee6923e70489ebe99418212d3b1a54ae65942123bf38c9949f359936e3245f1e5ef15bb9b09e788ed527f637d92b26fd15c42b39afedda22ee74786551e26aee4ba231785e4204a8c351dca0db528e915f86aa36b2045d6fcd19c77048de2e61a13aee46ce1f4a02c13152aeb634235394823ee3e07f312a85e7a67917583c5e1145c188ce8f2070a5b21f621e57af6af2f6663d5bd3e87816e37cb2866759e880d90219cd3bd255c257e29c7c7433769dd32f87a3157b5e23b855a61663eac8da02d93d59d612818e5a2ab132986632329589788b9492ae77f8288ef4adf38b49fcc29b9a7f9bfbab95c055c14b2b4499c38f8117417ec7b570517862c26f2ddcdf7763ab852872890bad5b226f11d58709941bfbe8dd6635067624c501d31a6e6e49585824dc98e3eff9180b04027d90593afb5b2b6dba088af0f15292a29'

# Função para renovar o access token usando o refresh token
def refresh_access_token(refresh_token):
    url = 'https://www.melhorenvio.com.br/oauth/token'
    payload = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': '15897',
        'client_secret': 'ik7qxO9M1iX2uvFYOQd7A3AjjjOPXtJ6vApjPnBt'
    }
    response = requests.post(url, data=payload)
    if response.status_code == 200:
        tokens = response.json()
        print("Novo Access Token:", tokens['access_token'])
        return tokens['access_token'], tokens.get('refresh_token', refresh_token)
    else:
        print("Erro ao renovar tokens:", response.json())
        return None, None

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

# Função para obter os dados de rastreamento de um pedido específico
def obter_rastreamento(pedido_id, access_token):
    url = f'https://www.melhorenvio.com.br/api/v2/me/shipment/track/{pedido_id}'
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        rastreio = response.json().get('tracking', {}).get('code', 'N/A')
        return rastreio
    elif response.status_code == 401:
        # Tentativa de renovar o token ao receber erro 401 (não autorizado)
        novo_access_token, novo_refresh_token = refresh_access_token(refresh_token)
        if novo_access_token:
            global access_token
            access_token = novo_access_token
            headers["Authorization"] = f"Bearer {access_token}"
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                rastreio = response.json().get('tracking', {}).get('code', 'N/A')
                return rastreio
        print(f"Erro ao obter rastreamento após renovação para o pedido {pedido_id}: {response.status_code} - {response.text}")
        return 'N/A'
    else:
        print(f"Erro ao obter rastreamento para o pedido {pedido_id}: {response.status_code} - {response.text}")
        return 'N/A'

# Função para obter os nomes e IDs dos clientes do Kommo
def obter_nomes_ids_clientes():
    api_url = 'https://creditoessencial.kommo.com/api/v4/contacts'
    headers = {
        'Authorization': f'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6ImM0YWU0ZDkzNjkxYjI1NzA3MmZlNzQwMjNjMWUzYjdhMzg3NjY1MWRhM2UyNGUwYjFkYTBkYzQyN2Q1NDZmNjYyNWZmYTdlNGVmZTk3NzUzIn0.eyJhdWQiOiIyYmFjMWY2OC1jMjA0LTQ3MmEtYWRmZS0wNjMwYTI1OTJjZDQiLCJqdGkiOiJjNGFlNGQ5MzY5MWIyNTcwNzJmZTc0MDIzYzFlM2I3YTM4NzY2NTFkYTNlMjRlMGIxZGEwZGM0MjdkNTQ2ZjY2MjVmZmE3ZTRlZmU5Nzc1MyIsImlhdCI6MTcyNTI5MzExNSwibmJmIjoxNzI1MjkzMTE1LCJleHAiOjE3OTg3NjE2MDAsInN1YiI6IjEwNDY2MDM1IiwiZ3JhbnRfdHlwZSI6IiIsImFjY291bnRfaWQiOjMyMDYxNDU1LCJiYXNlX2RvbWFpbiI6ImtvbW1vLmNvbSIsInZlcnNpb24iOjIsInNjb3BlcyI6WyJjcm0iLCJmaWxlcyIsImZpbGVzX2RlbGV0ZSIsIm5vdGlmaWNhdGlvbnMiLCJwdXNoX25vdGlmaWNhdGlvbnMiXSwiaGFzaF91dWlkIjoiZTRmZWYxMDgtODViZi00ZmY0LTllOWYtMGRmZDAwNWYzNWJlIiwiYXBpX2RvbWFpbiI6ImFwaS1nLmtvbW1vLmNvbSJ9.rW4oG9BwkoTSdLi6DLFlCiL0wE8tMPN5dBnNAZtnQYGjdOe1kSKx4fU2s3Tm8vrHF0aI7_1YlubA85ty4uhsh4x_1IGC9593zmOKN-Z2nkK0qSaX0ANQwTNST5XjuhF03FcLEpnqJSb-bjPW-U15vg2SIwR0qezbrPuJMKtjFdiGNwWQW3Jjx2VogZzRQuuRXA30VT8bdDtzySnSQnG0NIb8wGie9QYsZPcYT3c4HQVlPHL8sr9OPhNujTi7YTpiCDnrwDQvO4JBt0CstD78X_Snf4bGQfSOUa8KoAX9DkrHz8-LDkhGc6O1Rwq92iZk6nANI34a8SVyz2oVXwntTw'
    }

    response = requests.get(api_url, headers=headers)

    if response.status_code == 200:
        contacts = response.json()['_embedded']['contacts']
        clientes = {contact['name']: contact['id'] for contact in contacts}
        print("Clientes obtidos do Kommo:", clientes)
        return clientes
    elif response.status_code == 401:
        # Tentativa de renovar o token ao receber erro 401 (não autorizado)
        novo_access_token, novo_refresh_token = refresh_access_token(refresh_token)
        if novo_access_token:
            global access_token
            access_token = novo_access_token
            headers['Authorization'] = f'Bearer {access_token}'
            response = requests.get(api_url, headers=headers)
            if response.status_code == 200:
                contacts = response.json()['_embedded']['contacts']
                clientes = {contact['name']: contact['id'] for contact in contacts}
                print("Clientes obtidos do Kommo após renovação:", clientes)
                return clientes
        print(f"Erro ao obter contatos após renovação: {response.status_code} - {response.text}")
    else:
        print(f"Erro ao obter contatos: {response.status_code} - {response.text}")
    return {}

# Função para obter todos os pedidos do Melhor Envio
def obter_todos_pedidos():
    base_url = "https://www.melhorenvio.com.br/api/v2/me/orders"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {access_token}",
        "User-Agent": "Planilha Crédito Essencial (julia.pimentel@creditoessencial.com.br)"
    }
    todos_pedidos = []
    pagina = 1

    while True:
        response = requests.get(f"{base_url}?page={pagina}", headers=headers)
        if response.status_code == 204:
            print(f"Página {pagina}: Sem conteúdo (204).")
            break
        elif response.status_code == 401:
            # Tentativa de renovar o token ao receber erro 401 (não autorizado)
            novo_access_token, novo_refresh_token = refresh_access_token(refresh_token)
            if novo_access_token:
                global access_token
                access_token = novo_access_token
                headers['Authorization'] = f'Bearer {access_token}'
                response = requests.get(f"{base_url}?page={pagina}", headers=headers)
                if response.status_code == 200:
                    dados = response.json()
                    pedidos = dados.get('data', [])
                    if not pedidos:
                        break
                    todos_pedidos.extend(pedidos)
                    pagina += 1
                else:
                    print(f"Erro ao fazer a requisição após renovação: {response.status_code} - {response.text}")
                    break
            else:
                print("Não foi possível renovar o access token.")
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

    print("Pedidos obtidos do Melhor Envio:", todos_pedidos)
    return todos_pedidos

# Função para encontrar o nome mais próximo usando fuzzy matching
def encontrar_nome_semelhante(nome_cliente_pedido, clientes):
    nomes_kommo = list(clientes.keys())
    nome_correspondente, pontuacao = process.extractOne(nome_cliente_pedido, nomes_kommo)
    if pontuacao > 85:
        return nome_correspondente
    return None

# Função para atualizar a planilha com os dados dos clientes e pedidos, incluindo o rastreamento
def atualizar_planilha_google_sheets(pedidos, clientes, worksheet):
    lista_pedidos = []
    ids_adicionados = set()

    for pedido in pedidos:
        nome_cliente_pedido = pedido.get('to', {}).get('name', 'N/A')
        nome_correspondente = encontrar_nome_semelhante(nome_cliente_pedido, clientes)
        id_cliente = clientes.get(nome_correspondente, '') if nome_correspondente else 'CLIENT ID NOT FOUND'
        id_pedido = pedido.get('id', 'N/A')
        rastreio = obter_rastreamento(id_pedido, access_token)  # Obtém o rastreamento usando a função

        if id_pedido not in ids_adicionados:
            lista_pedidos.append([
                id_cliente,
                nome_correspondente or nome_cliente_pedido,
                id_pedido,
                pedido.get('status', 'N/A'),
                pedido.get('service', {}).get('company', {}).get('name', 'N/A'),
                pedido.get('updated_at', 'N/A'),
                pedido.get('to', {}).get('phone', 'N/A'),
                rastreio  # Adiciona o rastreamento à lista de pedidos
            ])
            ids_adicionados.add(id_pedido)

    # Limpa a planilha e adiciona os cabeçalhos, incluindo o rastreamento
    worksheet.clear()
    worksheet.append_row(["ID do Cliente", "Nome do Cliente", "ID do Pedido", "Status do Pedido", "Transportadora",
                          "Data de Atualização", "Telefone do Cliente", "Rastreamento"])  # Cabeçalho com Rastreamento
    worksheet.append_rows(lista_pedidos, value_input_option='USER_ENTERED')

# Executando as funções para obter dados e atualizar a planilha
clientes = obter_nomes_ids_clientes()
pedidos = obter_todos_pedidos()

if pedidos and clientes:
    atualizar_planilha_google_sheets(pedidos, clientes, worksheet)
    print("Planilha atualizada com sucesso!")
else:
    print("Nenhum pedido ou cliente encontrado para salvar.")
