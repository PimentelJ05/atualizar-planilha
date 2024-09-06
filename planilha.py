import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
from fuzzywuzzy import process

# Configuração dos tokens e credenciais
access_token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxNTg5NyIsImp0aSI6ImIwNTU3ZjViYWJiZmQ4Y2JlMTM1MGYzMTBiODk5OTQ5OTk1ODgxYzIzNzhiMjcwMGI4ZTA2MTVmOWUyNDRjYTJiODc0YmM1OGM0YTc0Mzk5IiwiaWF0IjoxNzI1NTQzMjA4Ljk1NjgsIm5iZiI6MTcyNTU0MzIwOC45NTY4MDIsImV4cCI6MTcyODEzNTIwOC44ODM0NjQsInN1YiI6IjljNzMxYmY1LTA5OGItNDFiMy1iYzczLWJjYzFiOWQ3MWFkMCIsInNjb3BlcyI6WyJjYXJ0LXJlYWQiLCJjYXJ0LXdyaXRlIl19.sDfnwkohO72MNIn5vd-2mDVl50ftG3SNxVlchOmgmCkiHD69PUkkNbiX702ZzvcJdRDoSzUWgdUGwHesISAUOwqXhDfB3RU7XEA1gIFtbTosNlgHksalbCU7Aev7FO0mdyahCAjzMo8nydGHqcsb2cuM02zC3I2O1esAi_ZmSoqUI37AyAyiwlpb46LDRobgi88qzoowiQb5FQ93yQjfI55XTtAXJ2WSa23r9LHSxROp8VU7jlVJH-tLtBYNHrfAK5YrNYl7j--ZzKch_4atQBKmLtDWNzchKmEzBu_IyUdBdee-nYnN6PXzwJgitjLXRS1aoSvAytAwyHZcllNys9qsKf8MgL7yqnY5rEBn3IHC9A6Cj6NQiyBm8x4J0K-H6lnmabLyQ0ZG6zWoYopT9oihYNUBueoYhczqxuwx4h0SGFpLGGWZtq_yGOJbOu5_2ujDG9NjHdKnuBEzJbCAEMLr1VMnwZzKBBq7KF6mfpa7njWrC6Shsj_Env_Z5o2WIctbuBuy6WBDJmd8VgR0lRdi8UXcicD-s9SPyToYDnvmTBmqdJOS7alj2cdX3viPw42dSJ3I1eRL19MXxdNkpm-JzEmFNU5uNJi5trxqG9FRhNN-bthRixhB6EhwDX9MHYawTX14Q2j4NZNeCu2QUDnUGn5dotG5I5BjiV8wQGE'
refresh_token = 'def50200cc347aa16bbf2326ff0a2b659e52abe511a1bc1d2bb773a9f8363fbdf46e9a1fc2744364fe8869365765829ff04caa4eac5c837302c1dc25a42ea0e7fd3bfd1ea02f9f32fbb4ab80a7cffbb38e1ce8ae1dd61b66ae4396abda49c41c495f8ee6923e70489ebe99418212d3b1a54ae65942123bf38c9949f359936e3245f1e5ef15bb9b09e788ed527f637d92b26fd15c42b39afedda22ee74786551e26aee4ba231785e4204a8c351dca0db528e915f86aa36b2045d6fcd19c77048de2e61a13aee46ce1f4a02c13152aeb634235394823ee3e07f312a85e7a67917583c5e1145c188ce8f2070a5b21f621e57af6af2f6663d5bd3e87816e37cb2866759e880d90219cd3bd255c257e29c7c7433769dd32f87a3157b5e23b855a61663eac8da02d93d59d612818e5a2ab132986632329589788b9492ae77f8288ef4adf38b49fcc29b9a7f9bfbab95c055c14b2b4499c38f8117417ec7b570517862c26f2ddcdf7763ab852872890bad5b226f11d58709941bfbe8dd6635067624c501d31a6e6e49585824dc98e3eff9180b04027d90593afb5b2b6dba088af0f15292a29'
kommo_access_token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsImp0aSI6ImM0YWU0ZDkzNjkxYjI1NzA3MmZlNzQwMjNjMWUzYjdhMzg3NjY1MWRhM2UyNGUwYjFkYTBkYzQyN2Q1NDZmNjYyNWZmYTdlNGVmZTk3NzUzIn0.eyJhdWQiOiIyYmFjMWY2OC1jMjA0LTQ3MmEtYWRmZS0wNjMwYTI1OTJjZDQiLCJqdGkiOiJjNGFlNGQ5MzY5MWIyNTcwNzJmZTc0MDIzYzFlM2I3YTM4NzY2NTFkYTNlMjRlMGIxZGEwZGM0MjdkNTQ2ZjY2MjVmZmE3ZTRlZmU5Nzc1MyIsImlhdCI6MTcyNTI5MzExNSwibmJmIjoxNzI1MjkzMTE1LCJleHAiOjE3OTg3NjE2MDAsInN1YiI6IjEwNDY2MDM1IiwiZ3JhbnRfdHlwZSI6IiIsImFjY291bnRfaWQiOjMyMDYxNDU1LCJiYXNlX2RvbWFpbiI6ImtvbW1vLmNvbSIsInZlcnNpb24iOjIsInNjb3BlcyI6WyJjcm0iLCJmaWxlcyIsImZpbGVzX2RlbGV0ZSIsIm5vdGlmaWNhdGlvbnMiLCJwdXNoX25vdGlmaWNhdGlvbnMiXSwiaGFzaF91dWlkIjoiZTRmZWYxMDgtODViZi00ZmY0LTllOWYtMGRmZDAwNWYzNWJlIiwiYXBpX2RvbWFpbiI6ImFwaS1nLmtvbW1vLmNvbSJ9.rW4oG9BwkoTSdLi6DLFlCiL0wE8tMPN5dBnNAZtnQYGjdOe1kSKx4fU2s3Tm8vrHF0aI7_1YlubA85ty4uhsh4x_1IGC9593zmOKN-Z2nkK0qSaX0ANQwTNST5XjuhF03FcLEpnqJSb-bjPW-U15vg2SIwR0qezbrPuJMKtjFdiGNwWQW3Jjx2VogZzRQuuRXA30VT8bdDtzySnSQnG0NIb8wGie9QYsZPcYT3c4HQVlPHL8sr9OPhNujTi7YTpiCDnrwDQvO4JBt0CstD78X_Snf4bGQfSOUa8KoAX9DkrHz8-LDkhGc6O1Rwq92iZk6nANI34a8SVyz2oVXwntTw'

# Lendo as credenciais do secret do GitHub Actions
credentials_json = os.environ.get("GOOGLE_CREDENTIALS")
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
        return tokens['access_token'], tokens.get('refresh_token')
    else:
        print("Erro ao renovar tokens:", response.json())
        return None, None

def obter_nomes_ids_clientes():
    api_url = 'https://creditoessencial.kommo.com/api/v4/contacts'
    headers = {
        'Authorization': f'Bearer {kommo_access_token}'
    }

    response = requests.get(api_url, headers=headers)

    if response.status_code == 200:
        contacts = response.json()['_embedded']['contacts']
        clientes = {contact['name']: contact['id'] for contact in contacts}
        print("Clientes obtidos do Kommo:", clientes)
        return clientes
    else:
        print(f"Erro ao obter contatos: {response.status_code} - {response.text}")
    return {}

def obter_detalhes_pedido(order_id):
    url = f"https://www.melhorenvio.com.br/api/v2/me/orders/{order_id}"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f'Bearer {access_token}',
        "User-Agent": "Planilha Crédito Essencial (julia.pimentel@creditoessencial.com.br)"
    }
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        pedido_detalhes = response.json()
        # Extrair o código de rastreamento corretamente
        tracking_code = pedido_detalhes.get('tracking_code', None)
        return tracking_code, pedido_detalhes  # Retorna o código de rastreamento e os detalhes do pedido
    elif response.status_code == 401:
        novo_access_token, novo_refresh_token = refresh_access_token(refresh_token)
        if novo_access_token:
            headers['Authorization'] = f'Bearer {novo_access_token}'
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                pedido_detalhes = response.json()
                tracking_code = pedido_detalhes.get('tracking_code', None)
                return tracking_code, pedido_detalhes
    print(f"Erro ao obter detalhes do pedido {order_id}: {response.status_code} - {response.text}")
    return None, {}


def obter_todos_pedidos():
    base_url = "https://www.melhorenvio.com.br/api/v2/me/orders"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f'Bearer {access_token}',
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
            novo_access_token, novo_refresh_token = refresh_access_token(refresh_token)
            if novo_access_token:
                headers['Authorization'] = f'Bearer {novo_access_token}'
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
        for pedido in pedidos:
            tracking_code, detalhes = obter_detalhes_pedido(pedido['id'])
            if tracking_code:  # Somente busca rastreamento se o código estiver disponível
                tracking_info = obter_codigo_rastreamento(tracking_code)
                pedido['tracking_info'] = tracking_info
            else:
                pedido['tracking_info'] = {}
            todos_pedidos.append(pedido)
        pagina += 1

    print("Pedidos obtidos do Melhor Envio com detalhes e rastreamento:", todos_pedidos)
    return todos_pedidos


def encontrar_nome_semelhante(nome_cliente_pedido, clientes):
    nomes_kommo = list(clientes.keys())
    nome_correspondente, pontuacao = process.extractOne(nome_cliente_pedido, nomes_kommo)
    if pontuacao > 85:
        return nome_correspondente
    return None

def atualizar_planilha_google_sheets(pedidos, clientes, worksheet):
    lista_pedidos = []
    ids_adicionados = set()
    
    for pedido in pedidos:
        nome_cliente_pedido = pedido.get('to', {}).get('name', 'N/A')
        nome_correspondente = encontrar_nome_semelhante(nome_cliente_pedido, clientes)
        id_cliente = clientes.get(nome_correspondente, '') if nome_correspondente else 'CLIENT ID NOT FOUND'
        id_pedido = pedido.get('id', 'N/A')
        
        # Verifica se o código de rastreamento está presente
        if 'tracking_info' in pedido and 'tracking_code' in pedido['tracking_info']:
            codigo_rastreamento = pedido['tracking_info'].get('tracking_code', 'N/A')
        else:
            codigo_rastreamento = "não foi possível localizar o código"

        # Verifica se o status de rastreamento está presente
        if 'tracking_info' in pedido and 'status' in pedido['tracking_info']:
            status_rastreamento = pedido['tracking_info'].get('status', 'N/A')
        else:
            status_rastreamento = "não foi possível localizar o status"

        if id_pedido not in ids_adicionados:
            lista_pedidos.append([
                id_cliente,
                nome_correspondente or nome_cliente_pedido,
                id_pedido,
                codigo_rastreamento,
                status_rastreamento,
                pedido.get('service', {}).get('company', {}).get('name', 'N/A'),
                pedido.get('updated_at', 'N/A'),
                pedido.get('to', {}).get('phone', 'N/A')
            ])
            ids_adicionados.add(id_pedido)

    # Limpar dados existentes e adicionar cabeçalhos
    worksheet.clear()  
    worksheet.append_row(["ID do Cliente", "Nome do Cliente", "ID do Pedido", "Código de Rastreamento", "Status de Rastreamento", "Transportadora", "Data de Atualização", "Telefone do Cliente"])
    worksheet.append_rows(lista_pedidos, value_input_option='USER_ENTERED')

    print("Planilha atualizada com sucesso!")

else:
    print("Nenhum pedido ou cliente encontrado para salvar.")
