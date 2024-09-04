import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials

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

# Acessar a planilha do Google Sheets (substitua pelo ID da sua planilha)
planilha_id = '1-_lGGdU1gH_IkehJrUPydfQ_KxW3J_R9rAS-5dLyGFA'
spreadsheet = client.open_by_key(planilha_id)
worksheet = spreadsheet.sheet1

# Exemplo de atualização na planilha
def atualizar_planilha_google_sheets(pedidos, worksheet):
    lista_pedidos = []
    for pedido in pedidos:
        lista_pedidos.append([
            pedido.get('to', {}).get('name', 'N/A'),            # Nome do Cliente
            pedido.get('id', 'N/A'),                            # ID do Pedido
            pedido.get('status', 'N/A'),                       # Status do Pedido
            pedido.get('service', {}).get('company', {}).get('name', 'N/A'),  # Transportadora
            pedido.get('updated_at', 'N/A'),                   # Data de Atualização
            pedido.get('to', {}).get('phone', 'N/A')           # Telefone do Cliente
        ])

    # Adicionar cabeçalhos
    worksheet.clear()  # Limpar dados existentes antes de adicionar novos
    worksheet.append_row(["Nome do Cliente", "ID do Pedido", "Status do Pedido", "Transportadora", "Data de Atualização", "Telefone do Cliente"])
    worksheet.append_rows(lista_pedidos, value_input_option='USER_ENTERED')  # Adicionar dados

# Função principal para obter pedidos e atualizar a planilha
def obter_todos_pedidos():
    base_url = "https://www.melhorenvio.com.br/api/v2/me/orders"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxNTg5NyIsImp0aSI6IjcyOGFjN2U2MDJmNDM1ZjE4YWJmY2ZjOWI2OTU0Nzk5MDgwMzI2OTE5YmMyNzQ2ZjMxYmZmMGY1OWYzZGE0M2Q1NjMyYWRkMzRkZThiNGU2IiwiaWF0IjoxNzI1MDI3ODAxLjA5NzQ3NiwibmJmIjoxNzI1MDI3ODAxLjA5NzQ3OCwiZXhwIjoxNzI3NjE5ODAxLjA1NTUzMywic3ViIjoiOWM3MzFiZjUtMDk4Yi00MWIzLWJjNzMtYmNjMWI5ZDcxYWQwIiwic2NvcGVzIjpbImNhcnQtcmVhZCIsImNhcnQtd3JpdGUiLCJjb21wYW5pZXMtcmVhZCIsImNvbXBhbmllcy13cml0ZSIsImNvdXBvbnMtcmVhZCIsImNvdXBvbnMtd3JpdGUiLCJub3RpZmljYXRpb25zLXJlYWQiLCJvcmRlcnMtcmVhZCIsInByb2R1Y3RzLXJlYWQiLCJwcm9kdWN0cy13cml0ZSIsInB1cmNoYXNlcy1yZWFkIiwic2hpcHBpbmctY2FsY3VsYXRlIiwic2hpcHBpbmctY2FuY2VsIiwic2hpcHBpbmctY2hlY2tvdXQiLCJzaGlwcGluZy1jb21wYW5pZXMiLCJzaGlwcGluZy1nZW5lcmF0ZSIsInNoaXBwaW5nLXByZXZpZXciLCJzaGlwcGluZy1wcmludCIsInNoaXBwaW5nLXNoYXJlIiwic2hpcHBpbmctdHJhY2tpbmciLCJlY29tbWVyY2Utc2hpcHBpbmciLCJ0cmFuc2FjdGlvbnMtcmVhZCIsInVzZXJzLXJlYWQiLCJ1c2Vycy13cml0ZSJdfQ.2uIiaMclQhDERAaQ_qSIShaS1j1UcM-RHrL23H21xk9v0u7M3kvvc4EKQi7U5cnUzi_-2IkJsnr4J7iF04j7ZmxZfmySvH6ttbnFDI6csThrZSw7NanibXIbRnrrAZ6_LmbSk61dj-M12KbPS7OaaMHYYsut04MjowixYlRPDlVHo7bglajHN_fRxTjlGPKIIltevnwtdXkHiIVvEhxavYQvo5PQTs-w6zMmqCAllYWxCJQAyQdtGIJaDoc-zpQEr2XWKIoy4aFmwA_W7fKDXayiPni-K_3J0hvr2wn6p1V-REHqJJ9Wg_RzswU6tL7trArFOghaR05TEYTq1fcLpz99ZjPyX7ES3VfGGXcNwmTzuo-0IWQOmman9_hTjL_YXO_uH-UL9C1p8KQRKVT7qQ2hG1JvzLvuZoOPIO4WKbFso1M4eKyool2_kKjnyKeUDDR1MHfCu_Elt9g_Oo-psHeK9ecubPIq1vjt_gNqKGNjIZidXpc3zmSjNNcQ7_ihVtXkOm0rZ5nw9PC4oFGic_Dj31KjaXawHgttp7tNZpV-LL8P3Zh2u7A6Oggayeaz4WtDTtzgJe1Le9lpGjhpED2xpTHlLmSiFmf0DQ1lqQ5O74aYQg4zY-yhMrgKcjxdHBfsZlixFbOD_O-iZ_roZqdw2Nk1r5Fwow0VIU4tkBM",
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

    return todos_pedidos

# Executando as funções para atualizar a planilha
pedidos = obter_todos_pedidos()
if pedidos:
    atualizar_planilha_google_sheets(pedidos, worksheet)
else:
    print("Nenhum pedido encontrado para salvar.")
