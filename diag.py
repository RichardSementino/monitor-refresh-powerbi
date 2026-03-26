import os
import sys
import requests
import jwt
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# Função para obter token via Client Credentials
def get_access_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://analysis.windows.net/powerbi/api/.default"
    }

    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except requests.exceptions.RequestException as e:
        print(f"Erro ao obter token: {e}")
        sys.exit(1)

# Função para decodificar e exibir claims do JWT
def validate_token(token):
    try:
        decoded = jwt.decode(token, options={"verify_signature": False})
        print("\nClaims importantes do token:")
        print(f"  - appid: {decoded.get('appid', 'N/A')}")
        print(f"  - roles: {decoded.get('roles', [])}")
        print(f"  - scp: {decoded.get('scp', 'N/A')}")
        print(f"  - exp: {datetime.fromtimestamp(decoded.get('exp', 0))}")
        print(f"  - iss: {decoded.get('iss', 'N/A')}")
        print(f"  - aud: {decoded.get('aud', 'N/A')}")
        return True
    except jwt.InvalidTokenError as e:
        print(f"Erro ao decodificar token: {e}")
        return False

# Função para testar um endpoint
def test_endpoint(url, headers, description):
    try:
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            return {"status": "Sucesso", "details": f"{description} - OK"}
        elif response.status_code == 401:
            return {"status": "Erro 401", "details": f"{description} - Unauthorized. Token aceito pelo Entra pode não estar aceito no Power BI/Fabric."}
        elif response.status_code == 403:
            return {"status": "Erro 403", "details": f"{description} - Forbidden. Verifique acesso da aplicação à workspace/dataset."}
        else:
            return {"status": f"Erro {response.status_code}", "details": f"{description} - {response.text}"}

    except requests.exceptions.RequestException as e:
        return {"status": "Erro de rede", "details": f"{description} - {str(e)}"}

# Função principal
def main():
    tenant_id = os.getenv("PBI_TENANT_ID")
    client_id = os.getenv("PBI_CLIENT_ID")
    client_secret = os.getenv("PBI_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        print("Erro: Variáveis de ambiente PBI_TENANT_ID, PBI_CLIENT_ID e PBI_CLIENT_SECRET devem estar definidas.")
        sys.exit(1)

    group_id = input("Digite o ID do grupo (workspace) do Power BI: ").strip()
    dataset_id = input("Digite o ID do dataset: ").strip()

    if not group_id or not dataset_id:
        print("Erro: IDs de grupo e dataset são obrigatórios.")
        sys.exit(1)

    token = get_access_token(tenant_id, client_id, client_secret)
    print("Token obtido com sucesso.")

    if not validate_token(token):
        print("Falha na validação do token. Abortando.")
        sys.exit(1)

    headers = {"Authorization": f"Bearer {token}"}

    tests = [
        {
            "url": "https://api.powerbi.com/v1.0/myorg/groups",
            "desc": "Listar grupos (workspaces)"
        },
        {
            "url": f"https://api.powerbi.com/v1.0/myorg/groups/{group_id}",
            "desc": "Consultar workspace específica"
        },
        {
            "url": f"https://api.powerbi.com/v1.0/myorg/groups/{group_id}/datasets",
            "desc": "Listar datasets da workspace"
        },
        {
            "url": f"https://api.powerbi.com/v1.0/myorg/groups/{group_id}/datasets/{dataset_id}/refreshes?$top=5",
            "desc": "Consultar histórico de refresh"
        }
    ]

    print("\nResultados dos testes:")
    for test in tests:
        result = test_endpoint(test["url"], headers, test["desc"])
        print(f"- {result['status']}: {result['details']}")

if __name__ == "__main__":
    main()
