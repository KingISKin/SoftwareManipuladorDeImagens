import requests
import os
import rarfile

# Função para Testar o inicio de download
def iniciando_download():
    print("Iniciando download")

# Função Principal
def baixar_imagens(
    client_id,
    client_secret,
    tenant_id,
    site_id,
    drive_id,
    main_folder_path,
    download_path,
):
    # Função para obter o token de acesso
    def get_access_token(client_id, client_secret, tenant_id):
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json()["access_token"]

    # Função para listar as subpastas na pasta principal e permitir a escolha do usuário
    def escolher_pasta(access_token, site_id, drive_id, folder_path):
        items = list_folder_items(access_token, site_id, drive_id, folder_path)
        folders = [item for item in items if "folder" in item]

        if not folders:
            print("Nenhuma subpasta encontrada.")
            return None

        print("Escolha uma pasta para download:")
        for i, folder in enumerate(folders, 1):
            print(f"{i}. {folder['name']}")

        escolha = int(input("Digite o número da pasta desejada: ")) - 1
        if 0 <= escolha < len(folders):
            return folders[escolha]["name"]
        else:
            print("Escolha inválida.")
            return None

    # Função para listar os itens da pasta atual
    def list_folder_items(access_token, site_id, drive_id, current_folder_path):
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:{current_folder_path}:/children"
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json().get("value", [])

    # Função para baixar um arquivo
    def download_file(access_token, download_url, local_path):
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.get(download_url, headers=headers, stream=True)
        response.raise_for_status()
        with open(local_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

    # Função para extrair um arquivo RAR
    def extrair_arquivo_rar(file_path, extract_path):
        with rarfile.RarFile(file_path) as rf:
            rf.extractall(extract_path)
        print(f"Arquivo RAR extraído em: {extract_path}")
        os.remove(file_path)  # Remove o arquivo RAR após a extração

    # Função recursiva para baixar itens de uma pasta
    def baixar_recursivamente(
        access_token, site_id, drive_id, current_folder_path, current_download_path
    ):
        if not os.path.exists(current_download_path):
            os.makedirs(current_download_path)

        items = list_folder_items(access_token, site_id, drive_id, current_folder_path)

        for item in items:
            if "file" in item:
                file_name = item["name"]
                download_url = item.get("@microsoft.graph.downloadUrl")
                if download_url:
                    local_file_path = os.path.join(current_download_path, file_name)
                    print(f"Baixando {file_name} para {local_file_path}...")
                    download_file(access_token, download_url, local_file_path)
                    print(f"{file_name} baixado com sucesso.")
                    if file_name.lower().endswith(".rar"):
                        extrair_arquivo_rar(local_file_path, current_download_path)
                else:
                    print(f"URL de download não encontrada para o arquivo: {file_name}")
            elif "folder" in item:
                subfolder_name = item["name"]
                subfolder_path = os.path.join(
                    current_folder_path, subfolder_name
                ).replace("\\", "/")
                sub_download_path = os.path.join(current_download_path, subfolder_name)
                print(
                    f"Encontrada subpasta: {subfolder_name}. Baixando recursivamente..."
                )
                baixar_recursivamente(
                    access_token, site_id, drive_id, subfolder_path, sub_download_path
                )

    # Executa o processo principal de download
    try:
        print("Obtendo token de acesso...")
        token = get_access_token(client_id, client_secret, tenant_id)
        print("Token de acesso obtido com sucesso.")

        print("Escolhendo pasta para download...")
        selected_folder = escolher_pasta(token, site_id, drive_id, main_folder_path)

        if selected_folder:
            print(f"Iniciando o download da pasta: {selected_folder}")
            selected_folder_path = os.path.join(
                main_folder_path, selected_folder
            ).replace("\\", "/")
            selected_download_path = os.path.join(download_path, selected_folder)
            baixar_recursivamente(
                token, site_id, drive_id, selected_folder_path, selected_download_path
            )
            print("Download concluído com sucesso.")
        else:
            print("Nenhuma pasta selecionada.")

    except requests.exceptions.HTTPError as http_err:
        print(f"Erro HTTP ocorreu: {http_err}")
    except Exception as err:
        print(f"Ocorreu um erro: {err}")

# Chamada principal
if __name__ == "__main__":
    CLIENT_ID = "SEU_CLIENT_ID"
    CLIENT_SECRET = "SEU_CLIENT_SECRET"
    TENANT_ID = "SEU_TENANT_ID"
    SITE_ID = "SEU_SITE_ID"
    DRIVE_ID = "SEU_DRIVE_ID"
    MAIN_FOLDER_PATH = "/Imagens_Processamento"
    DOWNLOAD_PATH = "./PDFs_Baixados"

    # Inicia o download das imagens
    baixar_imagens(
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        tenant_id=TENANT_ID,
        site_id=SITE_ID,
        drive_id=DRIVE_ID,
        main_folder_path=MAIN_FOLDER_PATH,
        download_path=DOWNLOAD_PATH,
    )
