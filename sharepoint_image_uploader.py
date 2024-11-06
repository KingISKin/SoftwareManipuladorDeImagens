import os
import shutil
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext


# Configurações do SharePoint
site_url = "https://suacempresa.sharepoint.com/sites/site"
username = "seu_email@empresarial.com"
password = "sua_senha"
base_folder_url = "/sites/site/Lotes_processados"

local_upload_folder_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "Lotes_Upload"
)

# Autenticação no SharePoint
ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    print("Autenticação falhou.")
else:
    ctx = ClientContext(site_url, ctx_auth)

    # Lista as pastas disponíveis para upload
    folders = [
        f
        for f in os.listdir(local_upload_folder_path)
        if os.path.isdir(os.path.join(local_upload_folder_path, f))
    ]
    print("Escolha uma pasta para compactar e fazer upload:")
    for i, folder_name in enumerate(folders, start=1):
        print(f"{i}. {folder_name}")

    # Seleciona a pasta a ser compactada e enviada
    choice = int(input("Digite o número da pasta desejada: ")) - 1
    chosen_folder = folders[choice]
    chosen_folder_path = os.path.join(local_upload_folder_path, chosen_folder)
    rar_path = os.path.join(local_upload_folder_path, f"{chosen_folder}.rar")

    # Compacta a pasta em formato .rar
    shutil.make_archive(chosen_folder, "rar", local_upload_folder_path, chosen_folder)
    print(f"Pasta '{chosen_folder}' compactada em '{rar_path}'.")

    # Configura o nome da pasta de destino no SharePoint
    current_date = datetime.now().strftime("%d_%m_%Y")
    folder_name = f"LOTE_{current_date}"
    target_folder_url = f"{base_folder_url}/{folder_name}"

    # Cria a pasta de destino no SharePoint, se necessário
    ctx.web.folders.add(target_folder_url).execute_query()
    print(f"Pasta '{folder_name}' criada no SharePoint.")

    # Upload do arquivo .rar para o SharePoint
    with open(rar_path, "rb") as file_content:
        ctx.web.get_folder_by_server_relative_url(target_folder_url).upload_file(
            f"{chosen_folder}.rar", file_content
        ).execute_query()
        print(
            f"Arquivo '{chosen_folder}.rar' enviado com sucesso para '{folder_name}' no SharePoint."
        )

print("Upload concluído.")
