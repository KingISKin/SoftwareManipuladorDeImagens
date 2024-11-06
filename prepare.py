import os
import subprocess

# Função para Listar pastas dentro de Lotes_Upload, e prepara-las
def listar_pastas():
    lotes_upload_path = "Lotes_Upload"

    if os.path.exists(lotes_upload_path):

        folders = [
            f
            for f in os.listdir(lotes_upload_path)
            if os.path.isdir(os.path.join(lotes_upload_path, f))
        ]
        return folders
    else:
        print("A pasta 'Lotes_Upload' não existe.")
        return []
# Função para converter em .rar
def converter_para_rar(pasta):

    lotes_upload_path = "Lotes_Upload"
    folder_path = os.path.join(lotes_upload_path, pasta)

    if os.path.exists(folder_path):
        rar_file_name = f"{pasta}.rar"
        rar_file_path = os.path.join(lotes_upload_path, rar_file_name)

        rar_path = r"C:\Program Files\WinRAR\rar.exe"  # Substitua com o caminho correto do rar.exe caso necessário

        try:
            subprocess.run(
                [rar_path, "a", rar_file_path, folder_path + "\\*"], check=True
            )
            os.system("cls")
        except subprocess.CalledProcessError as e:
            print(f"Erro ao converter a pasta: {str(e)}")
    else:
        print(f"A pasta '{pasta}' não existe.")
# Função principal
def RarConverter():
    pastas = listar_pastas()

    if not pastas:
        return

    print("Pastas disponíveis para conversão:")

    for i, pasta in enumerate(pastas, 1):
        print(f"{i}. {pasta}")

    escolha = input("Digite o número da pasta que deseja converter: ")

    try:
        escolha = int(escolha)
        if 1 <= escolha <= len(pastas):
            pasta_selecionada = pastas[escolha - 1]
            converter_para_rar(pasta_selecionada)
        else:
            print("Escolha inválida.")
    except ValueError:
        print("Entrada inválida. Por favor, digite um número.")

if __name__ == "__main__":
    RarConverter()
