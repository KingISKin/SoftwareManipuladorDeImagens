import os
import re
from pdf2image import convert_from_path

# Função para criar uma pasta se ela não existir
def criar_pasta(path):
    if not os.path.exists(path):
        os.makedirs(path)

# Função para listar arquivos PDF em uma pasta de origem
def listar_pdfs(pasta_origem):
    arquivos = os.listdir(pasta_origem)
    return sorted(
        [os.path.join(pasta_origem, f) for f in arquivos if f.lower().endswith(".pdf")]
    )

# Função para obter o nome base de um arquivo PDF
def obter_nome_base(pdf_name):
    return re.split(r"_[A-Z]+", pdf_name)[0]

# Função para converter um PDF em imagens TIFF
def converter_pdf_para_tiff(pdf_path, poppler_path, dpi=300):
    try:
        return convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)
    except Exception as e:
        print(f"Erro ao converter {pdf_path}: {e}")
        return []
    
# Função para escolher uma pasta para processamento
def escolher_pasta():
    pasta_principal = os.path.join(os.getcwd(), "PDFs_Baixados")
    if not os.path.isdir(pasta_principal):
        print("Erro: A pasta 'PDFS PARA PROCESSAMENTO' não existe.")
        return None

    pastas = [
        pasta
        for pasta in os.listdir(pasta_principal)
        if os.path.isdir(os.path.join(pasta_principal, pasta))
    ]
    print("Escolha uma pasta para processar:")
    for i, pasta in enumerate(pastas, start=1):
        print(f"{i}. {pasta}")

    escolha_pasta = int(input("Digite o número da pasta desejada: "))
    if not (1 <= escolha_pasta <= len(pastas)):
        print("Erro: Número inválido.")
        return None

    return os.path.join(
        pasta_principal, pastas[escolha_pasta - 1]
    )  # Retorna o caminho da pasta escolhida

# Função para criar uma pasta de saída única, evitando conflitos de nome
def criar_pasta_saida_unica(pasta_principal, nome_pasta_saida):
    contador = 1
    caminho_pasta_saida = os.path.join(pasta_principal, nome_pasta_saida)

    # Cria uma nova pasta caso o nome já exista
    while os.path.exists(caminho_pasta_saida):
        caminho_pasta_saida = os.path.join(
            pasta_principal, f"{nome_pasta_saida}{contador}"
        )
        contador += 1

    os.makedirs(caminho_pasta_saida)  # Cria a pasta final
    return caminho_pasta_saida

# Função principal para processar PDFs e converter para TIFF
def ProcessamentoDePDFs():
    print("Bem-vindo ao conversor de PDFs para TIFF.")
    pasta_origem = escolher_pasta()
    if not pasta_origem:
        return

    nome_pasta_saida = f"{os.path.basename(os.path.normpath(pasta_origem))}_Processada"
    numero_inicial = int(
        input("Digite o número inicial para as subpastas (com 8 dígitos): ")
    )

    os.system("cls")  # Limpa a tela do console

    # Configurações para a conversão
    poppler_path = r"C:\poppler\poppler-24.08.0\Library\bin"
    pasta_lotes_tif = os.path.join(os.getcwd(), "Imagens_TIFF")
    criar_pasta(pasta_lotes_tif)  # Cria a pasta para lotes de TIFF

    caminho_pasta_saida = criar_pasta_saida_unica(pasta_lotes_tif, nome_pasta_saida)
    lista_pdfs = listar_pdfs(pasta_origem)  # Lista PDFs na pasta de origem

    if not lista_pdfs:
        print(f"Nenhum arquivo PDF encontrado na pasta '{pasta_origem}'.")
        return

    # Agrupa PDFs pelo nome base
    grupos_pdfs = {}
    for pdf_path in lista_pdfs:
        nome_base = obter_nome_base(os.path.basename(pdf_path))
        grupos_pdfs.setdefault(nome_base, []).append(pdf_path)

    numero_atual = numero_inicial
    for nome_base, pdfs in grupos_pdfs.items():
        nome_subpasta = f"{numero_atual:08}"  # Formata o nome da subpasta
        caminho_subpasta = os.path.join(caminho_pasta_saida, nome_subpasta)
        criar_pasta(caminho_subpasta)  # Cria a subpasta para as TIFFs

        os.system("cls")
        print(f"Status Conversão {numero_atual}/{len(grupos_pdfs)}...")

        for contador_local, pdf_path in enumerate(pdfs, start=1):
            paginas = converter_pdf_para_tiff(pdf_path, poppler_path=poppler_path)
            for pagina in paginas:
                caminho_tiff = os.path.join(
                    caminho_subpasta, f"{contador_local:08}.tif"
                )
                try:
                    pagina.save(caminho_tiff, "TIFF", compression="none")
                except Exception as e:
                    print(f"Erro ao salvar {caminho_tiff}: {e}")

        numero_atual += 1  # Incrementa o número para a próxima subpasta

    print(
        f"\nConversão concluída! As imagens TIFF estão agrupadas na pasta '{caminho_pasta_saida}'."
    )
# Execução do script
if __name__ == "__main__":
    ProcessamentoDePDFs()
