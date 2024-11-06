import os
import json
import sqlite3
import numpy as np
from PIL import Image
from pyzbar.pyzbar import decode
from fpdf import FPDF
import json

# Função para conectar ao banco de dados SQLite e criar a tabela se o banco não existir
def conectar_bd():
    db_path = "automacao.db"
    conexao = sqlite3.connect(db_path)
    cursor = conexao.cursor()

    # Cria a tabela 'informacoes' caso o banco não exista, porém o objetivo do software é salvar numa Nuvem MySql
    cursor.execute(
        """CREATE TABLE IF NOT EXISTS informacoes (
                      id INTEGER PRIMARY KEY AUTOINCREMENT,
                      batchtrack TEXT,
                      codigo_empresa TEXT,
                      codigo_funcionario TEXT,
                      data_admissao TEXT,
                      cpf TEXT,
                      situacao_funcionario TEXT,
                      data_guia TEXT,
                      tipo_exame_num TEXT,
                      cnpj TEXT,
                      razao_social TEXT,
                      nome_funcionario TEXT,
                      tipo_exame TEXT,
                      exames TEXT,
                      numero_guia TEXT,
                      tipo_prontuario TEXT
                  )"""
    )
    conexao.commit()
    return conexao, cursor

# Função para processar cada imagem TIF e decodificar o QR Code
def processar_imagem(filepath):
    imagem = Image.open(filepath)
    imagem = np.array(imagem)
    qr_codes = decode(imagem)

    def decodificar_qr(codes):
        for qr in codes:
            json_content = qr.data.decode("utf-8")
            try:
                return json.loads(json_content)
            except json.JSONDecodeError:
                continue

    return decodificar_qr(qr_codes)

# Função para salvar as informações no banco de dados
def salvar_dados_bd(dados, batchtrack, cursor, conexao):
    # Converter '17' (Dados) para JSON atráves do dicionário
    exames = dados.get("17", "")
    if isinstance(exames, dict):
        exames = json.dumps(exames)  # Converte o dicionário para uma string JSON

    cursor.execute(
        """INSERT INTO informacoes (
                      batchtrack, codigo_empresa, codigo_funcionario, data_admissao, cpf, 
                      situacao_funcionario, data_guia, tipo_exame_num, cnpj, razao_social,
                      nome_funcionario, tipo_exame, exames, numero_guia, tipo_prontuario) 
                      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            batchtrack,
            dados.get("0", ""),
            dados.get("1", ""),
            dados.get("2", ""),
            dados.get("3", ""),
            dados.get("6", ""),
            dados.get("8", ""),
            dados.get("9", ""),
            dados.get("12", ""),
            dados.get("13", ""),
            dados.get("14", ""),
            dados.get("15", ""),
            exames,  # Armazena a string JSON no banco, dando a entender que podemos variar o tipo de informação recebida no json, como um resultado de exame
            dados.get("19", ""),
            dados.get("99", ""),
        ),
    )
    conexao.commit()

# Função para criar um PDF com a imagem TIFF
def criar_pdf(nome_arquivo, imagem_tiff):
    pdf = FPDF()
    pdf.add_page()
    imagem_tiff.convert("RGB").save("temp_image.jpg")
    pdf.image("temp_image.jpg", x=10, y=10, w=pdf.w - 20)
    pdf.output(nome_arquivo)
    os.remove("temp_image.jpg")

# Função principal para percorrer as pastas e processar as imagens, onde tratamos os erros e exceções
def processar_pastas(pasta_escolhida):
    conexao, cursor = conectar_bd()
    pasta_upload = os.path.join(".\\Lotes_Upload", os.path.basename(pasta_escolhida))
    os.makedirs(pasta_upload, exist_ok=True)
    erro_log_path = os.path.join(pasta_upload, "erros.txt")

    with open(erro_log_path, "w") as erro_log:
        for subdir, dirs, files in os.walk(pasta_escolhida):
            batchtrack = os.path.basename(subdir)
            for file in files:
                if file.lower().endswith(".tif"):
                    filepath = os.path.join(subdir, file)

                    dados_qrcode = processar_imagem(filepath)

                    if dados_qrcode:
                        if isinstance(dados_qrcode, dict):
                            salvar_dados_bd(dados_qrcode, batchtrack, cursor, conexao)
                            nome_funcionario = (
                                dados_qrcode.get("14", "FUNCIONARIO_DESCONHECIDO")
                                .upper()
                                .replace(" ", "_")
                            )
                            tipo_exame_num = dados_qrcode.get("9", None)
                            tipo_exame_mapeado = {
                                "1": "ADMISIONAL",
                                "2": "PERIODICO",
                                "3": "MUDANCA_DE_FUNCAO",
                                "4": "RETORNO_AO_TRABALHO",
                                "5": "DEMISIONAL",
                            }
                            tipo_exame = tipo_exame_mapeado.get(
                                str(tipo_exame_num), "TIPO_DE_EXAME_INVALIDO"
                            )
                            nome_pdf = os.path.join(
                                pasta_upload, f"{nome_funcionario}_ASO_{tipo_exame}.pdf"
                            )
                            criar_pdf(nome_pdf, Image.open(filepath))
                            os.system("cls")
                        else:
                            pass
                    else:
                        erro_log.write(filepath + "\n")

    conexao.close()
    print("Processamento concluído!")

# Função principal que inicia o processamento
def MainQrcode():
    pasta_principal = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "Imagens_TIFF"
    )
    subpastas = [
        d
        for d in os.listdir(pasta_principal)
        if os.path.isdir(os.path.join(pasta_principal, d))
    ]
    print("Escolha uma pasta para processar:")
    for i, subpasta in enumerate(subpastas, start=1):
        print(f"{i}: {subpasta}")

    escolha = int(input("Digite o número da pasta escolhida: ")) - 1

    if 0 <= escolha < len(subpastas):
        pasta_escolhida = os.path.join(pasta_principal, subpastas[escolha])
        processar_pastas(pasta_escolhida)
    else:
        print("Escolha inválida!Tente novamente")

if __name__ == "__main__":
    MainQrcode()
