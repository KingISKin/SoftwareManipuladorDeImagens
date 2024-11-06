Software de Conversão de Imagens
----------------
Este software realiza a conversão de imagens, integrando funcionalidades para baixar arquivos do SharePoint, 
transformar PDFs em imagens TIFF, ler QR codes, renomear arquivos, e enviar documentos processados de volta ao 
SharePoint.

==================================================================================================================================
Por algumas limitações, algumas interações são feitas com o terminal, observe isso ao utilizar o executavel e/ou o código fonte.
==================================================================================================================================

Funcionalidades
----------------
Download de PDFs: Baixa arquivos PDF do SharePoint para uma pasta local chamada PDFs_Baixados.
Conversão para TIFF: Converte os PDFs da pasta PDFs_Baixados em imagens TIFF e armazena-os na pasta Imagens_TIFF.
Leitura e Decodificação de QR Codes: Lê e decodifica QR codes nas imagens TIFF para extrair informações específicas.
Renomeação Automática: Renomeia as imagens TIFF com base nos dados decodificados do QR code.
Reconversão e Salvamento em PDF: Converte as imagens TIFF processadas de volta para PDF, salvando os arquivos finais 
em uma pasta chamada Lotes_Upload.
Upload para o SharePoint: Possibilita o upload dos PDFs da pasta Lotes_Upload para o SharePoint, com a criação de 
subpastas automáticas com base na data do upload.

Estrutura do Projeto
---------------------
O projeto está organizado nas seguintes pastas principais:

PDFs_Baixados: Armazena os PDFs baixados do SharePoint.
Imagens_TIFF: Armazena as imagens TIFF convertidas dos PDFs.
Lotes_Upload: Armazena os PDFs processados, prontos para serem enviados de volta ao SharePoint.

Instalação
------------
Este software requer Python e algumas bibliotecas específicas. Abaixo estão os passos para instalar e configurar o ambiente:

Pré-requisitos
--------------
Python (versão 3.8 ou superior)

Bibliotecas do Python: Execute o comando abaixo para instalar todas as dependências necessárias:

pip install pdf2image pillow numpy pyzbar fpdf office365-runtime PyQt5 rarfile
(insira no seu terminal para  instalar as bibliotecas necessárias)

Configuração do SharePoint
----------------------------
O software usa autenticação via office365-runtime para realizar operações de download e upload no SharePoint. 
Certifique-se de ter as credenciais do SharePoint configuradas para que a autenticação funcione corretamente.
(Você recebeu um .exe e os códigos fontes que fazem o código funcionar internamente, para rodar esse código na sua máquina,
você poderá executa-lo normalmente com o LoteTeste, porém o software foi pensado para armazenar os dados em um banco de dados 
mysql e realizar o upload/download diretamente do SharePoint. No entanto, suas credenciais deverão ser implementadas manualmente
por questões de segurança.)

Bibliotecas Utilizadas
----------------------
O software utiliza várias bibliotecas para manipulação de imagens, interface gráfica, conexão com SharePoint e outros 
processamentos. Abaixo está a lista de bibliotecas usadas:

os: Manipulação de arquivos e pastas no sistema.
re: Operações com expressões regulares.
pdf2image: Conversão de PDF para imagem.
json: Manipulação de dados no formato JSON.
sqlite3: Operações com banco de dados SQLite.(emulação de servidor mysql)
numpy: Manipulações numéricas e de arrays.
PIL (Pillow): Manipulação de imagens, redimensionamento e conversão.
pyzbar: IA para leitura e decodificação de QR codes.
fpdf: Criação de arquivos PDF.
datetime: Operações com data e hora.
office365-runtime: Autenticação e conexão com o SharePoint.
PyQt5: Criação de interface gráfica do usuário (GUI) com Qt.
rarfile:   biblioteca para lidar com arquivos RAR

Módulos Customizados
--------------------
Além das bibliotecas listadas, o software utiliza módulos personalizados:

sharepoint_image_downloader/uploader: Para fluxo de imagens a partir do SharePoint. (Desativado por questões de segurança)
qrcode_scanner: Para processamento de QR codes nas imagens TIFF.
pdf_to_tif_converter: Para converter PDFs em imagens TIFF.

Uso
---
Iniciar o Software: Execute o arquivo principal para abrir a interface gráfica.

Download de Arquivos: O botão de download permite baixar PDFs do SharePoint e armazená-los na pasta PDFs_Baixados.

Conversão para TIFF: Clique no botão de conversão para gerar arquivos TIF a partir dos PDFs baixados, eles serão depositados na pasta Imagens_TIFF.

Processamento de QR Codes: Use o botão para ler e decodificar QR codes, renomeando os arquivos automaticamente e os transformando em PDFs.
Esses arquivos serão tratados, de acordo com a necessidade do cliente/empresa. E por fim depositados na pasta de Lotes_Upload.

Upload de arquivos: Faça o upload dos arquivos para o SharePoint através da interface.
