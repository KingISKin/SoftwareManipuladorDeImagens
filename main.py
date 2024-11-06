import sys
from PyQt5.QtWidgets import (QApplication,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QFileDialog,
    QGraphicsView,
    QGraphicsScene,
    QGridLayout,
    QLabel
)
from PyQt5.QtGui import QPixmap, QImage, QIcon, QDesktopServices
from PyQt5.QtCore import Qt ,QThread, pyqtSignal, QDateTime, QUrl
from sharepoint_image_downloader import iniciando_download
from qrcode_scanner import MainQrcode
from pdf_to_tif_converter import ProcessamentoDePDFs
from prepare import RarConverter
from PIL import Image
import subprocess
import os

class QrCodeThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def run(self):
        try:
            MainQrcode()
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class ProcessThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def run(self):
        try:
            ProcessamentoDePDFs()
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class RarThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def run(self):
        try:
            RarConverter()
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Logo
        logo_label = QLabel(self)
        pixmap = QPixmap(r".\Icones\logo.png")  
        pixmap = pixmap.scaled(180, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation)  
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignCenter)  
        main_layout.addWidget(logo_label)

        # Layout para os botões principais
        button_layout = QVBoxLayout()
        
        # Botão "Iniciar processamento" centralizado
        self.botao_iniciar = self.create_button("Iniciar processamento", self.open_second_window,  "#007BFF", "#0056b3")
        self.botao_iniciar.setFixedSize(150, 45)
        button_layout.addWidget(self.botao_iniciar, alignment=Qt.AlignCenter)

        # Adiciona os botões principais ao layout principal
        main_layout.addLayout(button_layout)

        # Layout para o botão "Abrir Readme" (como link)
        bottom_layout = QHBoxLayout()

        # Botão "Abrir Readme" como link
        self.botao_readme = QLabel(self)
        self.botao_readme.setText('<a href="#" style="text-decoration: underline; color: blue;">Abrir Readme</a>')
        self.botao_readme.setAlignment(Qt.AlignLeft)  # Alinha à esquerda
        self.botao_readme.linkActivated.connect(self.open_readme)

        # Alinha o "Abrir Readme" no canto inferior esquerdo
        bottom_layout.addWidget(self.botao_readme, alignment=Qt.AlignLeft)

        # Botão "Encerrar Programa" no canto inferior direito
        self.botao_encerrar = self.create_button("Encerrar Programa", self.close_program, "#8B1A1A", "#A52424")
        bottom_layout.addWidget(self.botao_encerrar, alignment=Qt.AlignRight)

        # Adiciona o layout dos botões na parte inferior
        main_layout.addLayout(bottom_layout)

        self.setWindowTitle('Conversor de Imagens')
        self.setFixedSize(400, 300)
        self.show()

    def create_button(self, text, function, background_color, hover_color):

        button = QPushButton(text)
        button.clicked.connect(function)
        button.setStyleSheet(self.create_button_style(background_color, hover_color))
        return button

    def create_button_style(self, background_color, hover_color):

        return f"""
            QPushButton {{
                background-color: {background_color};
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
        """
    
    def open_readme(self):
        # Caminho absoluto para o arquivo readme.txt
        readme_path = os.path.join(os.getcwd(),'readme.txt')

        # Verifica se o arquivo existe
        if os.path.isfile(readme_path):
            try:
                # Usa subprocess para abrir o arquivo com o Bloco de Notas (notepad.exe)
                subprocess.run(['notepad.exe', readme_path], check=True)
            except FileNotFoundError:
                print("Erro: notepad.exe não encontrado.")
            except Exception as e:
                print(f"Erro ao tentar abrir o arquivo no Bloco de Notas: {e}")
        else:
            print(f"Arquivo não encontrado: {readme_path}")

    def close_program(self):

        print("Encerrando...")
        self.close()

    def open_second_window(self):
        # Obtém o diretório onde o programa foi executado
        base_dir = os.getcwd()

        # Defina os nomes das pastas relativas ao diretório atual
        folders_to_create = [
            os.path.join(base_dir, "Imagens_TIFF"),  # Pasta 1
            os.path.join(base_dir, "Lotes_Upload"),  # Pasta 2
            os.path.join(base_dir, "PDFs_Baixados")   # Pasta 3
        ]

        # Verifique se as pastas existem, caso contrário, crie-as
        for folder in folders_to_create:
            if not os.path.exists(folder):
                try:
                    os.makedirs(folder)
                    print(f"Pasta criada: {folder}")
                except Exception as e:
                    print(f"Erro ao criar a pasta {folder}: {e}")

        # Agora abre a segunda janela
        self.SecondWindow = SecondWindow()
        self.SecondWindow.show()
        self.close()

class SecondWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Processamento de Imagens")
        self.setFixedSize(400, 300)
        self.is_dark_mode = False
        self.setWindowIcon(QIcon(r".\Icones\icone.png"))
        self.initUI()

    def initUI(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        logo_label = QLabel(self)
        pixmap = QPixmap(r".\Icones\logo.png")  
        pixmap = pixmap.scaled(200, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)  
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignCenter)  
        main_layout.addWidget(logo_label)  

        self.botao_download = self.create_button(
            "Baixar Lotes", self.downloadAPI, "#003366", "#3366CC"
        )
        self.botao_processar = self.create_button(
            "Processar PDFs", self.ProcessarPdfs, "#0066CC", "#66B2FF"
        )
        self.botao_read = self.create_button(
            "Ler QrCodes", self.LerQrCodes, "#0066CC", "#66B2FF"
        )
        self.botao_corrigir = self.create_button(
            "Corrigir Lotes", self.Correcao, "#66CCFF", "#CCE5FF"
        )
        self.botao_upload = self.create_button(
            "Upload de Imagens", self.uploadAPI, "#66CCFF", "#CCE5FF"
        )
        self.botao_preparar = self.create_button(
            "Preparar Lotes", self.prepararo_winrar, "#0066CC", "#66B2FF"
        )
        self.botao_retornar = self.create_button(
            "Retornar", self.Return, "#0066CC", "#66B2FF"
        )
        self.botao_encerrar = self.create_button(
            "Encerrar Programa", self.close_program, "#8B1A1A", "#A52424"
        )

        button_layout = QGridLayout()
        button_layout.addWidget(self.botao_download, 0, 0)  
        button_layout.addWidget(self.botao_processar, 0, 1) 
        button_layout.addWidget(self.botao_read, 1, 0)       
        button_layout.addWidget(self.botao_corrigir, 1, 1) 
        button_layout.addWidget(self.botao_upload, 2, 0)  
        button_layout.addWidget(self.botao_preparar, 2, 1)
        button_layout.addWidget(self.botao_retornar, 3, 0)  
        button_layout.addWidget(self.botao_encerrar, 3, 1)

        main_layout.addLayout(button_layout)

    def create_button(self, text, function, background_color, hover_color):

        button = QPushButton(text)
        button.clicked.connect(function)
        button.setStyleSheet(self.create_button_style(background_color, hover_color))
        return button

    def create_button_style(self, background_color, hover_color):

        return f"""
            QPushButton {{
                background-color: {background_color};
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
            }}
        """

    def downloadAPI(self):
    # Diálogo de progresso inicial
        dialogo_inicio = QMessageBox()
        dialogo_inicio.setIcon(QMessageBox.Information)
        dialogo_inicio.setWindowTitle("Download em Progresso")
        dialogo_inicio.setText("Buscando Lotes disponíveis...")
        dialogo_inicio.setStandardButtons(QMessageBox.Ok)
        dialogo_inicio.exec_()

        # Executa a função de download
        iniciando_download()

        # Diálogo de conclusão
        dialogo_concluido = QMessageBox()
        dialogo_concluido.setIcon(QMessageBox.Information)
        dialogo_concluido.setWindowTitle("Download Concluído")
        dialogo_concluido.setText("Download Concluído com sucesso.")
        dialogo_concluido.setStandardButtons(QMessageBox.Ok)
        dialogo_concluido.exec_()
    
    def uploadAPI(self):

        dialogo_inicio = QMessageBox()
        dialogo_inicio.setIcon(QMessageBox.Information)
        dialogo_inicio.setWindowTitle("Upload em Progresso")
        dialogo_inicio.setText("Realizando Upload...")
        dialogo_inicio.setStandardButtons(QMessageBox.Ok)
        dialogo_inicio.exec_()

        print("Iniciando upload")

        dialogo_concluido = QMessageBox()
        dialogo_concluido.setIcon(QMessageBox.Information)
        dialogo_concluido.setWindowTitle("Upload Concluído")
        dialogo_concluido.setText("Upload Concluído com sucesso.")
        dialogo_concluido.setStandardButtons(QMessageBox.Ok)
        dialogo_concluido.exec_()

    def ProcessarPdfs(self):

        dialogo_inicio = QMessageBox()
        dialogo_inicio.setIcon(QMessageBox.Information)
        dialogo_inicio.setWindowTitle("Processamento em Progresso")
        dialogo_inicio.setText("Processando Lotes ...")
        dialogo_inicio.setStandardButtons(QMessageBox.Ok)
        dialogo_inicio.exec_()

        self.pr_thread = ProcessThread()
        self.pr_thread.finished.connect(self.processo_concluido)
        self.pr_thread.error.connect(self.mostrar_erro)
        self.pr_thread.start()

    def processo_concluido(self):
        QMessageBox.information(self, "Processamento Concluído", "Todos os PDFs foram processados com sucesso.")

    def LerQrCodes(self):
        # Diálogo de progresso inicial
        dialogo_inicio = QMessageBox()
        dialogo_inicio.setIcon(QMessageBox.Information)
        dialogo_inicio.setWindowTitle("Lendo QRCodes...")
        dialogo_inicio.setText("Buscando Lotes disponíveis...")
        dialogo_inicio.setStandardButtons(QMessageBox.Ok)
        dialogo_inicio.exec_()

        # Configura e inicia a thread para evitar bloqueio
        self.qr_thread = QrCodeThread()
        self.qr_thread.finished.connect(self.processo_concluido)
        self.qr_thread.error.connect(self.mostrar_erro)
        self.qr_thread.start()

    def mostrar_erro(self, mensagem):
        # Exibe o erro, se houver
        dialogo_erro = QMessageBox()
        dialogo_erro.setIcon(QMessageBox.Critical)
        dialogo_erro.setWindowTitle("Erro")
        dialogo_erro.setText(f"Ocorreu um erro: {mensagem}")
        dialogo_erro.setStandardButtons(QMessageBox.Ok)
        dialogo_erro.exec_()

    def Return(self):
        self.MainWindow = MainWindow()
        self.MainWindow.show()
        self.close()

    def prepararo_winrar(self):

        dialogo_inicio = QMessageBox()
        dialogo_inicio.setIcon(QMessageBox.Information)
        dialogo_inicio.setWindowTitle("Preparando Lote para Upload")
        dialogo_inicio.setText("Aguarde um instante...")
        dialogo_inicio.setStandardButtons(QMessageBox.Ok)
        dialogo_inicio.exec_()

        self.rar_thread = RarThread()
        self.rar_thread.finished.connect(self.processo_concluido)
        self.rar_thread.error.connect(self.mostrar_erro)
        self.rar_thread.start()

    def Correcao(self):

        print("Iniciando módulo de correção...")
        self.Corr_Window = Corr_Window()
        self.Corr_Window.show()
        self.close()

    def close_program(self):
        dialogo_inicio = QMessageBox()
        dialogo_inicio.setIcon(QMessageBox.Information)
        dialogo_inicio.setWindowTitle("Até Mais!!")
        dialogo_inicio.setText("Encerrando Software")
        dialogo_inicio.setStandardButtons(QMessageBox.Ok)
        dialogo_inicio.exec_()
        self.close()

class Corr_Window(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Processamento de Imagens")
        self.setWindowIcon(QIcon(r".\Icones\icone.png"))
        self.showMaximized()
        self.image_loaded = False
        self.image_paths = []
        self.current_index = 0
        self.initUI()

    def initUI(self):
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setDragMode(QGraphicsView.ScrollHandDrag)

        self.previous_button = self.create_button(
            "Imagem Anterior", self.show_previous_image, "#5A5A5A", "#767676"
        )
        self.next_button = self.create_button(
            "Próxima Imagem", self.show_next_image, "#5A5A5A", "#767676"
        )

        self.load_txt_button = self.create_button(
            "Carregar Arquivo .txt",
            self.load_images_from_txt_file,
            "#5A5A5A",
            "#767676",
        )

        self.rename_input = QLineEdit(self)
        self.rename_input.setPlaceholderText("Novo nome do arquivo (sem extensão)")

        self.rename_button = self.create_button(
            "Renomear", self.rename_file, "#5A5A5A", "#767676"
        )

        self.convert_button = self.create_button(
            "Converter em PDF", self.convert_to_pdf, "#5A5A5A", "#767676"
        )

        # Adicionando o botão de retorno
        self.return_button = self.create_button(
            "Voltar", self.Return, "#5A5A5A", "#767676"
        )

        main_layout = QHBoxLayout()

        image_layout = QVBoxLayout()
        image_layout.addWidget(self.view)

        navigation_layout = QHBoxLayout()
        navigation_layout.addWidget(self.load_txt_button)
        navigation_layout.addWidget(self.previous_button)
        navigation_layout.addWidget(self.next_button)

        rename_layout = QHBoxLayout()
        rename_layout.addWidget(self.rename_input)
        rename_layout.addWidget(self.rename_button)
        rename_layout.addWidget(self.convert_button)

        # Layout para o botão de retorno
        return_layout = QVBoxLayout()
        return_layout.addWidget(self.return_button)

        image_layout.addLayout(navigation_layout)
        image_layout.addLayout(rename_layout)
        image_layout.addLayout(return_layout)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        main_layout.addLayout(image_layout, 3)

        self.close_button = self.create_button(
            "Fechar", self.close, "#8B1A1A", "#A52424"
        )
        button_layout = QVBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)
        main_layout.addLayout(button_layout, 1)

    def create_button(self, text, slot, normal_color, hover_color):
        button = QPushButton(text)
        button.setStyleSheet(f"background-color: {normal_color};")
        button.setMouseTracking(True)
        button.clicked.connect(slot)
        button.setMinimumHeight(30)

        return button

    def load_images_from_txt_file(self):
        txt_file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo .txt", "", "Text Files (*.txt)"
        )
        if txt_file_path:
            self.image_paths = []
            self.extract_image_paths_from_txt(txt_file_path)

            if not self.image_paths:
                print("Nenhum caminho de imagem encontrado no arquivo .txt.")
                return

            self.current_index = 0
            self.display_image()

    def extract_image_paths_from_txt(self, txt_file_path):
        try:
            with open(txt_file_path, "r") as file:
                for line in file.readlines():
                    path = line.strip()
                    if os.path.exists(path):
                        self.image_paths.append(path)
        except Exception as e:
            print(f"Erro ao ler o arquivo {txt_file_path}: {e}")

    def display_image(self):
        if self.image_paths:
            path = self.image_paths[self.current_index]
            if os.path.exists(path):
                image = QImage(path)
                pixmap = QPixmap.fromImage(image)
                self.scene.clear()
                self.scene.addPixmap(pixmap)
            else:
                print(f"Imagem não encontrada: {path}")

    def show_previous_image(self):
        if self.image_paths and self.current_index > 0:
            self.current_index -= 1
            self.display_image()

    def show_next_image(self):
        if self.image_paths and self.current_index < len(self.image_paths) - 1:
            self.current_index += 1
            self.display_image()

    def rename_file(self):
        if not self.image_paths:
            QMessageBox.warning(
                self, "Aviso", "Nenhuma imagem carregada para renomear."
            )
            return

        new_name = self.rename_input.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Aviso", "Por favor, insira um novo nome.")
            return

        current_image_path = self.image_paths[self.current_index]
        directory = os.path.dirname(current_image_path)
        new_image_path = os.path.join(directory, f"{new_name}.TIF")

        try:
            os.rename(current_image_path, new_image_path)
            QMessageBox.information(
                self, "Sucesso", f"Arquivo renomeado para: {new_name}.tif"
            )
            self.image_paths[self.current_index] = new_image_path
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao renomear o arquivo: {e}")

    def convert_to_pdf(self):
        if not self.image_paths:
            QMessageBox.warning(
                self, "Aviso", "Nenhuma imagem carregada para converter."
            )
            return

        # Cria uma pasta com a data e hora atuais dentro de Lotes_Upload
        now = QDateTime.currentDateTime()
        formatted_datetime = now.toString("yyyy_MM_dd_hh_mm_ss")  # Ex: 20231106_123456
        upload_folder = os.path.join(os.getcwd(), "Lotes_Upload", f"Lote_{formatted_datetime}")

        # Cria a pasta se não existir
        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)

        current_image_path = self.image_paths[self.current_index]
        pdf_filename = os.path.basename(current_image_path).replace(".tif", ".pdf").replace(".TIF", ".pdf")
        pdf_path = os.path.join(upload_folder, pdf_filename)

        try:
            # Converte a imagem para PDF
            image = Image.open(current_image_path)
            image.save(pdf_path, "PDF", resolution=100.0)

            QMessageBox.information(
                self, "Sucesso", f"Imagem convertida para PDF: {pdf_path}"
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erro", f"Erro ao converter a imagem para PDF: {e}"
            )

    def Return(self):
        self.SecondWindow = SecondWindow()
        self.SecondWindow.show()
        self.close()

    def wheelEvent(self, event):
        if self.image_loaded:
            zoom_in = event.angleDelta().y() > 0
            factor = 1.25 if zoom_in else 0.8
            self.view.scale(factor, factor)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
