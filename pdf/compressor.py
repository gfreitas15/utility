from __future__ import annotations

import os

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
    QComboBox,
)

from shared.pdf_deps import PIL_AVAILABLE, Image, A4, canvas


class CompressorPDFWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_path = ""
        self.output_path = ""
        self.poppler_path = None  # Caminho do Poppler (None = usar PATH do sistema)
        self.init_ui()
        self.verificar_poppler()

    def init_ui(self):
        if not PIL_AVAILABLE:
            self._mostrar_erro_dependencias()
            return

        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Título
        titulo = QLabel("🗜️ Compressor de PDF")
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setFont(QFont("Arial", 16, QFont.Bold))
        layout.addWidget(titulo)

        # Instruções
        instrucoes = QLabel("Comprima PDFs com imagens escaneadas para reduzir o tamanho do arquivo")
        instrucoes.setAlignment(Qt.AlignCenter)
        instrucoes.setFont(QFont("Arial", 10))
        instrucoes.setStyleSheet("color: #666; margin: 10px;")
        layout.addWidget(instrucoes)

        # Seleção do PDF
        pdf_group = QGroupBox("📄 Arquivo PDF")
        pdf_layout = QVBoxLayout()

        pdf_input_layout = QHBoxLayout()
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("Selecione um arquivo PDF...")
        self.pdf_path_edit.setReadOnly(True)
        pdf_input_layout.addWidget(self.pdf_path_edit)

        self.btn_selecionar_pdf = QPushButton("📁 Selecionar PDF")
        self.btn_selecionar_pdf.clicked.connect(self.selecionar_pdf)
        self.btn_selecionar_pdf.setStyleSheet("background-color: #3498db; color: white;")
        pdf_input_layout.addWidget(self.btn_selecionar_pdf)
        pdf_layout.addLayout(pdf_input_layout)

        # Informações do arquivo
        self.lbl_info_arquivo = QLabel("Nenhum arquivo selecionado")
        self.lbl_info_arquivo.setWordWrap(True)
        self.lbl_info_arquivo.setStyleSheet("color: #7f8c8d; padding: 5px;")
        pdf_layout.addWidget(self.lbl_info_arquivo)

        pdf_group.setLayout(pdf_layout)
        layout.addWidget(pdf_group)

        # Opções de compressão
        opcoes_group = QGroupBox("⚙️ Opções de Compressão")
        opcoes_layout = QVBoxLayout()

        # Qualidade da imagem
        qualidade_layout = QHBoxLayout()
        qualidade_layout.addWidget(QLabel("Qualidade da imagem:"))
        self.qualidade_spin = QSpinBox()
        self.qualidade_spin.setMinimum(10)
        self.qualidade_spin.setMaximum(100)
        self.qualidade_spin.setValue(75)
        self.qualidade_spin.setSuffix("%")
        qualidade_layout.addWidget(self.qualidade_spin)
        qualidade_layout.addStretch()
        opcoes_layout.addLayout(qualidade_layout)

        # Resolução DPI
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(QLabel("Resolução (DPI):"))
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setMinimum(72)
        self.dpi_spin.setMaximum(300)
        self.dpi_spin.setValue(150)
        self.dpi_spin.setSuffix(" DPI")
        dpi_layout.addWidget(self.dpi_spin)
        dpi_layout.addStretch()
        opcoes_layout.addLayout(dpi_layout)

        # Formato de compressão
        formato_layout = QHBoxLayout()
        formato_layout.addWidget(QLabel("Formato de compressão:"))
        self.formato_combo = QComboBox()
        self.formato_combo.addItems(["JPEG (melhor compressão)", "PNG (melhor qualidade)"])
        formato_layout.addWidget(self.formato_combo)
        formato_layout.addStretch()
        opcoes_layout.addLayout(formato_layout)

        opcoes_group.setLayout(opcoes_layout)
        layout.addWidget(opcoes_group)

        # Configuração do Poppler
        poppler_group = QGroupBox("🔧 Configuração do Poppler")
        poppler_layout = QVBoxLayout()

        poppler_input_layout = QHBoxLayout()
        self.poppler_path_edit = QLineEdit()
        self.poppler_path_edit.setPlaceholderText("Deixe vazio para usar PATH do sistema...")
        self.poppler_path_edit.setReadOnly(True)
        poppler_input_layout.addWidget(self.poppler_path_edit)

        self.btn_selecionar_poppler = QPushButton("📁 Selecionar Poppler")
        self.btn_selecionar_poppler.clicked.connect(self.selecionar_poppler)
        self.btn_selecionar_poppler.setStyleSheet("background-color: #9b59b6; color: white;")
        poppler_input_layout.addWidget(self.btn_selecionar_poppler)

        self.btn_limpar_poppler = QPushButton("🔄 Usar PATH")
        self.btn_limpar_poppler.clicked.connect(self.limpar_poppler)
        self.btn_limpar_poppler.setStyleSheet("background-color: #95a5a6; color: white;")
        poppler_input_layout.addWidget(self.btn_limpar_poppler)
        poppler_layout.addLayout(poppler_input_layout)

        # Status do Poppler
        self.lbl_status_poppler = QLabel("Verificando Poppler...")
        self.lbl_status_poppler.setWordWrap(True)
        self.lbl_status_poppler.setStyleSheet("color: #7f8c8d; padding: 5px;")
        poppler_layout.addWidget(self.lbl_status_poppler)

        poppler_group.setLayout(poppler_layout)
        layout.addWidget(poppler_group)

        # Pasta de saída
        output_group = QGroupBox("💾 Arquivo de Saída")
        output_layout = QVBoxLayout()

        output_input_layout = QHBoxLayout()
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("Arquivo será salvo na mesma pasta do original...")
        self.output_path_edit.setReadOnly(True)
        output_input_layout.addWidget(self.output_path_edit)

        self.btn_selecionar_output = QPushButton("📁 Escolher Local")
        self.btn_selecionar_output.clicked.connect(self.selecionar_output)
        self.btn_selecionar_output.setStyleSheet("background-color: #3498db; color: white;")
        output_input_layout.addWidget(self.btn_selecionar_output)
        output_layout.addLayout(output_input_layout)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Botão de compressão
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_comprimir = QPushButton("🗜️ Comprimir PDF")
        self.btn_comprimir.clicked.connect(self.comprimir_pdf)
        self.btn_comprimir.setStyleSheet(
            "background-color: #2ecc71; color: white; padding: 12px 24px; font-size: 14px; font-weight: bold;"
        )
        self.btn_comprimir.setEnabled(False)
        btn_layout.addWidget(self.btn_comprimir)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Status
        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setWordWrap(True)
        layout.addWidget(self.lbl_status)

        layout.addStretch()

    def _mostrar_erro_dependencias(self):
        layout = QVBoxLayout()
        msg = QLabel(
            "❌ Dependências não instaladas!\n\nPara usar o Compressor de PDF, instale as dependências:\n\npip install Pillow reportlab PyPDF2 pdf2image"
        )
        msg.setAlignment(Qt.AlignCenter)
        msg.setFont(QFont("Segoe UI", 12))
        msg.setStyleSheet("color: red; padding: 20px;")
        layout.addWidget(msg)
        self.setLayout(layout)

    def verificar_poppler(self):
        """Verifica se o Poppler está disponível"""
        try:
            import subprocess
            import shutil

            # Tentar encontrar pdftoppm no PATH
            poppler_exe = shutil.which("pdftoppm")

            if poppler_exe:
                # Testar se funciona
                try:
                    result = subprocess.run(
                        ["pdftoppm", "-h"], capture_output=True, timeout=5, text=True
                    )
                    if result.returncode == 0 or "pdftoppm" in result.stderr.lower():
                        self.lbl_status_poppler.setText("✅ Poppler encontrado no PATH do sistema")
                        self.lbl_status_poppler.setStyleSheet(
                            "color: #2ecc71; padding: 5px; font-weight: bold;"
                        )
                        self.poppler_path = None
                        return
                except Exception:
                    pass

            # Tentar encontrar em locais comuns do Windows
            locais_comuns = [
                r"C:\poppler\Library\bin",
                r"C:\Program Files\poppler\Library\bin",
                r"C:\Program Files (x86)\poppler\Library\bin",
                os.path.join(os.path.expanduser("~"), "poppler", "Library", "bin"),
            ]

            for local in locais_comuns:
                pdftoppm_path = os.path.join(local, "pdftoppm.exe")
                if os.path.exists(pdftoppm_path):
                    self.poppler_path = local
                    self.poppler_path_edit.setText(local)
                    self.lbl_status_poppler.setText(f"✅ Poppler encontrado em: {local}")
                    self.lbl_status_poppler.setStyleSheet(
                        "color: #2ecc71; padding: 5px; font-weight: bold;"
                    )
                    return

            # Não encontrado
            self.lbl_status_poppler.setText(
                "⚠️ Poppler não encontrado automaticamente.\n"
                "Clique em 'Selecionar Poppler' para especificar o caminho manualmente."
            )
            self.lbl_status_poppler.setStyleSheet("color: #e67e22; padding: 5px;")

        except Exception as e:
            self.lbl_status_poppler.setText(f"Erro ao verificar Poppler: {str(e)}")
            self.lbl_status_poppler.setStyleSheet("color: red; padding: 5px;")

    def selecionar_poppler(self):
        """Permite ao usuário selecionar a pasta do Poppler manualmente"""
        pasta = QFileDialog.getExistingDirectory(
            self, "Selecionar pasta do Poppler (deve conter pdftoppm.exe)", ""
        )
        if pasta:
            # Verificar se contém pdftoppm.exe
            pdftoppm_path = os.path.join(pasta, "pdftoppm.exe")
            if os.path.exists(pdftoppm_path):
                self.poppler_path = pasta
                self.poppler_path_edit.setText(pasta)
                self.lbl_status_poppler.setText(f"✅ Poppler configurado: {pasta}")
                self.lbl_status_poppler.setStyleSheet(
                    "color: #2ecc71; padding: 5px; font-weight: bold;"
                )
            else:
                QMessageBox.warning(
                    self,
                    "Caminho Inválido",
                    "A pasta selecionada não contém pdftoppm.exe!\n\n"
                    "Por favor, selecione a pasta 'bin' do Poppler que contém os executáveis.",
                )

    def limpar_poppler(self):
        """Limpa o caminho do Poppler e usa o PATH do sistema"""
        self.poppler_path = None
        self.poppler_path_edit.clear()
        self.verificar_poppler()

    def _verificar_poppler_disponivel(self):
        """Verifica se o Poppler está disponível para uso"""
        try:
            import subprocess
            import shutil

            # Se o caminho foi especificado, verificar diretamente
            if self.poppler_path:
                pdftoppm_path = os.path.join(self.poppler_path, "pdftoppm.exe")
                if os.path.exists(pdftoppm_path):
                    return True

            # Verificar no PATH
            poppler_exe = shutil.which("pdftoppm")
            if poppler_exe:
                try:
                    subprocess.run(["pdftoppm", "-h"], capture_output=True, timeout=5)
                    return True
                except Exception:
                    pass

            return False
        except Exception:
            return False

    def selecionar_pdf(self):
        arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        if arquivo:
            self.pdf_path = arquivo
            self.pdf_path_edit.setText(arquivo)

            # Mostrar informações do arquivo
            try:
                tamanho = os.path.getsize(arquivo) / (1024 * 1024)  # MB
                self.lbl_info_arquivo.setText(f"Tamanho: {tamanho:.2f} MB")
                self.lbl_info_arquivo.setStyleSheet(
                    "color: #2ecc71; padding: 5px; font-weight: bold;"
                )
            except Exception:
                self.lbl_info_arquivo.setText("Arquivo selecionado")

            # Atualizar caminho de saída padrão
            if not self.output_path:
                nome_base = os.path.splitext(arquivo)[0]
                self.output_path = f"{nome_base}_comprimido.pdf"
                self.output_path_edit.setText(self.output_path)

            self.btn_comprimir.setEnabled(True)
            self.lbl_status.setText("")

    def selecionar_output(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "Aviso", "Selecione primeiro o arquivo PDF!")
            return

        nome_base = os.path.splitext(os.path.basename(self.pdf_path))[0]
        arquivo, _ = QFileDialog.getSaveFileName(
            self, "Salvar PDF Comprimido", f"{nome_base}_comprimido.pdf", "Arquivos PDF (*.pdf)"
        )
        if arquivo:
            self.output_path = arquivo
            self.output_path_edit.setText(arquivo)

    def comprimir_pdf(self):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            QMessageBox.warning(self, "Erro", "Selecione um arquivo PDF válido!")
            return

        if not self.output_path:
            QMessageBox.warning(self, "Erro", "Selecione um local para salvar o arquivo comprimido!")
            return

        # Verificar se o Poppler está disponível
        if not self._verificar_poppler_disponivel():
            resposta = QMessageBox.question(
                self,
                "Poppler não encontrado",
                "O Poppler não foi encontrado. Deseja continuar mesmo assim?\n\n"
                "Recomendamos configurar o Poppler antes de continuar.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if resposta == QMessageBox.No:
                return

        # Verificar se o arquivo de saída já existe
        if os.path.exists(self.output_path) and self.output_path != self.pdf_path:
            resposta = QMessageBox.question(
                self,
                "Arquivo Existente",
                f"O arquivo {os.path.basename(self.output_path)} já existe.\nDeseja substituí-lo?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if resposta == QMessageBox.No:
                return

        # Desabilitar botão durante processamento
        self.btn_comprimir.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Iniciando compressão...")
        self.lbl_status.setStyleSheet("color: #3498db; font-weight: bold;")

        # Executar compressão em thread separada
        self.worker = CompressorPDFWorker(
            self.pdf_path,
            self.output_path,
            self.qualidade_spin.value(),
            self.dpi_spin.value(),
            self.formato_combo.currentText(),
            self.poppler_path,
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(self._compressao_finalizada)
        self.worker.error.connect(self._erro_compressao)
        self.worker.start()

    def _compressao_finalizada(self, resultado):
        self.progress_bar.setVisible(False)
        self.btn_comprimir.setEnabled(True)

        tamanho_original = resultado["tamanho_original"] / (1024 * 1024)  # MB
        tamanho_comprimido = resultado["tamanho_comprimido"] / (1024 * 1024)  # MB
        reducao = ((tamanho_original - tamanho_comprimido) / tamanho_original) * 100

        mensagem = (
            "✅ Compressão concluída!\n\n"
            f"Tamanho original: {tamanho_original:.2f} MB\n"
            f"Tamanho comprimido: {tamanho_comprimido:.2f} MB\n"
            f"Redução: {reducao:.1f}%\n\n"
            f"Arquivo salvo em:\n{self.output_path}"
        )

        self.lbl_status.setText(mensagem)
        self.lbl_status.setStyleSheet("color: #2ecc71; font-weight: bold;")

        QMessageBox.information(self, "Sucesso", mensagem)

    def _erro_compressao(self, mensagem_erro):
        self.progress_bar.setVisible(False)
        self.btn_comprimir.setEnabled(True)
        self.lbl_status.setText(f"❌ Erro: {mensagem_erro}")
        self.lbl_status.setStyleSheet("color: red; font-weight: bold;")
        QMessageBox.critical(self, "Erro", f"Erro ao comprimir PDF:\n\n{mensagem_erro}")


class CompressorPDFWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, pdf_path, output_path, qualidade, dpi, formato, poppler_path=None):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_path = output_path
        self.qualidade = qualidade
        self.dpi = dpi
        self.formato = formato
        self.poppler_path = poppler_path

    def run(self):
        try:
            import tempfile
            import shutil

            # Obter tamanho original
            tamanho_original = os.path.getsize(self.pdf_path)

            # Ler o PDF
            from PyPDF2 import PdfReader

            reader = PdfReader(self.pdf_path)
            total_paginas = len(reader.pages)

            if total_paginas == 0:
                self.error.emit("O PDF não contém páginas!")
                return

            # Criar diretório temporário
            temp_dir = tempfile.mkdtemp()

            try:
                # Converter cada página em imagem
                imagens_comprimidas = []

                for i, page in enumerate(reader.pages):
                    self.progress.emit(int((i / total_paginas) * 50))

                    # Converter página em imagem
                    try:
                        from pdf2image import convert_from_path

                        # Preparar parâmetros para convert_from_path
                        kwargs = {"dpi": self.dpi, "first_page": i + 1, "last_page": i + 1}

                        # Se o caminho do Poppler foi especificado, adicionar ao kwargs
                        if self.poppler_path:
                            kwargs["poppler_path"] = self.poppler_path

                        imagens = convert_from_path(self.pdf_path, **kwargs)
                        if not imagens:
                            continue
                        imagem = imagens[0]
                    except Exception as e:
                        # Verificar se é erro de poppler
                        erro_msg = str(e).lower()
                        if (
                            "poppler" in erro_msg
                            or "pdftoppm" in erro_msg
                            or "cannot find" in erro_msg
                            or "not found" in erro_msg
                        ):
                            mensagem_erro = (
                                "Erro: Poppler não encontrado!\n\n"
                                "O pdf2image requer o Poppler instalado no sistema.\n\n"
                                "SOLUÇÕES:\n"
                                "1. Na aba do Compressor, clique em 'Selecionar Poppler' e escolha a pasta 'bin' do Poppler\n"
                                "2. Ou adicione o Poppler ao PATH do sistema e reinicie o programa\n\n"
                                "Download: https://github.com/oschwartz10612/poppler-windows/releases\n\n"
                                f"Erro detalhado: {str(e)}"
                            )
                            self.error.emit(mensagem_erro)
                        else:
                            self.error.emit(f"Erro ao converter página {i+1}: {str(e)}")
                        return

                    # Comprimir imagem
                    if self.formato == "JPEG (melhor compressão)":
                        # Converter para RGB se necessário
                        if imagem.mode != "RGB":
                            imagem = imagem.convert("RGB")

                        # Salvar temporariamente
                        temp_path = os.path.join(temp_dir, f"page_{i}.jpg")
                        imagem.save(temp_path, "JPEG", quality=self.qualidade, optimize=True)
                    else:  # PNG
                        # Salvar temporariamente
                        temp_path = os.path.join(temp_dir, f"page_{i}.png")
                        imagem.save(temp_path, "PNG", optimize=True)

                    # Fechar a imagem explicitamente para liberar o arquivo
                    imagem.close()
                    del imagem
                    imagens_comprimidas.append(temp_path)

                # Criar novo PDF com imagens comprimidas
                self.progress.emit(60)

                from reportlab.pdfgen import canvas
                from reportlab.lib.pagesizes import A4

                # Determinar tamanho da página baseado na primeira imagem
                primeira_imagem = None
                if imagens_comprimidas:
                    primeira_imagem = Image.open(imagens_comprimidas[0])
                    largura, altura = primeira_imagem.size
                    # Converter pixels para pontos (1 ponto = 1/72 polegada)
                    largura_pt = largura * 72 / self.dpi
                    altura_pt = altura * 72 / self.dpi
                    primeira_imagem.close()
                    del primeira_imagem
                else:
                    largura_pt, altura_pt = A4

                # Criar PDF
                c = canvas.Canvas(self.output_path, pagesize=(largura_pt, altura_pt))

                for i, img_path in enumerate(imagens_comprimidas):
                    self.progress.emit(60 + int((i / len(imagens_comprimidas)) * 35))

                    # Abrir imagem apenas para obter dimensões
                    imagem = Image.open(img_path)
                    largura, altura = imagem.size
                    largura_pt = largura * 72 / self.dpi
                    altura_pt = altura * 72 / self.dpi
                    imagem.close()
                    del imagem

                    # Ajustar tamanho da página se necessário
                    c.setPageSize((largura_pt, altura_pt))

                    # Adicionar imagem (reportlab abre o arquivo internamente)
                    c.drawImage(img_path, 0, 0, width=largura_pt, height=altura_pt)
                    c.showPage()

                c.save()

                # Limpar diretório temporário
                # Adicionar um pequeno delay para garantir que todos os arquivos foram fechados
                import time

                time.sleep(0.1)

                # Tentar remover com retry em caso de erro
                max_tentativas = 5
                for tentativa in range(max_tentativas):
                    try:
                        shutil.rmtree(temp_dir)
                        break
                    except PermissionError:
                        if tentativa < max_tentativas - 1:
                            time.sleep(0.2)
                        else:
                            # Se ainda não conseguir, tentar remover arquivos individualmente
                            try:
                                for arquivo in os.listdir(temp_dir):
                                    arquivo_path = os.path.join(temp_dir, arquivo)
                                    try:
                                        os.remove(arquivo_path)
                                    except Exception:
                                        pass
                                os.rmdir(temp_dir)
                            except Exception:
                                # Se ainda falhar, apenas avisar mas não bloquear
                                pass

                # Obter tamanho final
                tamanho_comprimido = os.path.getsize(self.output_path)

                self.progress.emit(100)

                self.finished.emit({"tamanho_original": tamanho_original, "tamanho_comprimido": tamanho_comprimido})

            except Exception as e:
                # Limpar diretório temporário em caso de erro
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                raise e

        except Exception as e:
            self.error.emit(str(e))

