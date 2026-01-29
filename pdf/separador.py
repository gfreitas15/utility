from __future__ import annotations

import os
from pathlib import Path

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QCheckBox,
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
)

from shared.pdf_deps import PdfReader, PdfWriter, xlsxwriter


class SeparadorPDFWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.senha_correta = "Netman50!"  # Senha para acessar a funcionalidade
        self.autenticado = False
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Tela de autenticação
        self.auth_widget = QWidget()
        auth_layout = QVBoxLayout(self.auth_widget)

        # Título
        titulo = QLabel("🔐 Acesso Restrito - Separador de PDF")
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setFont(QFont("Arial", 16, QFont.Bold))
        auth_layout.addWidget(titulo)

        # Instruções
        instrucoes = QLabel("Esta funcionalidade requer autenticação.\nDigite a senha para continuar:")
        instrucoes.setAlignment(Qt.AlignCenter)
        instrucoes.setFont(QFont("Arial", 12))
        auth_layout.addWidget(instrucoes)

        # Campo de senha
        senha_layout = QHBoxLayout()
        senha_layout.addStretch()
        self.senha_input = QLineEdit()
        self.senha_input.setEchoMode(QLineEdit.Password)
        self.senha_input.setPlaceholderText("Digite a senha...")
        self.senha_input.setMaximumWidth(200)
        self.senha_input.returnPressed.connect(self.verificar_senha)
        senha_layout.addWidget(self.senha_input)
        senha_layout.addStretch()
        auth_layout.addLayout(senha_layout)

        # Botão de login
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_login = QPushButton("🔓 Acessar")
        self.btn_login.clicked.connect(self.verificar_senha)
        self.btn_login.setStyleSheet(
            "background-color: #2ecc71; color: white; padding: 10px 20px; font-size: 14px;"
        )
        btn_layout.addWidget(self.btn_login)
        btn_layout.addStretch()
        auth_layout.addLayout(btn_layout)

        # Status
        self.lbl_status_auth = QLabel("")
        self.lbl_status_auth.setAlignment(Qt.AlignCenter)
        self.lbl_status_auth.setStyleSheet("color: red; font-weight: bold;")
        auth_layout.addWidget(self.lbl_status_auth)

        auth_layout.addStretch()

        # Widget principal (inicialmente oculto)
        self.main_widget = QWidget()
        self.setup_main_ui()

        layout.addWidget(self.auth_widget)
        layout.addWidget(self.main_widget)

        self.main_widget.setVisible(False)

    def verificar_senha(self):
        senha = self.senha_input.text()
        if senha == self.senha_correta:
            self.autenticado = True
            self.auth_widget.setVisible(False)
            self.main_widget.setVisible(True)
            self.lbl_status_auth.setText("")
        else:
            self.lbl_status_auth.setText("❌ Senha incorreta!")
            self.senha_input.clear()
            self.senha_input.setFocus()

    def setup_main_ui(self):
        layout = QVBoxLayout(self.main_widget)

        # Título
        titulo = QLabel("📄 Separador de PDF por Marcadores")
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setFont(QFont("Arial", 16, QFont.Bold))
        layout.addWidget(titulo)

        # Instruções
        instrucoes = QLabel(
            "Selecione um PDF com marcadores (bookmarks) para separar em arquivos menores"
        )
        instrucoes.setAlignment(Qt.AlignCenter)
        instrucoes.setFont(QFont("Arial", 10))
        instrucoes.setStyleSheet("color: #666; margin: 10px;")
        layout.addWidget(instrucoes)

        # Seleção do PDF
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(QLabel("Arquivo PDF:"))
        self.pdf_path = QLineEdit()
        self.pdf_path.setPlaceholderText("Selecione um arquivo PDF...")
        pdf_layout.addWidget(self.pdf_path)
        self.btn_selecionar_pdf = QPushButton("📁 Selecionar PDF")
        self.btn_selecionar_pdf.clicked.connect(self.selecionar_pdf)
        self.btn_selecionar_pdf.setStyleSheet("background-color: #3498db; color: white;")
        pdf_layout.addWidget(self.btn_selecionar_pdf)
        layout.addLayout(pdf_layout)

        # Pasta de saída
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Pasta de saída:"))
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Selecione a pasta de destino...")
        output_layout.addWidget(self.output_path)
        self.btn_selecionar_pasta = QPushButton("📁 Selecionar Pasta")
        self.btn_selecionar_pasta.clicked.connect(self.selecionar_pasta)
        self.btn_selecionar_pasta.setStyleSheet("background-color: #3498db; color: white;")
        output_layout.addWidget(self.btn_selecionar_pasta)
        layout.addLayout(output_layout)

        # Opções avançadas
        opcoes_group = QGroupBox("Opções Avançadas")
        opcoes_layout = QVBoxLayout(opcoes_group)

        # Comprimir PDFs
        self.checkbox_comprimir = QCheckBox("Comprimir PDFs gerados (reduz tamanho)")
        self.checkbox_comprimir.setChecked(True)
        opcoes_layout.addWidget(self.checkbox_comprimir)

        # Páginas por seção (quando não há marcadores)
        paginas_layout = QHBoxLayout()
        paginas_layout.addWidget(QLabel("Páginas por seção (sem marcadores):"))
        self.spin_paginas = QSpinBox()
        self.spin_paginas.setRange(10, 200)
        self.spin_paginas.setValue(50)
        paginas_layout.addWidget(self.spin_paginas)
        paginas_layout.addStretch()
        opcoes_layout.addLayout(paginas_layout)

        layout.addWidget(opcoes_group)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress_bar)

        # Status
        self.lbl_status = QLabel("Pronto para processar")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_status)

        # Botões
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_processar = QPushButton("⚡ Processar PDF")
        self.btn_processar.clicked.connect(self.processar_pdf)
        self.btn_processar.setStyleSheet(
            "background-color: #2ecc71; color: white; padding: 10px 20px; font-size: 14px;"
        )
        self.btn_processar.setEnabled(False)
        btn_layout.addWidget(self.btn_processar)

        self.btn_limpar = QPushButton("🗑️ Limpar")
        self.btn_limpar.clicked.connect(self.limpar_campos)
        self.btn_limpar.setStyleSheet(
            "background-color: #e74c3c; color: white; padding: 10px 20px; font-size: 14px;"
        )
        btn_layout.addWidget(self.btn_limpar)

        self.btn_logout = QPushButton("🔒 Sair")
        self.btn_logout.clicked.connect(self.logout)
        self.btn_logout.setStyleSheet(
            "background-color: #95a5a6; color: white; padding: 10px 20px; font-size: 14px;"
        )
        btn_layout.addWidget(self.btn_logout)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def selecionar_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo PDF", "", "Arquivos PDF (*.pdf);;Todos os arquivos (*.*)"
        )
        if file_path:
            self.pdf_path.setText(file_path)
            # Definir pasta de saída baseada no nome do PDF
            pdf_name = Path(file_path).stem
            output_dir = os.path.join(os.path.dirname(file_path), f"{pdf_name}_secoes")
            self.output_path.setText(output_dir)
            self.btn_processar.setEnabled(True)

    def selecionar_pasta(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar pasta de saída")
        if folder_path:
            self.output_path.setText(folder_path)

    def limpar_campos(self):
        self.pdf_path.clear()
        self.output_path.clear()
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Pronto para processar")
        self.btn_processar.setEnabled(False)

    def logout(self):
        self.autenticado = False
        self.auth_widget.setVisible(True)
        self.main_widget.setVisible(False)
        self.limpar_campos()
        self.senha_input.clear()
        self.lbl_status_auth.setText("")

    def processar_pdf(self):
        if not self.pdf_path.text():
            QMessageBox.warning(self, "Aviso", "Selecione um arquivo PDF")
            return

        if not self.output_path.text():
            QMessageBox.warning(self, "Aviso", "Selecione uma pasta de saída")
            return

        # Executar processamento em thread separada
        self.worker = SeparadorPDFWorker(
            self.pdf_path.text(),
            self.output_path.text(),
            self.checkbox_comprimir.isChecked(),
            self.spin_paginas.value(),
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.status.connect(self.lbl_status.setText)
        self.worker.finished.connect(self.processamento_finalizado)
        self.worker.error.connect(self.erro_processamento)

        self.progress_bar.setVisible(True)
        self.btn_processar.setEnabled(False)
        self.worker.start()

    def processamento_finalizado(self, resultado):
        self.progress_bar.setVisible(False)
        self.btn_processar.setEnabled(True)

        if resultado["sucesso"]:
            QMessageBox.information(
                self,
                "Sucesso",
                f"PDF processado com sucesso!\n\n"
                f"Seções separadas: {resultado['total_secoes']}\n"
                f"Pasta de saída: {resultado['pasta_saida']}\n"
                f"Arquivo Excel: {resultado['excel_path']}",
            )
        else:
            QMessageBox.critical(self, "Erro", f"Erro ao processar PDF:\n{resultado['erro']}")

    def erro_processamento(self, erro):
        self.progress_bar.setVisible(False)
        self.btn_processar.setEnabled(True)
        QMessageBox.critical(self, "Erro", f"Erro durante processamento:\n{erro}")


class SeparadorPDFWorker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, pdf_path, output_path, comprimir, paginas_por_secao):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_path = output_path
        self.comprimir = comprimir
        self.paginas_por_secao = paginas_por_secao

    def run(self):
        try:
            self.status.emit("Abrindo PDF...")
            self.progress.emit(0)

            # Criar diretório de saída
            output_dir = Path(self.output_path)
            output_dir.mkdir(parents=True, exist_ok=True)

            # Abrir PDF
            with open(self.pdf_path, "rb") as file:
                pdf_reader = PdfReader(file)
                total_pages = len(pdf_reader.pages)

                self.status.emit("Extraindo marcadores...")

                # Extrair marcadores
                marcadores = self.extrair_marcadores(pdf_reader)

                if not marcadores:
                    self.status.emit("Nenhum marcador encontrado. Criando seções por páginas...")
                    marcadores = self.criar_marcadores_por_pagina(total_pages)

                total_secoes = len(marcadores)
                self.status.emit(f"Processando {total_secoes} seções...")

                # Lista para informações dos arquivos
                arquivos_info = []

                for i, marcador in enumerate(marcadores):
                    progress = int((i / total_secoes) * 100)
                    self.progress.emit(progress)
                    self.status.emit(
                        f"Processando seção {i + 1} de {total_secoes}: {marcador['titulo'][:50]}..."
                    )

                    # Nome do arquivo
                    nome_arquivo = self.sanitizar_nome_arquivo(f"{i+1:03d}_{marcador['titulo']}.pdf")
                    caminho_arquivo = output_dir / nome_arquivo

                    # Criar PDF otimizado
                    self.criar_pdf_otimizado(pdf_reader, marcador, caminho_arquivo)

                    # Adicionar à lista
                    arquivos_info.append(
                        {
                            "numero": i + 1,
                            "titulo": marcador["titulo"],
                            "pagina_inicio": marcador["pagina_inicio"],
                            "pagina_fim": marcador["pagina_fim"],
                            "total_paginas": marcador["pagina_fim"] - marcador["pagina_inicio"] + 1,
                            "arquivo": nome_arquivo,
                            "caminho": str(caminho_arquivo),
                        }
                    )

                # Gerar Excel
                self.status.emit("Gerando tabela Excel...")
                excel_path = self.gerar_excel(arquivos_info, output_dir)

                self.progress.emit(100)
                self.status.emit(f"Concluído! {total_secoes} seções processadas")

                self.finished.emit(
                    {
                        "sucesso": True,
                        "total_secoes": total_secoes,
                        "pasta_saida": str(output_dir),
                        "excel_path": str(excel_path),
                    }
                )

        except Exception as e:
            self.finished.emit({"sucesso": False, "erro": str(e)})

    def extrair_marcadores(self, pdf_reader):
        marcadores = []
        try:
            if pdf_reader.outline:
                self.processar_outline(pdf_reader.outline, marcadores, 0, pdf_reader)
        except Exception as e:
            print(f"Erro ao extrair marcadores: {e}")
        return marcadores

    def processar_outline(self, outline, marcadores, nivel, pdf_reader):
        for item in outline:
            if isinstance(item, list):
                self.processar_outline(item, marcadores, nivel + 1, pdf_reader)
            else:
                try:
                    if hasattr(item, "page") and item.page is not None:
                        pagina = pdf_reader.get_destination_page_number(item) + 1
                        marcador = {
                            "titulo": item.title.strip(),
                            "pagina_inicio": pagina,
                            "pagina_fim": pagina,
                            "nivel": nivel,
                        }
                        marcadores.append(marcador)
                except Exception as e:
                    print(f"Erro ao processar item do outline: {e}")

        # Calcular página final
        for i in range(len(marcadores)):
            if i < len(marcadores) - 1:
                marcadores[i]["pagina_fim"] = marcadores[i + 1]["pagina_inicio"] - 1
            else:
                marcadores[i]["pagina_fim"] = len(pdf_reader.pages)

    def criar_marcadores_por_pagina(self, total_pages):
        marcadores = []
        for i in range(0, total_pages, self.paginas_por_secao):
            pagina_inicio = i + 1
            pagina_fim = min(i + self.paginas_por_secao, total_pages)
            marcador = {
                "titulo": f"Seção {len(marcadores) + 1} (Páginas {pagina_inicio}-{pagina_fim})",
                "pagina_inicio": pagina_inicio,
                "pagina_fim": pagina_fim,
                "nivel": 0,
            }
            marcadores.append(marcador)
        return marcadores

    def sanitizar_nome_arquivo(self, nome):
        import re

        caracteres_invalidos = r'[<>:"/\\|?*#%&+=\s]'
        nome_limpo = re.sub(caracteres_invalidos, "_", nome)
        nome_limpo = re.sub(r"_+", "_", nome_limpo)
        nome_limpo = nome_limpo.strip("_")
        if len(nome_limpo) > 200:
            nome_limpo = nome_limpo[:200]
        return nome_limpo

    def criar_pdf_otimizado(self, pdf_reader, marcador, caminho_arquivo):
        try:
            pdf_writer = PdfWriter()

            for page_num in range(marcador["pagina_inicio"] - 1, marcador["pagina_fim"]):
                if page_num < len(pdf_reader.pages):
                    page = pdf_reader.pages[page_num]

                    if self.comprimir:
                        try:
                            page.compress_content_streams()
                        except Exception:
                            pass

                    pdf_writer.add_page(page)

            with open(caminho_arquivo, "wb") as output_file:
                pdf_writer.write(output_file)

        except Exception as e:
            print(f"Erro ao criar PDF otimizado: {e}")
            # Fallback
            pdf_writer = PdfWriter()
            for page_num in range(marcador["pagina_inicio"] - 1, marcador["pagina_fim"]):
                if page_num < len(pdf_reader.pages):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

            with open(caminho_arquivo, "wb") as output_file:
                pdf_writer.write(output_file)

    def gerar_excel(self, arquivos_info, output_dir):
        excel_path = output_dir / "indice_secoes.xlsx"

        try:
            workbook = xlsxwriter.Workbook(
                str(excel_path),
                {"strings_to_urls": False, "constant_memory": True, "tmpdir": str(output_dir)},
            )
            worksheet = workbook.add_worksheet("Índice de Seções")

            # Formatações
            titulo_format = workbook.add_format(
                {"font_color": "red", "bold": True, "underline": 1, "font_size": 16, "align": "center"}
            )

            header_format = workbook.add_format({"bold": True, "bg_color": "#D3D3D3", "border": 1})

            link_format = workbook.add_format({"font_color": "blue", "underline": 1})

            # Configurar colunas
            worksheet.set_column("A:A", 8)
            worksheet.set_column("B:B", 15)
            worksheet.set_column("C:C", 10)
            worksheet.set_column("D:D", 30)

            # Título
            worksheet.merge_range("A1:D1", "ÍNDICE DE SEÇÕES", titulo_format)

            # Cabeçalhos
            worksheet.write("A3", "Seção", header_format)
            worksheet.write("B3", "Páginas", header_format)
            worksheet.write("C3", "Total Págs", header_format)
            worksheet.write("D3", "Arquivo", header_format)

            # Dados
            row = 4
            for info in arquivos_info:
                worksheet.write(f"A{row}", info["numero"])
                worksheet.write(f"B{row}", f"{info['pagina_inicio']}-{info['pagina_fim']}")
                worksheet.write(f"C{row}", info["total_paginas"])

                # Criar link
                caminho_absoluto = os.path.abspath(info["caminho"])
                try:
                    caminho_para_link = caminho_absoluto.replace(os.sep, "/")
                    worksheet.write_url(
                        f"D{row}",
                        f"file:///{caminho_para_link}",
                        string=info["arquivo"],
                        cell_format=link_format,
                    )
                except Exception:
                    worksheet.write(f"D{row}", info["arquivo"], link_format)
                    worksheet.write_comment(f"D{row}", f"Caminho: {caminho_absoluto}")

                row += 1

            worksheet.freeze_panes(3, 0)
            workbook.close()

        except Exception as e:
            print(f"Erro ao criar Excel: {e}")
            # Criar versão simples
            self.gerar_excel_simples(arquivos_info, output_dir)

        return excel_path

    def gerar_excel_simples(self, arquivos_info, output_dir):
        excel_path = output_dir / "indice_secoes_simples.xlsx"

        try:
            workbook = xlsxwriter.Workbook(str(excel_path))
            worksheet = workbook.add_worksheet("Índice de Seções")

            # Formatações
            titulo_format = workbook.add_format(
                {"font_color": "red", "bold": True, "font_size": 16, "align": "center"}
            )

            header_format = workbook.add_format({"bold": True, "bg_color": "#D3D3D3"})

            # Configurar colunas
            worksheet.set_column("A:A", 8)
            worksheet.set_column("B:B", 15)
            worksheet.set_column("C:C", 10)
            worksheet.set_column("D:D", 30)
            worksheet.set_column("E:E", 50)

            # Título
            worksheet.merge_range("A1:E1", "ÍNDICE DE SEÇÕES (SEM LINKS)", titulo_format)

            # Cabeçalhos
            worksheet.write("A3", "Seção", header_format)
            worksheet.write("B3", "Páginas", header_format)
            worksheet.write("C3", "Total Págs", header_format)
            worksheet.write("D3", "Arquivo", header_format)
            worksheet.write("E3", "Caminho Completo", header_format)

            # Dados
            row = 4
            for info in arquivos_info:
                worksheet.write(f"A{row}", info["numero"])
                worksheet.write(f"B{row}", f"{info['pagina_inicio']}-{info['pagina_fim']}")
                worksheet.write(f"C{row}", info["total_paginas"])
                worksheet.write(f"D{row}", info["arquivo"])
                worksheet.write(f"E{row}", info["caminho"])
                row += 1

            workbook.close()
            print(f"Excel simples criado: {excel_path}")

        except Exception as e:
            print(f"Erro ao criar Excel simples: {e}")

