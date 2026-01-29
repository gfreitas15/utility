from __future__ import annotations

import os
import re
import unicodedata

import pandas as pd
from rapidfuzz import fuzz

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAbstractScrollArea,
    QComboBox,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class ComparadorPlanilhasWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.df1 = None
        self.df2 = None
        self.nome_arquivo1 = ""
        self.nome_arquivo2 = ""

        # Define caminho padrão de saída (Desktop)
        import os

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.caminho_saida_padrao = os.path.join(desktop_path, "planilha_comparacao.xlsx")

        self.init_ui()

        # Aplicar tema escuro por padrão na inicialização
        self.aplicar_tema(True)

    def init_ui(self):
        fonte_label = QFont("Segoe UI", 10)
        fonte_botao = QFont("Segoe UI", 10)

        layout = QVBoxLayout()
        layout.setSpacing(10)  # Espaçamento menor para otimizar espaço

        # --- Primeira planilha ---
        # Label "Planilha 1"
        lbl_planilha1 = QLabel("Planilha 1")
        lbl_planilha1.setFont(fonte_label)
        layout.addWidget(lbl_planilha1)

        layout1 = QHBoxLayout()
        self.btn_arquivo1 = QPushButton("📂 Selecionar Planilha")
        self.btn_arquivo1.setFont(fonte_botao)
        self.btn_arquivo1.clicked.connect(lambda: self.selecionar_planilha(1))
        # Lista de colunas (multi-seleção) para a primeira planilha (sempre visível)
        self.lst_colunas1 = QListWidget()
        self.lst_colunas1.setMinimumWidth(220)
        layout1.addWidget(self.btn_arquivo1)
        layout1.addWidget(self.lst_colunas1)
        layout.addLayout(layout1)
        self.lbl_arquivo1 = QLabel("Nenhum arquivo selecionado")
        self.lbl_arquivo1.setFont(fonte_label)
        layout.addWidget(self.lbl_arquivo1)
        self.tabela_preview1 = QTableWidget()
        self.tabela_preview1.setMinimumHeight(120)
        layout.addWidget(self.tabela_preview1)

        # --- Segunda planilha ---
        # Label "Planilha 2"
        lbl_planilha2 = QLabel("Planilha 2")
        lbl_planilha2.setFont(fonte_label)
        layout.addWidget(lbl_planilha2)

        layout2 = QHBoxLayout()
        self.btn_arquivo2 = QPushButton("📂 Selecionar Planilha")
        self.btn_arquivo2.setFont(fonte_botao)
        self.btn_arquivo2.clicked.connect(lambda: self.selecionar_planilha(2))
        # Lista de colunas (multi-seleção) para a segunda planilha (sempre visível)
        self.lst_colunas2 = QListWidget()
        self.lst_colunas2.setMinimumWidth(220)
        layout2.addWidget(self.btn_arquivo2)
        layout2.addWidget(self.lst_colunas2)
        layout.addLayout(layout2)
        self.lbl_arquivo2 = QLabel("Nenhum arquivo selecionado")
        self.lbl_arquivo2.setFont(fonte_label)
        layout.addWidget(self.lbl_arquivo2)
        self.tabela_preview2 = QTableWidget()
        self.tabela_preview2.setMinimumHeight(120)
        layout.addWidget(self.tabela_preview2)

        # --- Similaridade e Normalização ---
        sim_layout = QHBoxLayout()
        sim_layout.addWidget(QLabel("Similaridade (0-100%):"))
        self.spin_similaridade = QSpinBox()
        self.spin_similaridade.setRange(0, 100)
        self.spin_similaridade.setValue(90)
        sim_layout.addWidget(self.spin_similaridade)

        # Regras de normalização
        self.cmb_normalizacao = QComboBox()
        self.cmb_normalizacao.addItems(
            [
                "Padrão (acentos+maiusc+espaços)",
                "Ignorar pontuação",
                "Remover stopwords (LTDA, ME, SA)",
                "Sem normalização",
            ]
        )
        sim_layout.addWidget(QLabel("Normalização:"))
        sim_layout.addWidget(self.cmb_normalizacao)
        layout.addLayout(sim_layout)

        # --- Botões de ação ---
        btn_layout = QHBoxLayout()
        self.btn_limpar = QPushButton("🗑️ Limpar tudo")
        self.btn_limpar.clicked.connect(self.limpar_campos)
        self.btn_limpar.setMaximumWidth(120)  # Botão menor
        self.btn_saida = QPushButton("📁 Selecionar Saída")
        self.btn_saida.clicked.connect(self.selecionar_saida)

        # Adiciona espaçamento entre os botões
        btn_layout.addWidget(self.btn_limpar)
        btn_layout.addStretch(1)  # Espaçamento médio
        btn_layout.addWidget(self.btn_saida)
        layout.addLayout(btn_layout)

        # --- Campo de saída ---
        out_layout = QHBoxLayout()
        self.txt_saida = QLineEdit()
        self.txt_saida.setPlaceholderText("Selecione local/arquivo de saída")
        self.txt_saida.setText(f"📁 {self.caminho_saida_padrao}")  # Define valor padrão
        out_layout.addWidget(self.txt_saida)
        layout.addLayout(out_layout)

        # --- Botão Comparar e Barra de Progresso ---
        compare_layout = QHBoxLayout()
        self.btn_comparar = QPushButton("✅ Comparar")
        self.btn_comparar.clicked.connect(self.comparar)
        self.btn_comparar.setMaximumWidth(120)  # Botão pequeno, um pouco maior que o texto
        self.btn_cancelar = QPushButton("❌ Cancelar")
        self.btn_cancelar.setMaximumWidth(120)
        self.btn_cancelar.setEnabled(False)
        self.btn_cancelar.clicked.connect(self.cancelar_comparacao)
        self.progress = QProgressBar()
        self.progress.setStyleSheet(
            """
            QProgressBar {
                border: 2px solid #7f8c8d;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 3px;
                margin: 1px;
            }
        """
        )
        compare_layout.addWidget(self.btn_comparar)
        compare_layout.addWidget(self.progress)
        compare_layout.addWidget(self.btn_cancelar)
        layout.addLayout(compare_layout)

        self.setLayout(layout)

        self.setAcceptDrops(True)

    # Drag & Drop
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            # aceita apenas arquivos .xlsx ou .xls
            for url in event.mimeData().urls():
                if str(url.toLocalFile()).lower().endswith((".xlsx", ".xls")):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            return
        arquivos = [str(url.toLocalFile()) for url in event.mimeData().urls()]
        excel_files = [p for p in arquivos if p.lower().endswith((".xlsx", ".xls"))]
        if not excel_files:
            return
        # Se vier 1 arquivo, coloca na próxima planilha vazia; se vierem 2, preenche ambas
        try:
            if self.df1 is None and len(excel_files) >= 1:
                self._carregar_arquivo_em_planilha(1, excel_files[0])
            elif self.df2 is None and len(excel_files) >= 1:
                self._carregar_arquivo_em_planilha(2, excel_files[0])
            elif len(excel_files) >= 2:
                # Se ambas estão vazias, preenche ambas
                if self.df1 is None and self.df2 is None:
                    self._carregar_arquivo_em_planilha(1, excel_files[0])
                    self._carregar_arquivo_em_planilha(2, excel_files[1])
                # Se apenas uma está vazia, preenche a vazia
                elif self.df1 is None:
                    self._carregar_arquivo_em_planilha(1, excel_files[0])
                elif self.df2 is None:
                    self._carregar_arquivo_em_planilha(2, excel_files[0])
                # Se ambas estão preenchidas, substitui a segunda
                else:
                    self._carregar_arquivo_em_planilha(2, excel_files[0])
        except Exception:
            self._mostrar_erro("Falha no arrastar e soltar", "Não foi possível carregar os arquivos arrastados.")

    def _carregar_arquivo_em_planilha(self, num, path):
        try:
            df = pd.read_excel(path)
        except Exception:
            raise
        if df is None or len(df) == 0:
            QMessageBox.warning(self, "Aviso", "A planilha arrastada está vazia.")
        preview = df.head(25)
        nome_arquivo = path.split("/")[-1].split("\\")[-1]
        if num == 1:
            self.df1 = df
            self.nome_arquivo1 = nome_arquivo
            self.lbl_arquivo1.setText(f"📁 {path}")
            self.lst_colunas1.clear()
            for col in df.columns:
                item = QListWidgetItem(str(col))
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.lst_colunas1.addItem(item)
            self.mostrar_preview(self.tabela_preview1, preview)
        else:
            self.df2 = df
            self.nome_arquivo2 = nome_arquivo
            self.lbl_arquivo2.setText(f"📁 {path}")
            self.lst_colunas2.clear()
            for col in df.columns:
                item = QListWidgetItem(str(col))
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                self.lst_colunas2.addItem(item)
            self.mostrar_preview(self.tabela_preview2, preview)

        # (ajuda já adicionada no layout principal)

    def _obter_colunas_selecionadas(self, lista_widget):
        colunas = []
        for i in range(lista_widget.count()):
            item = lista_widget.item(i)
            if item.checkState() == Qt.Checked:
                colunas.append(item.text())
        return colunas

    def _mostrar_erro(self, titulo, mensagem):
        QMessageBox.critical(self, titulo, mensagem)

    def remover_sucessao(self, texto):
        """Remove ocorrências de '( SUCESSÃO DE )' ou '( SUCESSAO DE )' do texto original (case-insensitive)."""
        if texto is None:
            return texto
        s = str(texto)
        # Remoções com e sem acento, tolerando espaços extras
        s = re.sub(r"\(\s*SUCESSÃO\s+DE\s*\)", " ", s, flags=re.IGNORECASE)
        s = re.sub(r"\(\s*SUCESSAO\s+DE\s*\)", " ", s, flags=re.IGNORECASE)
        # Normaliza espaços
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _normalize_cpf(self, valor):
        if valor is None or pd.isna(valor):
            return ""
        return re.sub(r"\D", "", str(valor))

    def _get_cpf_column(self, df):
        for col in df.columns:
            header = str(col)
            header_norm = unicodedata.normalize("NFD", header)
            header_norm = "".join(c for c in header_norm if unicodedata.category(c) != "Mn")
            header_norm = re.sub(r"[^A-Z0-9]", "", header_norm.upper())
            if header_norm == "CPF":
                return col
        return None

    def _build_cpf_set_df1(self, cpf_col_df1):
        cpfs = set()
        if cpf_col_df1 is None or cpf_col_df1 not in self.df1.columns:
            return cpfs
        for _, row in self.df1.iterrows():
            cpf_val = self._normalize_cpf(row[cpf_col_df1]) if cpf_col_df1 in self.df1.columns else ""
            if cpf_val:
                cpfs.add(cpf_val)
        return cpfs

    def _preparar_compostos_df1(self, colunas1):
        df1_compostos_norm = []
        df1_compostos_exibicao = []
        for _, row in self.df1.iterrows():
            partes_exibicao = [str(row[col]) if col in self.df1.columns else "" for col in colunas1]
            composto_exibicao = " | ".join(partes_exibicao)
            composto_exibicao = self.remover_sucessao(composto_exibicao)
            composto_normalizado = self.normalizar_texto(composto_exibicao)
            df1_compostos_exibicao.append(composto_exibicao)
            df1_compostos_norm.append(composto_normalizado)
        return df1_compostos_exibicao, df1_compostos_norm

    def _calcular_resultado_linha(
        self,
        row,
        colunas2,
        similaridade_min,
        df1_compostos_exibicao,
        df1_compostos_norm,
        cpf_col_df1,
        cpf_col_df2,
        df1_cpfs_set,
    ):
        partes2_exibicao = [str(row[col]) if col in self.df2.columns else "" for col in colunas2]
        valor = " | ".join(partes2_exibicao)
        valor = self.remover_sucessao(valor)
        valor_normalizado = self.normalizar_texto(valor)
        valor_exibicao = valor_normalizado
        encontrado_exato = False
        nomes_similares = []

        # 1) Se ambas possuem coluna CPF, prioriza match exato por CPF
        if cpf_col_df1 and cpf_col_df2 and cpf_col_df2 in self.df2.columns:
            cpf_val = self._normalize_cpf(row[cpf_col_df2]) if not pd.isna(row[cpf_col_df2]) else ""
            if cpf_val and cpf_val in df1_cpfs_set:
                encontrado_exato = True
                return valor_exibicao, encontrado_exato, ""

        for nome_sistema_normalizado in df1_compostos_norm:
            if nome_sistema_normalizado == valor_normalizado:
                encontrado_exato = True
                break

        if not encontrado_exato:
            for composto_exibicao, composto_norm in zip(df1_compostos_exibicao, df1_compostos_norm):
                score = fuzz.token_sort_ratio(composto_norm, valor_normalizado)
                if score >= similaridade_min:
                    score_formatado = f"{score:.1f}".replace(".", ",")
                    nomes_similares.append(f"{self.normalizar_texto(composto_exibicao)} ({score_formatado}%)")

        return valor_exibicao, encontrado_exato, ", ".join(nomes_similares) if nomes_similares else ""

    def _mostrar_preview_dialog(self, df_preview, titulo="Pré-visualização (até 20 linhas)"):
        from PyQt5.QtWidgets import QDialog, QDialogButtonBox

        dialog = QDialog(self)
        dialog.setWindowTitle(titulo)
        dialog.setMinimumWidth(900)
        dialog.resize(1000, dialog.height())
        v = QVBoxLayout(dialog)
        lbl = QLabel("Confira uma amostra dos resultados antes de rodar tudo:")
        v.addWidget(lbl)
        tabela = QTableWidget()
        tabela.setRowCount(len(df_preview))
        tabela.setColumnCount(len(df_preview.columns))
        tabela.setHorizontalHeaderLabels(df_preview.columns)
        for i in range(len(df_preview)):
            for j, col in enumerate(df_preview.columns):
                tabela.setItem(i, j, QTableWidgetItem(str(df_preview.iloc[i, j])))
        # Melhorias de visualização para textos longos
        tabela.setWordWrap(True)
        tabela.setTextElideMode(Qt.ElideNone)
        tabela.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        tabela.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        tabela.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        header = tabela.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)
        tabela.resizeColumnsToContents()
        tabela.resizeRowsToContents()
        # Calcula altura exata da tabela (até 20 linhas) para evitar sobra
        header_h = tabela.horizontalHeader().height()
        rows_h = sum(tabela.rowHeight(r) for r in range(tabela.rowCount()))
        hscroll_h = tabela.horizontalScrollBar().sizeHint().height() if tabela.horizontalScrollBar().isVisible() else 0
        table_h = header_h + rows_h + (tabela.frameWidth() * 2) + hscroll_h
        tabela.setFixedHeight(table_h)
        v.addWidget(tabela)
        botoes = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        v.addWidget(botoes)
        botoes.accepted.connect(dialog.accept)
        botoes.rejected.connect(dialog.reject)
        # Ajusta a altura do diálogo para encaixar exatamente o conteúdo
        margins = v.contentsMargins().top() + v.contentsMargins().bottom()
        spacing = v.spacing() * 2  # label->table e table->buttons
        dialog_h = lbl.sizeHint().height() + table_h + botoes.sizeHint().height() + margins + spacing
        dialog.setMinimumHeight(dialog_h)
        dialog.setMaximumHeight(dialog_h)
        dialog.resize(dialog.width(), dialog_h)
        return dialog.exec() == 1

    def aplicar_tema(self, tema_escuro):
        if tema_escuro:
            self.setStyleSheet(
                """
                QWidget {background-color: #2c3e50; color: white;}
                QTableWidget {background-color: #34495e; color: white;}
                QTableWidget::item {background-color: #34495e; color: white;}
                QHeaderView::section {background-color: #2c3e50; color: white; border: 1px solid #7f8c8d; padding: 4px;}
                QLineEdit {background-color: #34495e; color: white;}
                QPushButton {border-radius: 8px; padding: 8px;}
                QProgressBar {background-color: #34495e; color: white; border-radius: 10px;}
                QLabel {color: white;}
                QListWidget {background-color: #34495e; color: white; border: 1px solid #7f8c8d;}
                QComboBox {background-color: #34495e; color: white; border: 1px solid #7f8c8d; padding: 4px;}
                QSpinBox {background-color: #34495e; color: white; border: 1px solid #7f8c8d; padding: 4px;}
            """
            )
        else:
            self.setStyleSheet(
                """
                QWidget {background-color: #ecf0f1; color: black;}
                QTableWidget {background-color: #bdc3c7; color: black;}
                QTableWidget::item {background-color: #bdc3c7; color: black;}
                QHeaderView::section {background-color: #95a5a6; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QLineEdit {background-color: #bdc3c7; color: black;}
                QPushButton {border-radius: 8px; padding: 8px;}
                QProgressBar {background-color: #bdc3c7; color: black; border-radius: 10px;}
                QLabel {color: black;}
                QListWidget {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d;}
                QComboBox {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QSpinBox {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
            """
            )

        # Cores fixas dos botões (não mudam com tema)
        self.btn_arquivo1.setStyleSheet("background-color: #3498db; color: white;")  # Azul - Selecionar
        self.btn_arquivo2.setStyleSheet("background-color: #3498db; color: white;")  # Azul - Selecionar
        self.btn_comparar.setStyleSheet("background-color: #2ecc71; color: white;")  # Verde - Comparar
        self.btn_limpar.setStyleSheet("background-color: #e74c3c; color: white;")  # Vermelho - Limpar
        self.btn_saida.setStyleSheet("background-color: #3498db; color: white;")  # Azul - Selecionar
        self.btn_cancelar.setStyleSheet("background-color: #e74c3c; color: white;")  # Vermelho - Cancelar

    def normalizar_texto(self, texto):
        """
        Normaliza texto removendo acentos, espaços extras e convertendo para maiúsculas,
        com variações baseadas na seleção do usuário.
        """
        if pd.isna(texto) or texto is None:
            return ""

        # Converte para string e remove espaços no início e fim
        texto = str(texto).strip()

        # Remove espaços duplos, triplos, etc. e substitui por espaço simples
        texto = re.sub(r"\s+", " ", texto)

        modo = (
            self.cmb_normalizacao.currentText()
            if hasattr(self, "cmb_normalizacao")
            else "Padrão (acentos+maiusc+espaços)"
        )

        # Base: remove acentos
        texto_base = unicodedata.normalize("NFD", texto)
        texto_base = "".join(c for c in texto_base if unicodedata.category(c) != "Mn")

        if modo == "Sem normalização":
            return texto

        if modo == "Ignorar pontuação":
            texto_base = re.sub(r"[\p{P}\p{S}]", " ", texto_base)
        else:
            # Remove apenas alguns sinais comuns
            texto_base = re.sub(r"[,;:.!?'\-]", " ", texto_base)

        if modo == "Remover stopwords (LTDA, ME, SA)":
            stop = {"LTDA", "ME", "S/A", "SA", "EIRELI", "EPP"}
            palavras = [p for p in texto_base.split() if p.upper() not in stop]
            texto_base = " ".join(palavras)

        # Regra específica: remover ocorrências de "( SUCESSÃO DE )"
        # Como já removemos acentos acima, procuramos por "SUCESSAO".
        # Aceita variações de espaço e caixa.
        texto_base = re.sub(r"\(\s*SUCESSAO\s+DE\s*\)", " ", texto_base, flags=re.IGNORECASE)
        # Normaliza espaços novamente após remoção
        texto_base = re.sub(r"\s+", " ", texto_base).strip()

        return texto_base.upper()

    def limpar_campos(self):
        """Limpa todos os campos da interface"""
        self.df1 = None
        self.df2 = None
        self.nome_arquivo1 = ""
        self.nome_arquivo2 = ""
        # Corrige: limpar listas de colunas reais
        if hasattr(self, "lst_colunas1"):
            self.lst_colunas1.clear()
        if hasattr(self, "lst_colunas2"):
            self.lst_colunas2.clear()
        self.lbl_arquivo1.setText("Nenhum arquivo selecionado")
        self.lbl_arquivo2.setText("Nenhum arquivo selecionado")
        self.txt_saida.setText(f"📁 {self.caminho_saida_padrao}")  # Restaura valor padrão

        # Limpa completamente as tabelas de preview
        self.tabela_preview1.clear()
        self.tabela_preview1.setRowCount(0)
        self.tabela_preview1.setColumnCount(0)
        self.tabela_preview2.clear()
        self.tabela_preview2.setRowCount(0)
        self.tabela_preview2.setColumnCount(0)

        self.progress.setValue(0)
        self.spin_similaridade.setValue(90)

    # --- Seleção de arquivos ---
    def selecionar_planilha(self, num):
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Excel Files (*.xlsx *.xls)")
        if path:
            try:
                df = pd.read_excel(path)
            except Exception as e:
                self._mostrar_erro(
                    "Falha ao abrir arquivo",
                    "Não foi possível ler a planilha.\nVerifique se o arquivo está corrompido ou protegido por senha.",
                )
                return

            if df is None or len(df) == 0:
                QMessageBox.warning(self, "Aviso", "A planilha selecionada está vazia.")
            preview = df.head(25)

            nome_arquivo = path.split("/")[-1].split("\\")[-1]

            if num == 1:
                self.df1 = df
                self.nome_arquivo1 = nome_arquivo
                self.lbl_arquivo1.setText(f"📁 {path}")
                self.lst_colunas1.clear()
                for col in df.columns:
                    item = QListWidgetItem(str(col))
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    self.lst_colunas1.addItem(item)
                self.mostrar_preview(self.tabela_preview1, preview)
            else:
                self.df2 = df
                self.nome_arquivo2 = nome_arquivo
                self.lbl_arquivo2.setText(f"📁 {path}")
                self.lst_colunas2.clear()
                for col in df.columns:
                    item = QListWidgetItem(str(col))
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    self.lst_colunas2.addItem(item)
                self.mostrar_preview(self.tabela_preview2, preview)

    def mostrar_preview(self, tabela, df_preview):
        tabela.clear()
        tabela.setRowCount(len(df_preview))
        tabela.setColumnCount(len(df_preview.columns))
        tabela.setHorizontalHeaderLabels(df_preview.columns)
        for i in range(len(df_preview)):
            for j, col in enumerate(df_preview.columns):
                tabela.setItem(i, j, QTableWidgetItem(str(df_preview.iloc[i, j])))
        tabela.resizeColumnsToContents()

    # --- Saída ---
    def selecionar_saida(self):
        path = QFileDialog.getSaveFileName(self, "Salvar Planilha", "", "Excel Files (*.xlsx)")[0]
        if path:
            if not path.lower().endswith(".xlsx"):
                path = path + ".xlsx"
            self.txt_saida.setText(f"📁 {path}")

    # --- Comparação ---
    def comparar(self):
        if self.df1 is None or self.df2 is None:
            QMessageBox.warning(self, "Erro", "Selecione as duas planilhas primeiro!")
            return
        if not self.txt_saida.text():
            QMessageBox.warning(self, "Erro", "Selecione o local de saída!")
            return
        if len(self.df1) == 0 or len(self.df2) == 0:
            QMessageBox.warning(self, "Erro", "Uma das planilhas está vazia. Importe arquivos com dados.")
            return

        # Obtém colunas selecionadas em cada planilha
        colunas1 = self._obter_colunas_selecionadas(self.lst_colunas1)
        colunas2 = self._obter_colunas_selecionadas(self.lst_colunas2)

        if not colunas1 or not colunas2:
            QMessageBox.warning(self, "Erro", "Selecione ao menos uma coluna em cada planilha!")
            return
        similaridade_min = self.spin_similaridade.value()

        resultados = []
        total = len(self.df2)
        self.progress.setValue(0)

        # Pré-calcula compostos normalizados da planilha 1 para acelerar buscas
        df1_compostos_exibicao, df1_compostos_norm = self._preparar_compostos_df1(colunas1)
        cpf_col_df1 = self._get_cpf_column(self.df1)
        cpf_col_df2 = self._get_cpf_column(self.df2)
        df1_cpfs_set = self._build_cpf_set_df1(cpf_col_df1)

        # Pré-visualização (até 20 linhas)
        nome_planilha2 = self.nome_arquivo2 if self.nome_arquivo2 else "PLANILHA 2"
        nome_planilha1 = self.nome_arquivo1 if self.nome_arquivo1 else "PLANILHA 1"
        nome_coluna_planilha2 = f"{' + '.join(colunas2)} NA PLANILHA {nome_planilha2}"
        nome_coluna_esta_na_planilha1 = f"ESTÁ NA PLANILHA {nome_planilha1}"
        nome_coluna_similares = f"{' + '.join(colunas1)} SIMILARES NA PLANILHA {nome_planilha1}"

        amostra = self.df2.head(20)
        prev_regs = []
        for _, r in amostra.iterrows():
            valor, ok, similares = self._calcular_resultado_linha(
                r,
                colunas2,
                similaridade_min,
                df1_compostos_exibicao,
                df1_compostos_norm,
                cpf_col_df1,
                cpf_col_df2,
                df1_cpfs_set,
            )
            prev_regs.append(
                {
                    nome_coluna_planilha2: valor,
                    nome_coluna_esta_na_planilha1: "Sim" if ok else "Não",
                    nome_coluna_similares: similares,
                }
            )
        df_preview = pd.DataFrame(prev_regs)
        if not self._mostrar_preview_dialog(df_preview):
            return

        # Processamento completo em thread
        self._worker = CompararWorker(
            df1_compostos_exibicao,
            df1_compostos_norm,
            self.df2.copy(),
            colunas2,
            similaridade_min,
            nome_coluna_planilha2,
            nome_coluna_esta_na_planilha1,
            nome_coluna_similares,
            self.normalizar_texto,
            cpf_col_df1,
            cpf_col_df2,
            df1_cpfs_set,
        )
        self._worker.progress.connect(self.progress.setValue)
        self._worker.finished.connect(self._comparacao_finalizada)
        self._worker.error.connect(lambda msg: self._mostrar_erro("Erro durante a comparação", msg))

        # UI state
        self.btn_comparar.setEnabled(False)
        self.btn_cancelar.setEnabled(True)
        self.progress.setValue(0)
        self._worker.start()

    def cancelar_comparacao(self):
        if hasattr(self, "_worker") and self._worker is not None:
            self._worker.cancel()

    def _comparacao_finalizada(self, payload):
        # payload: dict com 'cancelado' bool e 'resultados' list
        self.btn_comparar.setEnabled(True)
        self.btn_cancelar.setEnabled(False)
        if payload.get("cancelado"):
            QMessageBox.information(self, "Cancelado", "A comparação foi cancelada pelo usuário.")
            return

        resultados = payload.get("resultados", [])
        df_result = pd.DataFrame(resultados)
        try:
            df_result.to_excel(self.txt_saida.text(), index=False)
        except PermissionError:
            self._mostrar_erro(
                "Não foi possível salvar",
                "O arquivo de saída está aberto em outro programa. Feche-o e tente novamente.",
            )
            return
        except Exception:
            self._mostrar_erro(
                "Erro ao salvar",
                "Não foi possível salvar o arquivo. Verifique permissões e espaço em disco.",
            )
            return
        QMessageBox.information(self, "Concluído", "Comparação finalizada e arquivo salvo!")
        self.limpar_campos()


class CompararWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(
        self,
        df1_compostos_exibicao,
        df1_compostos_norm,
        df2,
        colunas2,
        similaridade_min,
        nome_coluna_planilha2,
        nome_coluna_esta_na_planilha1,
        nome_coluna_similares,
        normalize_func,
        cpf_col_df1,
        cpf_col_df2,
        df1_cpfs_set,
    ):
        super().__init__()
        self.df1_compostos_exibicao = df1_compostos_exibicao
        self.df1_compostos_norm = df1_compostos_norm
        self.df2 = df2
        self.colunas2 = colunas2
        self.similaridade_min = similaridade_min
        self.nome_coluna_planilha2 = nome_coluna_planilha2
        self.nome_coluna_esta_na_planilha1 = nome_coluna_esta_na_planilha1
        self.nome_coluna_similares = nome_coluna_similares
        self.normalize_func = normalize_func
        self.cpf_col_df1 = cpf_col_df1
        self.cpf_col_df2 = cpf_col_df2
        self.df1_cpfs_set = df1_cpfs_set
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            resultados = []
            total = len(self.df2)
            for i, row in self.df2.iterrows():
                if self._cancel:
                    self.finished.emit({"cancelado": True})
                    return
                partes2_exibicao = [str(row[col]) if col in self.df2.columns else "" for col in self.colunas2]
                valor = " | ".join(partes2_exibicao)
                # Remover '( SUCESSÃO DE )' e normalizar para exibição/saída
                valor = re.sub(r"\(\s*SUCESS[ÃA]O\s+DE\s*\)", " ", valor, flags=re.IGNORECASE)
                valor = re.sub(r"\s+", " ", valor).strip()
                valor_normalizado = self.normalize_func(valor)
                valor_exibicao = valor_normalizado

                encontrado_exato = False
                nomes_similares = []

                # 1) Match por CPF (se ambas possuem CPF)
                if self.cpf_col_df1 and self.cpf_col_df2 and self.cpf_col_df2 in self.df2.columns:
                    cpf_val = (
                        str(row[self.cpf_col_df2])
                        if self.cpf_col_df2 in self.df2.columns and not pd.isna(row[self.cpf_col_df2])
                        else ""
                    )
                    cpf_val = re.sub(r"\D", "", cpf_val)
                    if cpf_val and cpf_val in self.df1_cpfs_set:
                        encontrado_exato = True
                        resultados.append(
                            {
                                self.nome_coluna_planilha2: valor_exibicao,
                                self.nome_coluna_esta_na_planilha1: "Sim",
                                self.nome_coluna_similares: "",
                            }
                        )
                        progress_value = int((i + 1) / total * 100) if total else 0
                        self.progress.emit(progress_value)
                        continue

                for nome_sistema_normalizado in self.df1_compostos_norm:
                    if nome_sistema_normalizado == valor_normalizado:
                        encontrado_exato = True
                        break

                if not encontrado_exato:
                    for composto_exibicao, composto_norm in zip(
                        self.df1_compostos_exibicao, self.df1_compostos_norm
                    ):
                        score = fuzz.token_sort_ratio(composto_norm, valor_normalizado)
                        if score >= self.similaridade_min:
                            score_formatado = f"{score:.1f}".replace(".", ",")
                            nomes_similares.append(f"{self.normalize_func(composto_exibicao)} ({score_formatado}%)")

                resultados.append(
                    {
                        self.nome_coluna_planilha2: valor_exibicao,
                        self.nome_coluna_esta_na_planilha1: "Sim" if encontrado_exato else "Não",
                        self.nome_coluna_similares: ", ".join(nomes_similares) if nomes_similares else "",
                    }
                )

                progress_value = int((i + 1) / total * 100) if total else 0
                self.progress.emit(progress_value)

            self.finished.emit({"cancelado": False, "resultados": resultados})
        except Exception as e:
            self.error.emit(str(e))

    def _toggle_lista(self, lista, botao, checked):
        lista.setVisible(checked)
        botao.setText("▼ Colunas" if checked else "▶ Colunas")

