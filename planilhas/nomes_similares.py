from __future__ import annotations

import pandas as pd
from rapidfuzz import fuzz

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QCursor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QHeaderView,
    QToolTip,
)


class NomesSimilaresWorker(QThread):
    """Thread para comparar nomes na planilha e encontrar similares (incluindo erros de digitação)."""

    progress = pyqtSignal(int)
    finished = pyqtSignal(list)  # lista de (nome1, nome2, similaridade_pct)
    error = pyqtSignal(str)

    def __init__(self, nomes, limite_similaridade):
        super().__init__()
        self.nomes = nomes  # lista de strings (nomes únicos ou com índice para exibição)
        self.limite_similaridade = limite_similaridade
        self._cancelado = False

    def cancel(self):
        self._cancelado = True

    def run(self):
        try:
            # rapidfuzz.ratio: considera ordem dos nomes (erros de digitação sim; troca de ordem não)
            n = len(self.nomes)
            total_pares = n * (n - 1) // 2 if n > 1 else 0
            resultados = []
            pares_processados = 0
            for i in range(n):
                if self._cancelado:
                    break
                nome_a = self.nomes[i]
                if nome_a is None or (isinstance(nome_a, float) and pd.isna(nome_a)):
                    nome_a = ""
                nome_a = str(nome_a).strip()
                if not nome_a:
                    continue
                for j in range(i + 1, n):
                    if self._cancelado:
                        break
                    nome_b = self.nomes[j]
                    if nome_b is None or (isinstance(nome_b, float) and pd.isna(nome_b)):
                        nome_b = ""
                    nome_b = str(nome_b).strip()
                    if not nome_b:
                        continue
                    if nome_a == nome_b:
                        continue
                    # Apenas ratio: ordem dos nomes é considerada (ex: "Joao Silva Souza" vs "Joao Souza Silva" fica menos similar)
                    similaridade = fuzz.ratio(nome_a, nome_b)
                    if similaridade >= self.limite_similaridade:
                        resultados.append((nome_a, nome_b, round(similaridade, 1)))
                    pares_processados += 1
                    if total_pares > 0 and pares_processados % max(1, total_pares // 100) == 0:
                        self.progress.emit(int(100 * pares_processados / total_pares))
            self.progress.emit(100)
            self.finished.emit(resultados)
        except Exception as e:
            self.error.emit(str(e))


class NomesSimilaresWidget(QWidget):
    """Sub-aba: selecionar uma planilha e detectar nomes similares (incluindo erros de digitação)."""

    def __init__(self):
        super().__init__()
        self.df = None
        self.caminho_planilha = ""
        self._worker = None
        self._ultimos_resultados = []  # lista de (nome1, nome2, similaridade) para reordenar
        self.init_ui()
        self.aplicar_tema(True)

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        fonte_label = QFont("Segoe UI", 10)
        fonte_botao = QFont("Segoe UI", 10)

        # --- Seleção da planilha ---
        lbl_planilha = QLabel("Planilha com nomes")
        lbl_planilha.setFont(fonte_label)
        layout.addWidget(lbl_planilha)
        planilha_layout = QHBoxLayout()
        self.btn_selecionar = QPushButton("📂 Selecionar Planilha")
        self.btn_selecionar.setFont(fonte_botao)
        self.btn_selecionar.clicked.connect(self.selecionar_planilha)
        planilha_layout.addWidget(self.btn_selecionar)
        layout.addLayout(planilha_layout)
        self.lbl_arquivo = QLabel("Nenhum arquivo selecionado")
        self.lbl_arquivo.setFont(fonte_label)
        self.lbl_arquivo.setWordWrap(True)
        layout.addWidget(self.lbl_arquivo)

        # --- Coluna com nomes ---
        col_layout = QHBoxLayout()
        col_layout.addWidget(QLabel("Coluna com os nomes:"))
        self.cmb_coluna = QComboBox()
        self.cmb_coluna.setMinimumWidth(220)
        self.cmb_coluna.setEnabled(False)
        col_layout.addWidget(self.cmb_coluna)
        col_layout.addStretch()
        layout.addLayout(col_layout)

        # --- Similaridade mínima ---
        sim_layout = QHBoxLayout()
        sim_layout.addWidget(QLabel("Similaridade mínima (0-100%):"))
        self.spin_similaridade = QSpinBox()
        self.spin_similaridade.setRange(50, 100)
        self.spin_similaridade.setValue(85)
        self.spin_similaridade.setSuffix("%")
        sim_layout.addWidget(self.spin_similaridade)
        sim_layout.addStretch()
        layout.addLayout(sim_layout)

        # --- Botões ---
        btn_layout = QHBoxLayout()
        self.btn_analisar = QPushButton("🔍 Detectar nomes similares")
        self.btn_analisar.clicked.connect(self.analisar_nomes)
        self.btn_analisar.setEnabled(False)
        self.btn_cancelar = QPushButton("❌ Cancelar")
        self.btn_cancelar.setEnabled(False)
        self.btn_cancelar.clicked.connect(self.cancelar_analise)
        self.progress = QProgressBar()
        self.progress.setStyleSheet(
            """
            QProgressBar { border: 2px solid #7f8c8d; border-radius: 5px; text-align: center; }
            QProgressBar::chunk { background-color: #3498db; border-radius: 3px; margin: 1px; }
        """
        )
        btn_layout.addWidget(self.btn_analisar)
        btn_layout.addWidget(self.progress)
        btn_layout.addWidget(self.btn_cancelar)
        layout.addLayout(btn_layout)

        # --- Ordenação e tabela de resultados ---
        ord_layout = QHBoxLayout()
        ord_layout.addWidget(QLabel("Nomes similares encontrados:"))
        ord_layout.addStretch()
        ord_layout.addWidget(QLabel("Ordenar por:"))
        self.cmb_ordenar = QComboBox()
        self.cmb_ordenar.addItems(
            [
                "Similaridade (maior primeiro)",
                "Similaridade (menor primeiro)",
                "Nome 1 (A-Z)",
                "Nome 1 (Z-A)",
                "Nome 2 (A-Z)",
                "Nome 2 (Z-A)",
            ]
        )
        self.cmb_ordenar.setMinimumWidth(220)
        self.cmb_ordenar.currentIndexChanged.connect(self._reordenar_e_preencher_tabela)
        ord_layout.addWidget(self.cmb_ordenar)
        layout.addLayout(ord_layout)
        self.tabela_resultados = QTableWidget()
        self.tabela_resultados.setColumnCount(3)
        self.tabela_resultados.setHorizontalHeaderLabels(["Nome 1", "Nome 2", "Similaridade (%)"])
        self.tabela_resultados.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tabela_resultados.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabela_resultados.setMinimumHeight(200)
        layout.addWidget(self.tabela_resultados)
        self.tabela_resultados.cellClicked.connect(self._copiar_nome_clicado)

        # Botão para exportar o grid para planilha
        export_layout = QHBoxLayout()
        export_layout.addStretch()
        self.btn_exportar = QPushButton("💾 Exportar para Excel")
        self.btn_exportar.clicked.connect(self.exportar_para_excel)
        self.btn_exportar.setEnabled(False)
        export_layout.addWidget(self.btn_exportar)
        layout.addLayout(export_layout)

        self.setLayout(layout)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                p = str(url.toLocalFile())
                if p.lower().endswith((".xlsx", ".xls")):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            return
        urls = [str(u.toLocalFile()) for u in event.mimeData().urls()]
        excel_files = [p for p in urls if p.lower().endswith((".xlsx", ".xls"))]
        if excel_files:
            self._carregar_planilha(excel_files[0])

    def selecionar_planilha(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "", "Excel (*.xlsx *.xls)")
        if path:
            self._carregar_planilha(path)

    def _carregar_planilha(self, path):
        try:
            self.df = pd.read_excel(path)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível ler a planilha.\n{str(e)}")
            return
        if self.df is None or len(self.df) == 0:
            QMessageBox.warning(self, "Aviso", "A planilha está vazia.")
            return
        self.caminho_planilha = path
        self.lbl_arquivo.setText(f"📁 {path}")
        self.cmb_coluna.clear()
        for col in self.df.columns:
            self.cmb_coluna.addItem(str(col))
        self.cmb_coluna.setEnabled(True)
        self.btn_analisar.setEnabled(True)
        self._ultimos_resultados = []
        self.tabela_resultados.setRowCount(0)
        self.btn_exportar.setEnabled(False)

    def analisar_nomes(self):
        if self.df is None or self.cmb_coluna.currentIndex() < 0:
            QMessageBox.warning(self, "Erro", "Selecione uma planilha e a coluna de nomes.")
            return
        col = self.cmb_coluna.currentText()
        if col not in self.df.columns:
            QMessageBox.warning(self, "Erro", "Coluna inválida.")
            return
        nomes = self.df[col].dropna().astype(str).str.strip()
        nomes = nomes[nomes != ""].tolist()
        if len(nomes) < 2:
            QMessageBox.warning(self, "Aviso", "É necessário ao menos 2 nomes na coluna para comparar.")
            return
        self.btn_analisar.setEnabled(False)
        self.btn_cancelar.setEnabled(True)
        self.progress.setValue(0)
        self._ultimos_resultados = []
        self.tabela_resultados.setRowCount(0)
        self._worker = NomesSimilaresWorker(nomes, self.spin_similaridade.value())
        self._worker.progress.connect(self.progress.setValue)
        self._worker.finished.connect(self._analise_finalizada)
        self._worker.error.connect(self._erro_analise)
        self._worker.start()

    def cancelar_analise(self):
        if self._worker is not None:
            self._worker.cancel()

    def _ordenacao_chave(self, item):
        """Retorna chave para ordenação conforme opção selecionada em cmb_ordenar."""
        nome1, nome2, pct = item
        idx = self.cmb_ordenar.currentIndex()
        if idx == 0:
            return (-pct, nome1, nome2)  # Similaridade maior primeiro
        if idx == 1:
            return (pct, nome1, nome2)  # Similaridade menor primeiro
        if idx == 2:
            return (nome1.lower(), nome2.lower(), pct)
        if idx == 3:
            return (nome1.lower()[::-1], nome2.lower()[::-1], pct)  # Z-A = inverter para sort
        if idx == 4:
            return (nome2.lower(), nome1.lower(), pct)
        # idx == 5: Nome 2 (Z-A)
        return (nome2.lower()[::-1], nome1.lower()[::-1], pct)

    def _reordenar_e_preencher_tabela(self):
        """Ordena _ultimos_resultados conforme cmb_ordenar e preenche a tabela."""
        if not self._ultimos_resultados:
            return
        # Ordenar: para Z-A usamos reverse na lista ordenada por A-Z
        idx = self.cmb_ordenar.currentIndex()
        if idx == 3:  # Nome 1 (Z-A)
            ordenados = sorted(
                self._ultimos_resultados, key=lambda x: (x[0].lower(), x[1].lower(), x[2]), reverse=True
            )
        elif idx == 5:  # Nome 2 (Z-A)
            ordenados = sorted(
                self._ultimos_resultados, key=lambda x: (x[1].lower(), x[0].lower(), x[2]), reverse=True
            )
        else:
            ordenados = sorted(self._ultimos_resultados, key=self._ordenacao_chave)
        self.tabela_resultados.setRowCount(len(ordenados))
        for row, (nome1, nome2, pct) in enumerate(ordenados):
            self.tabela_resultados.setItem(row, 0, QTableWidgetItem(nome1))
            self.tabela_resultados.setItem(row, 1, QTableWidgetItem(nome2))
            self.tabela_resultados.setItem(row, 2, QTableWidgetItem(str(pct)))

    def _analise_finalizada(self, resultados):
        self.btn_analisar.setEnabled(True)
        self.btn_cancelar.setEnabled(False)
        self._ultimos_resultados = list(resultados)
        self._reordenar_e_preencher_tabela()
        self.btn_exportar.setEnabled(bool(resultados))
        if resultados:
            QMessageBox.information(self, "Concluído", f"Foram encontrados {len(resultados)} par(es) de nomes similares.")
        else:
            QMessageBox.information(
                self,
                "Concluído",
                "Nenhum par de nomes similares encontrado com o limite de similaridade definido.",
            )

    def _erro_analise(self, msg):
        self.btn_analisar.setEnabled(True)
        self.btn_cancelar.setEnabled(False)
        QMessageBox.critical(self, "Erro", f"Erro ao analisar nomes:\n{msg}")

    def _copiar_nome_clicado(self, row, column):
        """Copia o nome clicado (colunas Nome 1 ou Nome 2) para a área de transferência e mostra popup."""
        if column not in (0, 1):
            return
        item = self.tabela_resultados.item(row, column)
        if not item:
            return
        texto = item.text()
        if not texto:
            return
        QApplication.clipboard().setText(texto)
        # Mini pop-up perto do cursor informando que foi copiado
        QToolTip.showText(QCursor.pos(), "Copiado!", self.tabela_resultados)

    def exportar_para_excel(self):
        """Exporta o grid atual de resultados para uma planilha Excel."""
        if not self._ultimos_resultados:
            QMessageBox.information(self, "Exportar", "Não há dados para exportar.")
            return
        caminho, _ = QFileDialog.getSaveFileName(self, "Salvar planilha", "", "Excel (*.xlsx)")
        if not caminho:
            return
        if not caminho.lower().endswith(".xlsx"):
            caminho += ".xlsx"
        try:
            df = pd.DataFrame(self._ultimos_resultados, columns=["Nome 1", "Nome 2", "Similaridade (%)"])
            df.to_excel(caminho, index=False)
            QMessageBox.information(self, "Sucesso", f"Planilha salva em:\n{caminho}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar a planilha.\n{str(e)}")

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
                QComboBox {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QSpinBox {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
            """
            )
        self.btn_selecionar.setStyleSheet("background-color: #3498db; color: white;")
        self.btn_analisar.setStyleSheet("background-color: #2ecc71; color: white;")
        self.btn_cancelar.setStyleSheet("background-color: #e74c3c; color: white;")

