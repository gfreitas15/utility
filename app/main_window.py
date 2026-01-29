from __future__ import annotations

import pandas as pd

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from pdf.compressor import CompressorPDFWidget
from pdf.conversor import ConversorPDFWidget
from pdf.separador import SeparadorPDFWidget
from planilhas.container import ComparadorPlanilhasContainerWidget


class AplicacaoPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ferramentas de Produtividade")
        self.setGeometry(100, 100, 1000, 700)
        self.setWindowIcon(QIcon("icone.ico"))
        self.tema_escuro = True

        self.init_ui()
        self.aplicar_tema()
        self.centralizar_janela()

    def init_ui(self):
        # Widget central com abas
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Layout principal
        layout = QVBoxLayout(self.central_widget)

        # Barra superior com tema e ajuda
        topo_layout = QHBoxLayout()
        topo_layout.addStretch()
        self.btn_ajuda = QPushButton("❓ Ajuda")
        self.btn_ajuda.clicked.connect(self.mostrar_ajuda)
        self.btn_tema = QPushButton("🌗 Alternar Tema")
        self.btn_tema.clicked.connect(self.alternar_tema)
        topo_layout.addWidget(self.btn_ajuda)
        topo_layout.addWidget(self.btn_tema)
        layout.addLayout(topo_layout)

        # Sistema de abas
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)

        # Aba do Comparador de Planilhas (container com sub-abas: Comparar duas planilhas | Nomes similares)
        self.comparador_widget = ComparadorPlanilhasContainerWidget()
        self.tab_widget.addTab(self.comparador_widget, "📊 Comparador de Planilhas")

        # Aba do Conversor de PDF
        self.conversor_widget = ConversorPDFWidget()
        self.tab_widget.addTab(self.conversor_widget, "📄 Conversor de PDF")

        # Aba do Separador de PDF (com autenticação)
        self.separador_widget = SeparadorPDFWidget()
        self.tab_widget.addTab(self.separador_widget, "🔐 Separador de PDF")

        # Aba do Compressor de PDF
        self.compressor_widget = CompressorPDFWidget()
        self.tab_widget.addTab(self.compressor_widget, "🗜️ Compressor de PDF")

        layout.addWidget(self.tab_widget)

        # Barra de status
        versao = pd.Timestamp.now().strftime("1.%Y.%m.%d")
        self.lbl_status = QLabel(f"Versão {versao}  |  Feito por GABRIEL")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        fonte_status = QFont("Segoe UI", 10)
        fonte_status.setBold(True)
        self.lbl_status.setFont(fonte_status)
        layout.addWidget(self.lbl_status)

    def mostrar_ajuda(self):
        texto = (
            "🔧 FERRAMENTAS DE PRODUTIVIDADE - GUIA COMPLETO\n\n"
            "📊 COMPARADOR DE PLANILHAS:\n"
            "• Funcionalidade principal para comparar dados entre duas planilhas Excel\n"
            "• Suporte a múltiplas colunas como chave de comparação\n"
            "• Algoritmo de similaridade configurável (0-100%)\n"
            "• Normalização automática de texto (acentos, maiúsculas, espaços)\n"
            "• Detecção automática de colunas CPF para match exato\n"
            "• Pré-visualização antes do processamento completo\n\n"
            "📄 CONVERSOR DE PDF:\n"
            "• Conversão de imagens para PDF (PNG, JPG, JPEG, GIF, BMP, TIFF, WEBP)\n"
            "• Monitoramento automático de pastas\n"
            "• Junção de múltiplos PDFs em um único documento\n"
            "• Conversões especiais: Excel→PDF, Word→PDF, PDF→Word, PDF→Imagem\n"
            "• Log de atividades em tempo real\n\n"
            "🔐 SEPARADOR DE PDF:\n"
            "• Separação de PDFs grandes por marcadores (bookmarks)\n"
            "• Geração automática de Excel com links clicáveis\n"
            "• Compressão opcional dos PDFs gerados\n"
            "• Acesso restrito por senha\n\n"
            "🗜️ COMPRESSOR DE PDF:\n"
            "• Compressão de PDFs com imagens escaneadas\n"
            "• Redução significativa do tamanho do arquivo\n"
            "• Controle de qualidade e resolução das imagens\n"
            "• Suporte a formatos JPEG e PNG\n"
            "💡 DICAS IMPORTANTES:\n"
            "• Use múltiplas colunas quando os dados precisarem de contexto (ex.: CPF + NOME)\n"
            "• A normalização remove acentos e espaços extras automaticamente\n"
            "• Se o Excel recusar salvar, feche o arquivo de destino e tente novamente\n"
            "• Arraste e solte arquivos diretamente na interface para facilitar o uso\n"
            "• O tema pode ser alternado entre claro e escuro usando o botão no canto superior\n\n"
            "🆘 SUPORTE:\n"
            "• Versão: 1.2026.01.29\n"
            "• Desenvolvido por: GABRIEL\n"
            "• Para problemas, verifique se todas as dependências estão instaladas"
        )
        QMessageBox.information(self, "Ajuda", texto)

    def aplicar_tema(self):
        if self.tema_escuro:
            self.setStyleSheet(
                """
                QMainWindow {background-color: #2c3e50; color: white;}
                QWidget {background-color: #2c3e50; color: white;}
                QTabWidget::pane {border: 1px solid #7f8c8d; background-color: #34495e;}
                QTabBar::tab {background-color: #2c3e50; color: white; padding: 8px 16px; margin-right: 2px;}
                QTabBar::tab:selected {background-color: #3498db; color: white;}
                QTabBar::tab:hover {background-color: #34495e;}
                QPushButton {border-radius: 8px; padding: 8px; background-color: #34495e; color: white; border: 1px solid #7f8c8d;}
                QPushButton:hover {background-color: #3498db;}
                QLabel {color: white; background-color: transparent;}
                QLineEdit {background-color: #34495e; color: white; border: 1px solid #7f8c8d; padding: 4px;}
                QComboBox {background-color: #34495e; color: white; border: 1px solid #7f8c8d; padding: 4px;}
                QSpinBox {background-color: #34495e; color: white; border: 1px solid #7f8c8d; padding: 4px;}
                QListWidget {background-color: #34495e; color: white; border: 1px solid #7f8c8d;}
                QTableWidget {background-color: #34495e; color: white; border: 1px solid #7f8c8d;}
                QTableWidget::item {background-color: #34495e; color: white;}
                QHeaderView::section {background-color: #2c3e50; color: white; border: 1px solid #7f8c8d; padding: 4px;}
                QProgressBar {background-color: #34495e; color: white; border: 2px solid #7f8c8d; border-radius: 5px;}
                QProgressBar::chunk {background-color: #3498db; border-radius: 3px; margin: 1px;}
            """
            )
            self.btn_tema.setStyleSheet("background-color: #7f8c8d; color: white; border: 1px solid #7f8c8d;")
            self.btn_ajuda.setStyleSheet("background-color: #7f8c8d; color: white; border: 1px solid #7f8c8d;")
        else:
            self.setStyleSheet(
                """
                QMainWindow {background-color: #ecf0f1; color: black;}
                QWidget {background-color: #ecf0f1; color: black;}
                QTabWidget::pane {border: 1px solid #7f8c8d; background-color: #bdc3c7;}
                QTabBar::tab {background-color: #95a5a6; color: black; padding: 8px 16px; margin-right: 2px;}
                QTabBar::tab:selected {background-color: #3498db; color: white;}
                QTabBar::tab:hover {background-color: #bdc3c7;}
                QPushButton {border-radius: 8px; padding: 8px; background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d;}
                QPushButton:hover {background-color: #3498db; color: white;}
                QLabel {color: black; background-color: transparent;}
                QLineEdit {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QComboBox {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QSpinBox {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QListWidget {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d;}
                QTableWidget {background-color: #bdc3c7; color: black; border: 1px solid #7f8c8d;}
                QTableWidget::item {background-color: #bdc3c7; color: black;}
                QHeaderView::section {background-color: #95a5a6; color: black; border: 1px solid #7f8c8d; padding: 4px;}
                QProgressBar {background-color: #bdc3c7; color: black; border: 2px solid #7f8c8d; border-radius: 5px;}
                QProgressBar::chunk {background-color: #3498db; border-radius: 3px; margin: 1px;}
            """
            )
            self.btn_tema.setStyleSheet("background-color: #7f8c8d; color: white; border: 1px solid #7f8c8d;")
            self.btn_ajuda.setStyleSheet("background-color: #7f8c8d; color: white; border: 1px solid #7f8c8d;")

        # Aplica tema nas abas
        self.comparador_widget.aplicar_tema(self.tema_escuro)  # container repassa para sub-abas
        self.conversor_widget.aplicar_tema(self.tema_escuro)

        # Aplica tema na barra de status
        if hasattr(self, "lbl_status"):
            if self.tema_escuro:
                self.lbl_status.setStyleSheet(
                    "background-color: #1f2a37; color: #ffd166; font-weight: bold; padding: 6px 0; border-top: 1px solid #7f8c8d;"
                )
            else:
                self.lbl_status.setStyleSheet(
                    "background-color: #e3e7ea; color: #1f2a37; font-weight: bold; padding: 6px 0; border-top: 1px solid #95a5a6;"
                )

    def alternar_tema(self):
        self.tema_escuro = not self.tema_escuro
        self.aplicar_tema()

    def centralizar_janela(self):
        """Centraliza a janela na tela"""
        # Obter geometria da tela (área disponível, sem barra de tarefas)
        tela = QApplication.desktop().availableGeometry()

        # Obter geometria da janela
        janela = self.geometry()

        # Calcular posição central
        x = (tela.width() - janela.width()) // 2
        y = (tela.height() - janela.height()) // 2

        # Ajustar para área disponível da tela
        x += tela.x()
        y += tela.y()

        # Mover janela para o centro
        self.move(x, y)

