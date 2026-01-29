from __future__ import annotations

from PyQt5.QtWidgets import QTabWidget, QVBoxLayout, QWidget

from pdf.separador import SeparadorPDFWidget
from pdf.separador_editor import PdfPageEditorWidget


class SeparadorContainerWidget(QWidget):
    """
    Aba principal "separador" com sub-abas:
    - separador: editor de páginas (preview tipo iLovePDF)
    - separador 2: funcionalidade antiga (com senha)
    """

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.sub_tabs = QTabWidget()
        self.sub_tabs.setTabPosition(QTabWidget.North)

        self.separador = PdfPageEditorWidget()
        self.separador2 = SeparadorPDFWidget()

        self.sub_tabs.addTab(self.separador, "🗂️ Separador")
        self.sub_tabs.addTab(self.separador2, "🔐 Separador 2")

        layout.addWidget(self.sub_tabs)
        self.setLayout(layout)

