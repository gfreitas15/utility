from __future__ import annotations

from PyQt5.QtWidgets import QTabWidget, QVBoxLayout, QWidget

from planilhas.comparador_duas import ComparadorPlanilhasWidget
from planilhas.nomes_similares import NomesSimilaresWidget


class ComparadorPlanilhasContainerWidget(QWidget):
    """Container da aba Comparador de Planilhas: sub-abas 'Comparar duas planilhas' e 'Nomes similares'."""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.sub_tabs = QTabWidget()
        self.sub_tabs.setTabPosition(QTabWidget.North)
        self.comparar_duas = ComparadorPlanilhasWidget()
        self.nomes_similares = NomesSimilaresWidget()
        self.sub_tabs.addTab(self.comparar_duas, "Comparar duas planilhas")
        self.sub_tabs.addTab(self.nomes_similares, "Nomes similares")
        layout.addWidget(self.sub_tabs)
        self.setLayout(layout)

    def aplicar_tema(self, tema_escuro):
        self.comparar_duas.aplicar_tema(tema_escuro)
        self.nomes_similares.aplicar_tema(tema_escuro)

