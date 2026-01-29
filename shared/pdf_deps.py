"""
Dependências opcionais (PDF/imagens/conversões).

O projeto original importava tudo num try/except e usava `PIL_AVAILABLE` para
habilitar/desabilitar funcionalidades na UI. Mantemos o mesmo padrão aqui.
"""

from __future__ import annotations

PIL_AVAILABLE = False

# Defaults para evitar NameError quando a UI só "checa" disponibilidade.
Image = None
canvas = None
letter = None
A4 = None
PdfMerger = None
PdfReader = None
PdfWriter = None
openpyxl = None
Document = None
pdf2image = None
xlsxwriter = None

try:
    from PIL import Image  # type: ignore
    from reportlab.pdfgen import canvas  # type: ignore
    from reportlab.lib.pagesizes import letter, A4  # type: ignore
    from PyPDF2 import PdfMerger, PdfReader, PdfWriter  # type: ignore
    import openpyxl  # type: ignore
    from docx import Document  # type: ignore
    import pdf2image  # type: ignore
    import xlsxwriter  # type: ignore

    PIL_AVAILABLE = True
except Exception:
    # Qualquer falha de import mantém PIL_AVAILABLE=False
    PIL_AVAILABLE = False

