from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QRectF
from PyQt5.QtGui import QColor, QIcon, QImage, QPainter, QPainterPath, QPen, QPixmap, QTransform
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from shared.pdf_deps import PIL_AVAILABLE, PdfReader, PdfWriter, pdf2image


@dataclass
class PageRef:
    src: Optional[str]  # caminho do PDF de origem (None para página em branco)
    index0: Optional[int]  # índice 0-based dentro do PDF de origem
    rotate: int = 0  # 0/90/180/270
    blank: bool = False
    stable_id: Optional[int] = None  # ID fixo da página (não muda ao reordenar)
    src_name: Optional[str] = None  # nome amigável do arquivo (sem caminho)


def _make_thumb_card(pix: QPixmap, size: QSize, radius: int = 10) -> QPixmap:
    """
    Produz uma miniatura "estilo iLovePDF": tamanho fixo, cantos arredondados e borda leve.
    Isso garante que TODAS as miniaturas ocupem o mesmo espaço visual.
    """
    w, h = size.width(), size.height()
    out = QPixmap(w, h)
    out.fill(Qt.transparent)

    p = QPainter(out)
    p.setRenderHint(QPainter.Antialiasing, True)
    p.setRenderHint(QPainter.SmoothPixmapTransform, True)

    rect = QRectF(out.rect().adjusted(2, 2, -2, -2))
    path = QPainterPath()
    path.addRoundedRect(rect, radius, radius)

    # Fundo branco
    p.fillPath(path, QColor("#ffffff"))

    # Conteúdo (imagem) com padding interno
    inner = rect.adjusted(6, 6, -6, -6)
    scaled = pix.scaled(inner.size().toSize(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
    x = int(inner.x() + (inner.width() - scaled.width()) / 2)
    y = int(inner.y() + (inner.height() - scaled.height()) / 2)

    p.save()
    p.setClipPath(path)
    p.drawPixmap(x, y, scaled)
    p.restore()

    # Borda leve
    pen = QPen(QColor("#cfd8dc"))
    pen.setWidth(1)
    p.setPen(pen)
    p.drawPath(path)
    p.end()
    return out


def _find_poppler_path() -> Optional[str]:
    """
    No Windows, pdf2image normalmente precisa do Poppler.
    Retorna:
    - None: usar PATH do sistema
    - str: pasta bin do poppler (contém pdftoppm.exe)
    """
    try:
        import shutil

        if shutil.which("pdftoppm"):
            return None
    except Exception:
        pass

    locais_comuns = [
        r"C:\poppler\Library\bin",
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files (x86)\poppler\Library\bin",
        os.path.join(os.path.expanduser("~"), "poppler", "Library", "bin"),
    ]
    for local in locais_comuns:
        if os.path.exists(os.path.join(local, "pdftoppm.exe")):
            return local
    return None


class _ThumbnailWorker(QThread):
    progress = pyqtSignal(int)
    page_ready = pyqtSignal(object, object)  # PageRef, QImage
    error = pyqtSignal(str)

    def __init__(self, pages: List[PageRef], dpi: int, poppler_path: Optional[str]):
        super().__init__()
        self._pages = pages
        self._dpi = dpi
        self._poppler_path = poppler_path
        self._cancel = False

    def cancel(self) -> None:
        self._cancel = True

    def run(self) -> None:
        try:
            if not PIL_AVAILABLE or pdf2image is None:
                self.error.emit("Dependências do preview não estão instaladas (pdf2image/Pillow).")
                return

            convert_from_path = pdf2image.convert_from_path

            total = len(self._pages) or 1
            for i, ref in enumerate(self._pages):
                if self._cancel:
                    return

                # Páginas em branco: gera thumbnail simples
                if ref.blank or not ref.src or ref.index0 is None:
                    img = QImage(400, 520, QImage.Format_RGB32)
                    img.fill(Qt.white)
                    self.page_ready.emit(ref, img)
                    self.progress.emit(int(((i + 1) / total) * 100))
                    continue

                kwargs = {
                    "dpi": self._dpi,
                    "first_page": ref.index0 + 1,
                    "last_page": ref.index0 + 1,
                }
                if self._poppler_path:
                    kwargs["poppler_path"] = self._poppler_path

                pil_imgs = convert_from_path(ref.src, **kwargs)
                if not pil_imgs:
                    self.progress.emit(int(((i + 1) / total) * 100))
                    continue

                pil = pil_imgs[0]
                # Reduz tamanho ainda no worker para evitar crash por memória (QPixmap de imagem gigante)
                try:
                    resample = getattr(Image, "LANCZOS", 1)  # type: ignore[name-defined]
                    # Mantém proporção; limita a um tamanho confortável para thumbnails
                    pil.thumbnail((900, 1200), resample=resample)
                except Exception:
                    # Se falhar, segue com imagem original
                    pass
                # Converte para bytes PNG e depois QImage
                from io import BytesIO

                buf = BytesIO()
                pil.save(buf, format="PNG")
                qimg = QImage.fromData(buf.getvalue(), "PNG")
                self.page_ready.emit(ref, qimg)
                self.progress.emit(int(((i + 1) / total) * 100))
        except Exception as e:
            self.error.emit(str(e))


class PdfPageEditorWidget(QWidget):
    """
    Editor simples de páginas (estilo iLovePDF):
    - thumbnails por página
    - reorder via drag&drop
    - remover páginas
    - inserir páginas (de outro PDF) + página em branco
    - rotacionar
    - salvar novo PDF
    """

    def __init__(self):
        super().__init__()
        self._poppler_path = _find_poppler_path()
        self._worker: Optional[_ThumbnailWorker] = None
        self._base_pdf_path: Optional[str] = None
        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)

        # Top bar
        top = QHBoxLayout()
        self.btn_open = QPushButton("📄 Selecionar PDF")
        self.btn_open.clicked.connect(self._selecionar_pdf)
        top.addWidget(self.btn_open)

        self.lbl_pdf = QLabel("Nenhum PDF selecionado")
        self.lbl_pdf.setWordWrap(True)
        top.addWidget(self.lbl_pdf, 1)

        self.btn_save = QPushButton("💾 Salvar PDF")
        self.btn_save.clicked.connect(self._salvar_pdf)
        self.btn_save.setEnabled(False)
        top.addWidget(self.btn_save)
        layout.addLayout(top)

        # Actions / reorder controls (sem drag&drop)
        actions = QHBoxLayout()
        self.btn_add_pdf = QPushButton("➕ Adicionar PDF…")
        self.btn_add_pdf.clicked.connect(self._adicionar_pdf)
        self.btn_add_pdf.setEnabled(False)
        actions.addWidget(self.btn_add_pdf)

        self.btn_add_blank = QPushButton("⬜ Adicionar página em branco")
        self.btn_add_blank.clicked.connect(self._adicionar_pagina_branca)
        self.btn_add_blank.setEnabled(False)
        actions.addWidget(self.btn_add_blank)

        self.btn_remove = QPushButton("🗑️ Remover selecionadas")
        self.btn_remove.clicked.connect(self._remover_selecionadas)
        self.btn_remove.setEnabled(False)
        actions.addWidget(self.btn_remove)

        self.btn_move_up = QPushButton("⬆️")
        self.btn_move_up.setToolTip("Mover seleção para cima")
        self.btn_move_up.clicked.connect(lambda: self._mover_delta(-1))
        self.btn_move_up.setEnabled(False)
        actions.addWidget(self.btn_move_up)

        self.btn_move_down = QPushButton("⬇️")
        self.btn_move_down.setToolTip("Mover seleção para baixo")
        self.btn_move_down.clicked.connect(lambda: self._mover_delta(1))
        self.btn_move_down.setEnabled(False)
        actions.addWidget(self.btn_move_down)

        actions.addWidget(QLabel("Mover para nº:"))
        self.spin_move_to = QSpinBox()
        self.spin_move_to.setMinimum(1)
        self.spin_move_to.setMaximum(1)
        self.spin_move_to.setEnabled(False)
        self.spin_move_to.valueChanged.connect(self._mover_para_posicao_selecionada)
        actions.addWidget(self.spin_move_to)

        self.btn_rotate_left = QPushButton("⟲ Girar -90°")
        self.btn_rotate_left.clicked.connect(lambda: self._rotacionar_selecionadas(-90))
        self.btn_rotate_left.setEnabled(False)
        actions.addWidget(self.btn_rotate_left)

        self.btn_rotate_right = QPushButton("⟳ Girar +90°")
        self.btn_rotate_right.clicked.connect(lambda: self._rotacionar_selecionadas(90))
        self.btn_rotate_right.setEnabled(False)
        actions.addWidget(self.btn_rotate_right)

        actions.addStretch()
        layout.addLayout(actions)

        # Progress
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Split: thumbnails (left) / preview (right)
        splitter = QSplitter(Qt.Horizontal)

        left = QWidget()
        left_l = QVBoxLayout(left)
        left_l.setContentsMargins(0, 0, 0, 0)

        self.list_pages = QListWidget()
        self.list_pages.setViewMode(QListView.IconMode)
        self.list_pages.setIconSize(QSize(140, 190))
        self.list_pages.setResizeMode(QListView.Adjust)
        self.list_pages.setMovement(QListView.Snap)
        self.list_pages.setSpacing(10)
        self.list_pages.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # Removido drag&drop: reordenação via botões + campo "Mover para nº"
        self.list_pages.setDragEnabled(False)
        self.list_pages.setAcceptDrops(False)
        self.list_pages.setDropIndicatorShown(False)
        self.list_pages.setDragDropMode(QAbstractItemView.NoDragDrop)
        self.list_pages.itemSelectionChanged.connect(self._on_selection_changed)
        # Estilo: seleção com borda azul arredondada
        self.list_pages.setStyleSheet(
            """
            QListWidget {
                background: transparent;
                outline: none;
            }
            QListWidget::item {
                border: 2px solid transparent;
                border-radius: 12px;
                padding: 6px;
            }
            QListWidget::item:selected {
                border: 2px solid #3498db;
                border-radius: 12px;
                background-color: rgba(52, 152, 219, 0.10);
            }
            """
        )
        left_l.addWidget(self.list_pages)

        splitter.addWidget(left)

        right = QWidget()
        right_l = QVBoxLayout(right)
        right_l.setContentsMargins(10, 0, 0, 0)

        self.lbl_preview_title = QLabel("Preview")
        self.lbl_preview_title.setStyleSheet("font-weight: bold;")
        right_l.addWidget(self.lbl_preview_title)

        self.lbl_preview = QLabel("Selecione uma página para visualizar")
        self.lbl_preview.setAlignment(Qt.AlignCenter)
        self.lbl_preview.setMinimumWidth(320)
        self.lbl_preview.setMinimumHeight(420)
        self.lbl_preview.setStyleSheet("border: 1px solid #7f8c8d; border-radius: 6px;")
        self.lbl_preview.setWordWrap(True)
        right_l.addWidget(self.lbl_preview, 1)

        self.lbl_hint = QLabel(
            "Dicas:\n"
            "- Use ⬆️/⬇️ ou o campo “Mover para nº” para reordenar\n"
            "- Use Ctrl/Shift para selecionar várias páginas\n"
            "- Salvar gera um novo PDF (não altera o original)"
        )
        self.lbl_hint.setStyleSheet("color: #7f8c8d;")
        right_l.addWidget(self.lbl_hint)

        splitter.addWidget(right)
        splitter.setSizes([700, 350])
        layout.addWidget(splitter, 1)

        # Dependências
        if not PIL_AVAILABLE or PdfReader is None or PdfWriter is None or pdf2image is None:
            self._set_disabled_state(
                "❌ Dependências do editor de páginas não estão instaladas.\n\n"
                "Instale: Pillow, PyPDF2, pdf2image (e Poppler no Windows)."
            )

    def _set_disabled_state(self, msg: str) -> None:
        self.btn_open.setEnabled(False)
        self.btn_save.setEnabled(False)
        self.btn_add_pdf.setEnabled(False)
        self.btn_add_blank.setEnabled(False)
        self.btn_remove.setEnabled(False)
        self.btn_move_up.setEnabled(False)
        self.btn_move_down.setEnabled(False)
        self.spin_move_to.setEnabled(False)
        self.btn_rotate_left.setEnabled(False)
        self.btn_rotate_right.setEnabled(False)
        self.lbl_preview.setText(msg)

    def _selecionar_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "PDF (*.pdf)")
        if not path:
            return
        self._carregar_pdf(path)

    def _carregar_pdf(self, path: str) -> None:
        try:
            reader = PdfReader(path)
            n = len(reader.pages)
            if n <= 0:
                QMessageBox.warning(self, "Aviso", "O PDF não contém páginas.")
                return
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível abrir o PDF.\n{str(e)}")
            return

        self.lbl_pdf.setText(f"📁 {path}")
        self._base_pdf_path = path
        self.list_pages.clear()
        self.lbl_preview.setText("Gerando miniaturas…")
        self.btn_save.setEnabled(True)
        self.btn_add_pdf.setEnabled(True)
        self.btn_add_blank.setEnabled(True)
        self.btn_remove.setEnabled(True)
        self.btn_move_up.setEnabled(True)
        self.btn_move_down.setEnabled(True)
        self.spin_move_to.setEnabled(True)
        self.btn_rotate_left.setEnabled(True)
        self.btn_rotate_right.setEnabled(True)

        # Cria itens "placeholder" com PageRef
        src_name = Path(path).name
        refs: List[PageRef] = [PageRef(src=path, index0=i, stable_id=i + 1, src_name=src_name) for i in range(n)]
        for ref in refs:
            item = QListWidgetItem("")
            item.setData(Qt.UserRole, ref)
            item.setFlags(
                item.flags()
                | Qt.ItemIsEnabled
                | Qt.ItemIsSelectable
                | Qt.ItemIsDragEnabled
                | Qt.ItemIsDropEnabled
            )
            # placeholder icon (card fixo)
            ph = QPixmap(400, 520)
            ph.fill(Qt.white)
            item.setIcon(QIcon(_make_thumb_card(ph, self.list_pages.iconSize())))
            self.list_pages.addItem(item)

        self._start_thumbnail_worker(refs)
        self._renumerar_labels()
        self._sync_move_spin()

    def _start_thumbnail_worker(self, refs: List[PageRef]) -> None:
        if self._worker is not None:
            try:
                self._worker.cancel()
            except Exception:
                pass

        self.progress.setVisible(True)
        self.progress.setValue(0)

        self._worker = _ThumbnailWorker(refs, dpi=110, poppler_path=self._poppler_path)
        self._worker.progress.connect(self.progress.setValue)
        self._worker.page_ready.connect(self._on_thumb_ready)
        self._worker.error.connect(self._on_thumb_error)
        self._worker.finished.connect(self._on_thumb_finished)
        self._worker.start()

    def _on_thumb_ready(self, ref: PageRef, img: QImage) -> None:
        # Guard: imagem inválida
        if img is None or img.isNull():
            return
        # encontra item correspondente pelo PageRef (src+index0+blank)
        for i in range(self.list_pages.count()):
            item = self.list_pages.item(i)
            it_ref: PageRef = item.data(Qt.UserRole)
            # Ignora rotação do ref do worker (pode ter sido alterada enquanto gerava)
            # IMPORTANT: usa stable_id para evitar colisão ao adicionar o mesmo PDF novamente
            if (
                it_ref.stable_id == ref.stable_id
                and it_ref.src == ref.src
                and it_ref.index0 == ref.index0
                and it_ref.blank == ref.blank
            ):
                # Converte para pixmap e reduz cedo para evitar alto consumo de memória
                pix = QPixmap.fromImage(img)
                if pix.isNull():
                    break
                if it_ref.rotate:
                    pix = pix.transformed(QTransform().rotate(it_ref.rotate), Qt.SmoothTransformation)
                card = _make_thumb_card(pix, self.list_pages.iconSize())
                item.setIcon(QIcon(card))
                # Guarda thumb (já com rotação aplicada) numa resolução moderada para preview
                thumb_img = pix.toImage()
                try:
                    thumb_img = thumb_img.scaled(
                        QSize(900, 1200), Qt.KeepAspectRatio, Qt.SmoothTransformation
                    )
                except Exception:
                    pass
                item.setData(Qt.UserRole + 1, thumb_img)
                break

        # atualiza preview se necessário
        self._update_preview_from_selection()

    def _on_thumb_error(self, msg: str) -> None:
        self.progress.setVisible(False)
        QMessageBox.warning(
            self,
            "Preview indisponível",
            "Não foi possível gerar miniaturas.\n\n"
            f"Detalhes: {msg}\n\n"
            "No Windows, instale o Poppler (pdftoppm) ou adicione ao PATH.",
        )

    def _on_thumb_finished(self) -> None:
        self.progress.setVisible(False)

    def _on_selection_changed(self) -> None:
        self._sync_move_spin()
        self._update_preview_from_selection()

    def _sync_move_spin(self) -> None:
        total = max(1, self.list_pages.count())
        self.spin_move_to.blockSignals(True)
        self.spin_move_to.setMaximum(total)
        items = self.list_pages.selectedItems()
        if items:
            row = self.list_pages.row(items[0])
            self.spin_move_to.setValue(row + 1)
        else:
            self.spin_move_to.setValue(1)
        self.spin_move_to.blockSignals(False)

    def _update_preview_from_selection(self) -> None:
        items = self.list_pages.selectedItems()
        if not items:
            return
        item = items[0]
        img: Optional[QImage] = item.data(Qt.UserRole + 1)
        if img is None:
            self.lbl_preview.setText("Carregando preview…")
            return
        pix = QPixmap.fromImage(img).scaled(
            QSize(320, 420), Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        self.lbl_preview.setPixmap(pix)

    def _renumerar_labels(self) -> None:
        # Mantém o número exibido de acordo com a ordem atual + mostra origem (página antiga)
        base = self._base_pdf_path
        for i in range(self.list_pages.count()):
            item = self.list_pages.item(i)
            ref: PageRef = item.data(Qt.UserRole)
            current = i + 1

            # Texto abaixo da miniatura: "Pag. <Atual> (<ID>)"
            if ref.blank:
                txt = f"Pag. {current} ({ref.stable_id or ''})".strip()
            elif ref.stable_id:
                txt = f"Pag. {current} ({ref.stable_id})"
            else:
                txt = f"Pag. {current}"
            item.setText(txt)

            # Tooltip com detalhes
            if ref.blank:
                item.setToolTip("Página em branco")
            elif ref.src and ref.stable_id:
                item.setToolTip(f"Origem: {ref.src_name or ref.src}\nID: {ref.stable_id}")
            else:
                item.setToolTip("")

    def _current_refs_in_order(self) -> List[PageRef]:
        refs: List[PageRef] = []
        for i in range(self.list_pages.count()):
            refs.append(self.list_pages.item(i).data(Qt.UserRole))
        return refs

    def _remover_selecionadas(self) -> None:
        items = self.list_pages.selectedItems()
        if not items:
            return
        for it in items:
            row = self.list_pages.row(it)
            self.list_pages.takeItem(row)
        self._renumerar_labels()
        self._sync_move_spin()
        self.lbl_preview.setText("Selecione uma página para visualizar")
        self.lbl_preview.setPixmap(QPixmap())

    def _adicionar_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Adicionar PDF", "", "PDF (*.pdf)")
        if not path:
            return
        try:
            reader = PdfReader(path)
            n = len(reader.pages)
            if n <= 0:
                QMessageBox.warning(self, "Aviso", "O PDF selecionado não contém páginas.")
                return
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível abrir o PDF.\n{str(e)}")
            return

        src_name = Path(path).name
        start_id = self.list_pages.count() + 1
        new_refs = [PageRef(src=path, index0=i, stable_id=start_id + i, src_name=src_name) for i in range(n)]
        for ref in new_refs:
            item = QListWidgetItem("")  # renumeração depois
            item.setData(Qt.UserRole, ref)
            item.setFlags(
                item.flags()
                | Qt.ItemIsEnabled
                | Qt.ItemIsSelectable
                | Qt.ItemIsDragEnabled
                | Qt.ItemIsDropEnabled
            )
            ph = QPixmap(400, 520)
            ph.fill(Qt.white)
            item.setIcon(QIcon(_make_thumb_card(ph, self.list_pages.iconSize())))
            self.list_pages.addItem(item)

        self._start_thumbnail_worker(new_refs)
        self._renumerar_labels()
        self._sync_move_spin()

    def _adicionar_pagina_branca(self) -> None:
        # Inserir após a seleção (ou no final)
        insert_row = self.list_pages.count()
        items = self.list_pages.selectedItems()
        if items:
            insert_row = self.list_pages.row(items[-1]) + 1

        ref = PageRef(src=None, index0=None, blank=True, stable_id=self.list_pages.count() + 1, src_name=None)
        item = QListWidgetItem("")
        item.setData(Qt.UserRole, ref)
        item.setFlags(
            item.flags()
            | Qt.ItemIsEnabled
            | Qt.ItemIsSelectable
            | Qt.ItemIsDragEnabled
            | Qt.ItemIsDropEnabled
        )
        # Thumb branco (card fixo)
        img = QImage(400, 520, QImage.Format_RGB32)
        img.fill(Qt.white)
        item.setData(Qt.UserRole + 1, img)
        item.setIcon(QIcon(_make_thumb_card(QPixmap.fromImage(img), self.list_pages.iconSize())))
        self.list_pages.insertItem(insert_row, item)
        self._renumerar_labels()
        self._sync_move_spin()

    def _mover_delta(self, delta: int) -> None:
        items = self.list_pages.selectedItems()
        if not items:
            return
        # Move bloco mantendo ordem relativa
        rows = sorted(self.list_pages.row(it) for it in items)
        if delta < 0 and rows[0] == 0:
            return
        if delta > 0 and rows[-1] == self.list_pages.count() - 1:
            return
        target = (rows[0] + 1 + delta)
        self._mover_bloco_para_posicao(target)

    def _mover_para_posicao_selecionada(self) -> None:
        # chamado quando o usuário digita/ajusta o novo número (1-based)
        items = self.list_pages.selectedItems()
        if not items:
            return
        target = int(self.spin_move_to.value())
        self._mover_bloco_para_posicao(target)

    def _mover_bloco_para_posicao(self, target_pos1: int) -> None:
        items = self.list_pages.selectedItems()
        if not items:
            return

        # Captura itens selecionados preservando ordem atual
        rows = sorted(self.list_pages.row(it) for it in items)
        selected_items: List[QListWidgetItem] = []
        for r in reversed(rows):
            selected_items.append(self.list_pages.takeItem(r))
        selected_items.reverse()

        # Agora a lista já está sem o bloco selecionado
        max_insert = self.list_pages.count() + 1
        target_pos1 = max(1, min(target_pos1, max_insert))
        insert_row = target_pos1 - 1

        for i, it in enumerate(selected_items):
            self.list_pages.insertItem(insert_row + i, it)
            it.setSelected(True)

        self._renumerar_labels()
        self._sync_move_spin()
        self._update_preview_from_selection()

    def _rotacionar_selecionadas(self, delta: int) -> None:
        items = self.list_pages.selectedItems()
        if not items:
            return
        for item in items:
            ref: PageRef = item.data(Qt.UserRole)
            ref.rotate = (ref.rotate + delta) % 360
            item.setData(Qt.UserRole, ref)

            # Rotaciona thumbnail se existir
            img: Optional[QImage] = item.data(Qt.UserRole + 1)
            if img is not None:
                t = QTransform().rotate(delta)
                pix = QPixmap.fromImage(img).transformed(t, Qt.SmoothTransformation)
                pix2 = pix.scaled(self.list_pages.iconSize(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
                item.setIcon(QIcon(pix2))
                # Atualiza imagem armazenada para refletir rotação acumulada
                item.setData(Qt.UserRole + 1, pix.toImage())

        self._update_preview_from_selection()

    def _salvar_pdf(self) -> None:
        if self.list_pages.count() == 0:
            QMessageBox.warning(self, "Aviso", "Não há páginas para salvar.")
            return

        out, _ = QFileDialog.getSaveFileName(self, "Salvar PDF", "", "PDF (*.pdf)")
        if not out:
            return
        if not out.lower().endswith(".pdf"):
            out += ".pdf"

        try:
            writer = PdfWriter()
            refs = self._current_refs_in_order()

            # Cache de readers por arquivo
            readers: Dict[str, PdfReader] = {}

            # Tamanho base para páginas em branco
            base_w, base_h = 595, 842  # A4 fallback
            for ref in refs:
                if ref.src and ref.index0 is not None and not ref.blank:
                    r = readers.get(ref.src)
                    if r is None:
                        readers[ref.src] = PdfReader(ref.src)
                        r = readers[ref.src]
                    try:
                        p0 = r.pages[ref.index0]
                        base_w = float(p0.mediabox.width)
                        base_h = float(p0.mediabox.height)
                    except Exception:
                        pass
                    break

            for ref in refs:
                if ref.blank or not ref.src or ref.index0 is None:
                    writer.add_blank_page(width=base_w, height=base_h)
                    continue

                r = readers.get(ref.src)
                if r is None:
                    readers[ref.src] = PdfReader(ref.src)
                    r = readers[ref.src]
                page = r.pages[ref.index0]

                if ref.rotate:
                    try:
                        page = page.rotate(ref.rotate)  # PyPDF2 3.x
                    except Exception:
                        try:
                            # fallback APIs
                            if ref.rotate == 90:
                                page.rotate_clockwise(90)
                            elif ref.rotate == 180:
                                page.rotate_clockwise(180)
                            elif ref.rotate == 270:
                                page.rotate_clockwise(270)
                        except Exception:
                            pass

                writer.add_page(page)

            with open(out, "wb") as f:
                writer.write(f)

            QMessageBox.information(self, "Sucesso", f"PDF salvo com sucesso!\n\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar o PDF.\n{str(e)}")

