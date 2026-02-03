from __future__ import annotations

import sys

from PyQt5.QtWidgets import QApplication

from app.main_window import AplicacaoPrincipal
from shared.win_subprocess import apply_no_window_patch


def main() -> None:
    # Evita janelas CMD piscando no Windows ao usar pdftoppm (compressor/separador)
    apply_no_window_patch()
    app = QApplication(sys.argv)
    window = AplicacaoPrincipal()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()

