from __future__ import annotations

import sys

from PyQt5.QtWidgets import QApplication

from app.main_window import AplicacaoPrincipal


def main() -> None:
    app = QApplication(sys.argv)
    window = AplicacaoPrincipal()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()

