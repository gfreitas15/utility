"""
Evita que subprocessos (pdftoppm, etc.) abram uma janela CMD no Windows.
Deve ser chamado uma vez no startup da aplicação.
"""
from __future__ import annotations

import sys
import subprocess

# No Python 3.7+ existe subprocess.CREATE_NO_WINDOW no Windows
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000 if sys.platform == "win32" else 0)

_original_popen = subprocess.Popen


def _popen_no_window(*args, **kwargs):
    if sys.platform == "win32":
        kwargs.setdefault("creationflags", CREATE_NO_WINDOW)
    return _original_popen(*args, **kwargs)


def apply_no_window_patch() -> None:
    """Faz com que todos os subprocessos (incl. pdf2image/pdftoppm) não abram janela CMD."""
    if sys.platform != "win32":
        return
    subprocess.Popen = _popen_no_window


def remove_no_window_patch() -> None:
    """Restaura o Popen original (útil para testes)."""
    subprocess.Popen = _original_popen
