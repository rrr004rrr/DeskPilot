"""
DeskPilot — Windows Desktop Automation Tool
Entry point.
"""
# Set DPI awareness before any Qt/Windows initialisation.
# Qt wants DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 (-4).
# If we set it first the later Qt call is a no-op and the warning disappears.
import ctypes
try:
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:
    pass

import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("DeskPilot")
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
