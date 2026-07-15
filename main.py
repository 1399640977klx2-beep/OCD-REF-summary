"""
OCD Toolbox - Main Entry Point
"""
import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

# Ensure high-DPI support
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

from ui.main_window import MainWindow

APP_NAME = 'OCD Toolbox'
APP_VERSION = '1.0.0'


def load_stylesheet(app):
    """Load global QSS stylesheet for application-wide styling."""
    qss_path = os.path.join(os.path.dirname(__file__), 'styles', 'main.qss')
    if os.path.exists(qss_path):
        with open(qss_path, 'r', encoding='utf-8') as f:
            app.setStyleSheet(f.read())
    return app


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)

    # Set app style
    app.setStyle('Fusion')
    load_stylesheet(app)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()


