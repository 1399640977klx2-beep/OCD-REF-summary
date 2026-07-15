"""
Tab 3: Wafer Map - opens original draw_wafer_map.py as a standalone window.
"""
import os
import sys
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, QLabel
from PyQt5.QtCore import Qt


class WaferMapTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        title = QLabel("Wafer Map")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        desc = QLabel("Click the button below to open the wafer map drawing tool\nin a separate window.")
        desc.setStyleSheet("color: gray; font-size: 13px;")
        desc.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc)

        self.open_btn = QPushButton("Open Wafer Map")
        self.open_btn.setMinimumSize(200, 60)
        self.open_btn.setStyleSheet("font-size: 16px;")
        self.open_btn.clicked.connect(self._open)
        layout.addWidget(self.open_btn, alignment=Qt.AlignCenter)

        self.setLayout(layout)
        self._win = None

    def _open(self):
        # Import locally to avoid loading matplotlib at startup
        from .draw_wafer_map import DrawMapSubWindow
        self._win = DrawMapSubWindow(self.main_window)
        self._win.show()
