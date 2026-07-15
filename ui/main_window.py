"""
Main window with tabbed interface for OCD Toolbox.
"""
from PyQt5.QtWidgets import (
    QMainWindow, QTabWidget, QMessageBox, QStatusBar, QLabel
)
from PyQt5.QtCore import Qt

from ui.tab_ref_aggregator import RefAggregatorTab
from ui.tab_match import MatchTab
from ui.tab_wafer_map import WaferMapTab
from ui.tab_folder_manager import FolderManagerTab
from ui.tab_organize_spectra import OrganizeSpectraTab

APP_NAME = 'OCD Toolbox'


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(700, 500)

        self.last_directory = ''

        self._init_ui()
        self._init_statusbar()

    def _init_ui(self):
        # Main tab widget
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.North)
        self.tabs.setDocumentMode(True)

        # Create tabs
        self.tab_ref = RefAggregatorTab(self)
        self.tab_match = MatchTab(self)
        self.tab_map = WaferMapTab(self)
        self.tab_folder = FolderManagerTab(self)
        self.tab_organize = OrganizeSpectraTab(self)

        self.tabs.addTab(self.tab_ref, 'REF 数据汇总')
        self.tabs.addTab(self.tab_match, 'Match 匹配')
        self.tabs.addTab(self.tab_map, 'Wafer Map')
        self.tabs.addTab(self.tab_folder, '文件夹管理')
        self.tabs.addTab(self.tab_organize, '光谱整理')

        self.setCentralWidget(self.tabs)

    def _init_statusbar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel('就绪')
        self.status_bar.addWidget(self.status_label)

    def set_status(self, message):
        self.status_label.setText(message)

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self, '确认退出',
            '确定要退出 OCD Toolbox 吗？',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
