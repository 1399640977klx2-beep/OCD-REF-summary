"""
Tab 1: REF Data Aggregation
"""
import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QFileDialog, QTableWidget, QTableWidgetItem,
    QGroupBox, QHeaderView, QAbstractItemView, QProgressBar,
    QMessageBox, QSplitter, QTextEdit
)
from PyQt5.QtCore import Qt

from parsers import parse_files
from utils.file_utils import scan_data_files
from utils.excel_writer import _style_header, _auto_width
from utils.excel_writer import write_ref_summary_excel, write_ref_summary_by_pad
import pandas as pd


class RefAggregatorTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.current_df = None
        self._pad_data = None
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        # === Control Panel ===
        ctrl_group = QGroupBox('控制面板')
        ctrl_layout = QVBoxLayout()

        # Row 1: Vendor select + folder browse
        row1 = QHBoxLayout()
        row1.addWidget(QLabel('厂商:'))
        self.vendor_combo = QComboBox()
        self.vendor_combo.setObjectName("cmbVendor")
        self.vendor_combo.addItems(['KLA', 'NOVA', 'PMISH'])
        row1.addWidget(self.vendor_combo)

        self.browse_btn = QPushButton('选择数据文件夹')
        self.browse_btn.clicked.connect(self._browse_folder)
        row1.addWidget(self.browse_btn)

        self.folder_label = QLabel('未选择文件夹')
        self.folder_label.setObjectName('folderLabel')
        self.folder_label.setProperty('active', False)
        row1.addWidget(self.folder_label, 1)

        ctrl_layout.addLayout(row1)

        # Row 2: Action buttons
        row2 = QHBoxLayout()
        self.parse_btn = QPushButton('开始解析')
        self.parse_btn.setEnabled(False)
        self.parse_btn.clicked.connect(self._parse_data)
        row2.addWidget(self.parse_btn)

        self.export_btn = QPushButton('导出汇总Excel')
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export_excel)
        row2.addWidget(self.export_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBarRef")
        self.progress_bar.setVisible(False)
        row2.addWidget(self.progress_bar, 1)

        ctrl_layout.addLayout(row2)
        ctrl_group.setLayout(ctrl_layout)
        layout.addWidget(ctrl_group)

        # === Info Area ===
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setObjectName("infoText")
        self.info_text.setPlaceholderText('解析信息将显示在这里...')
        layout.addWidget(self.info_text)

        # === Data Table Preview ===
        table_label = QLabel('数据预览')
        table_label.setObjectName('tablePreviewLabel')
        layout.addWidget(table_label)

        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table, 1)

        self.setLayout(layout)

    def _browse_folder(self):
        directory = QFileDialog.getExistingDirectory(self, '选择 REF 数据文件夹',
                                                      self.main_window.last_directory)
        if directory:
            self.main_window.last_directory = directory
            self.folder_label.setText(directory)
            self.folder_label.setProperty('active', True)
            self.folder_label.style().unpolish(self.folder_label)
            self.folder_label.style().polish(self.folder_label)
            self.parse_btn.setEnabled(True)
            self.current_dir = directory

    def _parse_data(self):
        vendor = self.vendor_combo.currentText()
        directory = self.folder_label.text()

        if not os.path.isdir(directory):
            QMessageBox.warning(self, '错误', '请选择有效的文件夹路径')
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.info_text.clear()
        self.main_window.set_status('正在扫描文件...')

        files = scan_data_files(directory)
        if not files:
            QMessageBox.warning(self, '提示', f'在所选文件夹中未找到数据文件 ({directory})')
            self.progress_bar.setVisible(False)
            return

        self.info_text.append(f'找到 {len(files)} 个文件')
        self.progress_bar.setMaximum(len(files))

        self.main_window.set_status(f'正在解析 {vendor} 格式...')
        def log_error(msg):
            self.info_text.append(msg)
        dfs = parse_files(files, vendor, error_handler=log_error)
        self.progress_bar.setValue(len(files))

        if dfs is None or len(dfs) == 0:
            QMessageBox.warning(self, '解析失败', '未能成功解析任何文件')
            self.progress_bar.setVisible(False)
            return

        self.info_text.append(f'成功解析 {len(dfs)} 个文件')

        # Combine DataFrames
        import pandas as pd
        if isinstance(dfs, dict):
            self._pad_data = dfs
            self.current_df = pd.concat(list(dfs.values()), ignore_index=True)
            self.info_text.append(f'成功解析 {len(dfs)} 个 PAD: {list(dfs.keys())}')
        else:
            self._pad_data = None
            self.current_df = pd.concat(dfs, ignore_index=True)
        self.info_text.append(f'总数据行数: {len(self.current_df)}')

        wafer_count = self.current_df.get('WaferID', self.current_df.get('Wafer_ID', pd.Series())).nunique()
        self.info_text.append(f'Wafer 数量: {wafer_count}')

        self._update_table()
        self.export_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.main_window.set_status(f'解析完成: {len(self.current_df)} 行数据')

    def _update_table(self):
        if self.current_df is None:
            return
        df = self.current_df.head(100)
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(list(df.columns))
        self.table.setRowCount(len(df))
        for i, (_, row) in enumerate(df.iterrows()):
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val) if val is not None else '')
                self.table.setItem(i, j, item)
        self.table.resizeColumnsToContents()

    def _export_excel(self):
        if self.current_df is None:
            return

        vendor = self.vendor_combo.currentText()
        default_name = f'{vendor}_REF_Summary.xlsx'
        filepath, _ = QFileDialog.getSaveFileName(
            self, '保存汇总 Excel',
            os.path.join(self.main_window.last_directory, default_name),
            'Excel Files (*.xlsx)'
        )
        if filepath:
            if self._pad_data is not None:
                with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                    for pad_name, pad_df in self._pad_data.items():
                        pad_df.to_excel(writer, sheet_name=str(pad_name)[:31], index=False)
                        _style_header(writer.sheets[str(pad_name)[:31]], pad_df.columns)
                        _auto_width(writer.sheets[str(pad_name)[:31]], pad_df)
            else:
                pad_col = 'Pad name' if 'Pad name' in self.current_df.columns else ('PadName' if 'PadName' in self.current_df.columns else None)
                if vendor in ('PMISH', 'KLA') and pad_col:
                    write_ref_summary_by_pad(self.current_df, filepath, group_col=pad_col)
                else:
                    write_ref_summary_excel(self.current_df, filepath, vendor)
            QMessageBox.information(self, '导出成功', f'汇总结果已保存到\n{filepath}')
            self.main_window.set_status(f'已导出: {filepath}')


