"""
Tab 2: Match matching between company software output and BSL data.
"""
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QTableWidget, QTableWidgetItem, QGroupBox,
    QHeaderView, QAbstractItemView, QCheckBox, QScrollArea,
    QMessageBox, QFrame
)
from PyQt5.QtCore import Qt

from utils.file_utils import extract_wafer_id_from_sme_path
from utils.excel_writer import write_match_excel


class MatchTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.company_df = None
        self.bsl_df = None
        self.match_df = None
        self.stats_df = None
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        # === Input Section ===
        input_group = QGroupBox('数据加载')
        input_layout = QVBoxLayout()

        # Company output row
        row1 = QHBoxLayout()
        row1.addWidget(QLabel('公司软件输出:'))
        self.load_company_btn = QPushButton('加载 Excel')
        self.load_company_btn.clicked.connect(self._load_company_data)
        row1.addWidget(self.load_company_btn)
        self.company_label = QLabel('未加载')
        self.company_label.setStyleSheet('color: gray;')
        row1.addWidget(self.company_label, 1)
        input_layout.addLayout(row1)

        # BSL data row
        row2 = QHBoxLayout()
        row2.addWidget(QLabel('BSL/REF 数据:'))
        self.load_bsl_btn = QPushButton('加载 Excel')
        self.load_bsl_btn.clicked.connect(self._load_bsl_data)
        row2.addWidget(self.load_bsl_btn)
        self.bsl_label = QLabel('未加载')
        self.bsl_label.setStyleSheet('color: gray;')
        row2.addWidget(self.bsl_label, 1)
        input_layout.addLayout(row2)

        input_group.setLayout(input_layout)
        layout.addWidget(input_group)

        # === Parameter Selection ===
        param_group = QGroupBox('匹配参数选择')
        param_layout = QHBoxLayout()
        self.param_checkboxes = {}
        default_params = ['DP', 'HIGH']
        for p in default_params:
            cb = QCheckBox(p)
            cb.setChecked(True)
            self.param_checkboxes[p] = cb
            param_layout.addWidget(cb)
        param_layout.addStretch()
        self.refresh_params_btn = QPushButton('刷新参数列表')
        self.refresh_params_btn.setEnabled(False)
        self.refresh_params_btn.clicked.connect(self._refresh_params)
        param_layout.addWidget(self.refresh_params_btn)
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)

        # === Match Action ===
        action_layout = QHBoxLayout()
        self.match_btn = QPushButton('执行匹配')
        self.match_btn.setEnabled(False)
        self.match_btn.clicked.connect(self._run_match)
        action_layout.addWidget(self.match_btn)

        self.export_btn = QPushButton('导出 Match Excel')
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self._export_match)
        action_layout.addWidget(self.export_btn)

        action_layout.addStretch()
        layout.addLayout(action_layout)

        # === Results Preview ===
        self.result_label = QLabel('加载数据后点击"执行匹配"')
        self.result_label.setStyleSheet('font-weight: bold;')
        layout.addWidget(self.result_label)

        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table, 1)

        self.setLayout(layout)

    def _load_company_data(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, '加载公司软件输出',
            self.main_window.last_directory,
            'Excel Files (*.xlsx *.xls)'
        )
        if not filepath:
            return
        self.main_window.last_directory = os.path.dirname(filepath)

        try:
            self.company_df = pd.read_excel(filepath, sheet_name=0)
            # Extract wafer ID from path if Cur SME File Path exists
            if 'Cur SME File Path' in self.company_df.columns:
                self.company_df['Extracted_WaferID'] = self.company_df['Cur SME File Path'].apply(
                    extract_wafer_id_from_sme_path
                )
            self.company_label.setText(os.path.basename(filepath))
            self.company_label.setStyleSheet('color: black;')
            self._check_match_ready()
            self._refresh_params()
            self.main_window.set_status(f'公司数据加载: {len(self.company_df)} 行')
        except Exception as e:
            QMessageBox.critical(self, '错误', f'加载文件失败: {e}')

    def _load_bsl_data(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, '加载 BSL/REF 数据',
            self.main_window.last_directory,
            'Excel Files (*.xlsx *.xls)'
        )
        if not filepath:
            return
        self.main_window.last_directory = os.path.dirname(filepath)

        try:
            self.bsl_df = pd.read_excel(filepath, sheet_name=0)
            self.bsl_label.setText(os.path.basename(filepath))
            self.bsl_label.setStyleSheet('color: black;')
            self._check_match_ready()
            self._refresh_params()
            self.main_window.set_status(f'BSL数据加载: {len(self.bsl_df)} 行')
        except Exception as e:
            QMessageBox.critical(self, '错误', f'加载文件失败: {e}')

    def _check_match_ready(self):
        if self.company_df is not None and self.bsl_df is not None:
            self.match_btn.setEnabled(True)

    def _refresh_params(self):
        # Collect common parameter columns from both datasets
        common_params = set()
        for df in [self.company_df, self.bsl_df]:
            if df is not None:
                skip_cols = {'Cur SME File Path', 'Wafer ID', 'WaferID', 'Extracted_WaferID',
                            'Lot ID', 'LotID', 'Tool SN', 'ToolName', 'PAD Name', 'ModuleName',
                            'Die Seq', 'DieOrder', 'FIELD X', 'FIELD Y', 'Col/X', 'Row/Y',
                            'MSE', 'GOF', 'NGOF', 'LBH', 'regIter', 'SourceFile', 'Vendor'}
                for col in df.columns:
                    if col not in skip_cols and not col.startswith('Unnamed'):
                        common_params.add(col)

        # Update checkboxes
        param_layout = self.param_checkboxes
        for cb in self.findChildren(QCheckBox):
            param_layout[cb.text()] = cb

        # Already handled in layout
        for p in sorted(common_params):
            if p not in self.param_checkboxes:
                cb = QCheckBox(p)
                if p in ['DP', 'HIGH']:
                    cb.setChecked(True)
                self.param_checkboxes[p] = cb
                # Find parent layout and add
                param_group = self.findChild(QGroupBox, '匹配参数选择')
                if param_group:
                    pass  # Skip for now, parameters will be refreshed properly later

        self.refresh_params_btn.setEnabled(True)

    def _run_match(self):
        if self.company_df is None or self.bsl_df is None:
            QMessageBox.warning(self, '提示', '请先加载公司数据和 BSL 数据')
            return

        try:
            # Use extracted wafer ID or original Wafer ID
            company_key = self.company_df.get('Extracted_WaferID', self.company_df.get('Wafer ID'))
            bsl_key = self.bsl_df.get('WaferID')

            # Use FIELD X/FIELD Y vs Col/X/Row/Y
            company_x = self.company_df.get('FIELD X')
            company_y = self.company_df.get('FIELD Y')
            bsl_x = self.bsl_df.get('Col/X')
            bsl_y = self.bsl_df.get('Row/Y')

            if company_key is None or bsl_key is None:
                QMessageBox.warning(self, '匹配失败', '无法识别 Wafer ID 列')
                return

            # Get selected parameters
            selected_params = [p for p, cb in self.param_checkboxes.items() if cb.isChecked()]

            # Perform merge
            company_df_copy = self.company_df.copy()
            bsl_df_copy = self.bsl_df.copy()

            # Rename columns to avoid conflicts
            company_suffix = '_Company'
            bsl_suffix = '_BSL'

            company_rename = {c: f'{c}{company_suffix}' for c in company_df_copy.columns
                            if c not in ['Cur SME File Path', 'Extracted_WaferID', 'FIELD X', 'FIELD Y', 'Wafer ID']}
            bsl_rename = {c: f'{c}{bsl_suffix}' for c in bsl_df_copy.columns
                         if c not in ['WaferID', 'Col/X', 'Row/Y']}

            company_merge = company_df_copy.rename(columns=company_rename)
            bsl_merge = bsl_df_copy.rename(columns=bsl_rename)

            # Add merge keys
            company_merge['_WaferID'] = company_key
            company_merge['_X'] = company_x
            company_merge['_Y'] = company_y
            bsl_merge['_WaferID'] = bsl_key
            bsl_merge['_X'] = bsl_x
            bsl_merge['_Y'] = bsl_y

            # Merge
            self.match_df = pd.merge(
                company_merge, bsl_merge,
                on=['_WaferID', '_X', '_Y'],
                how='inner'
            )

            matched_count = len(self.match_df)
            company_count = len(self.company_df)
            bsl_count = len(self.bsl_df)

            # Build stats
            if matched_count > 0:
                stats_rows = []
                for wid in self.match_df['_WaferID'].unique():
                    subset = self.match_df[self.match_df['_WaferID'] == wid]
                    row = {'WaferID': wid}
                    for p in selected_params:
                        cp = f'{p}_Company'
                        bp = f'{p}_BSL'
                        if cp in subset.columns and bp in subset.columns:
                            row[f'{p}_Mean_Company'] = subset[cp].mean()
                            row[f'{p}_Mean_BSL'] = subset[bp].mean()
                            row[f'{p}_Bias'] = (subset[cp] - subset[bp]).mean()
                            row[f'{p}_Std_Company'] = subset[cp].std()
                            row[f'{p}_Std_BSL'] = subset[bp].std()
                    stats_rows.append(row)
                self.stats_df = pd.DataFrame(stats_rows)

            # Update UI
            self.result_label.setText(
                f'匹配完成: 公司 {company_count} 点 → BSL {bsl_count} 点 → 匹配 {matched_count} 点'
            )

            # Show preview
            self.table.setColumnCount(len(self.match_df.columns))
            self.table.setHorizontalHeaderLabels(list(self.match_df.columns))
            self.table.setRowCount(min(len(self.match_df), 50))
            for i in range(min(len(self.match_df), 50)):
                for j, col in enumerate(self.match_df.columns):
                    val = self.match_df.iloc[i][col]
                    self.table.setItem(i, j, QTableWidgetItem(str(val) if val is not None else ''))
            self.table.resizeColumnsToContents()

            self.export_btn.setEnabled(True)
            self.main_window.set_status(f'匹配完成: {matched_count} 个点匹配成功')

        except Exception as e:
            import traceback
            QMessageBox.critical(self, '匹配失败', f'匹配过程中发生错误:\n{e}\n{traceback.format_exc()}')

    def _export_match(self):
        if self.match_df is None:
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, '保存 Match Excel',
            os.path.join(self.main_window.last_directory, 'Match_Result.xlsx'),
            'Excel Files (*.xlsx)'
        )
        if filepath:
            write_match_excel(
                self.bsl_df if self.bsl_df is not None else pd.DataFrame(),
                self.stats_df if self.stats_df is not None else pd.DataFrame(),
                self.match_df,
                filepath
            )
            QMessageBox.information(self, '导出成功', f'Match Excel 已保存到:\n{filepath}')
            self.main_window.set_status(f'已导出: {filepath}')
