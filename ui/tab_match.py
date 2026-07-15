# Tab 2: Match Sheet Generation
import os, shutil
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QComboBox, QGroupBox, QCheckBox, QMessageBox,
    QProgressBar, QScrollArea, QGridLayout
)
from PyQt5.QtCore import Qt
from utils.excel_writer_match import generate_match_sheet


class MatchTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.filepath = None
        self.bsl_sheet = None
        self.company_sheet = None
        self.bsl_df = None
        self.company_df = None
        self.param_checkboxes = {}
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        # Step 1: Load Excel
        s1 = QGroupBox("Step 1: Load Excel")
        s1l = QVBoxLayout()
        r1 = QHBoxLayout()
        self.load_btn = QPushButton("Load Excel")
        self.load_btn.setObjectName("btnLoadMatch")
        self.load_btn.clicked.connect(self._load_file)
        r1.addWidget(self.load_btn)
        self.file_label = QLabel("Not loaded")
        self.file_label.setObjectName("fileLabel")
        self.file_label.setProperty("active", False)
        r1.addWidget(self.file_label, 1)
        s1l.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("BSL Sheet:"))
        self.bsl_combo = QComboBox()
        self.bsl_combo.setObjectName("cmbBslSheet")
        self.bsl_combo.setEnabled(False)
        self.bsl_combo.currentTextChanged.connect(self._on_bsl_changed)
        r2.addWidget(self.bsl_combo)
        r2.addWidget(QLabel("Company Sheet:"))
        self.company_combo = QComboBox()
        self.company_combo.setObjectName("cmbCompanySheet")
        self.company_combo.setEnabled(False)
        self.company_combo.currentTextChanged.connect(self._on_company_changed)
        r2.addWidget(self.company_combo)
        r2.addStretch()
        s1l.addLayout(r2)

        # Confirm button
        self.confirm_btn = QPushButton("Confirm Selection -> Load Parameters")
        self.confirm_btn.setObjectName("btnConfirm")
        self.confirm_btn.setEnabled(False)
        self.confirm_btn.clicked.connect(self._confirm_sheets)
        s1l.addWidget(self.confirm_btn)

        r3 = QHBoxLayout()
        self.bsl_info = QLabel("")
        self.bsl_info.setObjectName("bslInfo")
        r3.addWidget(self.bsl_info)
        self.company_info = QLabel("")
        self.company_info.setObjectName("companyInfo")
        r3.addWidget(self.company_info)
        r3.addStretch()
        s1l.addLayout(r3)
        s1.setLayout(s1l)
        layout.addWidget(s1)

        # Step 2: Parameters
        s2 = QGroupBox("Step 2: Select Parameters")
        s2l = QVBoxLayout()
        self.params_scroll = QScrollArea()
        self.params_scroll.setWidgetResizable(True)
        self.params_scroll.setObjectName("paramScroll")
        self.params_container = QWidget()
        self.params_layout = QGridLayout(self.params_container)
        self.params_layout.setAlignment(Qt.AlignTop)
        self.params_scroll.setWidget(self.params_container)
        s2l.addWidget(self.params_scroll)

        r4 = QHBoxLayout()
        self.sel_all_btn = QPushButton("Select All")
        self.sel_all_btn.setObjectName("btnSelectAll")
        self.sel_all_btn.setEnabled(False)
        self.sel_all_btn.clicked.connect(lambda: self._toggle_all(True))
        r4.addWidget(self.sel_all_btn)
        self.desel_btn = QPushButton("Deselect All")
        self.desel_btn.setObjectName("btnDeselectAll")
        self.desel_btn.setEnabled(False)
        self.desel_btn.clicked.connect(lambda: self._toggle_all(False))
        r4.addWidget(self.desel_btn)
        r4.addStretch()
        self.param_info = QLabel("")
        r4.addWidget(self.param_info)
        s2l.addLayout(r4)
        s2.setLayout(s2l)
        layout.addWidget(s2)

        # Step 3: Generate
        s3 = QGroupBox("Step 3: Generate")
        s3l = QVBoxLayout()
        r5 = QHBoxLayout()
        self.gen_btn = QPushButton("Generate Match Sheet -> Save As New File")
        self.gen_btn.setObjectName("btnGenerateMatch")
        self.gen_btn.setEnabled(False)
        self.gen_btn.clicked.connect(self._generate)
        r5.addWidget(self.gen_btn)
        self.progress = QProgressBar()
        self.progress.setObjectName("progressBarMatch")
        self.progress.setVisible(False)
        r5.addWidget(self.progress, 1)
        s3l.addLayout(r5)
        s3.setLayout(s3l)
        layout.addWidget(s3)
        layout.addStretch()
        self.setLayout(layout)

    def _load_file(self):
        fp, _ = QFileDialog.getOpenFileName(self, "Open Excel", self.main_window.last_directory, "Excel Files (*.xlsx *.xls)")
        if not fp: return
        self.main_window.last_directory = os.path.dirname(fp)
        self.filepath = fp
        self.file_label.setText(os.path.basename(fp))
        self.file_label.setProperty("active", True)
        self.file_label.style().unpolish(self.file_label)
        self.file_label.style().polish(self.file_label)
        try:
            xl = pd.ExcelFile(fp)
            sheets = xl.sheet_names
            self.bsl_combo.clear()
            self.company_combo.clear()
            self.bsl_combo.addItems(sheets)
            self.company_combo.addItems(sheets)
            self.bsl_combo.setEnabled(True)
            self.company_combo.setEnabled(True)
            self._clear_params()
            self.gen_btn.setEnabled(False)
            self.main_window.set_status(f"Loaded: {len(sheets)} sheets")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _on_bsl_changed(self, name):
        if not name or not self.filepath: return
        try:
            self.bsl_sheet = name
            self.bsl_df = pd.read_excel(self.filepath, sheet_name=name)
            self.bsl_info.setText(f"BSL: {len(self.bsl_df)} x {len(self.bsl_df.columns)}")
            self.confirm_btn.setEnabled(True)
            self._check_ready()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _on_company_changed(self, name):
        if not name or not self.filepath: return
        try:
            self.company_sheet = name
            self.company_df = pd.read_excel(self.filepath, sheet_name=name)
            self.company_info.setText(f"Company: {len(self.company_df)} x {len(self.company_df.columns)}")
            self.confirm_btn.setEnabled(True)
            self._check_ready()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _confirm_sheets(self):
        self._refresh_params()
        self._validate_params()
        self.bsl_info.setText(self.bsl_info.text() + " (loaded)")

    def _clear_params(self):
        while self.params_layout.count():
            item = self.params_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.param_checkboxes.clear()
        self.param_info.setText("")
        self.sel_all_btn.setEnabled(False)
        self.desel_btn.setEnabled(False)

    def _refresh_params(self):
        self._clear_params()
        if self.bsl_df is None: return
        for idx, col in enumerate(self.bsl_df.columns):
            cb = QCheckBox(str(col))
            cb.setObjectName("paramCheckbox")
            cb.stateChanged.connect(self._check_ready)
            self.param_checkboxes[str(col)] = cb
            self.params_layout.addWidget(cb, idx // 2, idx % 2)
        self.sel_all_btn.setEnabled(True)
        self.desel_btn.setEnabled(True)

    def _validate_params(self):
        if self.company_df is None: return
        cc = set(self.company_df.columns)
        ok = 0; bad = 0
        for name, cb in self.param_checkboxes.items():
            if name in cc:
                cb.setProperty("valid", True)
                cb.style().unpolish(cb)
                cb.style().polish(cb)
                ok += 1
            else:
                cb.setProperty("valid", False)
                cb.style().unpolish(cb)
                cb.style().polish(cb)
                bad += 1
        self.param_info.setText(f"Match: {ok} | Missing: {bad}")

    def _toggle_all(self, checked):
        for cb in self.param_checkboxes.values():
            cb.setChecked(checked)

    def _check_ready(self):
        has_checked = any(cb.isChecked() for cb in self.param_checkboxes.values())
        if self.bsl_sheet and self.company_sheet and has_checked:
            self.gen_btn.setEnabled(True)
        else:
            self.gen_btn.setEnabled(False)

    def _generate(self):
        sel = [p for p, cb in self.param_checkboxes.items() if cb.isChecked()]
        if not sel:
            QMessageBox.warning(self, "Error", "Please select at least one parameter")
            return
        cc = set(self.company_df.columns)
        missing = [p for p in sel if p not in cc]
        if missing:
            QMessageBox.warning(self, "Missing", "Not in company data: " + ", ".join(missing))
            return
        base = os.path.splitext(os.path.basename(self.filepath))[0] + "_match.xlsx"
        sp, _ = QFileDialog.getSaveFileName(self, "Save As", os.path.join(self.main_window.last_directory, base), "Excel Files (*.xlsx)")
        if not sp: return
        self.progress.setVisible(True)
        self.progress.setValue(30)
        try:
            generate_match_sheet(self.filepath, sp, self.bsl_sheet, self.company_sheet, sel)
            self.progress.setValue(100)
            QMessageBox.information(self, "Done", f"Saved to:\n{sp}")
            self.main_window.set_status(f"Exported: {os.path.basename(sp)}")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Failed", str(e) + "\n" + traceback.format_exc())
        finally:
            self.progress.setVisible(False)

