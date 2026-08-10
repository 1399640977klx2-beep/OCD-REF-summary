"""
Tab 6: Organize Spectra - wraps organize_spectra.py
"""
import sys
import re
from pathlib import Path
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QLineEdit, QGroupBox, QMessageBox,
    QProgressBar, QTextEdit, QCheckBox
)
from PyQt5.QtCore import Qt

from utils.organize_spectra import (
    organize,
    DEFAULT_PRODUCT_PATTERN,
    DEFAULT_MACHINE_PATTERN,
    DEFAULT_WAFER_PATTERNS,
)


class OrganizeSpectraTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        # Step 1: Source
        s1 = QGroupBox("Step 1: Source Directory")
        s1l = QVBoxLayout()
        r1 = QHBoxLayout()
        self.src_btn = QPushButton("Browse Source")
        self.src_btn.clicked.connect(self._browse_src)
        r1.addWidget(self.src_btn)
        self.src_label = QLabel("Not selected")
        self.src_label.setStyleSheet("color: gray;")
        r1.addWidget(self.src_label, 1)
        s1l.addLayout(r1)
        s1.setLayout(s1l)
        layout.addWidget(s1)

        # Step 2: Destination
        s2 = QGroupBox("Step 2: Destination Directory")
        s2l = QVBoxLayout()
        r2 = QHBoxLayout()
        self.dst_btn = QPushButton("Browse Destination")
        self.dst_btn.clicked.connect(self._browse_dst)
        r2.addWidget(self.dst_btn)
        self.dst_label = QLabel("Not selected")
        self.dst_label.setStyleSheet("color: gray;")
        r2.addWidget(self.dst_label, 1)
        s2l.addLayout(r2)
        s2.setLayout(s2l)
        layout.addWidget(s2)

        # Step 3: Config
        s3 = QGroupBox("Step 3: Configuration")
        s3l = QVBoxLayout()
        r3 = QHBoxLayout()
        r3.addWidget(QLabel("Product Regex:"))
        self.product_re = QLineEdit(DEFAULT_PRODUCT_PATTERN)
        r3.addWidget(self.product_re)
        s3l.addLayout(r3)

        r3b = QHBoxLayout()
        r3b.addWidget(QLabel("Machine Regex:"))
        self.machine_re = QLineEdit(DEFAULT_MACHINE_PATTERN)
        r3b.addWidget(self.machine_re)
        s3l.addLayout(r3b)

        r3c = QHBoxLayout()
        r3c.addWidget(QLabel("Wafer Regex (comma separated):"))
        self.wafer_re = QLineEdit(", ".join(DEFAULT_WAFER_PATTERNS))
        r3c.addWidget(self.wafer_re)
        s3l.addLayout(r3c)

        r4 = QHBoxLayout()
        self.dry_run_cb = QCheckBox("Dry Run (preview only)")
        r4.addWidget(self.dry_run_cb)
        self.overwrite_cb = QCheckBox("Overwrite existing")
        r4.addWidget(self.overwrite_cb)
        r4.addStretch()
        s3l.addLayout(r4)

        # Show detected profiles
        self.profile_label = QLabel("")
        self.profile_label.setStyleSheet("color: #1565c0; font-size: 12px;")
        s3l.addWidget(self.profile_label)
        s3.setLayout(s3l)
        layout.addWidget(s3)

        # Step 4: Execute
        s4 = QGroupBox("Step 4: Execute")
        s4l = QVBoxLayout()
        r5 = QHBoxLayout()
        self.exec_btn = QPushButton("Organize Spectra")
        self.exec_btn.setEnabled(False)
        self.exec_btn.setMinimumHeight(40)
        self.exec_btn.clicked.connect(self._execute)
        r5.addWidget(self.exec_btn)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        r5.addWidget(self.progress, 1)
        s4l.addLayout(r5)
        s4.setLayout(s4l)
        layout.addWidget(s4)

        # Log area
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(150)
        self.log.setPlaceholderText("Execution log...")
        layout.addWidget(self.log)

        self.setLayout(layout)

    def _browse_src(self):
        d = QFileDialog.getExistingDirectory(self, "Select Source",
                                              self.main_window.last_directory)
        if d:
            self.src_label.setText(d)
            self.src_label.setStyleSheet("color: black;")
            self._update_profiles(d)
            self._check_ready()

    def _browse_dst(self):
        d = QFileDialog.getExistingDirectory(self, "Select Destination",
                                              self.main_window.last_directory)
        if d:
            self.dst_label.setText(d)
            self.dst_label.setStyleSheet("color: black;")
            self._check_ready()

    def _update_profiles(self, src_dir):
        # Show which machine folders are detected under the source
        try:
            p = Path(src_dir)
            machine_pat = self.machine_re.text().strip() or DEFAULT_MACHINE_PATTERN
            found = []
            for entry in p.iterdir():
                if entry.is_dir() and re.search(machine_pat, entry.name):
                    found.append(entry.name)
            self.profile_label.setText("Detected machine folders:\n" + "\n".join(found))
        except Exception:
            self.profile_label.setText("")

    def _check_ready(self):
        if self.src_label.text() != "Not selected" and self.dst_label.text() != "Not selected":
            self.exec_btn.setEnabled(True)

    def _execute(self):
        src = self.src_label.text()
        dst = self.dst_label.text()
        if src == dst:
            QMessageBox.warning(self, "Error", "Source and destination cannot be the same")
            return

        dry_run = self.dry_run_cb.isChecked()
        overwrite = self.overwrite_cb.isChecked()
        product_pat = self.product_re.text().strip() or DEFAULT_PRODUCT_PATTERN
        product_re = re.compile(product_pat)

        machine_pat = self.machine_re.text().strip() or DEFAULT_MACHINE_PATTERN
        machine_filters = [re.compile(machine_pat)]

        wafer_text = self.wafer_re.text().strip()
        if wafer_text:
            wafer_patterns = [
                re.compile(p.strip(), re.IGNORECASE)
                for p in wafer_text.split(",")
                if p.strip()
            ]
        else:
            wafer_patterns = [
                re.compile(p, re.IGNORECASE) for p in DEFAULT_WAFER_PATTERNS
            ]

        self.progress.setVisible(True)
        self.progress.setValue(50)
        self.log.clear()

        try:
            # Redirect print to log
            import io
            buf = io.StringIO()
            old_stdout = sys.stdout
            sys.stdout = buf

            organize(Path(src), Path(dst), product_re, machine_filters,
                     wafer_patterns=wafer_patterns, dry_run=dry_run,
                     overwrite=overwrite)

            sys.stdout = old_stdout
            self.log.setText(buf.getvalue())

            self.progress.setValue(100)
            msg = "Preview complete" if dry_run else "Organization complete"
            QMessageBox.information(self, "Done", msg)
            self.main_window.set_status(msg)
        except Exception as e:
            import traceback
            sys.stdout = sys.__stdout__
            QMessageBox.critical(self, "Error", str(e) + "\n" + traceback.format_exc())
        finally:
            self.progress.setVisible(False)
