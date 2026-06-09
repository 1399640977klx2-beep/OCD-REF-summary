"""
Tab 4: Folder Management
Integrated from renamefolder.py and re_renamefolder.py
"""
import os
import re
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QListWidget, QListWidgetItem, QLineEdit,
    QGroupBox, QMessageBox, QSplitter, QProgressBar
)
from PyQt5.QtCore import Qt


class FolderManagerTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.folder_data = []
        self.base_path = ""
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        ctrl_group = QGroupBox("\u63a7\u5236\u9762\u677f")
        ctrl_layout = QVBoxLayout()

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("\u76ee\u6807\u6587\u4ef6\u5939:"))
        self.folder_label = QLabel("\u672a\u9009\u62e9")
        self.folder_label.setStyleSheet("color: gray;")
        row1.addWidget(self.folder_label, 1)
        self.browse_btn = QPushButton("\u9009\u62e9\u6587\u4ef6\u5939")
        self.browse_btn.clicked.connect(self._browse)
        row1.addWidget(self.browse_btn)
        self.scan_btn = QPushButton("\u626b\u63cf\u6587\u4ef6\u5939")
        self.scan_btn.setEnabled(False)
        self.scan_btn.clicked.connect(self._scan_folders)
        row1.addWidget(self.scan_btn)
        ctrl_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.auto_btn = QPushButton("\u81ea\u52a8\u91cd\u547d\u540d")
        self.auto_btn.setEnabled(False)
        self.auto_btn.clicked.connect(self._auto_rename)
        row2.addWidget(self.auto_btn)
        ctrl_layout.addLayout(row2)

        ctrl_group.setLayout(ctrl_layout)
        layout.addWidget(ctrl_group)

        splitter = QSplitter(Qt.Horizontal)

        left_w = QWidget()
        left_l = QVBoxLayout(left_w)
        left_l.addWidget(QLabel("\u6587\u4ef6\u5939\u5217\u8868:"))
        self.folder_list = QListWidget()
        self.folder_list.itemClicked.connect(self._on_select)
        left_l.addWidget(self.folder_list)
        splitter.addWidget(left_w)

        right_w = QWidget()
        right_l = QVBoxLayout(right_w)
        right_l.addWidget(QLabel("\u7f16\u8f91\u524d\u7f00:"))
        self.info = QLabel("\u672a\u9009\u62e9")
        right_l.addWidget(self.info)
        right_l.addWidget(QLabel("\u65b0\u524d\u7f00:"))
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("\u8f93\u5165\u65b0\u524d\u7f00")
        right_l.addWidget(self.prefix_edit)
        self.apply_btn = QPushButton("\u5e94\u7528\u5230\u6b64\u9879")
        self.apply_btn.setEnabled(False)
        self.apply_btn.clicked.connect(self._apply)
        right_l.addWidget(self.apply_btn)

        right_l.addWidget(QLabel("\u6279\u91cf\u66ff\u6362:"))
        bl = QHBoxLayout()
        self.find_e = QLineEdit(); self.find_e.setPlaceholderText("\u67e5\u627e")
        self.repl_e = QLineEdit(); self.repl_e.setPlaceholderText("\u66ff\u6362\u4e3a")
        self.batch_btn = QPushButton("\u66ff\u6362")
        self.batch_btn.clicked.connect(self._batch)
        bl.addWidget(self.find_e); bl.addWidget(self.repl_e); bl.addWidget(self.batch_btn)
        right_l.addLayout(bl)
        right_l.addStretch()
        splitter.addWidget(right_w)
        splitter.setSizes([400, 400])
        layout.addWidget(splitter, 1)

        bl2 = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        bl2.addWidget(self.progress, 1)
        self.exec_btn = QPushButton("\u6267\u884c\u91cd\u547d\u540d")
        self.exec_btn.setEnabled(False)
        self.exec_btn.clicked.connect(self._execute)
        bl2.addWidget(self.exec_btn)
        layout.addLayout(bl2)

        self.setLayout(layout)

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "", self.main_window.last_directory)
        if d:
            self.main_window.last_directory = d
            self.base_path = d
            self.folder_label.setText(d)
            self.folder_label.setStyleSheet("color: black;")
            self.scan_btn.setEnabled(True)
            self.auto_btn.setEnabled(True)

    def _scan_folders(self):
        if not os.path.isdir(self.base_path): return
        self.folder_data = []
        self.folder_list.clear()
        p = re.compile(r"^(.+)_(\d{8}_\d{6})$")
        for root, dirs, _ in os.walk(self.base_path):
            for dn in dirs:
                m = p.match(dn)
                if m:
                    self.folder_data.append({
                        "orig": dn, "pref": m.group(1), "new_pref": m.group(1),
                        "dt": m.group(2), "path": os.path.join(root, dn)
                    })
                    self.folder_list.addItem(f"{m.group(1)} | {m.group(2)}")
        self.exec_btn.setEnabled(len(self.folder_data) > 0)
        if not self.folder_data:
            QMessageBox.information(self, "Info", "No matching folders found")

    def _auto_rename(self):
        if not os.path.isdir(self.base_path): return
        dp = re.compile(r"^\d{8}_\d{6}$")
        n = 0
        for root, dirs, _ in os.walk(self.base_path, topdown=True):
            for dn in dirs:
                if dp.match(dn):
                    fp = os.path.join(root, dn)
                    sd = [d for d in os.listdir(fp) if os.path.isdir(os.path.join(fp, d))]
                    if sd:
                        fs = [f for f in os.listdir(os.path.join(fp, sd[0])) if os.path.isfile(os.path.join(fp, sd[0], f))]
                        if fs:
                            px = fs[0].split("_")[0]
                            if px:
                                np = os.path.join(root, f"{px}_{dn}")
                                if not os.path.exists(np):
                                    os.rename(fp, np); n += 1
        QMessageBox.information(self, "Done", f"Renamed {n} folders")
        self._scan_folders()

    def _on_select(self, item):
        idx = self.folder_list.row(item)
        if idx < len(self.folder_data):
            d = self.folder_data[idx]
            self.info.setText(f"{d['pref']}_{d['dt']}")
            self.prefix_edit.setText(d["new_pref"])
            self.apply_btn.setEnabled(True)

    def _apply(self):
        idx = self.folder_list.currentRow()
        if idx < 0 or idx >= len(self.folder_data): return
        np = self.prefix_edit.text().strip()
        if np:
            self.folder_data[idx]["new_pref"] = np
            d = self.folder_data[idx]
            self.folder_list.item(idx).setText(f"{np} | {d['dt']}")

    def _batch(self):
        f, r = self.find_e.text(), self.repl_e.text()
        if not f: return
        c = 0
        for i, d in enumerate(self.folder_data):
            np = d["new_pref"].replace(f, r)
            if np != d["new_pref"]:
                self.folder_data[i]["new_pref"] = np
                self.folder_list.item(i).setText(f"{np} | {d['dt']}")
                c += 1
        QMessageBox.information(self, "Done", f"Updated {c} prefixes")

    def _execute(self):
        if not self.folder_data: return
        names = {}
        for d in self.folder_data:
            nn = f"{d['new_pref']}_{d['dt']}"
            if nn in names:
                QMessageBox.warning(self, "Error", f"Duplicate: {nn}"); return
            names[nn] = True
        self.progress.setVisible(True)
        self.progress.setMaximum(len(self.folder_data))
        s = 0
        for i, d in enumerate(self.folder_data):
            self.progress.setValue(i+1)
            nn = f"{d['new_pref']}_{d['dt']}"
            np = os.path.join(os.path.dirname(d["path"]), nn)
            if d["path"] != np and not os.path.exists(np):
                try: os.rename(d["path"], np); s += 1
                except: pass
        self.progress.setVisible(False)
        QMessageBox.information(self, "Done", f"Success: {s}/{len(self.folder_data)}")
        self._scan_folders()
