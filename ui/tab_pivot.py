"""Tab: 数据透视表 - two pivot summary tools in one UI."""
import os
import re
import traceback

import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QComboBox, QFileDialog, QGroupBox, QMessageBox, QTextEdit,
)

from utils.pivot_summary import (
    load_dataframe as load_summary_df,
    compute_pivot_summary,
    save_summary,
    DEFAULT_PARAMETERS,
)
from utils.pivot_tool_summary import (
    load_dataframe as load_tool_df,
    compute_tool_pivot,
    save_tool_pivot,
)

TYPE_SUMMARY = 0
TYPE_TOOL = 1


class PivotTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self._last_auto_save = ""
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        type_group = QGroupBox("功能类型")
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("生成类型:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems([
            "第一种：Tool/Product/Wafer 汇总",
            "第二种：Product/Wafer × Tool 透视表",
        ])
        self.type_combo.currentIndexChanged.connect(self._update_default_save)
        type_layout.addWidget(self.type_combo, 1)
        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        input_group = QGroupBox("数据输入")
        input_layout = QVBoxLayout()

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("数据文件:"))
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("选择从 Excel 导出并复制出来的数据文件")
        self.input_edit.textChanged.connect(self._update_default_save)
        row1.addWidget(self.input_edit, 1)
        self.input_btn = QPushButton("浏览")
        self.input_btn.clicked.connect(self._browse_input)
        row1.addWidget(self.input_btn)
        input_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        row2.addWidget(self.sheet_combo, 1)
        input_layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("参数名:"))
        self.param_edit = QLineEdit(", ".join(DEFAULT_PARAMETERS))
        self.param_edit.setPlaceholderText(
            "逗号分隔，留空使用默认：CD_Bot, CD_Top, HIGH, SPA, THK")
        row3.addWidget(self.param_edit, 1)
        input_layout.addLayout(row3)

        input_group.setLayout(input_layout)
        layout.addWidget(input_group)

        output_group = QGroupBox("输出")
        output_layout = QVBoxLayout()

        row4 = QHBoxLayout()
        row4.addWidget(QLabel("保存路径:"))
        self.save_edit = QLineEdit()
        self.save_edit.setPlaceholderText("选择输出 xlsx 文件")
        row4.addWidget(self.save_edit, 1)
        self.save_btn = QPushButton("浏览")
        self.save_btn.clicked.connect(self._browse_save)
        row4.addWidget(self.save_btn)
        output_layout.addLayout(row4)

        gen_row = QHBoxLayout()
        self.generate_btn = QPushButton("生成透视表")
        self.generate_btn.setObjectName("btnGeneratePivot")
        self.generate_btn.clicked.connect(self._generate)
        gen_row.addWidget(self.generate_btn)
        gen_row.addStretch()
        output_layout.addLayout(gen_row)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(180)
        self.log.setPlaceholderText("生成日志与警告将显示在这里...")
        layout.addWidget(self.log, 1)

        self.setLayout(layout)

    def _browse_input(self):
        fp, _ = QFileDialog.getOpenFileName(
            self, "选择数据文件", self.main_window.last_directory,
            "Excel/CSV 文件 (*.xlsx *.csv)")
        if not fp:
            return
        self.main_window.last_directory = os.path.dirname(fp)
        self.input_edit.setText(fp)
        self._load_sheets(fp)

    def _load_sheets(self, filepath):
        self.sheet_combo.clear()
        if filepath.lower().endswith(".csv"):
            self.sheet_combo.addItem("CSV（无 Sheet）")
            self.sheet_combo.setEnabled(False)
            return
        try:
            xl = pd.ExcelFile(filepath)
            sheets = xl.sheet_names
            self.sheet_combo.addItems(sheets)
            self.sheet_combo.setEnabled(True)
            if sheets:
                self.sheet_combo.setCurrentIndex(0)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))

    def _suggest_save_path(self):
        fp = self.input_edit.text().strip()
        if not fp:
            return ""
        stem = os.path.splitext(os.path.basename(fp))[0]
        if self.type_combo.currentIndex() == TYPE_TOOL:
            name = stem + "_tool_pivot.xlsx"
        else:
            name = stem + "_pivot_summary.xlsx"
        return os.path.join(os.path.dirname(fp), name)

    def _update_default_save(self):
        suggested = self._suggest_save_path()
        if not suggested:
            return
        current = self.save_edit.text().strip()
        if current == "" or current == self._last_auto_save:
            self.save_edit.setText(suggested)
            self._last_auto_save = suggested

    def _browse_save(self):
        current = self.save_edit.text().strip()
        default_path = current or self._suggest_save_path()
        if not default_path:
            default_path = self.main_window.last_directory
        fp, _ = QFileDialog.getSaveFileName(
            self, "选择保存路径", default_path, "Excel 文件 (*.xlsx)")
        if not fp:
            return
        if not fp.lower().endswith(".xlsx"):
            fp += ".xlsx"
        self.save_edit.setText(fp)
        self._last_auto_save = fp

    def _parse_parameters(self):
        text = self.param_edit.text().strip()
        if not text:
            return list(DEFAULT_PARAMETERS)
        return [p.strip() for p in re.split(r"[,，;；]+", text) if p.strip()]

    def _selected_sheet(self):
        fp = self.input_edit.text().strip().lower()
        if fp.endswith(".csv"):
            return None
        return self.sheet_combo.currentText().strip() or None

    def _generate(self):
        input_path = self.input_edit.text().strip()
        save_path = self.save_edit.text().strip()
        if not input_path or not os.path.isfile(input_path):
            QMessageBox.warning(self, "错误", "请选择有效的输入文件")
            return
        if not save_path:
            QMessageBox.warning(self, "错误", "请选择保存路径")
            return
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        parameters = self._parse_parameters()
        if not parameters:
            QMessageBox.warning(self, "错误", "请输入至少一个参数名")
            return

        sheet = self._selected_sheet()
        self.log.clear()
        self.log.append(f"输入: {input_path}")
        self.log.append(f"参数: {', '.join(parameters)}")
        self.main_window.set_status("正在生成透视表...")

        try:
            if self.type_combo.currentIndex() == TYPE_TOOL:
                df = load_tool_df(input_path, sheet)
                column_names, rows, blocks, warnings = compute_tool_pivot(
                    df, parameters)
                save_tool_pivot(save_path, column_names, rows, blocks, warnings)
                self.log.append(f"完成: {len(rows)} 行，{len(column_names)} 列")
            else:
                df = load_summary_df(input_path, sheet)
                summary, warnings = compute_pivot_summary(df, parameters)
                save_summary(summary, save_path, warnings)
                self.log.append(
                    f"完成: {len(summary)} 行，{len(summary.columns)} 列")

            if warnings:
                self.log.append("警告:")
                for w in warnings:
                    self.log.append("  - " + w)
            else:
                self.log.append("无缺失后缀警告")
            self.log.append(f"已保存: {save_path}")
            self.main_window.set_status(
                f"透视表已生成: {os.path.basename(save_path)}")
            QMessageBox.information(
                self, "完成", f"透视表已生成：\n{save_path}")
        except Exception as e:
            self.log.append("错误: " + str(e))
            self.log.append(traceback.format_exc())
            self.main_window.set_status("透视表生成失败")
            QMessageBox.critical(self, "失败", str(e))
