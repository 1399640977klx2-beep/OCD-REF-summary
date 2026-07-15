import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.gridspec import GridSpec
from scipy.interpolate import griddata
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QFileDialog, QScrollArea, QComboBox, QLabel, QCheckBox,
                             QGroupBox, QGridLayout, QMessageBox, QMainWindow, QSpinBox,
                             QDoubleSpinBox, QFrame)
from PyQt5.QtCore import Qt
import os
import re


class DataSelectionWindow(QMainWindow):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df
        self.parent = parent
        self.selected_wafers = set()
        self.selected_params = set()
        self.setWindowTitle("选择要绘制的Wafer和参数")
        self.setGeometry(200, 200, 400, 400)
        self.initUI()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Wafer选择区域
        wafer_group = QGroupBox("选择Wafer")
        wafer_layout = QVBoxLayout()

        self.wafer_select_all = QCheckBox("全选")
        self.wafer_select_all.setChecked(True)
        self.wafer_select_all.stateChanged.connect(self.toggle_all_wafers)
        wafer_layout.addWidget(self.wafer_select_all)

        # Wafer复选框容器
        wafer_scroll = QScrollArea()
        wafer_scroll.setWidgetResizable(True)
        wafer_content = QWidget()
        self.wafer_layout = QVBoxLayout(wafer_content)

        wafers = self.df.iloc[:, 0].unique()
        self.wafer_checkboxes = {}
        for wafer in wafers:
            cb = QCheckBox(str(wafer))
            cb.setChecked(True)
            cb.stateChanged.connect(lambda state, cb_type='wafer': self.update_select_all_state(cb_type))
            self.wafer_checkboxes[wafer] = cb
            self.wafer_layout.addWidget(cb)
            self.selected_wafers.add(wafer)

        wafer_scroll.setWidget(wafer_content)
        wafer_layout.addWidget(wafer_scroll)
        wafer_group.setLayout(wafer_layout)
        layout.addWidget(wafer_group)

        # 参数选择区域
        param_group = QGroupBox("选择参数")
        param_layout = QVBoxLayout()

        self.param_select_all = QCheckBox("全选")
        self.param_select_all.setChecked(True)
        self.param_select_all.stateChanged.connect(self.toggle_all_params)
        param_layout.addWidget(self.param_select_all)

        # 参数复选框容器
        param_scroll = QScrollArea()
        param_scroll.setWidgetResizable(True)
        param_content = QWidget()
        self.param_layout = QVBoxLayout(param_content)

        params = self.df.columns[5:]
        self.param_checkboxes = {}
        for param in params:
            cb = QCheckBox(str(param))
            cb.setChecked(True)
            cb.stateChanged.connect(lambda state, cb_type='param': self.update_select_all_state(cb_type))
            self.param_checkboxes[param] = cb
            self.param_layout.addWidget(cb)
            self.selected_params.add(param)

        param_scroll.setWidget(param_content)
        param_layout.addWidget(param_scroll)
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)

        # 确认按钮
        confirm_btn = QPushButton("确认选择")
        confirm_btn.clicked.connect(self.confirm_selection)
        layout.addWidget(confirm_btn)

    def toggle_all_wafers(self, state):
        for cb in self.wafer_checkboxes.values():
            cb.setChecked(state == Qt.Checked)
        if state == Qt.Checked:
            self.selected_wafers = set(self.wafer_checkboxes.keys())
        else:
            self.selected_wafers.clear()

    def toggle_all_params(self, state):
        for cb in self.param_checkboxes.values():
            cb.setChecked(state == Qt.Checked)
        if state == Qt.Checked:
            self.selected_params = set(self.param_checkboxes.keys())
        else:
            self.selected_params.clear()

    def update_select_all_state(self, checkbox_type):
        """根据子复选框状态更新全选复选框"""
        if checkbox_type == 'wafer':
            checkboxes = self.wafer_checkboxes.values()
            select_all_cb = self.wafer_select_all
        else:  # param
            checkboxes = self.param_checkboxes.values()
            select_all_cb = self.param_select_all

        # 临时断开全选复选框的信号，避免循环触发
        select_all_cb.blockSignals(True)

        # 如果所有子复选框都被选中，则选中全选复选框
        if all(cb.isChecked() for cb in checkboxes):
            select_all_cb.setChecked(True)
        # 如果有任何一个子复选框未选中，则取消全选复选框
        else:
            select_all_cb.setChecked(False)

        # 重新连接信号
        select_all_cb.blockSignals(False)

    def confirm_selection(self):
        # 更新选中的wafer和参数
        self.selected_wafers = {wafer for wafer, cb in self.wafer_checkboxes.items() if cb.isChecked()}
        self.selected_params = {param for param, cb in self.param_checkboxes.items() if cb.isChecked()}

        if not self.selected_wafers or not self.selected_params:
            QMessageBox.warning(self, "警告", "请至少选择一个Wafer和一个参数")
            return

        self.parent.selected_wafers = self.selected_wafers
        self.parent.selected_params = self.selected_params
        self.close()


class DrawMapSubWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.df = None
        self.file_path = None
        self.selected_wafers = None  # 存储选中的wafer
        self.selected_params = None  # 存储选中的参数
        self.setWindowTitle("Draw Contour Maps")
        self.setGeometry(100, 100, 1200, 800)
        self.initUI()

    def initUI(self):
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # 顶部控制区域
        control_group = QGroupBox("控制面板")
        control_layout = QGridLayout()  # 改为网格布局

        # 第一行：文件操作
        self.load_button = QPushButton("加载Excel文件")
        self.load_button.clicked.connect(self.load_excel_file)
        control_layout.addWidget(self.load_button, 0, 0)

        control_layout.addWidget(QLabel("     选择Sheet:"), 0, 1)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        control_layout.addWidget(self.sheet_combo, 0, 2, 1, 1)
        control_layout.addWidget(QLabel("(注意:Sheet内XY坐标单位须为mm, 前5列依次为Wafer/die/X/Y/R, 第六列开始为数据, die和R列可以空着)"), 0, 4, 1, 3)

        self.select_data_button = QPushButton("选择数据")
        self.select_data_button.setEnabled(False)
        self.select_data_button.clicked.connect(self.open_data_selection)
        control_layout.addWidget(self.select_data_button, 0, 3)

        # 第二行：数据点显示开关
        data_points_label = QLabel("数据点设置:")
        data_points_label.setStyleSheet("font-weight: bold;")
        control_layout.addWidget(data_points_label, 1, 0)

        self.show_datapoints_check = QCheckBox("显示数据点")
        self.show_datapoints_check.setChecked(True)
        control_layout.addWidget(self.show_datapoints_check, 1, 1)

        self.color_datapoints_check = QCheckBox("彩色数据点")
        self.color_datapoints_check.setChecked(True)
        control_layout.addWidget(self.color_datapoints_check, 1, 2)

        # 第三行：数据点详细设置（点大小、透明度、边缘线宽）
        control_layout.addWidget(QLabel("点大小:"), 2, 0)
        self.datapoint_size_spin = QSpinBox()
        self.datapoint_size_spin.setRange(1, 100)
        self.datapoint_size_spin.setValue(30)
        control_layout.addWidget(self.datapoint_size_spin, 2, 1)

        control_layout.addWidget(QLabel("透明度:"), 2, 2)
        self.datapoint_alpha_spin = QDoubleSpinBox()
        self.datapoint_alpha_spin.setRange(0.1, 1.0)
        self.datapoint_alpha_spin.setSingleStep(0.1)
        self.datapoint_alpha_spin.setValue(0.7)
        control_layout.addWidget(self.datapoint_alpha_spin, 2, 3)

        control_layout.addWidget(QLabel("边缘线宽:"), 2, 4)
        self.datapoint_linewidth_spin = QDoubleSpinBox()
        self.datapoint_linewidth_spin.setRange(0.1, 3.0)
        self.datapoint_linewidth_spin.setSingleStep(0.1)
        self.datapoint_linewidth_spin.setValue(0.5)
        control_layout.addWidget(self.datapoint_linewidth_spin, 2, 5)

        # 添加水平分隔线
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.HLine)
        separator1.setFrameShadow(QFrame.Sunken)
        control_layout.addWidget(separator1, 3, 0, 1, 6)

        # 第四行：绘图设置（五个选项在同一行）
        plot_settings_label = QLabel("绘图设置:")
        plot_settings_label.setStyleSheet("font-weight: bold;")
        control_layout.addWidget(plot_settings_label, 4, 0)

        # 晶圆边界
        control_layout.addWidget(QLabel("边界:"), 5, 0)
        self.wafer_edge_combo = QComboBox()
        self.wafer_edge_combo.addItems(["none", "12英寸", "8英寸"])
        self.wafer_edge_combo.setCurrentText("none")
        control_layout.addWidget(self.wafer_edge_combo, 5, 1)

        # 颜色映射
        control_layout.addWidget(QLabel("颜色:"), 5, 2)
        self.cmap_combo = QComboBox()
        cmap_options = [
            'rainbow', 'viridis', 'plasma', 'inferno', 'magma', 'coolwarm',
            'jet', 'seismic', 'hot', 'cool', 'spring',
            'summer', 'autumn', 'winter', 'bone', 'pink', 'gray'
        ]
        self.cmap_combo.addItems(cmap_options)
        self.cmap_combo.setCurrentText('rainbow')
        control_layout.addWidget(self.cmap_combo, 5, 3)

        # 标题字体
        control_layout.addWidget(QLabel("标题字体:"), 5, 4)
        self.title_fontsize_spin = QSpinBox()
        self.title_fontsize_spin.setRange(6, 24)
        self.title_fontsize_spin.setValue(14)
        control_layout.addWidget(self.title_fontsize_spin, 5, 5)

        # 图像大小
        control_layout.addWidget(QLabel("图像大小:"), 5, 6)
        self.figsize_spin = QSpinBox()
        self.figsize_spin.setRange(4, 12)
        self.figsize_spin.setValue(8)
        control_layout.addWidget(self.figsize_spin, 5, 7)

        # 等高线层级
        control_layout.addWidget(QLabel("等高线:"), 5, 8)
        self.level_spin = QSpinBox()
        self.level_spin.setRange(5, 30)
        self.level_spin.setValue(8)
        control_layout.addWidget(self.level_spin, 5, 9)

        # 添加水平分隔线
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.HLine)
        separator2.setFrameShadow(QFrame.Sunken)
        control_layout.addWidget(separator2, 6, 0, 1, 11)

        # 第六行：操作按钮
        self.draw_button = QPushButton("绘制所有图像")
        self.draw_button.setEnabled(False)
        self.draw_button.clicked.connect(self.draw_all_maps)
        control_layout.addWidget(self.draw_button, 7, 11, 1, 1)

        self.save_button = QPushButton("保存所有图像")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_all_images)
        control_layout.addWidget(self.save_button, 7, 12, 1, 1)

        # 设置列宽比例
        control_layout.setColumnStretch(0, 1)  # 标签列 - 较小
        control_layout.setColumnStretch(1, 2)  # 下拉框/输入框 - 较大
        control_layout.setColumnStretch(2, 1)
        control_layout.setColumnStretch(3, 2)
        control_layout.setColumnStretch(4, 1)
        control_layout.setColumnStretch(5, 2)
        control_layout.setColumnStretch(6, 1)
        control_layout.setColumnStretch(7, 2)
        control_layout.setColumnStretch(8, 1)
        control_layout.setColumnStretch(9, 1)
        control_layout.setColumnStretch(10, 1)
        control_layout.setColumnStretch(11, 1)
        control_layout.setColumnStretch(12, 1)


        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        # 图像显示区域
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.scroll_content = QWidget()
        self.scroll_layout = QGridLayout(self.scroll_content)
        self.scroll_layout.setAlignment(Qt.AlignTop)

        self.scroll_area.setWidget(self.scroll_content)
        main_layout.addWidget(self.scroll_area)

        self.setLayout(main_layout)

    def open_data_selection(self):
        if self.df is None:
            return

        self.selection_window = DataSelectionWindow(self.df, self)
        self.selection_window.show()

    def get_selected_data(self):
        """获取当前选中的数据"""
        if self.selected_wafers is None:
            wafers = self.df.iloc[:, 0].unique()
        else:
            wafers = self.selected_wafers

        if self.selected_params is None:
            params = self.df.columns[5:]
        else:
            params = self.selected_params

        return wafers, params

    def load_excel_file(self):
        file_dialog = QFileDialog()
        if hasattr(self.parent, 'last_directory'):
            initial_dir = self.parent.last_directory
        else:
            initial_dir = os.getcwd()

        file_path, _ = file_dialog.getOpenFileName(
            self, 'Open Excel File', initial_dir, 'Excel Files (*.xlsx *.xls)'
        )

        if file_path:
            self.file_path = file_path
            if hasattr(self.parent, 'last_directory'):
                self.parent.last_directory = os.path.dirname(file_path)

            # 获取所有sheet名称
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names

                self.sheet_combo.clear()
                self.sheet_combo.addItem("请选择")
                self.sheet_combo.addItems(sheet_names)
                self.sheet_combo.setEnabled(True)
                self.draw_button.setEnabled(False)

            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败: {str(e)}")

    def on_sheet_changed(self, sheet_name):
        # 如果选择的是"请选择"或者空值，禁用按钮
        if sheet_name == "请选择" or not sheet_name:
            self.draw_button.setEnabled(False)
            self.select_data_button.setEnabled(True)
            return

        if sheet_name and self.file_path:
            try:
                self.df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                if len(self.df.columns) < 6:
                    QMessageBox.warning(self, "警告", "数据列数不足，请确保数据格式正确")
                    self.draw_button.setEnabled(False)
                else:
                    self.draw_button.setEnabled(True)  # 选择有效sheet后启用按钮

            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Sheet失败: {str(e)}")
                self.draw_button.setEnabled(False)

    def draw_all_maps(self):
        if self.df is None:
            return

        # 获取用户输入的参数
        title_fontsize = self.title_fontsize_spin.value()
        figsize = self.figsize_spin.value()
        level = self.level_spin.value()
        cmap = self.cmap_combo.currentText()
        show_datapoints = self.show_datapoints_check.isChecked()
        color_datapoints = self.color_datapoints_check.isChecked()
        datapoint_size = self.datapoint_size_spin.value()
        datapoint_alpha = self.datapoint_alpha_spin.value()
        datapoint_linewidth = self.datapoint_linewidth_spin.value()
        wafer_edge = self.wafer_edge_combo.currentText()

        # 清空之前的图像
        self.clear_scroll_area()

        # 获取wafer名称和参数名称（使用选择功能）
        if self.selected_wafers is not None:
            wafer_names = list(self.selected_wafers)
        else:
            wafer_names = self.df.iloc[:, 0].unique()

        if self.selected_params is not None:
            parameter_columns = list(self.selected_params)
        else:
            parameter_columns = self.df.columns[5:]

        # 计算总图数并提示
        total_plots = len(wafer_names) * len(parameter_columns)
        if total_plots > 16:
            reply = QMessageBox.question(
                self,
                "确认绘制",
                f"当前选择将绘制 {total_plots} 张图像，可能会耗时过久。\n是否继续？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        if len(wafer_names) == 0 or len(parameter_columns) == 0:
            QMessageBox.warning(self, "警告", "未找到有效的wafer或参数数据")
            return

        current_row = 0

        for wafer_name in wafer_names:
            # 为每个wafer创建一个水平布局来放置所有参数的图像
            wafer_group = QGroupBox(f"Wafer: {wafer_name}")
            wafer_layout = QHBoxLayout()

            for param_name in parameter_columns:
                try:
                    # 创建matplotlib图形
                    fig = Figure(figsize=(figsize, figsize * 0.75))
                    canvas = FigureCanvas(fig)

                    # 调用绘图函数，传入用户设置的参数
                    self.create_individual_contour_map(
                        self.df, wafer_name, param_name, fig=fig,
                        figsize=figsize, cmap=cmap, title_fontsize=title_fontsize,
                        level=level, show_datapoints=show_datapoints,
                        color_datapoints=color_datapoints,
                        datapoint_size=datapoint_size,
                        datapoint_alpha=datapoint_alpha,
                        datapoint_linewidth=datapoint_linewidth,
                        wafer_edge=wafer_edge,
                    )

                    # 设置canvas大小
                    canvas_width = int(figsize * 100)
                    canvas_height = int(figsize * 75)
                    canvas.setFixedSize(canvas_width, canvas_height)
                    wafer_layout.addWidget(canvas)

                except Exception as e:
                    print(f"绘制 {wafer_name} - {param_name} 时出错: {e}")

            # 如果没有成功绘制任何图像，跳过这个wafer
            if wafer_layout.count() > 0:
                wafer_group.setLayout(wafer_layout)
                self.scroll_layout.addWidget(wafer_group, current_row, 0)
                current_row += 1

        if current_row > 0:
            self.save_button.setEnabled(True)

        if current_row == 0:
            QMessageBox.warning(self, "警告", "未能成功绘制任何图像，请检查数据格式")


    def clear_scroll_area(self):
        # 移除所有子组件
        for i in reversed(range(self.scroll_layout.count())):
            widget = self.scroll_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

    def save_all_images(self):
        if self.df is None:
            return

        # 选择保存目录
        file_dialog = QFileDialog()
        save_dir = file_dialog.getExistingDirectory(
            self, "选择保存目录",
            self.parent.last_directory if hasattr(self.parent, 'last_directory') else os.getcwd()
        )

        if not save_dir:
            return

        # 获取当前参数设置
        title_fontsize = self.title_fontsize_spin.value()
        figsize = self.figsize_spin.value()
        level = self.level_spin.value()
        cmap = self.cmap_combo.currentText()
        show_datapoints = self.show_datapoints_check.isChecked()
        color_datapoints = self.color_datapoints_check.isChecked()
        datapoint_size = self.datapoint_size_spin.value()
        datapoint_alpha = self.datapoint_alpha_spin.value()
        datapoint_linewidth = self.datapoint_linewidth_spin.value()
        wafer_edge = self.wafer_edge_combo.currentText()

        # 获取wafer名称和参数名称（使用选择功能）
        if self.selected_wafers is not None:
            wafer_names = list(self.selected_wafers)
        else:
            wafer_names = self.df.iloc[:, 0].unique()

        if self.selected_params is not None:
            parameter_columns = list(self.selected_params)
        else:
            parameter_columns = self.df.columns[5:]

        saved_count = 0

        for wafer_name in wafer_names:
            for param_name in parameter_columns:
                try:
                    # 创建安全的文件名
                    safe_wafer_name = re.sub(r'[<>:"/\\|?*]', '_', str(wafer_name))
                    safe_param_name = re.sub(r'[<>:"/\\|?*]', '_', str(param_name))
                    filename = f"wafer_{safe_wafer_name}_{safe_param_name}.png"
                    output_path = os.path.join(save_dir, filename)

                    # 创建图形并直接保存
                    fig = Figure(figsize=(figsize, figsize * 0.75))

                    # 调用绘图函数
                    self.create_individual_contour_map(
                        self.df, wafer_name, param_name, fig=fig,
                        figsize=figsize, cmap=cmap, title_fontsize=title_fontsize,
                        level=level, show_datapoints=show_datapoints,
                        color_datapoints=color_datapoints,
                        datapoint_size=datapoint_size,
                        datapoint_alpha=datapoint_alpha,
                        datapoint_linewidth=datapoint_linewidth,
                        wafer_edge=wafer_edge,
                    )

                    # 独立保存图形
                    fig.tight_layout()  # 先用tight_layout调整布局
                    plt.subplots_adjust(left=0.05, right=0.95, bottom=0.05, top=0.95)
                    fig.savefig(output_path, dpi=200)
                    plt.close(fig)

                    saved_count += 1

                except Exception as e:
                    print(f"保存 {wafer_name} - {param_name} 时出错: {e}")

        # 显示保存结果
        QMessageBox.information(
            self,
            "保存完成",
            f"成功保存 {saved_count} 张图像到:\n{save_dir}"
        )

    def create_individual_contour_map(self, df, wafer_name, param_name, fig=None,
                                      output_path=None, show_plot=False, figsize=8,
                                      cmap='viridis', title_fontsize=14, level=10,
                                      show_datapoints=True, datapoint_size=30,
                                      datapoint_alpha=0.7, datapoint_edgecolor='black',
                                      datapoint_linewidth=0.5, color_datapoints=True,
                                      wafer_edge="none"):

        if fig is None:
            fig = plt.figure(figsize=(figsize, figsize * 0.75))
        else:
            # 清除现有图形
            fig.clear()

        # 使用GridSpec来更好地控制布局
        gs = GridSpec(2, 1, height_ratios=[35, 1], hspace=0.35)

        ax = fig.add_subplot(gs[0])  # 主图区域
        stats_ax = fig.add_subplot(gs[1])  # 统计信息区域

        # 筛选当前wafer的数据
        wafer_data = df[df.iloc[:, 0] == wafer_name]

        # 提取坐标和数据
        x = wafer_data.iloc[:, 2].values  # X坐标
        y = wafer_data.iloc[:, 3].values  # Y坐标
        z = wafer_data[param_name].values  # 参数值

        # 计算统计信息
        z_min = np.min(z)
        z_max = np.max(z)
        z_range = z_max - z_min
        z_std = np.std(z)

        # 创建统计信息文本
        stats_text = f"Min: {z_min:.2f} | Max: {z_max:.2f} | Range: {z_range:.2f} | STD: {z_std:.2f}"

        # 创建网格
        xi = np.linspace(min(x), max(x), 100)
        yi = np.linspace(min(y), max(y), 100)
        xi, yi = np.meshgrid(xi, yi)

        # 插值到规则网格
        zi = griddata((x, y), z, (xi, yi), method='cubic')

        # 绘制filled contour
        contourf = ax.contourf(xi, yi, zi, levels=level, cmap=cmap, alpha=0.6)

        # 绘制contour线（使用相同的colormap但颜色更深）
        contour = ax.contour(xi, yi, zi, levels=level, cmap=cmap, alpha=0.8, linewidths=1.3)

        # 添加数据点
        if show_datapoints:
            if color_datapoints:
                # 根据参数值设置数据点颜色（与等高线图一致）
                scatter = ax.scatter(x, y, c=z, cmap=cmap, s=datapoint_size,
                                     alpha=datapoint_alpha, edgecolors=datapoint_edgecolor,
                                     linewidth=datapoint_linewidth, vmin=z.min(), vmax=z.max())
            else:
                # 使用单一颜色
                ax.scatter(x, y, c='red', s=datapoint_size, alpha=datapoint_alpha,
                           edgecolors=datapoint_edgecolor, linewidth=datapoint_linewidth)

        # 添加颜色条
        cbar = fig.colorbar(contourf, ax=ax)
        cbar.set_label(param_name, rotation=270, labelpad=20, fontsize=title_fontsize - 2)

        # 计算坐标轴字体大小
        axis_fontsize = max(title_fontsize - 1, 8)

        # 添加标题和标签
        ax.set_title(f'Wafer {wafer_name} - {param_name}\nContour Map',
                     fontsize=title_fontsize, fontweight='bold')
        ax.set_xlabel('X (mm)', fontsize=axis_fontsize)
        ax.set_ylabel('Y (mm)', fontsize=axis_fontsize)

        # 设置等比例坐标轴
        ax.axis('equal')
        ax.grid(True, alpha=0.3)

        # 设置坐标轴刻度字体大小
        ax.tick_params(axis='both', which='major', labelsize=axis_fontsize - 2)

        # 绘制晶圆边界
        if wafer_edge != "none":
            if wafer_edge == "12英寸":
                radius = 150  # 12英寸晶圆半径约150mm
            else:  # 8英寸
                radius = 100  # 8英寸晶圆半径约100mm

            # 绘制圆形边界
            circle = plt.Circle((0, 0), radius, fill=False, color='black',
                                linestyle='--', linewidth=2, alpha=0.7)
            ax.add_patch(circle)

            # 设置坐标轴范围以适应晶圆边界
            ax.set_xlim(-radius * 1.1, radius * 1.1)
            ax.set_ylim(-radius * 1.1, radius * 1.1)

        # 在单独的轴中显示统计信息（无边框，只有文字）
        stats_ax.text(0.45, 0.2, stats_text,
                      ha='center', va='center',
                      fontsize=title_fontsize - 2,
                      fontweight='bold')

        # 隐藏统计信息轴的边框和刻度
        stats_ax.axis('off')

        fig.tight_layout()

        return fig

