"""
Tab 3: Wafer Map Drawing
Migrated from data/Py test/draw_wafer_map.py
"""
import os
import re
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Qt5Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.gridspec import GridSpec
from scipy.interpolate import griddata
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QScrollArea, QComboBox, QLabel, QCheckBox, QGroupBox, QGridLayout,
    QMessageBox, QSpinBox, QDoubleSpinBox, QFrame
)
from PyQt5.QtCore import Qt

class WaferMapTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self.df = None
        self.file_path = None
        self.selected_wafers = None
        self.selected_params = None
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        ctrl_group = QGroupBox('\u63a7\u5236\u9762\u677f')
        ctrl_layout = QGridLayout()

        self.load_btn = QPushButton('\u52a0\u8f7dExcel\u6587\u4ef6')
        self.load_btn.clicked.connect(self._load_file)
        ctrl_layout.addWidget(self.load_btn, 0, 0)
        ctrl_layout.addWidget(QLabel('Sheet:'), 0, 1)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        ctrl_layout.addWidget(self.sheet_combo, 0, 2)
        self.select_btn = QPushButton('\u9009\u62e9\u6570\u636e')
        self.select_btn.setEnabled(False)
        self.select_btn.clicked.connect(self._open_selection)
        ctrl_layout.addWidget(self.select_btn, 0, 3)

        ctrl_layout.addWidget(QLabel('\u6570\u636e\u70b9\u8bbe\u7f6e:'), 1, 0)
        self.show_dp = QCheckBox('\u663e\u793a\u6570\u636e\u70b9')
        self.show_dp.setChecked(True)
        ctrl_layout.addWidget(self.show_dp, 1, 1)
        self.color_dp = QCheckBox('\u5f69\u8272\u6570\u636e\u70b9')
        self.color_dp.setChecked(True)
        ctrl_layout.addWidget(self.color_dp, 1, 2)

        ctrl_layout.addWidget(QLabel('\u70b9\u5927\u5c0f:'), 2, 0)
        self.dp_size = QSpinBox(); self.dp_size.setRange(1,100); self.dp_size.setValue(30)
        ctrl_layout.addWidget(self.dp_size, 2, 1)
        ctrl_layout.addWidget(QLabel('\u900f\u660e\u5ea6:'), 2, 2)
        self.dp_alpha = QDoubleSpinBox(); self.dp_alpha.setRange(0.1,1.0); self.dp_alpha.setValue(0.7)
        ctrl_layout.addWidget(self.dp_alpha, 2, 3)
        ctrl_layout.addWidget(QLabel('\u8fb9\u7f18\u7ebf\u5bbd:'), 2, 4)
        self.dp_lw = QDoubleSpinBox(); self.dp_lw.setRange(0.1,3.0); self.dp_lw.setValue(0.5)
        ctrl_layout.addWidget(self.dp_lw, 2, 5)

        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setFrameShadow(QFrame.Sunken)
        ctrl_layout.addWidget(sep, 3, 0, 1, 6)

        ctrl_layout.addWidget(QLabel('\u7ed8\u56fe\u8bbe\u7f6e:'), 4, 0)
        ctrl_layout.addWidget(QLabel('\u8fb9\u754c:'), 5, 0)
        self.edge_combo = QComboBox()
        self.edge_combo.addItems(['none', '12\u82f1\u5bf8', '8\u82f1\u5bf8'])
        ctrl_layout.addWidget(self.edge_combo, 5, 1)

        ctrl_layout.addWidget(QLabel('\u989c\u8272:'), 5, 2)
        self.cmap_combo = QComboBox()
        for c in ['rainbow','viridis','plasma','coolwarm','jet']:
            self.cmap_combo.addItem(c)
        self.cmap_combo.setCurrentText('rainbow')
        ctrl_layout.addWidget(self.cmap_combo, 5, 3)

        ctrl_layout.addWidget(QLabel('\u6807\u9898\u5b57\u4f53:'), 5, 4)
        self.title_fs = QSpinBox(); self.title_fs.setRange(6,24); self.title_fs.setValue(14)
        ctrl_layout.addWidget(self.title_fs, 5, 5)

        ctrl_layout.addWidget(QLabel('\u56fe\u50cf\u5927\u5c0f:'), 5, 6)
        self.figsize_spin = QSpinBox(); self.figsize_spin.setRange(4,12); self.figsize_spin.setValue(8)
        ctrl_layout.addWidget(self.figsize_spin, 5, 7)

        ctrl_layout.addWidget(QLabel('\u7b49\u9ad8\u7ebf\u5c42\u7ea7:'), 5, 8)
        self.level_spin = QSpinBox(); self.level_spin.setRange(5,30); self.level_spin.setValue(8)
        ctrl_layout.addWidget(self.level_spin, 5, 9)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.HLine); sep2.setFrameShadow(QFrame.Sunken)
        ctrl_layout.addWidget(sep2, 6, 0, 1, 12)

        self.draw_btn = QPushButton('\u7ed8\u5236\u6240\u6709\u56fe\u50cf')
        self.draw_btn.setEnabled(False)
        self.draw_btn.clicked.connect(self._draw_all)
        ctrl_layout.addWidget(self.draw_btn, 7, 0)

        self.save_btn = QPushButton('\u4fdd\u5b58\u6240\u6709\u56fe\u50cf')
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self._save_all)
        ctrl_layout.addWidget(self.save_btn, 7, 1)

        ctrl_group.setLayout(ctrl_layout)
        layout.addWidget(ctrl_group)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QGridLayout(self.scroll_content)
        self.scroll_layout.setAlignment(Qt.AlignTop)
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area, 1)
        self.setLayout(layout)
        self.selection_dialog = None

    def _load_file(self):
        fp, _ = QFileDialog.getOpenFileName(self, '', self.main_window.last_directory if hasattr(self.main_window,'last_directory') else '', 'Excel Files (*.xlsx *.xls)')
        if not fp: return
        if hasattr(self.main_window,'last_directory'): self.main_window.last_directory = os.path.dirname(fp)
        try:
            xl = pd.ExcelFile(fp)
            self.file_path = fp
            self.sheet_combo.clear()
            self.sheet_combo.addItem('')
            self.sheet_combo.addItems(xl.sheet_names)
            self.sheet_combo.setEnabled(True)
            self.draw_btn.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))

    def _on_sheet_changed(self, name):
        if not name or not self.file_path: return
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=name)
            self.draw_btn.setEnabled(len(self.df.columns) > 5)
            self.select_btn.setEnabled(len(self.df.columns) > 5)
            self.selected_wafers = None
            self.selected_params = None
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))

    def _open_selection(self):
        if self.df is not None:
            from PyQt5.QtWidgets import QMainWindow
            dialog = QMainWindow(self)
            dialog.setWindowTitle('Select')
            dialog.setGeometry(200,200,500,400)
            cw = QWidget(); layout = QVBoxLayout(cw)
            self.w_cbs = {}; self.p_cbs = {}
            g1 = QGroupBox('Wafers')
            l1 = QVBoxLayout()
            for w in self.df.iloc[:,0].unique():
                cb = QCheckBox(str(w)); cb.setChecked(True); l1.addWidget(cb)
                self.w_cbs[w] = cb
            g1.setLayout(l1); layout.addWidget(g1)
            g2 = QGroupBox('Params')
            l2 = QVBoxLayout()
            for c in self.df.columns[5:]:
                if c in ('Vendor','SourceFile'): continue
                cb = QCheckBox(str(c)); cb.setChecked(True); l2.addWidget(cb)
                self.p_cbs[c] = cb
            g2.setLayout(l2); layout.addWidget(g2)
            btn = QPushButton('OK')
            def confirm():
                self.selected_wafers = set(w for w,cb in self.w_cbs.items() if cb.isChecked())
                self.selected_params = set(p for p,cb in self.p_cbs.items() if cb.isChecked())
                dialog.close()
            btn.clicked.connect(confirm)
            layout.addWidget(btn)
            dialog.setCentralWidget(cw)
            self.selection_dialog = dialog
            dialog.show()

    def _draw_all(self):
        if self.df is None: return
        while self.scroll_layout.count():
            w = self.scroll_layout.itemAt(0).widget()
            if w: w.deleteLater()
        wafers = list(self.selected_wafers) if self.selected_wafers else list(self.df.iloc[:,0].unique())
        params = list(self.selected_params) if self.selected_params else list(self.df.columns[5:])
        for wafer in wafers:
            wg = QGroupBox('Wafer: '+str(wafer))
            wl = QHBoxLayout()
            for param in params:
                try:
                    fig = Figure(figsize=(self.figsize_spin.value(), self.figsize_spin.value()*0.75))
                    canvas = FigureCanvas(fig)
                    self._draw_map(fig, wafer, param)
                    canvas.setFixedSize(int(self.figsize_spin.value()*100), int(self.figsize_spin.value()*75))
                    wl.addWidget(canvas)
                except Exception as e:
                    print('Error:', wafer, param, e)
            if wl.count():
                wg.setLayout(wl)
                self.scroll_layout.addWidget(wg, self.scroll_layout.count(), 0)
        self.save_btn.setEnabled(self.scroll_layout.count() > 0)

    def _draw_map(self, fig, wafer, param):
        data = self.df[self.df.iloc[:,0]==wafer]
        x, y, z = data.iloc[:,2].values, data.iloc[:,3].values, data[param].values
        gs = GridSpec(2,1,height_ratios=[35,1],hspace=0.35)
        ax, sa = fig.add_subplot(gs[0]), fig.add_subplot(gs[1])
        xi = np.linspace(x.min(),x.max(),100)
        yi = np.linspace(y.min(),y.max(),100)
        xi, yi = np.meshgrid(xi, yi)
        zi = griddata((x,y), z, (xi,yi), method='cubic')
        lv, cm = self.level_spin.value(), self.cmap_combo.currentText()
        cf = ax.contourf(xi,yi,zi,levels=lv,cmap=cm,alpha=0.6)
        ax.contour(xi,yi,zi,levels=lv,cmap=cm,alpha=0.8,linewidths=1.3)
        if self.show_dp.isChecked():
            kw = dict(s=self.dp_size.value(), alpha=self.dp_alpha.value(), edgecolors='black', linewidth=self.dp_lw.value())
            if self.color_dp.isChecked(): ax.scatter(x,y,c=z,cmap=cm,**kw)
            else: ax.scatter(x,y,c='red',**kw)
        fig.colorbar(cf,ax=ax).set_label(param,rotation=270,labelpad=20)
        ax.set_title(f'Wafer:{wafer} - {param}',fontsize=self.title_fs.value())
        ax.set_xlabel('X(mm)'); ax.set_ylabel('Y(mm)'); ax.axis('equal'); ax.grid(True,alpha=0.3)
        e = self.edge_combo.currentText()
        if e != 'none':
            r = 150.4 if '12' in e else 100
            ax.add_patch(plt.Circle((0,0),r,fill=False,color='black',linestyle='--',linewidth=2,alpha=0.7))
            ax.set_xlim(-r*1.1,r*1.1); ax.set_ylim(-r*1.1,r*1.1)
        sa.text(0.45,0.2,f'Min:{z.min():.2f} Max:{z.max():.2f} Range:{z.max()-z.min():.2f} Std:{z.std():.2f}',ha='center',va='center')
        sa.axis('off')

    def _save_all(self):
        if self.df is None: return
        d = QFileDialog.getExistingDirectory(self, '')
        if not d: return
        wafers = list(self.selected_wafers) if self.selected_wafers else list(self.df.iloc[:,0].unique())
        params = list(self.selected_params) if self.selected_params else list(self.df.columns[5:])
        for wafer in wafers:
            for param in params:
                try:
                    fig = Figure(figsize=(self.figsize_spin.value(),self.figsize_spin.value()*0.75))
                    self._draw_map(fig, wafer, param)
                    fig.savefig(os.path.join(d,f'{wafer}_{param}.png'),dpi=200)
                    import matplotlib.pyplot as plt2; plt2.close(fig)
                except Exception as e: print(e)
        QMessageBox.information(self,'Done','Saved to: '+d)
