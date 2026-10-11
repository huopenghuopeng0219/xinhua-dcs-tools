# -*- coding: utf-8 -*-
"""
新华DCS辅助器 V2.2.2
窗口界面版本（PySide6）
"""
import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                QHBoxLayout, QLabel, QPushButton, QRadioButton, 
                                QFileDialog, QProgressBar, QButtonGroup, QMessageBox,
                                QGroupBox)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

# 获取基础路径（支持打包后的exe）
if getattr(sys, 'frozen', False):
    # 打包后的exe
    BASE_PATH = os.path.dirname(sys.executable)
    # 临时解压目录
    TEMP_PATH = sys._MEIPASS
else:
    # 开发环境
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))
    TEMP_PATH = BASE_PATH

sys.path.insert(0, BASE_PATH)
from generate_wiring_table_new import WiringTableGenerator, ConfigGenerator
from material_difference import generate_difference, output_path_for


def get_resource_path(filename):
    """获取资源文件路径（优先从exe同目录找，否则从临时目录找）"""
    # 先检查exe同目录
    path1 = os.path.join(BASE_PATH, filename)
    if os.path.exists(path1):
        return path1
    # 再检查临时目录
    path2 = os.path.join(TEMP_PATH, filename)
    if os.path.exists(path2):
        return path2
    return path1


def output_path_for_mode(file_path, mode):
    """按源文件名保留项目名称和版本，只替换输出类型。"""
    if mode == "difference":
        return str(output_path_for(file_path, BASE_PATH))
    source_label, output_label, extension = {
        "wiring": ("机柜设计", "接线总表", ".xlsx"),
        "config": ("接线总表", "转组态", ".xls"),
    }[mode]
    stem = os.path.splitext(os.path.basename(file_path))[0]
    if source_label in stem:
        stem = stem.replace(source_label, output_label, 1)
    else:
        stem = f"{stem}-{output_label}"
    return os.path.join(BASE_PATH, stem + extension)


class GenerateThread(QThread):
    """后台生成线程"""
    finished = Signal(bool, str, int)
    
    def __init__(self, file_path, output_format, mode="wiring"):
        super().__init__()
        self.file_path = file_path
        self.output_format = output_format
        self.mode = mode
    
    def run(self):
        try:
            if self.mode == "difference":
                target = output_path_for_mode(self.file_path, self.mode)
                output_path, count = generate_difference(self.file_path, target)
                self.finished.emit(True, str(output_path), count)
            elif self.mode == "config":
                # 生成转组态
                output_path = output_path_for_mode(self.file_path, self.mode)
                generator = ConfigGenerator(self.file_path)
                generator.generate(output_path)
                self.finished.emit(True, output_path, 0)
            else:
                # 生成接线总表
                generator = WiringTableGenerator(self.file_path)
                output_path = output_path_for_mode(self.file_path, self.mode)
                count = generator.generate(output_path)
                self.finished.emit(True, output_path, count)
        except Exception as e:
            self.finished.emit(False, str(e), 0)


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("新华DCS辅助器 V2.2.2")
        self.setFixedSize(560, 360)
        self.file_path = None
        self._init_ui()
        self._update_prompt()
    
    def _init_ui(self):
        """初始化界面"""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        # 标题
        title = QLabel("新华DCS辅助器")
        title.setFont(QFont("微软雅黑", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # 文件选择
        file_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray; padding: 5px;")
        file_layout.addWidget(self.file_label, 1)
        
        self.select_btn = QPushButton("选择文件")
        self.select_btn.clicked.connect(self._select_file)
        file_layout.addWidget(self.select_btn)
        layout.addLayout(file_layout)
        
        # 功能选择
        func_group = QGroupBox("选择功能")
        func_layout = QHBoxLayout()
        self.func_group = QButtonGroup()
        
        self.wiring_radio = QRadioButton("生成接线总表")
        self.wiring_radio.setChecked(True)
        self.func_group.addButton(self.wiring_radio, 0)
        self.wiring_radio.toggled.connect(self._update_prompt)
        func_layout.addWidget(self.wiring_radio)
        
        self.config_radio = QRadioButton("生成转组态")
        self.func_group.addButton(self.config_radio, 1)
        self.config_radio.toggled.connect(self._update_prompt)
        func_layout.addWidget(self.config_radio)

        self.difference_radio = QRadioButton("生成差异表")
        self.func_group.addButton(self.difference_radio, 2)
        self.difference_radio.toggled.connect(self._update_prompt)
        func_layout.addWidget(self.difference_radio)
        
        func_group.setLayout(func_layout)
        layout.addWidget(func_group)
        
        # 生成按钮
        self.generate_btn = QPushButton("生成")
        self.generate_btn.setFont(QFont("微软雅黑", 12, QFont.Bold))
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.setEnabled(False)
        self.generate_btn.clicked.connect(self._generate)
        layout.addWidget(self.generate_btn)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)
        
        # 状态
        self.status_label = QLabel("请先选择Excel源文件")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: blue; padding: 5px;")
        layout.addWidget(self.status_label)
        
        # 底部联系信息
        footer = QLabel("发现BUG请联系：霍 15847214219")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("color: gray; font-size: 11px; padding: 8px;")
        layout.addWidget(footer)
    
    def _update_prompt(self):
        """更新提示文字"""
        if self.config_radio.isChecked():
            self.status_label.setText("请选择接线总表excel")
        else:
            self.status_label.setText("请选择机柜设计excel")
        self.status_label.setStyleSheet("color: blue; padding: 5px;")
    
    def _select_file(self):
        """选择文件"""
        if self.config_radio.isChecked():
            # 生成转组态时，提示选择接线总表
            filter_msg = "请选择接线总表"
            file_types = "Excel文件 (*.xlsx);;所有文件 (*.*)"
        else:
            # 生成接线总表时，提示选择卡件排布
            filter_msg = "请选择EXCEL机柜设计"
            file_types = "Excel文件 (*.xlsm *.xlsx *.xls);;所有文件 (*.*)"
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, filter_msg, "", file_types
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.file_label.setStyleSheet("color: black; padding: 5px;")
            self.generate_btn.setEnabled(True)
            self.status_label.setText(f"已选择: {os.path.basename(file_path)}")
            self.status_label.setStyleSheet("color: green; padding: 5px;")
    
    def _generate(self):
        """生成文件"""
        if not self.file_path:
            QMessageBox.warning(self, "提示", "请先选择源文件！")
            return
        
        mode = {0: "wiring", 1: "config", 2: "difference"}[self.func_group.checkedId()]
        if mode == "difference" and not self.file_path.lower().endswith((".xls", ".xlsx", ".xlsm")):
            QMessageBox.warning(self, "提示", "差异表需要 .xls、.xlsx 或 .xlsm 机柜设计文件")
            return
        output_path = output_path_for_mode(self.file_path, mode)
        if os.path.exists(output_path) and QMessageBox.question(
            self, "覆盖文件", f"文件已存在，是否覆盖？\n{output_path}"
        ) != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度
        self.generate_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        
        self.status_label.setText({
            "wiring": "正在生成接线总表，请稍候...",
            "config": "正在生成转组态文件，请稍候...",
            "difference": "正在生成差异表，请稍候...",
        }[mode])
        self.status_label.setStyleSheet("color: orange; padding: 5px;")
        
        # 启动后台线程（固定输出xlsx）
        self.thread = GenerateThread(self.file_path, "xlsx", mode)
        self.thread.finished.connect(self._on_finished)
        self.thread.start()
    
    def _on_finished(self, success, result, count):
        """生成完成"""
        self.progress.setVisible(False)
        self.generate_btn.setEnabled(True)
        
        if success:
            if count > 0:
                self.status_label.setText(f"生成完成！共 {count} 行数据")
            else:
                self.status_label.setText("生成完成！")
            self.status_label.setStyleSheet("color: green; padding: 5px;")
            QMessageBox.information(self, "成功", 
                f"文件已生成！\n\n文件位置: {result}")
        else:
            self.status_label.setText(f"生成失败: {result}")
            self.status_label.setStyleSheet("color: red; padding: 5px;")
            QMessageBox.critical(self, "错误", f"生成失败:\n{result}")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
