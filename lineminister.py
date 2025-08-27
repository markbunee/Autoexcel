import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QLineEdit, QListWidget, QGroupBox, 
                             QTabWidget, QCheckBox, QSpinBox, QComboBox,
                             QMessageBox, QProgressBar, QFormLayout, QSizePolicy)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QRectF, QPointF
from PyQt5.QtGui import QFont, QPainter, QColor, QPen
import pandas as pd
import math
from particleanimation import ParticleAnimation

class WorkerThread(QThread):
    """工作线程类，用于在后台执行耗时操作"""
    output_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs
    
    def run(self):
        try:
            # 为支持输出回调的函数添加output_callback参数
            if self.function.__name__ in ['classify_files', 'classify_files_by_keywords', 'classify_files_by_extension', 
                                         'merge_excel_files_by_column', 'split_excel_by_column', 'rename_files_sequentially',
                                         'clean_excel_data', 'process_indication_standardization', 'update_file_comparison',
                                         'process_excel']:
                self.kwargs['output_callback'] = self._output_callback
            
            result = self.function(*self.args, **self.kwargs)
            if result:
                self.finished_signal.emit(True, "操作成功完成")
            else:
                self.finished_signal.emit(False, "操作失败")
        except Exception as e:
            self.finished_signal.emit(False, f"操作出错: {str(e)}")
    
    def _output_callback(self, message):
        """输出回调函数，将消息发送到GUI"""
        self.output_signal.emit(message)
