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
# 导入各个功能模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from file_classify import classify_files, classify_files_by_keywords, classify_files_by_extension
from file_Merge import merge_excel_files_by_column
from file_Splitting import split_excel_by_column
from file_rename import rename_files_sequentially
from file_clean import clean_excel_data
from menet_file_normalize import process_indication_standardization
from menet_update import update_file_comparison
from file_Mulc_sim_match import process_excel
from lineminister import WorkerThread


class AutoExcelGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle('AutoExcel')
        self.setGeometry(100, 100, 1100, 750)
        
        # 设置全局字体
        font = QFont("微软雅黑", 9)
        self.setFont(font)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # 创建标签页
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
            QTabBar::tab {
                background: #F0F0F0;
                border: 1px solid #CCCCCC;
                border-bottom-color: #CCCCCC;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 8ex;
                padding: 8px;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                border-bottom-color: #FFFFFF;
            }
        """)
        main_layout.addWidget(tab_widget)
        
        # 主页标签页
        self.home_tab = self.create_home_tab()
        tab_widget.addTab(self.home_tab, "主页")
        
        # 文件分类标签页
        self.classify_tab = self.create_classify_tab()
        tab_widget.addTab(self.classify_tab, "文件分类")
        
        # 文件合并标签页
        self.merge_tab = self.create_merge_tab()
        tab_widget.addTab(self.merge_tab, "合并报表")
        
        # 文件拆分标签页
        self.split_tab = self.create_split_tab()
        tab_widget.addTab(self.split_tab, "拆分报表")
        
        # 文件重命名标签页
        self.rename_tab = self.create_rename_tab()
        tab_widget.addTab(self.rename_tab, "批量重命名")
        
        # 数据清洗标签页
        self.clean_tab = self.create_clean_tab()
        tab_widget.addTab(self.clean_tab, "数据清洗")
        
        # Menet文件标准化标签页
        self.normalize_tab = self.create_normalize_tab()
        tab_widget.addTab(self.normalize_tab, "适应症写法规范化")
        
        # Menet文件对比标签页
        self.compare_tab = self.create_compare_tab()
        tab_widget.addTab(self.compare_tab, "一致性评价进度月更新")
        
        # 关于作者标签页
        self.about_tab = self.create_about_tab()
        tab_widget.addTab(self.about_tab, "关于作者")
        
        # 创建输出区域
        output_group = QGroupBox("输出日志")
        output_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        output_layout = QVBoxLayout()
        output_layout.setSpacing(5)
        output_layout.setContentsMargins(10, 10, 10, 10)
        
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                background-color: #FFFFFF;
            }
        """)
        self.output_text.setMinimumHeight(120)
        output_layout.addWidget(self.output_text)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 20px;
            }
        """)
        output_layout.addWidget(self.progress_bar)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # 设置状态栏
        self.statusBar().showMessage('就绪')
        
    def create_home_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 创建动画背景
        self.animation_widget = ParticleAnimation()
        self.animation_widget.setMinimumSize(800, 480)
        layout.addWidget(self.animation_widget)
        
        # 创建标题标签
        title_label = QLabel("AutoExcel")
        title_label.setStyleSheet("""
            QLabel {
                color: #1685a9;
                font-size: 100px;
                font-weight: bold;
                qproperty-alignment: AlignCenter;
            }
        """)
        title_label.setGeometry(0, 200, 1050, 300)
        title_label.setParent(self.animation_widget)
        
        # 创建副标题
        subtitle_label = QLabel("Excel文件自动化处理工具")
        subtitle_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 25px;
                font-weight: bold;
                qproperty-alignment: AlignCenter;
            }
        """)
        subtitle_label.setGeometry(0, 270, 990, 340)
        subtitle_label.setParent(self.animation_widget)
        
        tab.setLayout(layout)
        return tab
        
    def create_about_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(50, 50, 50, 50)
        
        # 标题
        title_label = QLabel("关于作者")
        title_label.setStyleSheet("""
            QLabel {
                color: #2196F3;
                font-size: 32px;
                font-weight: bold;
                qproperty-alignment: AlignCenter;
            }
        """)
        layout.addWidget(title_label)
        
        # 信息展示区域
        info_group = QGroupBox()
        info_group.setStyleSheet("""
            QGroupBox {
                border: none;
                background-color: #F9F9F9;
                border-radius: 10px;
            }
        """)
        info_layout = QVBoxLayout()
        info_layout.setSpacing(15)
        info_layout.setContentsMargins(30, 30, 30, 30)
        # 0824
        # 版本号
        version_label = QLabel("版本号：autoexcel_20250826")
        version_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #333333;
            }
        """)
        info_layout.addWidget(version_label)
        
        # 作者
        author_label = QLabel("作者：白怿")
        author_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #333333;
            }
        """)
        info_layout.addWidget(author_label)
        
        # 邮箱
        email_label = QLabel("邮箱：3195824330@qq.com")
        email_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #333333;
            }
        """)
        info_layout.addWidget(email_label)
        
        # GitHub
        github_label = QLabel("web：<a href='https://markbunee.github.io/'>https://markbunee.github.io/</a>")
        github_label.setStyleSheet("""
            QLabel {
                font-size: 17px;
                color: #333333;
            }
        """)
        github_label.setOpenExternalLinks(True)
        info_layout.addWidget(github_label)

        #其他
        content_label = QLabel("持续更新" \
        " 未来新增数据分析与挖掘其他功能" \
        " 新增更多图标" \
        " 新增自动辅助办公" \
        " 新增小型人工智能辅助")
        content_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                color: #333333;
            }
        """)
        info_layout.addWidget(content_label)
        
        # 添加垂直弹性空间
        info_layout.addStretch()
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 添加垂直弹性空间
        layout.addStretch()
        
        tab.setLayout(layout)
        return tab
        
    def create_classify_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 源目录选择
        source_layout = QHBoxLayout()
        source_layout.setSpacing(10)
        source_layout.addWidget(QLabel("源目录:"), 0)
        self.classify_source_edit = QLineEdit()
        self.classify_source_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        source_layout.addWidget(self.classify_source_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_classify_source)
        source_layout.addWidget(browse_btn)
        layout.addLayout(source_layout)
        
        # 目标目录选择
        target_layout = QHBoxLayout()
        target_layout.setSpacing(10)
        target_layout.addWidget(QLabel("目标目录:"), 0)
        self.classify_target_edit = QLineEdit()
        self.classify_target_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        target_layout.addWidget(self.classify_target_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_classify_target)
        target_layout.addWidget(browse_btn)
        layout.addLayout(target_layout)
        
        # 分类方式选择
        method_layout = QHBoxLayout()
        method_layout.setSpacing(10)
        method_layout.addWidget(QLabel("分类方式:"), 0)
        self.classify_method_combo = QComboBox()
        self.classify_method_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                border-radius: 4px;
            }
        """)
        self.classify_method_combo.addItem("按文件类型分类")
        self.classify_method_combo.addItem("按关键词分类")
        self.classify_method_combo.addItem("按文件扩展名分类")
        method_layout.addWidget(self.classify_method_combo, 1)
        layout.addLayout(method_layout)
        
        # 文件类型选择（仅在按文件类型分类时启用）
        self.type_layout = QHBoxLayout()
        self.type_layout.setSpacing(10)
        self.type_layout.addWidget(QLabel("文件类型:"), 0)
        self.classify_types_edit = QLineEdit()
        self.classify_types_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        self.classify_types_edit.setPlaceholderText("输入文件类型，如: word excel pdf (留空为全部)")
        self.type_layout.addWidget(self.classify_types_edit, 1)
        layout.addLayout(self.type_layout)
        
        # 关键词选择（仅在按关键词分类时启用）
        self.keyword_layout = QHBoxLayout()
        self.keyword_layout.setSpacing(10)
        self.keyword_layout.addWidget(QLabel("关键词:"), 0)
        self.classify_keywords_edit = QLineEdit()
        self.classify_keywords_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        self.classify_keywords_edit.setPlaceholderText("输入关键词，如: 广东,浙江,上海 (以英文逗号分隔)")
        self.keyword_layout.addWidget(self.classify_keywords_edit, 1)
        layout.addLayout(self.keyword_layout)
        
        # 连接分类方式选择变化信号
        self.classify_method_combo.currentIndexChanged.connect(self.on_classify_method_changed)
        
        # 执行按钮
        execute_btn = QPushButton("执行文件分类")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_classify)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 添加用法说明
        usage_group = QGroupBox("用法说明")
        usage_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(5)
        usage_layout.setContentsMargins(10, 10, 10, 10)
        
        usage_text = QTextEdit()
        usage_text.setReadOnly(True)
        usage_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #FAFAFA;
            }
        """)
        usage_text.setMaximumHeight(150)
        usage_text.setHtml("""
        <b>功能说明：</b><br>
        将指定目录下的文件按类型分类到不同的文件夹中，支持按文件类型、关键词和文件扩展名三种模式分类<br><br>
        
        <b>输入：</b><br>
        - 源目录：包含需要分类的文件的目录<br>
        - 目标目录：分类后文件存放的目录<br>
        - 分类方式：选择一种分类方式（按文件类型、按关键词、按文件扩展名）<br>
        - 文件类型：指定要分类的文件类型，如word、excel、pdf等，留空则分类所有类型<br>
        - 关键词：按文件名中包含的关键词进行分类，每个关键词创建一个独立文件夹<br><br>
        
        <b>输出：</b><br>
        在目标目录下创建不同类型的文件夹，并将文件移动到对应文件夹中<br><br>
        
        <b>分类方式说明：</b><br>
        1. <b>按文件类型分类</b>：根据文件扩展名将文件分类到对应类型文件夹中<br>
        2. <b>按关键词分类</b>：根据文件名中包含的关键词进行分类，创建keyword/{关键词}文件夹<br>
        3. <b>按文件扩展名分类</b>：根据文件扩展名创建文件夹，如.png、.json等<br><br>
        
        <b>示例：</b><br>
        源目录：D:/报告文件<br>
        目标目录：D:/分类结果<br>
        分类方式：按关键词分类<br>
        关键词：广东,浙江,上海<br>
        执行后，将在D:/分类结果下创建对应文件夹：<br>
        - keyword/广东、keyword/浙江、keyword/上海（按关键词分类）<br>
        - other（未匹配的文件）<br><br>
        """)
        usage_layout.addWidget(usage_text)
        usage_group.setLayout(usage_layout)
        layout.addWidget(usage_group)
        
        tab.setLayout(layout)
        
        # 初始化界面状态
        self.on_classify_method_changed(0)
        
        return tab
        
    def create_merge_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 文件列表
        file_list_layout = QHBoxLayout()
        file_list_layout.setSpacing(10)
        file_list_layout.addWidget(QLabel("Excel文件列表:"), 0)
        self.merge_file_list = QListWidget()
        self.merge_file_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                background-color: #FFFFFF;
            }
        """)
        self.merge_file_list.setMinimumHeight(100)
        file_list_layout.addWidget(self.merge_file_list, 1)
        layout.addLayout(file_list_layout)
        
        # 添加和删除按钮
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        add_btn = QPushButton("添加文件")
        add_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        add_btn.clicked.connect(self.add_merge_files)
        button_layout.addWidget(add_btn)
        
        remove_btn = QPushButton("删除选中")
        remove_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        remove_btn.clicked.connect(self.remove_merge_files)
        button_layout.addWidget(remove_btn)
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # 匹配列设置
        match_layout = QHBoxLayout()
        match_layout.setSpacing(10)
        match_layout.addWidget(QLabel("匹配列序号 (以空格分隔):"), 0)
        self.merge_match_edit = QLineEdit()
        self.merge_match_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        self.merge_match_edit.setPlaceholderText("例如: 1 2 1")
        match_layout.addWidget(self.merge_match_edit, 1)
        layout.addLayout(match_layout)
        
        # 输出文件
        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)
        output_layout.addWidget(QLabel("输出文件:"), 0)
        self.merge_output_edit = QLineEdit()
        self.merge_output_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        self.merge_output_edit.setPlaceholderText("输出文件路径")
        output_layout.addWidget(self.merge_output_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_merge_output)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)
        
        # 执行按钮
        execute_btn = QPushButton("执行合并")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_merge)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 添加用法说明
        usage_group = QGroupBox("用法说明")
        usage_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(5)
        usage_layout.setContentsMargins(10, 10, 10, 10)
        
        usage_text = QTextEdit()
        usage_text.setReadOnly(True)
        usage_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #FAFAFA;
            }
        """)
        usage_text.setMaximumHeight(150)
        usage_text.setHtml("""
        <b>功能说明：</b><br>
        根据指定列将多个Excel文件合并成一个文件<br><br>
        
        <b>输入：</b><br>
        - Excel文件列表：需要合并的Excel文件<br>
        - 匹配列序号：每个文件用于匹配的列序号（从1开始）<br>
        - 输出文件：合并后文件的保存路径<br><br>
        
        <b>输出：</b><br>
        一个包含所有输入文件数据的合并Excel文件<br><br>
        
        <b>示例：</b><br>
        文件1：2023年广东省经济报告.xlsx（企业名称在第1列）<br>
        文件2：2023年浙江省经济报告.xlsx（企业名称在第2列）<br>
        匹配列序号：1 2<br>
        输出文件：D:/合并结果/2023年经济报告合并.xlsx<br>
        执行后，将根据企业名称匹配两个文件的数据并合并到一个文件中
        """)
        usage_layout.addWidget(usage_text)
        usage_group.setLayout(usage_layout)
        layout.addWidget(usage_group)
        
        tab.setLayout(layout)
        return tab
        
    def create_split_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 源文件选择
        source_layout = QHBoxLayout()
        source_layout.setSpacing(10)
        source_layout.addWidget(QLabel("源Excel文件:"), 0)
        self.split_source_edit = QLineEdit()
        self.split_source_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        source_layout.addWidget(self.split_source_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_split_source)
        source_layout.addWidget(browse_btn)
        layout.addLayout(source_layout)
        
        # 拆分模式选择
        mode_layout = QHBoxLayout()
        mode_layout.setSpacing(10)
        mode_layout.addWidget(QLabel("拆分模式:"), 0)
        self.split_mode_combo = QComboBox()
        self.split_mode_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                border-radius: 4px;
            }
        """)
        self.split_mode_combo.addItem("按行拆分（根据列值）")
        self.split_mode_combo.addItem("按列拆分（选择列组合）")
        self.split_mode_combo.currentIndexChanged.connect(self.on_split_mode_changed)
        mode_layout.addWidget(self.split_mode_combo, 1)
        layout.addLayout(mode_layout)
        
        # 分割列设置（按行拆分时使用）
        self.split_column_layout = QHBoxLayout()
        self.split_column_layout.setSpacing(10)
        self.split_column_layout.addWidget(QLabel("分割列序号:"), 0)
        self.split_column_spin = QSpinBox()
        self.split_column_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.split_column_spin.setMinimum(1)
        self.split_column_spin.setMaximum(100)
        self.split_column_layout.addWidget(self.split_column_spin, 1)
        layout.addLayout(self.split_column_layout)
        
        # 输出列设置（按列拆分时使用）
        self.output_columns_layout = QHBoxLayout()
        self.output_columns_layout.setSpacing(10)
        self.output_columns_layout.addWidget(QLabel("输出列组合 (例如: 1,2 1,3 1,4,5):"), 0)
        self.split_output_edit = QLineEdit()
        self.split_output_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        self.output_columns_layout.addWidget(self.split_output_edit, 1)
        layout.addLayout(self.output_columns_layout)
        
        # 输出目录
        output_dir_layout = QHBoxLayout()
        output_dir_layout.setSpacing(10)
        output_dir_layout.addWidget(QLabel("输出目录:"), 0)
        self.split_output_dir_edit = QLineEdit()
        self.split_output_dir_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        output_dir_layout.addWidget(self.split_output_dir_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_split_output_dir)
        output_dir_layout.addWidget(browse_btn)
        layout.addLayout(output_dir_layout)
        
        # 执行按钮
        execute_btn = QPushButton("执行拆分")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_split)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 添加用法说明
        usage_group = QGroupBox("用法说明")
        usage_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(5)
        usage_layout.setContentsMargins(10, 10, 10, 10)
        
        usage_text = QTextEdit()
        usage_text.setReadOnly(True)
        usage_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #FAFAFA;
            }
        """)
        usage_text.setMaximumHeight(150)
        usage_text.setHtml("""
        <b>功能说明：</b><br>
        根据指定条件将一个Excel文件拆分成多个文件<br><br>
        
        <b>拆分模式说明：</b><br>
        1. <b>按行拆分</b>：根据某一列的不同值将文件拆分成多个文件，每个文件包含该列值相同的行<br>
        2. <b>按列拆分</b>：选择特定的列组合来创建不同的文件，每个文件包含指定的列<br><br>
        
        <b>输入：</b><br>
        - 源Excel文件：需要拆分的Excel文件<br>
        - 拆分模式：选择按行拆分或按列拆分<br>
        - 分割列序号：用于分割的列序号（从1开始）<br>
        - 输出列组合：每个拆分文件包含的列，如"1,2 1,3 1,4,5"<br>
        - 输出目录：拆分后文件的保存目录<br><br>
        
        <b>输出：</b><br>
        多个Excel文件，每个文件包含按指定条件分割的数据<br><br>
        
        <b>示例：</b><br>
        源文件：2023年全国经济报告.xlsx<br>
        拆分模式：按行拆分<br>
        分割列序号：1（企业所在省份）<br>
        输出目录：D:/拆分结果<br>
        执行后，将按省份拆分文件，每个省份一个文件，包含所有列数据<br><br>
        
        源文件：2023年全国经济报告.xlsx<br>
        拆分模式：按列拆分<br>
        输出列组合：1,2,3 1,4 1,5,6（分别生成包含不同列组合的文件）<br>
        输出目录：D:/拆分结果<br>
        执行后，将生成多个文件，每个文件包含指定的列组合
        """)
        usage_layout.addWidget(usage_text)
        usage_group.setLayout(usage_layout)
        layout.addWidget(usage_group)
        
        tab.setLayout(layout)
        
        # 初始化界面状态
        self.on_split_mode_changed(0)
        
        return tab
        
    def create_rename_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 源目录选择
        source_layout = QHBoxLayout()
        source_layout.setSpacing(10)
        source_layout.addWidget(QLabel("源目录:"), 0)
        self.rename_source_edit = QLineEdit()
        self.rename_source_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        source_layout.addWidget(self.rename_source_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_rename_source)
        source_layout.addWidget(browse_btn)
        layout.addLayout(source_layout)
        
        # 前缀设置
        prefix_layout = QHBoxLayout()
        prefix_layout.setSpacing(10)
        prefix_layout.addWidget(QLabel("文件名前缀:"), 0)
        self.rename_prefix_edit = QLineEdit()
        self.rename_prefix_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        prefix_layout.addWidget(self.rename_prefix_edit, 1)
        layout.addLayout(prefix_layout)
        
        # 起始编号和位数
        number_layout = QHBoxLayout()
        number_layout.setSpacing(10)
        number_layout.addWidget(QLabel("起始编号:"), 0)
        self.rename_start_spin = QSpinBox()
        self.rename_start_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.rename_start_spin.setMinimum(1)
        self.rename_start_spin.setValue(1)
        number_layout.addWidget(self.rename_start_spin, 1)
        
        number_layout.addWidget(QLabel("编号位数:"), 0)
        self.rename_digits_spin = QSpinBox()
        self.rename_digits_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.rename_digits_spin.setMinimum(1)
        self.rename_digits_spin.setMaximum(10)
        self.rename_digits_spin.setValue(4)
        number_layout.addWidget(self.rename_digits_spin, 1)
        layout.addLayout(number_layout)
        
        # 关键词设置
        keyword_layout = QHBoxLayout()
        keyword_layout.setSpacing(10)
        keyword_layout.addWidget(QLabel("关键词:"), 0)
        self.rename_keyword_edit = QLineEdit()
        self.rename_keyword_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        keyword_layout.addWidget(self.rename_keyword_edit, 1)
        layout.addLayout(keyword_layout)
        
        # 执行按钮
        execute_btn = QPushButton("执行重命名")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_rename)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 添加用法说明
        usage_group = QGroupBox("用法说明")
        usage_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(5)
        usage_layout.setContentsMargins(10, 10, 10, 10)
        
        usage_text = QTextEdit()
        usage_text.setReadOnly(True)
        usage_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #FAFAFA;
            }
        """)
        usage_text.setMaximumHeight(150)
        usage_text.setHtml("""
        <b>功能说明：</b><br>
        对指定目录下的文件进行批量重命名<br><br>
        
        <b>输入：</b><br>
        - 源目录：包含需要重命名文件的目录<br>
        - 文件名前缀：重命名后的文件前缀<br>
        - 起始编号：重命名开始的编号<br>
        - 编号位数：编号的位数，如4位数就是0001, 0002...<br>
        - 关键词：添加在序号后的关键词<br><br>
        
        <b>输出：</b><br>
        按照指定规则重命名后的文件<br><br>
        
        <b>示例：</b><br>
        源目录：D:/经济报告<br>
        文件名前缀：经济报告_<br>
        起始编号：1<br>
        编号位数：4<br>
        关键词：广东省<br>
        执行后，文件将被重命名为：经济报告_0001_广东省.xlsx、经济报告_0002_广东省.xlsx...
        """)
        usage_layout.addWidget(usage_text)
        usage_group.setLayout(usage_layout)
        layout.addWidget(usage_group)
        
        tab.setLayout(layout)
        return tab
        
    def create_clean_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 源文件选择
        source_layout = QHBoxLayout()
        source_layout.setSpacing(10)
        source_layout.addWidget(QLabel("源Excel文件:"), 0)
        self.clean_source_edit = QLineEdit()
        self.clean_source_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        source_layout.addWidget(self.clean_source_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_clean_source)
        source_layout.addWidget(browse_btn)
        layout.addLayout(source_layout)
        
        # 输出文件
        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)
        output_layout.addWidget(QLabel("输出文件:"), 0)
        self.clean_output_edit = QLineEdit()
        self.clean_output_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        output_layout.addWidget(self.clean_output_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_clean_output)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)
        
        # 清洗选项
        options_group = QGroupBox("清洗选项")
        options_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        options_layout = QVBoxLayout()
        options_layout.setSpacing(10)
        options_layout.setContentsMargins(10, 10, 10, 10)
        
        self.clean_symbols_check = QCheckBox("清除指定符号")
        self.clean_symbols_check.setChecked(True)
        options_layout.addWidget(self.clean_symbols_check)
        
        self.symbols_edit = QLineEdit()
        self.symbols_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        self.symbols_edit.setPlaceholderText("输入要清除的符号，以空格分隔，例如: * # @")
        options_layout.addWidget(self.symbols_edit)
        
        self.mark_empty_check = QCheckBox("标记空值单元格为黄色")
        self.mark_empty_check.setChecked(True)
        options_layout.addWidget(self.mark_empty_check)
        
        self.mark_duplicates_check = QCheckBox("标记重复行整行蓝色")
        self.mark_duplicates_check.setChecked(True)
        options_layout.addWidget(self.mark_duplicates_check)
        
        self.clean_internal_spaces_check = QCheckBox("清除单元格内部空格和回车")
        self.clean_internal_spaces_check.setChecked(False)
        options_layout.addWidget(self.clean_internal_spaces_check)
        
        # 中英文空格和符号差异处理选项
        self.clean_chinese_space_check = QCheckBox("同时清除中文全角空格")
        self.clean_chinese_space_check.setChecked(False)
        self.clean_chinese_space_check.setEnabled(False)  # 默认禁用，只有当清除内部空格选项启用时才可用
        options_layout.addWidget(self.clean_chinese_space_check)
        
        self.clean_english_punctuation_check = QCheckBox("处理英文标点符号差异")
        self.clean_english_punctuation_check.setChecked(False)
        self.clean_english_punctuation_check.setEnabled(False)  # 默认禁用，只有当清除内部空格选项启用时才可用
        options_layout.addWidget(self.clean_english_punctuation_check)
        
        # 连接信号槽，使子选项在主选项启用时才可用
        self.clean_internal_spaces_check.stateChanged.connect(self.on_clean_internal_spaces_changed)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # 执行按钮
        execute_btn = QPushButton("执行数据清洗")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_clean)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 添加用法说明
        usage_group = QGroupBox("用法说明")
        usage_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(5)
        usage_layout.setContentsMargins(10, 10, 10, 10)
        
        usage_text = QTextEdit()
        usage_text.setReadOnly(True)
        usage_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #FAFAFA;
            }
        """)
        usage_text.setMaximumHeight(150)
        usage_text.setHtml("""
        <b>功能说明：</b><br>
        对Excel文件进行数据清洗，标记空值和重复行，清除指定符号<br><br>
        
        <b>输入：</b><br>
        - 源Excel文件：需要清洗的Excel文件<br>
        - 输出文件：清洗后文件的保存路径<br>
        - 清洗选项：选择需要执行的清洗操作<br><br>
        
        <b>输出：</b><br>
        清洗后的Excel文件，包含标记和清除操作的结果<br><br>
        
        <b>示例：</b><br>
        源文件：原始经济报告.xlsx<br>
        输出文件：D:/清洗结果/清洗后经济报告.xlsx<br>
        清洗选项：标记空值单元格为黄色、标记重复行整行蓝色、清除指定符号(*)<br>
        执行后，将生成一个清洗后的文件，空值单元格标记为黄色，重复行标记为蓝色，*号被清除
        """)
        usage_layout.addWidget(usage_text)
        usage_group.setLayout(usage_layout)
        layout.addWidget(usage_group)
        
        tab.setLayout(layout)
        return tab
        
    def create_normalize_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 输入文件选择
        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)
        input_layout.addWidget(QLabel("输入Excel文件:"), 0)
        self.normalize_input_edit = QLineEdit()
        self.normalize_input_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        input_layout.addWidget(self.normalize_input_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_normalize_input)
        input_layout.addWidget(browse_btn)
        layout.addLayout(input_layout)
        
        # 输出目录选择
        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)
        output_layout.addWidget(QLabel("输出目录:"), 0)
        self.normalize_output_edit = QLineEdit()
        self.normalize_output_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        output_layout.addWidget(self.normalize_output_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_normalize_output)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)
        
        # 列索引设置
        column_layout = QHBoxLayout()
        column_layout.setSpacing(10)
        column_layout.addWidget(QLabel("需归一化的列索引:"), 0)
        self.normalize_column_spin = QSpinBox()
        self.normalize_column_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.normalize_column_spin.setMinimum(0)
        self.normalize_column_spin.setMaximum(100)
        self.normalize_column_spin.setValue(3)  # 默认第四列
        column_layout.addWidget(self.normalize_column_spin, 1)
        layout.addLayout(column_layout)
        
        # 分组列索引设置
        group_layout = QHBoxLayout()
        group_layout.setSpacing(10)
        group_layout.addWidget(QLabel("分组列索引:"), 0)
        self.normalize_group_spin = QSpinBox()
        self.normalize_group_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.normalize_group_spin.setMinimum(0)
        self.normalize_group_spin.setMaximum(100)
        self.normalize_group_spin.setValue(1)  # 默认第二列
        group_layout.addWidget(self.normalize_group_spin, 1)
        layout.addLayout(group_layout)
        
        # 阈值设置
        threshold_group = QGroupBox("阈值设置")
        threshold_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        threshold_layout = QFormLayout()
        threshold_layout.setSpacing(10)
        threshold_layout.setContentsMargins(10, 10, 10, 10)
        
        threshold_layout.addRow(QLabel("相似度阈值 (0-100):"))
        self.normalize_similarity_spin = QSpinBox()
        self.normalize_similarity_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.normalize_similarity_spin.setMinimum(0)
        self.normalize_similarity_spin.setMaximum(100)
        self.normalize_similarity_spin.setValue(85)
        threshold_layout.addRow(self.normalize_similarity_spin)
        
        threshold_layout.addRow(QLabel("编辑距离阈值:"))
        self.normalize_edit_distance_spin = QSpinBox()
        self.normalize_edit_distance_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.normalize_edit_distance_spin.setMinimum(0)
        self.normalize_edit_distance_spin.setMaximum(20)
        self.normalize_edit_distance_spin.setValue(3)
        threshold_layout.addRow(self.normalize_edit_distance_spin)
        
        threshold_layout.addRow(QLabel("最小文本长度:"))
        self.normalize_min_length_spin = QSpinBox()
        self.normalize_min_length_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.normalize_min_length_spin.setMinimum(1)
        self.normalize_min_length_spin.setMaximum(20)
        self.normalize_min_length_spin.setValue(4)
        threshold_layout.addRow(self.normalize_min_length_spin)
        
        threshold_group.setLayout(threshold_layout)
        layout.addWidget(threshold_group)
        
        # 执行按钮
        execute_btn = QPushButton("执行文件内容标准化")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_normalize)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 添加用法说明
        usage_group = QGroupBox("用法说明")
        usage_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(5)
        usage_layout.setContentsMargins(10, 10, 10, 10)
        
        usage_text = QTextEdit()
        usage_text.setReadOnly(True)
        usage_text.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #FAFAFA;
            }
        """)
        usage_text.setMaximumHeight(150)
        usage_text.setHtml("""
        <b>功能说明：</b><br>
        对Excel文件内容进行文本相似度归一化处理<br><br>
        
        <b>输入：</b><br>
        - 输入Excel文件：需要标准化的Excel文件<br>
        - 输出目录：处理结果文件的保存目录<br>
        - 需归一化的列索引：需要进行标准化处理的列（从0开始）<br>
        - 分组列索引：用于分组的列（从0开始）<br><br>
        
        <b>输出：</b><br>
        标准化后的Excel文件，相似文本被归一化为统一格式<br><br>
        
        <b>示例：</b><br>
        输入文件：D:/经济报告/原始经济报告.xlsx<br>
        输出目录：D:/标准化结果<br>
        需归一化的列索引：3（企业名称列）<br>
        分组列索引：1（省份列）<br>
        执行后，将在同一省份内对企业名称进行标准化，如"广东电力公司"和"广东电力"将被归一化
        """)
        usage_layout.addWidget(usage_text)
        usage_group.setLayout(usage_layout)
        layout.addWidget(usage_group)
        
        tab.setLayout(layout)
        return tab
     
    def create_compare_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # 文件1选择（历史文件）
        file1_layout = QHBoxLayout()
        file1_layout.setSpacing(10)
        file1_layout.addWidget(QLabel("历史文件:"), 0)
        self.compare_file1_edit = QLineEdit()
        self.compare_file1_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        file1_layout.addWidget(self.compare_file1_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_compare_file1)
        file1_layout.addWidget(browse_btn)
        layout.addLayout(file1_layout)
        
        # 文件2选择（当前文件）
        file2_layout = QHBoxLayout()
        file2_layout.setSpacing(10)
        file2_layout.addWidget(QLabel("当前文件:"), 0)
        self.compare_file2_edit = QLineEdit()
        self.compare_file2_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        file2_layout.addWidget(self.compare_file2_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_compare_file2)
        file2_layout.addWidget(browse_btn)
        layout.addLayout(file2_layout)
        
        # 输出文件选择
        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)
        output_layout.addWidget(QLabel("输出文件:"), 0)
        self.compare_output_edit = QLineEdit()
        self.compare_output_edit.setStyleSheet("QLineEdit { padding: 5px; border: 1px solid #CCCCCC; border-radius: 4px; }")
        output_layout.addWidget(self.compare_output_edit, 1)
        browse_btn = QPushButton("浏览")
        browse_btn.setStyleSheet("QPushButton { padding: 5px 15px; }")
        browse_btn.clicked.connect(self.browse_compare_output)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)
        
        # 阈值设置
        threshold_group = QGroupBox("阈值设置")
        threshold_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        threshold_layout = QFormLayout()
        threshold_layout.setSpacing(10)
        threshold_layout.setContentsMargins(10, 10, 10, 10)
        
        threshold_layout.addRow(QLabel("名称相似度阈值 (0-100):"))
        self.compare_name_similarity_spin = QSpinBox()
        self.compare_name_similarity_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_name_similarity_spin.setMinimum(0)
        self.compare_name_similarity_spin.setMaximum(100)
        self.compare_name_similarity_spin.setValue(80)
        threshold_layout.addRow(self.compare_name_similarity_spin)
        
        threshold_layout.addRow(QLabel("文本相似度阈值 (0-1.0):"))
        self.compare_text_similarity_spin = QSpinBox()
        self.compare_text_similarity_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_text_similarity_spin.setMinimum(0)
        self.compare_text_similarity_spin.setMaximum(100)
        self.compare_text_similarity_spin.setValue(40)  # 0.4 * 100
        self.compare_text_similarity_spin.setSingleStep(5)
        threshold_layout.addRow(self.compare_text_similarity_spin)
        
        threshold_group.setLayout(threshold_layout)
        layout.addWidget(threshold_group)
        
        # 列索引设置
        column_group = QGroupBox("列索引设置")
        column_group.setStyleSheet("""
            QGroupBox {
                border: 1px solid #CCCCCC;
                border-radius: 4px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
            }
        """)
        column_layout = QFormLayout()
        column_layout.setSpacing(10)
        column_layout.setContentsMargins(10, 10, 10, 10)
        
        # 文件1列索引
        column_layout.addRow(QLabel("文件1 - 药品名称列索引:"))
        self.compare_file1_drug_spin = QSpinBox()
        self.compare_file1_drug_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file1_drug_spin.setMinimum(0)
        self.compare_file1_drug_spin.setMaximum(100)
        self.compare_file1_drug_spin.setValue(1)
        column_layout.addRow(self.compare_file1_drug_spin)
        
        column_layout.addRow(QLabel("文件1 - 企业名称列索引:"))
        self.compare_file1_company_spin = QSpinBox()
        self.compare_file1_company_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file1_company_spin.setMinimum(0)
        self.compare_file1_company_spin.setMaximum(100)
        self.compare_file1_company_spin.setValue(4)
        column_layout.addRow(self.compare_file1_company_spin)
        
        column_layout.addRow(QLabel("文件1 - 状态列索引:"))
        self.compare_file1_status_spin = QSpinBox()
        self.compare_file1_status_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file1_status_spin.setMinimum(0)
        self.compare_file1_status_spin.setMaximum(100)
        self.compare_file1_status_spin.setValue(7)
        column_layout.addRow(self.compare_file1_status_spin)
        
        column_layout.addRow(QLabel("文件1 - 内容列索引:"))
        self.compare_file1_content_spin = QSpinBox()
        self.compare_file1_content_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file1_content_spin.setMinimum(0)
        self.compare_file1_content_spin.setMaximum(100)
        self.compare_file1_content_spin.setValue(23)
        column_layout.addRow(self.compare_file1_content_spin)
        
        # 文件2列索引
        column_layout.addRow(QLabel("文件2 - 药品名称列索引:"))
        self.compare_file2_drug_spin = QSpinBox()
        self.compare_file2_drug_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file2_drug_spin.setMinimum(0)
        self.compare_file2_drug_spin.setMaximum(100)
        self.compare_file2_drug_spin.setValue(0)
        column_layout.addRow(self.compare_file2_drug_spin)
        
        column_layout.addRow(QLabel("文件2 - 企业名称列索引:"))
        self.compare_file2_company_spin = QSpinBox()
        self.compare_file2_company_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file2_company_spin.setMinimum(0)
        self.compare_file2_company_spin.setMaximum(100)
        self.compare_file2_company_spin.setValue(3)
        column_layout.addRow(self.compare_file2_company_spin)
        
        column_layout.addRow(QLabel("文件2 - 状态列索引:"))
        self.compare_file2_status_spin = QSpinBox()
        self.compare_file2_status_spin.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #CCCCCC;
                border-radius: 4px;
            }
        """)
        self.compare_file2_status_spin.setMinimum(0)
        self.compare_file2_status_spin.setMaximum(100)
        self.compare_file2_status_spin.setValue(6)
        column_layout.addRow(self.compare_file2_status_spin)
        
        column_group.setLayout(column_layout)
        layout.addWidget(column_group)
        
        # 执行按钮
        execute_btn = QPushButton("执行文件对比分析")
        execute_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        execute_btn.clicked.connect(self.execute_compare)
        layout.addWidget(execute_btn)
        
        # 添加弹性空间
        layout.addStretch()
        
        tab.setLayout(layout)
        return tab
        
    # 文件浏览函数
    def browse_classify_source(self):
        directory = QFileDialog.getExistingDirectory(self, "选择源目录")
        if directory:
            self.classify_source_edit.setText(directory)
            
    def browse_classify_target(self):
        directory = QFileDialog.getExistingDirectory(self, "选择目标目录")
        if directory:
            self.classify_target_edit.setText(directory)
            
    def on_classify_method_changed(self, index):
        """当分类方式改变时，显示相应的配置项"""
        # 根据选择的分类方式显示或隐藏相应配置项
        if index == 0:  # 按文件类型分类
            self.type_layout.itemAt(0).widget().show()  # QLabel
            self.type_layout.itemAt(1).widget().show()  # QLineEdit
        else:
            self.type_layout.itemAt(0).widget().hide()  # QLabel
            self.type_layout.itemAt(1).widget().hide()  # QLineEdit
            
        if index == 1:  # 按关键词分类
            self.keyword_layout.itemAt(0).widget().show()  # QLabel
            self.keyword_layout.itemAt(1).widget().show()  # QLineEdit
        else:
            self.keyword_layout.itemAt(0).widget().hide()  # QLabel
            self.keyword_layout.itemAt(1).widget().hide()  # QLineEdit
            
    def on_split_mode_changed(self, index):
        """当拆分模式改变时，显示相应的配置项"""
        # 根据选择的拆分模式显示或隐藏相应配置项
        if index == 0:  # 按行拆分
            self.split_column_layout.itemAt(0).widget().show()  # QLabel
            self.split_column_layout.itemAt(1).widget().show()  # QSpinBox
            self.output_columns_layout.itemAt(0).widget().hide()  # QLabel
            self.output_columns_layout.itemAt(1).widget().hide()  # QLineEdit
        else:  # 按列拆分
            self.split_column_layout.itemAt(0).widget().hide()  # QLabel
            self.split_column_layout.itemAt(1).widget().hide()  # QSpinBox
            self.output_columns_layout.itemAt(0).widget().show()  # QLabel
            self.output_columns_layout.itemAt(1).widget().show()  # QLineEdit
            
    def add_merge_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if files:
            for file in files:
                self.merge_file_list.addItem(file)
                
    def remove_merge_files(self):
        for item in self.merge_file_list.selectedItems():
            self.merge_file_list.takeItem(self.merge_file_list.row(item))
            
    def browse_merge_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "保存合并文件", "", "Excel Files (*.xlsx)")
        if file:
            self.merge_output_edit.setText(file)
            
    def browse_split_source(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择源Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.split_source_edit.setText(file)
            
    def browse_split_output_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if directory:
            self.split_output_dir_edit.setText(directory)
            
    def browse_rename_source(self):
        directory = QFileDialog.getExistingDirectory(self, "选择源目录")
        if directory:
            self.rename_source_edit.setText(directory)
            
    def browse_clean_source(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择源Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.clean_source_edit.setText(file)
            
    def browse_clean_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "保存清洗后文件", "", "Excel Files (*.xlsx)")
        if file:
            self.clean_output_edit.setText(file)
            
    def browse_normalize_input(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择输入Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.normalize_input_edit.setText(file)
            
    def browse_normalize_output(self):
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if directory:
            self.normalize_output_edit.setText(directory)
            
    def browse_compare_file1(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择历史文件", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.compare_file1_edit.setText(file)
            
    def browse_compare_file2(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择当前文件", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.compare_file2_edit.setText(file)
            
    def browse_compare_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "保存对比结果文件", "", "Excel Files (*.xlsx)")
        if file:
            self.compare_output_edit.setText(file)
            
    # 执行功能函数
    def execute_classify(self):
        source_path = self.classify_source_edit.text()
        target_path = self.classify_target_edit.text()
        method_index = self.classify_method_combo.currentIndex()
        
        if not source_path or not target_path:
            QMessageBox.warning(self, "警告", "请填写源目录和目标目录")
            return
            
        # 根据选择的分类方式执行相应的分类函数
        if method_index == 0:  # 按文件类型分类
            types_text = self.classify_types_edit.text()
            selected_types = types_text.split() if types_text else None
            
            # 在工作线程中执行
            self.worker_thread = WorkerThread(
                classify_files, source_path, target_path, selected_types, None, None
            )
        elif method_index == 1:  # 按关键词分类
            keywords_text = self.classify_keywords_edit.text()
            keywords = keywords_text.split(',') if keywords_text else None
            if keywords:
                # 去除关键词前后的空格
                keywords = [keyword.strip() for keyword in keywords]
            
            # 在工作线程中执行关键词分类
            self.worker_thread = WorkerThread(
                classify_files_by_keywords, source_path, target_path, keywords
            )
        elif method_index == 2:  # 按文件扩展名分类
            # 在工作线程中执行扩展名分类
            self.worker_thread = WorkerThread(
                classify_files_by_extension, source_path, target_path
            )
            
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行文件分类...")
        
    def execute_merge(self):
        file_count = self.merge_file_list.count()
        if file_count < 2:
            QMessageBox.warning(self, "警告", "请至少添加两个Excel文件")
            return
            
        file_paths = [self.merge_file_list.item(i).text() for i in range(file_count)]
        match_columns_text = self.merge_match_edit.text()
        output_path = self.merge_output_edit.text()
        
        if not match_columns_text:
            QMessageBox.warning(self, "警告", "请输入匹配列序号")
            return
            
        if not output_path:
            QMessageBox.warning(self, "警告", "请选择输出文件路径")
            return
            
        try:
            match_columns = [int(x) for x in match_columns_text.split()]
        except ValueError:
            QMessageBox.warning(self, "警告", "匹配列序号格式错误，请输入数字并以空格分隔")
            return
            
        if len(match_columns) != len(file_paths):
            QMessageBox.warning(self, "警告", "匹配列数量与文件数量不一致")
            return
            
        # 在工作线程中执行
        self.worker_thread = WorkerThread(
            merge_excel_files_by_column, file_paths, match_columns, None, output_path
        )
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行文件合并...")
        
    def execute_split(self):
        source_file = self.split_source_edit.text()
        output_dir = self.split_output_dir_edit.text()
        mode_index = self.split_mode_combo.currentIndex()
        
        if not source_file or not output_dir:
            QMessageBox.warning(self, "警告", "请填写源文件和输出目录")
            return
            
        if mode_index == 0:  # 按行拆分
            split_column = self.split_column_spin.value()
            
            # 在工作线程中执行
            self.worker_thread = WorkerThread(
                split_excel_by_column, source_file, split_column, None, output_dir
            )
        else:  # 按列拆分
            output_columns_text = self.split_output_edit.text()
            
            if not output_columns_text:
                QMessageBox.warning(self, "警告", "请输入输出列组合")
                return
                
            try:
                output_columns = []
                for part in output_columns_text.split():
                    cols = [int(x) for x in part.split(',')]
                    output_columns.append(cols)
            except ValueError:
                QMessageBox.warning(self, "警告", "输出列组合格式错误，请按示例格式输入")
                return
                
            # 在工作线程中执行（这里我们创建一个简化版本的按列拆分功能）
            self.worker_thread = WorkerThread(
                self.split_excel_by_columns_only, source_file, output_columns, output_dir
            )
            
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行文件拆分...")
        
    def execute_rename(self):
        source_path = self.rename_source_edit.text()
        prefix = self.rename_prefix_edit.text()
        start_number = self.rename_start_spin.value()
        digits = self.rename_digits_spin.value()
        keyword = self.rename_keyword_edit.text()
        
        if not source_path:
            QMessageBox.warning(self, "警告", "请选择源目录")
            return
            
        # 在工作线程中执行
        self.worker_thread = WorkerThread(
            rename_files_sequentially, source_path, prefix, start_number, digits, keyword
        )
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行文件重命名...")
        
    def execute_clean(self):
        source_file = self.clean_source_edit.text()
        output_file = self.clean_output_edit.text()
        clean_symbols = self.clean_symbols_check.isChecked()
        symbols_text = self.symbols_edit.text()
        mark_empty = self.mark_empty_check.isChecked()
        mark_duplicates = self.mark_duplicates_check.isChecked()
        clean_internal_spaces = self.clean_internal_spaces_check.isChecked()
        clean_chinese_space = self.clean_chinese_space_check.isChecked()
        clean_english_punctuation = self.clean_english_punctuation_check.isChecked()
        
        if not source_file:
            QMessageBox.warning(self, "警告", "请选择源Excel文件")
            return
            
        symbols_to_remove = symbols_text.split() if symbols_text else None
        
        # 在工作线程中执行
        self.worker_thread = WorkerThread(
            clean_excel_data, source_file, output_file, clean_symbols, symbols_to_remove, 
            mark_empty, mark_duplicates, clean_internal_spaces, clean_chinese_space, clean_english_punctuation
        )
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行数据清洗...")
        
    def execute_normalize(self):
        input_file = self.normalize_input_edit.text()
        output_folder = self.normalize_output_edit.text()
        column_index = self.normalize_column_spin.value()
        group_column_index = self.normalize_group_spin.value()
        similarity_threshold = self.normalize_similarity_spin.value()
        edit_distance_threshold = self.normalize_edit_distance_spin.value()
        min_text_length = self.normalize_min_length_spin.value()
        
        if not input_file or not output_folder:
            QMessageBox.warning(self, "警告", "请填写输入文件和输出目录")
            return
            
        # 在工作线程中执行
        self.worker_thread = WorkerThread(
            process_indication_standardization, input_file, output_folder, column_index, group_column_index,
            similarity_threshold, edit_distance_threshold, min_text_length
        )
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行适应症写法规范化...")
        
    def execute_compare(self):
        file1_path = self.compare_file1_edit.text()
        file2_path = self.compare_file2_edit.text()
        output_path = self.compare_output_edit.text()
        name_similarity_threshold = self.compare_name_similarity_spin.value()
        text_similarity_threshold = self.compare_text_similarity_spin.value() / 100.0
        
        # 文件1列索引
        file1_drug_col = self.compare_file1_drug_spin.value()
        file1_company_col = self.compare_file1_company_spin.value()
        file1_status_col = self.compare_file1_status_spin.value()
        file1_content_col = self.compare_file1_content_spin.value()
        
        # 文件2列索引
        file2_drug_col = self.compare_file2_drug_spin.value()
        file2_company_col = self.compare_file2_company_spin.value()
        file2_status_col = self.compare_file2_status_spin.value()
        
        if not file1_path or not file2_path or not output_path:
            QMessageBox.warning(self, "警告", "请填写所有文件路径")
            return
            
        # 在工作线程中执行
        self.worker_thread = WorkerThread(
            update_file_comparison, file1_path, file2_path, output_path,
            name_similarity_threshold, text_similarity_threshold,
            file1_drug_col, file1_company_col, file1_status_col, file1_content_col,
            file2_drug_col, file2_company_col, file2_status_col
        )
        self.worker_thread.output_signal.connect(self.append_output)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_operation_finished)
        self.worker_thread.start()
        
        self.set_ui_disabled(True)
        self.append_output("开始执行文件对比分析...")
        
    def append_output(self, text):
        self.output_text.append(text)
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def on_operation_finished(self, success, message):
        self.set_ui_disabled(False)
        if success:
            self.append_output(f"操作成功: {message}")
            QMessageBox.information(self, "成功", message)
        else:
            self.append_output(f"操作失败: {message}")
            QMessageBox.critical(self, "失败", message)
        self.statusBar().showMessage("就绪")
        
    def on_clean_internal_spaces_changed(self, state):
        """当清除内部空格选项状态改变时，启用或禁用子选项"""
        enabled = state == Qt.Checked
        self.clean_chinese_space_check.setEnabled(enabled)
        self.clean_english_punctuation_check.setEnabled(enabled)
    def set_ui_disabled(self, disabled):
        # 禁用/启用所有标签页的控件
        for i in range(self.centralWidget().layout().itemAt(0).widget().count()):
            tab = self.centralWidget().layout().itemAt(0).widget().widget(i)
            self.set_widget_disabled(tab, disabled)
            
        self.progress_bar.setVisible(disabled)
        
    def set_widget_disabled(self, widget, disabled):
        if hasattr(widget, 'setLayout'):
            layout = widget.layout()
            if layout:
                for i in range(layout.count()):
                    item = layout.itemAt(i)
                    if item.widget():
                        if isinstance(item.widget(), QWidget):
                            self.set_widget_disabled(item.widget(), disabled)
                        else:
                            item.widget().setDisabled(disabled)
                            
    def split_excel_by_columns_only(self, source_file, output_columns, output_dir, output_callback=None):
        """
        仅按列拆分Excel文件的辅助函数
        
        Args:
            source_file (str): 源Excel文件路径
            output_columns (list): 要输出的列组合列表
            output_dir (str): 输出目录路径
            output_callback (function): 输出回调函数
            
        Returns:
            bool: 拆分是否成功
        """
        def _print(msg):
            if output_callback:
                output_callback(msg)
            else:
                print(msg)
                
        try:
            import pandas as pd
            import os
            
            # 检查源文件是否存在
            if not os.path.exists(source_file):
                raise FileNotFoundError(f"源文件不存在: {source_file}")
            
            # 创建输出目录
            os.makedirs(output_dir, exist_ok=True)
            
            # 读取Excel文件
            df = pd.read_excel(source_file, dtype=str)
            
            _print(f"源文件共有 {len(df)} 行, {len(df.columns)} 列")
            
            # 为每个输出列组合创建文件
            for i, columns in enumerate(output_columns):
                # 转换列索引为从0开始
                zero_based_columns = [col - 1 for col in columns]
                
                # 检查列索引是否有效
                invalid_cols = [col for col in zero_based_columns if col >= len(df.columns) or col < 0]
                if invalid_cols:
                    _print(f"警告: 列索引 {invalid_cols} 超出范围，跳过该列组合")
                    
                else:
                    # 选择列
                    selected_df = df.iloc[:, zero_based_columns]
                    
                    # 生成输出文件名
                    output_file = os.path.join(output_dir, f"output_{i + 1}.xlsx")
                    
                    # 保存文件
                    selected_df.to_excel(output_file, index=False)
                    _print(f"已保存 {output_file}")
                    
            return True
        except Exception as e:
            _print(f"错误: {str(e)}")
            return False

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用现代样式
    window = AutoExcelGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()