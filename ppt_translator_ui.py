import sys
import os
import json
import requests
import platform
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QComboBox, 
                           QFileDialog, QProgressBar, QLineEdit, QMessageBox,
                           QGroupBox, QMenu)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QAction
from ppt_xml_translator import PPTXMLTranslator

def open_file_location(file_path):
    """跨平台打开文件所在位置
    Args:
        file_path: 文件路径
    """
    if platform.system() == "Windows":
        os.system(f'explorer /select,"{file_path}"')
    elif platform.system() == "Darwin":  # macOS
        os.system(f'open -R "{file_path}"')
    else:  # Linux 或其他系统
        os.system(f'xdg-open "{os.path.dirname(file_path)}"')

class TranslationWorker(QThread):
    """翻译工作线程"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    slide_progress = pyqtSignal(int, int)  # 当前页数，总页数

    def __init__(self, translator, input_file, output_file, from_lang, to_lang):
        super().__init__()
        self.translator = translator
        self.input_file = input_file
        self.output_file = output_file
        self.from_lang = from_lang
        self.to_lang = to_lang

    def run(self):
        try:
            self.progress.emit("开始翻译...")
            output_path = self.translator.translate_pptx_file(
                self.input_file,
                self.output_file,
                self.from_lang,
                self.to_lang,
                progress_callback=self.handle_progress
            )
            self.finished.emit(output_path)
        except Exception as e:
            self.error.emit(str(e))

    def handle_progress(self, current_slide, total_slides):
        """处理翻译进度"""
        self.slide_progress.emit(current_slide, total_slides)
        self.progress.emit(f"正在翻译第 {current_slide}/{total_slides} 页...")

class PPTTranslatorUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.translator = None
        self.worker = None
        self.recent_files = self.load_recent_files()
        self.create_menu_bar()
        
        # 根据操作系统调整界面
        self.adjust_for_platform()

    def adjust_for_platform(self):
        """根据操作系统调整界面"""
        if platform.system() == "Darwin":  # macOS
            # 调整字体大小
            self.setStyleSheet("QLabel { font-size: 13px; } QLineEdit { font-size: 13px; } QPushButton { font-size: 13px; }")
            # 调整按钮大小
            self.translate_btn.setMinimumHeight(32)
        else:
            self.translate_btn.setMinimumHeight(40)

    def init_ui(self):
        """初始化UI界面"""
        self.setWindowTitle('PPT翻译工具')
        self.setMinimumWidth(600)
        
        # 创建主窗口部件和布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # 输入文件
        input_layout = QHBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("选择要翻译的PPT文件...")
        input_btn = QPushButton("浏览...")
        input_btn.clicked.connect(self.select_input_file)
        input_layout.addWidget(QLabel("输入文件:"))
        input_layout.addWidget(self.input_path)
        input_layout.addWidget(input_btn)
        file_layout.addLayout(input_layout)

        # 输出文件
        output_layout = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("选择翻译后的保存位置...")
        output_btn = QPushButton("浏览...")
        output_btn.clicked.connect(self.select_output_file)
        output_layout.addWidget(QLabel("输出文件:"))
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(output_btn)
        file_layout.addLayout(output_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 翻译设置区域
        settings_group = QGroupBox("翻译设置")
        settings_layout = QVBoxLayout()

        # 语言方向设置
        lang_layout = QHBoxLayout()
        self.from_lang = QComboBox()
        self.from_lang.addItems(['中文', '英文'])
        self.to_lang = QComboBox()
        self.to_lang.addItems(['英文', '中文'])
        # 添加语言切换按钮
        switch_btn = QPushButton("⇄")
        switch_btn.setFixedWidth(30)
        switch_btn.clicked.connect(self.switch_languages)
        lang_layout.addWidget(QLabel("从:"))
        lang_layout.addWidget(self.from_lang)
        lang_layout.addWidget(switch_btn)
        lang_layout.addWidget(QLabel("翻译到:"))
        lang_layout.addWidget(self.to_lang)
        settings_layout.addLayout(lang_layout)

        # 模型选择
        model_layout = QHBoxLayout()
        self.model_select = QComboBox()
        self.model_select.addItems(['llama3:8b', 'qwen:7b', 'qwen:1.8b'])
        model_layout.addWidget(QLabel("模型:"))
        model_layout.addWidget(self.model_select)
        settings_layout.addLayout(model_layout)

        # 服务器设置
        server_layout = QHBoxLayout()
        self.server_url = QLineEdit("http://localhost:2342")
        test_connection_btn = QPushButton("测试连接")
        test_connection_btn.clicked.connect(self.test_server_connection)
        server_layout.addWidget(QLabel("服务器地址:"))
        server_layout.addWidget(self.server_url)
        server_layout.addWidget(test_connection_btn)
        settings_layout.addLayout(server_layout)

        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)

        # 进度显示
        progress_group = QGroupBox("翻译进度")
        progress_layout = QVBoxLayout()
        
        # 页面进度
        self.slide_progress = QProgressBar()
        self.slide_progress.setTextVisible(True)
        self.slide_progress.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.slide_progress.setFormat("准备就绪")
        progress_layout.addWidget(self.slide_progress)
        
        # 详细信息标签
        self.status_label = QLabel("等待开始...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.status_label)
        
        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)

        # 翻译按钮
        self.translate_btn = QPushButton("开始翻译")
        self.translate_btn.setMinimumHeight(40)
        self.translate_btn.clicked.connect(self.start_translation)
        layout.addWidget(self.translate_btn)

        # 设置窗口大小和位置
        self.setGeometry(100, 100, 800, 500)

    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu('文件')
        
        # 最近使用的文件
        self.recent_menu = QMenu('最近使用', self)
        self.update_recent_files_menu()
        file_menu.addMenu(self.recent_menu)
        
        # 清除最近使用记录
        clear_recent = QAction('清除最近使用记录', self)
        clear_recent.triggered.connect(self.clear_recent_files)
        file_menu.addAction(clear_recent)
        
        file_menu.addSeparator()
        
        # 退出
        exit_action = QAction('退出', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 帮助菜单
        help_menu = menubar.addMenu('帮助')
        
        # 关于
        about_action = QAction('关于', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def load_recent_files(self):
        """加载最近使用的文件列表"""
        try:
            config_dir = self.get_config_dir()
            os.makedirs(config_dir, exist_ok=True)
            recent_files_path = os.path.join(config_dir, 'recent_files.json')
            if os.path.exists(recent_files_path):
                with open(recent_files_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载最近文件列表失败: {e}")
        return []

    def save_recent_files(self):
        """保存最近使用的文件列表"""
        try:
            config_dir = self.get_config_dir()
            os.makedirs(config_dir, exist_ok=True)
            recent_files_path = os.path.join(config_dir, 'recent_files.json')
            with open(recent_files_path, 'w', encoding='utf-8') as f:
                json.dump(self.recent_files, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存最近文件列表失败: {e}")

    def get_config_dir(self):
        """获取配置文件目录（跨平台）"""
        if platform.system() == "Windows":
            return os.path.join(os.getenv('APPDATA'), 'PPTTranslator')
        elif platform.system() == "Darwin":  # macOS
            return os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', 'PPTTranslator')
        else:  # Linux 或其他系统
            return os.path.join(os.path.expanduser('~'), '.config', 'PPTTranslator')

    def update_recent_files_menu(self):
        """更新最近使用的文件菜单"""
        self.recent_menu.clear()
        for file_path in self.recent_files:
            action = QAction(file_path, self)
            action.triggered.connect(lambda checked, path=file_path: self.open_recent_file(path))
            self.recent_menu.addAction(action)

    def add_recent_file(self, file_path):
        """添加文件到最近使用列表"""
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        self.recent_files.insert(0, file_path)
        self.recent_files = self.recent_files[:10]  # 只保留最近10个
        self.save_recent_files()
        self.update_recent_files_menu()

    def open_recent_file(self, file_path):
        """打开最近使用的文件"""
        if os.path.exists(file_path):
            self.input_path.setText(file_path)
            base_name = os.path.splitext(file_path)[0]
            self.output_path.setText(f"{base_name}_translated.pptx")
        else:
            QMessageBox.warning(self, "错误", f"文件不存在：{file_path}")
            self.recent_files.remove(file_path)
            self.save_recent_files()
            self.update_recent_files_menu()

    def clear_recent_files(self):
        """清除最近使用的文件记录"""
        self.recent_files = []
        self.save_recent_files()
        self.update_recent_files_menu()

    def show_about(self):
        """显示关于对话框"""
        QMessageBox.about(self, 
            "关于 PPT翻译工具",
            "PPT翻译工具 v1.0.0\n\n"
            "基于本地Ollama服务的PPT翻译工具\n"
            "支持中英互译，保持原始格式\n\n"
            "作者: enzo X cursor\n"
            "最后更新: 2024-12-21"
        )

    def switch_languages(self):
        """切换源语言和目标语言"""
        from_idx = self.from_lang.currentIndex()
        to_idx = self.to_lang.currentIndex()
        self.from_lang.setCurrentIndex(to_idx)
        self.to_lang.setCurrentIndex(from_idx)

    def test_server_connection(self):
        """测试服务器连接"""
        try:
            # 创建翻译器实例
            translator = PPTXMLTranslator(host=self.server_url.text())
            # 发送测试请求
            response = requests.get(f"{self.server_url.text()}/api/version")
            if response.status_code == 200:
                QMessageBox.information(self, "连接测试", "服务器连接成功！")
            else:
                QMessageBox.warning(self, "连接测试", "服务器连接失败！")
        except Exception as e:
            QMessageBox.critical(self, "连接测试", f"连接错误：{str(e)}")

    def select_input_file(self):
        """选择输入文件"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择PPT文件",
            "",
            "PowerPoint Files (*.pptx);;All Files (*)"
        )
        if file_name:
            self.input_path.setText(file_name)
            # 自动生成输出文件路径
            base_name = os.path.splitext(file_name)[0]
            self.output_path.setText(f"{base_name}_translated.pptx")
            # 添加到最近使用的文件列表
            self.add_recent_file(file_name)

    def select_output_file(self):
        """选择输出文件"""
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "保存翻译后的文件",
            "",
            "PowerPoint Files (*.pptx);;All Files (*)"
        )
        if file_name:
            self.output_path.setText(file_name)

    def start_translation(self):
        """开始翻译"""
        # 检查输入
        if not self.input_path.text():
            QMessageBox.warning(self, "警告", "请选择要翻译的PPT文件！")
            return
        if not self.output_path.text():
            QMessageBox.warning(self, "警告", "请选择输出文件位置！")
            return

        # 禁用界面
        self.setEnabled(False)
        self.status_label.setText("正在准备翻译...")
        self.slide_progress.setRange(0, 100)
        self.slide_progress.setValue(0)

        # 准备翻译参数
        from_lang = 'zh' if self.from_lang.currentText() == '中文' else 'en'
        to_lang = 'en' if self.to_lang.currentText() == '英文' else 'zh'

        # 创建翻译器实例
        self.translator = PPTXMLTranslator(
            model_name=self.model_select.currentText(),
            host=self.server_url.text()
        )

        # 创建并启动工作线程
        self.worker = TranslationWorker(
            self.translator,
            self.input_path.text(),
            self.output_path.text(),
            from_lang,
            to_lang
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.slide_progress.connect(self.update_slide_progress)
        self.worker.finished.connect(self.translation_finished)
        self.worker.error.connect(self.translation_error)
        self.worker.start()

    def update_progress(self, message):
        """更新进度信息"""
        self.status_label.setText(message)

    def update_slide_progress(self, current, total):
        """更新页面进度"""
        progress = int((current / total) * 100)
        self.slide_progress.setValue(progress)
        self.slide_progress.setFormat(f"进度: {progress}% ({current}/{total}页)")

    def translation_finished(self, output_path):
        """翻译完成处理"""
        self.slide_progress.setValue(100)
        self.slide_progress.setFormat("翻译完成！")
        self.status_label.setText("翻译已完成")
        self.setEnabled(True)
        
        reply = QMessageBox.information(
            self,
            "完成",
            f"翻译已完成！\n输出文件：{output_path}\n\n是否打开输出文件夹？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            open_file_location(output_path)

    def translation_error(self, error_message):
        """翻译错误处理"""
        self.slide_progress.setValue(0)
        self.slide_progress.setFormat("翻译失败")
        self.status_label.setText("翻译失败")
        self.setEnabled(True)
        QMessageBox.critical(
            self,
            "错误",
            f"翻译过程中出错：\n{error_message}"
        )

def main():
    app = QApplication(sys.argv)
    # 设置应用程序样式
    app.setStyle('Fusion')
    window = PPTTranslatorUI()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main() 