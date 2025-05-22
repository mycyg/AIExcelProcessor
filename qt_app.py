import sys
import json
from pathlib import Path
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget,
    QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QSpinBox, QTextEdit, QFileDialog,
    QComboBox, QCheckBox, QProgressBar, QGroupBox,
    QListWidget, QMessageBox
)
from PySide6.QtCore import Qt, QThread, Signal as pyqtSignal, QPoint, QTimer, Slot, QDateTime
from PySide6.QtGui import QKeyEvent, QFocusEvent, QCloseEvent

# 假设 processor.py 位于名为 'gui' 的子目录中
# 如果 processor.py 与 qt_app.py 在同一目录, 请改回:
# from processor import ExcelProcessor
from gui.processor import ExcelProcessor

class CustomTextEditWithSuggestions(QTextEdit):
    """
    一个自定义的QTextEdit，当输入'/'或手动触发时，
    会显示一个建议弹出列表 (QListWidget) 作为其子控件。
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        # 将建议列表作为 QTextEdit 的直接子控件
        self._suggestions_list_widget = QListWidget(self) 
        # 无需设置 Qt.Popup，它是一个普通的子控件，但会被置顶显示

        self._suggestions_list_widget.hide()
        self._suggestion_items = []
        
        self._show_popup_timer = QTimer(self) 
        self._show_popup_timer.setSingleShot(True)
        self._show_popup_timer.setInterval(50) 
        self._show_popup_timer.timeout.connect(self._trigger_suggestions_popup_display)

        self._suggestions_list_widget.itemClicked.connect(self._on_suggestion_item_clicked)
        
        self._suggestions_list_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #A9A9A9; 
                background-color: white;
                font-size: 10pt;
                margin: 0px; /* 子控件不需要额外边距来避免裁剪 */
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:hover {
                background-color: #E0E0E0; 
            }
            QListWidget::item:selected {
                background-color: #C0C0C0; 
                color: black;
            }
        """)

    def set_suggestion_items(self, items):
        self._suggestion_items = items
        if self.hasFocus() and self._suggestions_list_widget.isVisible():
            self._update_suggestions_list_content()

    def _update_suggestions_list_content(self, filter_text=""):
        self._suggestions_list_widget.clear()
        
        if not self._suggestion_items:
            self._suggestions_list_widget.addItem("无可用列名")
        else:
            added_items = 0
            for item_text in self._suggestion_items:
                if filter_text.lower() in item_text.lower(): 
                    self._suggestions_list_widget.addItem(item_text)
                    added_items +=1
            if added_items == 0:
                self._suggestions_list_widget.addItem("无匹配项")
        
        if self._suggestions_list_widget.count() == 0: 
            self._suggestions_list_widget.hide()

    def _trigger_suggestions_popup_display(self):
        # 检查 QTextEdit 是否有焦点，对于子控件方式，这仍然是一个好习惯
        if not self.hasFocus():
            # 如果是由 '/' 触发但焦点已失，则不显示
            # 如果是由按钮触发，按钮点击处理函数应确保焦点
            self._suggestions_list_widget.hide() # 确保隐藏
            return

        self._update_suggestions_list_content()
        
        if self._suggestions_list_widget.count() == 0 or \
           (self._suggestions_list_widget.count() == 1 and \
            self._suggestions_list_widget.item(0).text() in ["无可用列名", "无匹配项"]):
            self._suggestions_list_widget.hide()
            return

        cursor = self.textCursor()
        rect = self.cursorRect(cursor) 
        
        popup_pos = QPoint(rect.left(), rect.bottom() + 2) 
        self._suggestions_list_widget.move(popup_pos)
        
        self._suggestions_list_widget.setMinimumWidth(max(150, self.width() // 3)) 
        self._suggestions_list_widget.setMaximumWidth(self.viewport().width() - rect.left() - 5) 

        if self._suggestions_list_widget.count() > 0:
            self._suggestions_list_widget.setFixedHeight(
                min(150, (self._suggestions_list_widget.sizeHintForRow(0) * self._suggestions_list_widget.count()) + 
                         2 * self._suggestions_list_widget.frameWidth() + 5) 
            )
        else:
            self._suggestions_list_widget.setFixedHeight(30) 

        self._suggestions_list_widget.show()
        self._suggestions_list_widget.raise_() 

        QTimer.singleShot(0, self._finish_popup_setup)

    def _finish_popup_setup(self):
        if not self._suggestions_list_widget.isVisible():
            return
        
        self._suggestions_list_widget.setFocus(Qt.OtherFocusReason) 
        
        if self._suggestions_list_widget.count() > 0: 
            first_item_text = self._suggestions_list_widget.item(0).text()
            if first_item_text not in ["无可用列名", "无匹配项"]:
                self._suggestions_list_widget.setCurrentRow(0)
            elif self._suggestions_list_widget.count() > 1: 
                second_item_text = self._suggestions_list_widget.item(1).text()
                if second_item_text not in ["无可用列名", "无匹配项"]: 
                    self._suggestions_list_widget.setCurrentRow(1)

    def keyPressEvent(self, event: QKeyEvent):
        if self._suggestions_list_widget.isVisible():
            key = event.key()
            if key in [Qt.Key_Enter, Qt.Key_Return, Qt.Key_Tab]: 
                if self._suggestions_list_widget.currentItem():
                    self._insert_selected_suggestion(self._suggestions_list_widget.currentItem().text())
                self._suggestions_list_widget.hide()
                self.setFocus() 
                event.accept()
                return
            elif key == Qt.Key_Escape:
                self._suggestions_list_widget.hide()
                self.setFocus() 
                event.accept()
                return
            elif key in [Qt.Key_Up, Qt.Key_Down, Qt.Key_PageUp, Qt.Key_PageDown]:
                self._suggestions_list_widget.keyPressEvent(event)
                event.accept()
                return

        if event.text() == '/':
            self._show_popup_timer.start() 
            super().keyPressEvent(event) 
            event.accept()
            return

        super().keyPressEvent(event)

    def _on_suggestion_item_clicked(self, item):
        self._insert_selected_suggestion(item.text())
        self._suggestions_list_widget.hide()
        self.setFocus() 

    def _insert_selected_suggestion(self, suggestion_text):
        if suggestion_text in ["无可用列名", "无匹配项"]: 
            return

        cursor = self.textCursor()
        current_text_before_cursor = self.toPlainText()[:cursor.position()]
        if current_text_before_cursor.endswith('/'):
             cursor.deletePreviousChar()

        cursor.insertText(f"{{row['{suggestion_text}']}}") 
        self.setTextCursor(cursor)

    def focusOutEvent(self, event: QFocusEvent): 
        if self._suggestions_list_widget.isVisible():
            if QApplication.focusWidget() != self._suggestions_list_widget and \
               (not hasattr(QApplication.focusWidget(), 'parentWidget') or \
                QApplication.focusWidget().parentWidget() != self._suggestions_list_widget):
                QTimer.singleShot(100, self._check_and_hide_popup_on_focus_out_child)
        super().focusOutEvent(event)

    def _check_and_hide_popup_on_focus_out_child(self):
        if self._suggestions_list_widget.isVisible():
            current_focus = QApplication.focusWidget()
            if current_focus != self._suggestions_list_widget and \
               (not hasattr(current_focus, 'parentWidget') or current_focus.parentWidget() != self._suggestions_list_widget):
                self._suggestions_list_widget.hide()


    def show_suggestions_manually(self):
        # 这个方法由按钮点击调用，按钮点击处理程序应确保此QTextEdit具有焦点
        if self.hasFocus():
             self._show_popup_timer.start()
        # else:
            # 如果没有焦点，可以考虑是否要强制显示，或者依赖按钮处理程序设置焦点
            # print("CustomTextEditWithSuggestions: show_suggestions_manually called but no focus.")


class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, int) 
    finished = pyqtSignal(bool, str)      

    def __init__(self, processor_instance, parent=None):
        super().__init__(parent)
        self.processor = processor_instance

    def run(self):
        try:
            result = self.processor.start_processing() 
            self.finished.emit(result, "处理完成" if result else "处理失败或中止")
        except Exception as e:
            self.finished.emit(False, f"处理线程出错: {str(e)}")


class ExcelProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_path = Path("config.json")
        self.processor_instance = None 
        self.processing_thread = None
        self.is_processing = False
        self.column_names = [] 
        
        self._init_ui()
        self.config = self._load_config() 
        self._apply_config_to_ui() 

    def _init_ui(self):
        self.setWindowTitle("Excel 批量处理工具")
        self.setMinimumSize(850, 700) 

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        tabs = QTabWidget()
        tabs.addTab(self._create_basic_settings_tab(), "基本设置")
        tabs.addTab(self._create_column_settings_tab(), "列设置")
        tabs.addTab(self._create_prompt_settings_tab(), "提示词设置")
        tabs.addTab(self._create_process_tab(), "处理与日志") 

        layout.addWidget(tabs)

    def _create_basic_settings_tab(self):
        tab = QWidget()
        main_tab_layout = QHBoxLayout(tab)

        file_group = QGroupBox("文件设置")
        file_layout = QVBoxLayout(file_group)

        input_file_layout = QHBoxLayout()
        input_file_layout.addWidget(QLabel("输入文件:"))
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setPlaceholderText("选择 Excel 输入文件 (.xlsx, .xls)")
        input_file_layout.addWidget(self.input_file_edit)
        input_browse_btn = QPushButton("浏览...")
        input_browse_btn.clicked.connect(self.on_browse_input_file_clicked)
        input_file_layout.addWidget(input_browse_btn)
        file_layout.addLayout(input_file_layout)

        output_file_layout = QHBoxLayout()
        output_file_layout.addWidget(QLabel("输出文件:"))
        self.output_file_edit = QLineEdit()
        self.output_file_edit.setPlaceholderText("选择或输入 Excel 输出文件 (.xlsx)")
        output_file_layout.addWidget(self.output_file_edit)
        output_browse_btn = QPushButton("浏览...")
        output_browse_btn.clicked.connect(self.on_browse_output_file_clicked)
        output_file_layout.addWidget(output_browse_btn)
        file_layout.addLayout(output_file_layout)
        
        file_layout.addWidget(QLabel("Sheet 名称:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setToolTip("输入文件加载后，此处会列出可用的Sheet")
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_selection_changed)
        file_layout.addWidget(self.sheet_combo)

        file_layout.addWidget(QLabel("判断空行的列:"))
        self.empty_column_combo = QComboBox()
        self.empty_column_combo.setToolTip("选择用于判断某行是否为空（并跳过处理）的依据列")
        file_layout.addWidget(self.empty_column_combo)
        
        file_layout.addStretch(1) 
        main_tab_layout.addWidget(file_group, 1) 

        api_process_group = QGroupBox("API 及处理设置")
        api_process_layout = QVBoxLayout(api_process_group)

        api_process_layout.addWidget(QLabel("API URL:"))
        self.api_url_edit = QLineEdit()
        api_process_layout.addWidget(self.api_url_edit)

        api_process_layout.addWidget(QLabel("API Key:"))
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        api_process_layout.addWidget(self.api_key_edit)

        api_process_layout.addWidget(QLabel("模型名称:"))
        self.model_edit = QLineEdit()
        api_process_layout.addWidget(self.model_edit)
        
        api_process_layout.addWidget(QLabel("批处理数量 (Batch Size):"))
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setRange(1, 1000)
        self.batch_size_spin.setSuffix(" 行/批")
        api_process_layout.addWidget(self.batch_size_spin)

        api_process_layout.addWidget(QLabel("并行处理数量 (Workers):"))
        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(1, 100)
        api_process_layout.addWidget(self.workers_spin)
        
        api_process_layout.addStretch(1) 
        main_tab_layout.addWidget(api_process_group, 1) 
        
        return tab

    def _create_column_settings_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab) 

        input_group = QGroupBox("输入列选择 (用于内容整合模板)")
        input_group_layout = QVBoxLayout(input_group)
        self.input_columns_dynamic_layout = QVBoxLayout() 
        input_group_layout.addLayout(self.input_columns_dynamic_layout)
        input_group_layout.addStretch(1) 
        layout.addWidget(input_group, 1)

        output_group = QGroupBox("输出列定义 (LLM生成)")
        output_group_layout = QVBoxLayout(output_group)
        output_group_layout.addWidget(QLabel("在此定义LLM需要输出的列名，每行一个:"))
        self.output_columns_edit = QTextEdit()
        self.output_columns_edit.setPlaceholderText("例如:\n摘要\n关键词\n分类")
        output_group_layout.addWidget(self.output_columns_edit)
        layout.addWidget(output_group, 1)
        
        return tab

    def _create_prompt_settings_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab) 

        content_group = QGroupBox("内容整合模板")
        content_layout = QVBoxLayout(content_group)
        
        content_header_layout = QHBoxLayout()
        content_header_layout.addWidget(QLabel("输入 '/' 或点击按钮插入列名:"))
        insert_column_btn = QPushButton("插入列引用")
        insert_column_btn.clicked.connect(self.on_insert_column_button_clicked)
        content_header_layout.addWidget(insert_column_btn)
        content_layout.addLayout(content_header_layout)

        self.content_template_edit = CustomTextEditWithSuggestions(self)
        self.content_template_edit.setPlaceholderText("例如: 【背景: {row['背景信息']}】 【内容: {row['发言内容']}】")
        content_layout.addWidget(self.content_template_edit)
        layout.addWidget(content_group, 1)

        llm_group = QGroupBox("LLM 提示词模板")
        llm_layout = QVBoxLayout(llm_group)
        llm_layout.addWidget(QLabel("使用 {{content}} 引用上方整合后的内容."))
        self.llm_template_edit = QTextEdit()
        self.llm_template_edit.setPlaceholderText("例如: 基于以下内容:\n{{content}}\n请提取关键信息并分类。")
        llm_layout.addWidget(self.llm_template_edit)
        layout.addWidget(llm_group, 1)

        return tab

    def _create_process_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px;")
        self.start_btn.clicked.connect(self.on_start_processing_clicked)
        
        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("background-color: #f44336; color: white; padding: 8px;")
        self.stop_btn.clicked.connect(self.on_stop_processing_clicked)
        
        control_layout.addWidget(self.start_btn)
        control_layout.addWidget(self.stop_btn)
        control_layout.addStretch(1)
        layout.addLayout(control_layout)

        layout.addWidget(QLabel("处理进度:"))
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p% (%v/%m)")
        layout.addWidget(self.progress_bar)

        layout.addWidget(QLabel("日志与结果:"))
        self.result_text_edit = QTextEdit() 
        self.result_text_edit.setReadOnly(True)
        layout.addWidget(self.result_text_edit)
        
        return tab

    def _load_config(self):
        default_config = {
            'input_file': '', 'output_file': '', 'sheet_name': '', 'empty_column': '',
            'batch_size': 30, 'workers': 10,
            'api_url': 'https://ark.cn-beijing.volces.com/api/v3/chat/completions',
            'api_key': '', 'model': 'doubao-1-5-pro-32k-250115',
            'content_template': "【内容: {row['发言内容']}】",
            'llm_template': "请处理以下内容:\n{{content}}",
            'input_columns': {}, 
            'output_columns': ["问题分类", "发言分类"] 
        }
        if not self.config_path.exists():
            self.log_message("配置文件 config.json 未找到，将使用默认设置。")
            return default_config
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                for key, value in default_config.items():
                    if key not in loaded_config:
                        loaded_config[key] = value
                self.log_message("配置文件加载成功。") 
                return loaded_config
        except Exception as e:
            self.log_message(f"加载配置文件失败: {e}。将使用默认设置。") 
            QMessageBox.warning(self, "配置加载错误", f"加载 config.json 失败: {e}\n将使用默认设置。")
            return default_config

    def _apply_config_to_ui(self):
        self.input_file_edit.setText(self.config.get('input_file', ''))
        if self.config.get('input_file'): 
            self._update_sheets_and_columns_from_file(self.config['input_file'])

        self.output_file_edit.setText(self.config.get('output_file', ''))
        
        self.api_url_edit.setText(self.config.get('api_url', ''))
        self.api_key_edit.setText(self.config.get('api_key', ''))
        self.model_edit.setText(self.config.get('model', ''))
        
        self.batch_size_spin.setValue(int(self.config.get('batch_size', 30)))
        self.workers_spin.setValue(int(self.config.get('workers', 10)))
        
        self.content_template_edit.setPlainText(self.config.get('content_template', ''))
        self.llm_template_edit.setPlainText(self.config.get('llm_template', ''))
        
        output_cols_list = self.config.get('output_columns', [])
        self.output_columns_edit.setPlainText("\n".join(output_cols_list))

    def _gather_config_from_ui(self):
        input_cols_config = {}
        for i in range(self.input_columns_dynamic_layout.count()):
            widget = self.input_columns_dynamic_layout.itemAt(i).widget()
            if isinstance(widget, QCheckBox):
                input_cols_config[widget.text()] = widget.isChecked()
        
        output_cols_list = [line.strip() for line in self.output_columns_edit.toPlainText().splitlines() if line.strip()]

        return {
            'input_file': self.input_file_edit.text(),
            'output_file': self.output_file_edit.text(),
            'sheet_name': self.sheet_combo.currentText(),
            'empty_column': self.empty_column_combo.currentText(),
            'batch_size': self.batch_size_spin.value(),
            'workers': self.workers_spin.value(),
            'api_url': self.api_url_edit.text(),
            'api_key': self.api_key_edit.text(),
            'model': self.model_edit.text(),
            'content_template': self.content_template_edit.toPlainText(),
            'llm_template': self.llm_template_edit.toPlainText(),
            'input_columns': input_cols_config,
            'output_columns': output_cols_list
        }

    def _save_config(self):
        current_ui_config = self._gather_config_from_ui()
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(current_ui_config, f, ensure_ascii=False, indent=4)
            self.log_message("配置已保存到 config.json")
        except Exception as e:
            self.log_message(f"保存配置失败: {e}")
            QMessageBox.critical(self, "保存配置错误", f"无法保存配置到 config.json: {e}")
    
    @Slot()
    def on_browse_input_file_clicked(self):
        filename, _ = QFileDialog.getOpenFileName(self, "选择输入 Excel 文件", "", "Excel Files (*.xlsx *.xls)")
        if filename:
            self.input_file_edit.setText(filename)
            self._update_sheets_and_columns_from_file(filename)

    @Slot()
    def on_browse_output_file_clicked(self):
        filename, _ = QFileDialog.getSaveFileName(self, "选择输出 Excel 文件", "", "Excel Files (*.xlsx)")
        if filename:
            self.output_file_edit.setText(filename)

    def _update_sheets_and_columns_from_file(self, filename):
        try:
            excel_file = pd.ExcelFile(filename)
            sheet_names = excel_file.sheet_names
            
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheet_names)
            
            config_sheet_name = self.config.get('sheet_name')
            if config_sheet_name in sheet_names:
                self.sheet_combo.setCurrentText(config_sheet_name)
            elif sheet_names:
                self.sheet_combo.setCurrentIndex(0)
            
            if self.sheet_combo.currentIndex() >= 0 :
                 self.on_sheet_selection_changed() 
            else: 
                self.column_names = []
                self.empty_column_combo.clear()
                self._update_input_column_checkboxes([])
                self.content_template_edit.set_suggestion_items([])

        except Exception as e:
            self.log_message(f"读取文件 '{filename}' 的 Sheet 失败: {e}")
            QMessageBox.warning(self, "文件读取错误", f"无法读取文件 '{filename}' 的 Sheet 列表: {e}")
            self.column_names = []
            self.sheet_combo.clear()
            self.empty_column_combo.clear()
            self._update_input_column_checkboxes([])
            self.content_template_edit.set_suggestion_items([])

    @Slot()
    def on_sheet_selection_changed(self):
        input_file = self.input_file_edit.text()
        selected_sheet = self.sheet_combo.currentText()

        if not input_file or not selected_sheet:
            self.column_names = []
            self.empty_column_combo.clear()
            self._update_input_column_checkboxes([])
            self.content_template_edit.set_suggestion_items([])
            return

        try:
            df = pd.read_excel(input_file, sheet_name=selected_sheet, nrows=0) 
            self.column_names = [str(col) for col in df.columns]

            self.empty_column_combo.clear()
            self.empty_column_combo.addItems(self.column_names)
            config_empty_col = self.config.get('empty_column')
            if config_empty_col in self.column_names:
                self.empty_column_combo.setCurrentText(config_empty_col)
            elif self.column_names:
                self.empty_column_combo.setCurrentIndex(0)

            self._update_input_column_checkboxes(self.column_names)
            self.content_template_edit.set_suggestion_items(self.column_names) 

        except Exception as e:
            self.log_message(f"读取 Sheet '{selected_sheet}' 的列失败: {e}")
            QMessageBox.warning(self, "Sheet 读取错误", f"无法读取 Sheet '{selected_sheet}' 的列: {e}")
            self.column_names = []
            self.empty_column_combo.clear()
            self._update_input_column_checkboxes([])
            self.content_template_edit.set_suggestion_items([])

    def _update_input_column_checkboxes(self, columns):
        while self.input_columns_dynamic_layout.count():
            item = self.input_columns_dynamic_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        configured_input_cols = self.config.get('input_columns', {})

        for col_name in columns:
            checkbox = QCheckBox(col_name)
            checkbox.setChecked(configured_input_cols.get(col_name, False)) 
            self.input_columns_dynamic_layout.addWidget(checkbox)

    @Slot()
    def on_insert_column_button_clicked(self):
        # 首先确保目标文本编辑框获得焦点
        self.content_template_edit.setFocus(Qt.OtherFocusReason) 
        # 然后再调用手动显示建议的方法
        self.content_template_edit.show_suggestions_manually()

    def _validate_inputs(self):
        config = self._gather_config_from_ui()
        error_messages = []

        if not config['input_file'] or not Path(config['input_file']).exists():
            error_messages.append("输入文件无效或不存在。")
        if not config['output_file']:
            error_messages.append("请指定输出文件。")
        if not config['sheet_name']:
            error_messages.append("请选择一个 Sheet 进行处理。")
        if not config['empty_column'] and self.column_names: 
             error_messages.append("请选择用于判断空行的列。")
        if not config['api_url']:
            error_messages.append("API URL 不能为空。")
        if not config['api_key'] or config['api_key'] == "你的APIkey":
            error_messages.append("API Key 不能为空，且不能使用默认占位符 '你的APIkey'。")
        if not config['model']:
            error_messages.append("模型名称不能为空。")
        if not config['content_template']:
            error_messages.append("内容整合模板不能为空。")
        if not config['llm_template']:
            error_messages.append("LLM 提示词模板不能为空。")
        if not any(config['input_columns'].values()):
             error_messages.append("请至少选择一个输入列。")
        if not config['output_columns']:
            error_messages.append("请至少定义一个输出列。")
        
        if error_messages:
            QMessageBox.warning(self, "输入校验失败", "\n".join(error_messages))
            return False
        return True

    @Slot()
    def on_start_processing_clicked(self):
        if self.is_processing:
            QMessageBox.information(self, "提示", "处理已在进行中。")
            return
        
        if not self._validate_inputs():
            return

        self._save_config() 
        self.current_config_for_processing = self._gather_config_from_ui() 

        try:
            def progress_callback_from_thread(current, total):
                if self.processing_thread: 
                     self.processing_thread.progress_updated.emit(current, total)

            self.processor_instance = ExcelProcessor(
                self.current_config_for_processing,
                progress_callback_from_thread 
            )
        except Exception as e:
            self.log_message(f"初始化处理器失败: {e}")
            QMessageBox.critical(self, "错误", f"初始化处理器失败: {e}")
            return

        self.processing_thread = ProcessingThread(self.processor_instance, self)
        self.processing_thread.progress_updated.connect(self.on_progress_updated_slot)
        self.processing_thread.finished.connect(self.on_processing_finished_slot)
        
        self.is_processing = True
        self._set_ui_processing_state(True)
        self.log_message("开始处理...")
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(100) 
        self.processing_thread.start()

    @Slot()
    def on_stop_processing_clicked(self):
        if self.processor_instance and self.is_processing:
            self.log_message("正在发送停止请求...")
            self.processor_instance.stop_processing()
            self.stop_btn.setEnabled(False) 
            self.stop_btn.setText("停止中...")
        else:
            self.log_message("没有正在进行的处理任务。")

    @Slot(int, int)
    def on_progress_updated_slot(self, current_value, total_value):
        if total_value > 0:
            self.progress_bar.setMaximum(total_value)
            self.progress_bar.setValue(current_value)
        else:
            self.progress_bar.setMaximum(100) 
            self.progress_bar.setValue(0)


    @Slot(bool, str)
    def on_processing_finished_slot(self, success, message):
        self.is_processing = False
        self._set_ui_processing_state(False)
        self.log_message(message)
        
        if success:
            QMessageBox.information(self, "处理完成", message)
            self.progress_bar.setValue(self.progress_bar.maximum()) 
        else:
            if "用户停止" in message or "中止" in message : 
                 QMessageBox.warning(self, "处理中止", message)
            else:
                 QMessageBox.critical(self, "处理失败", message)
            self.progress_bar.setValue(0) 

        self.processor_instance = None
        if self.processing_thread:
            self.processing_thread.quit()
            self.processing_thread.wait()
        self.processing_thread = None


    def _set_ui_processing_state(self, processing: bool):
        self.start_btn.setEnabled(not processing)
        self.stop_btn.setEnabled(processing)
        if not processing:
            self.stop_btn.setText("停止处理") 

        for i in range(self.centralWidget().layout().itemAt(0).widget().count()):
            tab_widget_main = self.centralWidget().layout().itemAt(0).widget() 
            tab_page = tab_widget_main.widget(i) 
            if tab_widget_main.tabText(i) != "处理与日志":
                 tab_page.setEnabled(not processing)

    def log_message(self, message):
        if hasattr(self, 'result_text_edit') and self.result_text_edit is not None:
            self.result_text_edit.append(f"[{QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] {message}")
            self.result_text_edit.ensureCursorVisible() 
        else:
            print(f"LOG (UI not ready): [{QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] {message}")


    def closeEvent(self, event: QCloseEvent):
        if self.is_processing:
            reply = QMessageBox.question(self, "退出确认",
                                         "处理仍在进行中。确定要退出吗？\n未完成的处理将会丢失。",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                if self.processor_instance:
                    self.processor_instance.stop_processing() 
                if self.processing_thread and self.processing_thread.isRunning():
                    self.processing_thread.quit() 
                    self.processing_thread.wait(1000) 
                self._save_config() 
                event.accept()
            else:
                event.ignore()
        else:
            self._save_config() 
            event.accept()

def main():
    app = QApplication(sys.argv)
    window = ExcelProcessorGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
