import sys
import json
import os
from pathlib import Path
import pandas as pd
import dataclasses
import multiprocessing
import threading
from typing import List
import requests

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QSpinBox, QTextEdit, QFileDialog, QComboBox, 
    QCheckBox, QProgressBar, QGroupBox, QListWidget, QMessageBox, QScrollArea
)
from PySide6.QtCore import Qt, QThread, Signal as pyqtSignal, QPoint, QTimer, Slot, QDateTime
from PySide6.QtGui import QKeyEvent, QFocusEvent

from config import ProcessingConfig
from processor import ExcelProcessor
from volcengine_processor import VolcengineProcessor

# CustomTextEditWithSuggestions remains the same
class CustomTextEditWithSuggestions(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._suggestions_list_widget = QListWidget(self)
        self._suggestions_list_widget.setWindowFlag(Qt.Popup)
        self._suggestions_list_widget.setFocusPolicy(Qt.NoFocus)
        self._suggestions_list_widget.setFocusProxy(self)
        self._suggestions_list_widget.hide()
        self._suggestion_items = []
        self._suggestions_list_widget.itemClicked.connect(self._on_suggestion_item_clicked)

    def set_suggestion_items(self, items):
        self._suggestion_items = items

    def _update_suggestions_list(self):
        self._suggestions_list_widget.clear()
        if not self._suggestion_items:
            self._suggestions_list_widget.addItem("无可用列名")
        else:
            self._suggestions_list_widget.addItems(self._suggestion_items)

    def keyPressEvent(self, event: QKeyEvent):
        if self._suggestions_list_widget.isVisible():
            if event.key() in (Qt.Key_Enter, Qt.Key_Return, Qt.Key_Tab):
                if self._suggestions_list_widget.currentItem():
                    self._insert_selected_suggestion(self._suggestions_list_widget.currentItem().text())
                self._suggestions_list_widget.hide()
                return
            elif event.key() == Qt.Key_Escape:
                self._suggestions_list_widget.hide()
                return
            elif event.key() in (Qt.Key_Up, Qt.Key_Down, Qt.Key_PageUp, Qt.Key_PageDown):
                self._suggestions_list_widget.keyPressEvent(event)
                return

        if event.text() == '/':
            super().keyPressEvent(event)
            self.show_suggestions()
            return

        super().keyPressEvent(event)

    def focusOutEvent(self, event: QFocusEvent):
        if self._suggestions_list_widget.isVisible() and not self._suggestions_list_widget.hasFocus():
             self._suggestions_list_widget.hide()
        super().focusOutEvent(event)

    def show_suggestions(self):
        self._update_suggestions_list()
        if self._suggestions_list_widget.count() == 0:
            return
        if self._suggestions_list_widget.count() == 1 and self._suggestions_list_widget.item(0).text() == "无可用列名":
            return

        cursor_rect = self.cursorRect()
        popup_pos = self.mapToGlobal(QPoint(cursor_rect.left(), cursor_rect.bottom()))
        self._suggestions_list_widget.move(popup_pos)
        self._suggestions_list_widget.setMinimumWidth(self.width() // 2)
        self._suggestions_list_widget.show()
        self._suggestions_list_widget.setFocus()

    def _on_suggestion_item_clicked(self, item):
        self._insert_selected_suggestion(item.text())
        self._suggestions_list_widget.hide()

    def _insert_selected_suggestion(self, suggestion_text):
        if suggestion_text == "无可用列名":
            return
        cursor = self.textCursor()
        if cursor.position() > 0 and self.toPlainText()[cursor.position() - 1] == '/':
            cursor.deletePreviousChar()
        cursor.insertText(f"{{row['{suggestion_text}']}}")
        self.setTextCursor(cursor)

class StandardProcessingThread(QThread):
    progress = pyqtSignal(str, object, object)

    def __init__(self, config: ProcessingConfig, parent=None):
        super().__init__(parent)
        self.config = config
        self.processor = ExcelProcessor(self.config)

    def run(self):
        for msg_type, data, total in self.processor.start_processing():
            self.progress.emit(msg_type, data, total)

    def stop(self):
        self.processor.stop()

class VolcengineProcessingThread(QThread):
    progress = pyqtSignal(str, object, object)

    def __init__(self, config: ProcessingConfig, parent=None):
        super().__init__(parent)
        self.config = config
        self.processor = None
        self._is_running = True

    def run(self):
        try:
            progress_queue = multiprocessing.Queue()
            self.processor = VolcengineProcessor(self.config, progress_queue)
            
            proc_manager_thread = threading.Thread(target=self.processor.run, daemon=True)
            proc_manager_thread.start()

            processed_count = 0
            total_rows = -1 # Will be set by a message from the queue

            while self._is_running:
                try:
                    # Poll the queue for updates
                    msg = progress_queue.get(timeout=1)
                    if isinstance(msg, tuple) and len(msg) == 3:
                        msg_type, data, total = msg
                        if msg_type == "total_rows":
                            total_rows = data
                            self.progress.emit("progress", processed_count, total_rows)
                        else:
                            self.progress.emit(msg_type, data, total)
                    elif isinstance(msg, int): # Progress update
                        processed_count += msg
                        if total_rows != -1:
                            self.progress.emit("progress", processed_count, total_rows)

                except multiprocessing.queues.Empty:
                    if not proc_manager_thread.is_alive():
                        self.progress.emit("info", "处理进程已完成。", 0)
                        break # Exit the polling loop
                    continue
            
            # Final update
            self.progress.emit("finish", processed_count, total_rows if total_rows != -1 else processed_count)

        except Exception as e:
            self.progress.emit("error", f"火山引擎处理器线程错误: {e}", 0)

    def stop(self):
        self._is_running = False
        self.progress.emit("info", "多进程模式不支持中途停止，将在当前任务完成后结束。", 0)

class ExcelProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_path = Path("config.json")
        self.processing_thread = None
        self.column_names = []
        self._init_ui()
        self._load_config_and_apply_to_ui()
        self.on_mode_changed() # Set initial UI state based on mode

    def _init_ui(self):
        self.setWindowTitle("Excel 批量处理工具 v2.1 (Scalable)")
        self.setMinimumSize(1000, 800)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        top_settings_layout = QHBoxLayout()
        top_settings_layout.addWidget(self._create_file_group(), stretch=2)
        top_settings_layout.addWidget(self._create_api_group(), stretch=2)
        top_settings_layout.addWidget(self._create_columns_group(), stretch=3)
        
        self.top_widget = QWidget()
        self.top_widget.setLayout(top_settings_layout)
        self.top_widget.setMaximumHeight(240)
        layout.addWidget(self.top_widget)

        self.prompts_group = self._create_prompts_group()
        layout.addWidget(self.prompts_group)

        self.process_group = self._create_process_group()
        self.process_group.setMaximumHeight(250)
        layout.addWidget(self.process_group)

        layout.setStretch(0, 0)
        layout.setStretch(1, 1)
        layout.setStretch(2, 0)

    def _create_file_group(self) -> QGroupBox:
        group = QGroupBox("第一步：文件设置")
        layout = QVBoxLayout(group)
        input_layout = QHBoxLayout()
        input_label = QLabel("输入:")
        input_label.setToolTip("选择要处理的源Excel文件 (.xlsx, .xls)。")
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setToolTip("选择要处理的源Excel文件 (.xlsx, .xls)。")
        input_browse_btn = QPushButton("浏览...")
        input_browse_btn.clicked.connect(self.browse_input_file)
        input_layout.addWidget(input_label)
        input_layout.addWidget(self.input_file_edit)
        input_layout.addWidget(input_browse_btn)
        layout.addLayout(input_layout)
        output_layout = QHBoxLayout()
        output_label = QLabel("输出:")
        output_label.setToolTip("设置处理结果要保存到的Excel文件路径。")
        self.output_file_edit = QLineEdit()
        self.output_file_edit.setToolTip("设置处理结果要保存到的Excel文件路径。")
        output_browse_btn = QPushButton("浏览...")
        output_browse_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_file_edit)
        output_layout.addWidget(output_browse_btn)
        layout.addLayout(output_layout)
        sheet_label = QLabel("Sheet 名称:")
        sheet_label.setToolTip("从输入文件中选择要处理的Sheet（工作表）。")
        layout.addWidget(sheet_label)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setToolTip("从输入文件中选择要处理的Sheet（工作表）。")
        self.sheet_combo.currentIndexChanged.connect(self.update_columns_from_sheet)
        layout.addWidget(self.sheet_combo)
        empty_col_label = QLabel("判断空行的列:")
        empty_col_label.setToolTip("选择一个列作为判断依据。如果这一列没有数据，则该行将被跳过，不进行处理。")
        layout.addWidget(empty_col_label)
        self.empty_column_combo = QComboBox()
        self.empty_column_combo.setToolTip("选择一个列作为判断依据。如果这一列没有数据，则该行将被跳过，不进行处理。")
        layout.addWidget(self.empty_column_combo)
        return group

    def _create_api_group(self) -> QGroupBox:
        group = QGroupBox("第二步：API与处理模式")
        layout = QVBoxLayout(group)

        mode_layout = QHBoxLayout()
        mode_label = QLabel("处理模式:")
        mode_label.setToolTip("标准模式使用通用requests库，火山引擎模式使用官方SDK。")
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["标准模式", "火山引擎SDK模式"])
        self.mode_combo.setToolTip("标准模式使用通用requests库，火山引擎模式使用官方SDK。")
        self.mode_combo.currentTextChanged.connect(self.on_mode_changed)
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.mode_combo)
        layout.addLayout(mode_layout)

        self.api_url_label = QLabel("API URL (仅标准模式):")
        layout.addWidget(self.api_url_label)
        self.api_url_edit = QLineEdit()
        layout.addWidget(self.api_url_edit)
        api_key_label = QLabel("API Key (火山模式下可留空):")
        api_key_label.setToolTip("火山模式下如果留空，将尝试读取环境变量 ARK_API_KEY。")
        layout.addWidget(api_key_label)
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.api_key_edit)
        model_label = QLabel("模型名称/Endpoint ID:")
        layout.addWidget(model_label)
        self.model_edit = QLineEdit()
        layout.addWidget(self.model_edit)
        
        proc_config_layout = QHBoxLayout()
        self.batch_label = QLabel("批处理:")
        proc_config_layout.addWidget(self.batch_label)
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setRange(1, 1000)
        proc_config_layout.addWidget(self.batch_size_spin)
        self.workers_label = QLabel("并行/进程数:")
        proc_config_layout.addWidget(self.workers_label)
        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(1, 1000)
        proc_config_layout.addWidget(self.workers_spin)
        layout.addLayout(proc_config_layout)

        self.timeout_label = QLabel("API超时(秒):" )
        self.api_timeout_spin = QSpinBox()
        self.api_timeout_spin.setRange(10, 9999)
        self.api_timeout_spin.setValue(180)
        self.timeout_widget = QWidget()
        timeout_layout = QHBoxLayout(self.timeout_widget)
        timeout_layout.addWidget(self.timeout_label)
        timeout_layout.addWidget(self.api_timeout_spin)
        timeout_layout.setContentsMargins(0,0,0,0)
        layout.addWidget(self.timeout_widget)

        return group

    def _create_columns_group(self) -> QGroupBox:
        group = QGroupBox("第三步：列设置")
        main_layout = QHBoxLayout(group)
        input_group = QGroupBox("输入列")
        input_group.setToolTip("勾选所有需要参与内容整合的列。")
        input_group_layout = QVBoxLayout(input_group)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        self.input_columns_layout = QVBoxLayout(scroll_widget)
        self.input_columns_layout.addStretch()
        scroll_area.setWidget(scroll_widget)
        input_group_layout.addWidget(scroll_area)
        main_layout.addWidget(input_group)

        output_group = QGroupBox("输出列")
        output_group.setToolTip("定义希望LLM为你生成的新列的名称，每行一个。")
        output_layout = QVBoxLayout(output_group)
        self.output_columns_edit = QTextEdit()
        self.output_columns_edit.setPlaceholderText("每行一个")
        output_layout.addWidget(self.output_columns_edit)
        main_layout.addWidget(output_group)
        return group

    def _create_prompts_group(self) -> QGroupBox:
        group = QGroupBox("第四步：提示词设置")
        main_layout = QHBoxLayout(group)

        content_group = QGroupBox("内容整合模板")
        content_group.setToolTip("定义如何将Excel单行数据整合成一段文本。\n使用 {row['列名']} 来引用列数据。")
        content_layout = QVBoxLayout(content_group)
        
        content_top_layout = QHBoxLayout()
        self.generate_prompt_btn = QPushButton("一键配置所有模板")
        self.generate_prompt_btn.setToolTip("根据Excel数据和列设置，自动生成下方两个模板的内容。")
        self.generate_prompt_btn.clicked.connect(self.generate_llm_template)
        content_top_layout.addWidget(self.generate_prompt_btn)
        content_top_layout.addStretch()
        insert_column_btn = QPushButton("插入列引用")
        insert_column_btn.clicked.connect(lambda: self.content_template_edit.show_suggestions())
        content_top_layout.addWidget(insert_column_btn)
        content_layout.addLayout(content_top_layout)

        self.content_template_edit = CustomTextEditWithSuggestions(self)
        content_layout.addWidget(self.content_template_edit)
        main_layout.addWidget(content_group)

        llm_group = QGroupBox("LLM 提示词模板")
        llm_group.setToolTip("定义发送给大模型的主提示词。\n使用 {{content}} 来引用由“内容整合模板”生成的那段文本。")
        llm_layout = QVBoxLayout(llm_group)
        self.llm_template_edit = QTextEdit()
        llm_layout.addWidget(self.llm_template_edit)
        main_layout.addWidget(llm_group)
        return group

    def _create_process_group(self) -> QGroupBox:
        group = QGroupBox("第五步：处理控制与日志")
        layout = QVBoxLayout(group)
        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始处理")
        self.start_btn.clicked.connect(self.start_processing)
        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_processing)
        control_layout.addWidget(self.start_btn)
        control_layout.addWidget(self.stop_btn)
        control_layout.addStretch()
        layout.addLayout(control_layout)
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        layout.addWidget(self.log_edit)
        return group

    def _call_llm_for_full_configuration(self, api_key: str, model: str, api_url: str, timeout: int, all_columns: List[str], input_columns: List[str], output_columns: List[str], raw_data_examples: str) -> str:
        """Calls the LLM to generate both the content and LLM prompt templates."""
        self.log("正在向LLM发送请求以生成完整配置...")

        system_prompt = (
            "你是一位专家级的系统架构师和提示词工程师。你的任务是为数据处理流水线自动完成设置。基于用户提供的原始数据结构以及他们期望的输入和输出，你必须生成两个组件：\n" 
            "1.  一个“内容整合模板” (content_integration_template): 这是一个字符串模板，用于将单行原始数据格式化为一段连贯的文本。此模板必须使用 `{row['列名']}` 的格式作为占位符。\n" 
            "2.  一个“LLM提示词模板” (llm_prompt_template): 这是一套给另一个LLM的详细指令，告诉它如何处理由“内容整合模板”生成的文本，以提取出所有期望的输出字段。此模板必须包含一个 `{{content}}` 占位符。\n\n" 
            "你的回复必须是一个单独的、格式严格的JSON对象，且只包含 `content_integration_template` 和 `llm_prompt_template` 这两个键。不要在JSON对象之外包含任何解释、标题或任何其他文字。"
        )

        all_columns_str = ", ".join(all_columns)
        input_columns_str = ", ".join(input_columns)
        output_columns_str = ", ".join(output_columns)

        user_prompt = (
            f"请根据以下关于数据处理任务的信息，为我生成所需的JSON配置对象：\n\n"
            f"1. **源文件中的所有可用列:**\n   {all_columns_str}\n\n"
            f"2. **用户选择用于任务的输入列:**\n   {input_columns_str}\n\n"
            f"3. **用户期望LLM生成的输出列:**\n   {output_columns_str}\n\n"
            f"4. **头两行原始数据样例 (JSON格式):**\n```json\n{raw_data_examples}\n```\n\n"
            "现在，请生成包含 `content_integration_template` 和 `llm_prompt_template` 键的JSON对象。"
        )

        # Combine prompts and structure for multimodal format
        final_prompt_text = f"{system_prompt}\n\n{user_prompt}"

        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        data = {
            'model': model,
            'messages': [
                {
                    'role': 'user',
                    'content': [
                        {'type': 'text', 'text': final_prompt_text}
                    ]
                }
            ],
            'response_format': {'type': 'json_object'}
        }

        try:
            response = requests.post(api_url, headers=headers, json=data, timeout=timeout)
            response.raise_for_status()
            self.log("成功收到LLM的响应。")
            return response.text # Return the raw text to be parsed
        except requests.exceptions.RequestException as e:
            error_message = f"调用LLM API失败: {e}"
            self.log(f"[错误] {error_message}")
            QMessageBox.critical(self, "API 调用失败", error_message)
            return f"API_ERROR: {e}"
        except Exception as e: # Catch other potential errors like JSON parsing in the response
            error_message = f"处理API响应时出错: {e}"
            self.log(f"[错误] {error_message}")
            QMessageBox.critical(self, "API 响应处理失败", error_message)
            return f"PARSE_ERROR: {e}"

    @Slot()
    def generate_llm_template(self):
        self.log("开始一键配置模板...")
        config = self._gather_config_from_ui()

        # 1. Validation
        if not all([config.input_file, config.sheet_name]):
            QMessageBox.warning(self, "校验失败", "请先完成“第一步：文件设置”中的所有必填项。")
            self.log("[警告] 用户未完成文件设置。")
            return

        if not any(config.input_columns.values()):
            QMessageBox.warning(self, "校验失败", "请在“第三步：列设置”中至少勾选一个输入列。")
            self.log("[警告] 用户未选择任何输入列。")
            return
        
        if not config.output_columns:
            QMessageBox.warning(self, "校验失败", "请在“第三步：列设置”中定义至少一个输出列。")
            self.log("[警告] 用户未定义任何输出列。")
            return
            
        if config.processing_mode == "标准模式" and not config.api_key:
            QMessageBox.warning(self, "校验失败", "标准模式下必须填写API Key才能与LLM通信。")
            self.log("[警告] 标准模式下缺少API Key。")
            return

        # 2. Read Excel Data
        try:
            self.log(f"正在读取文件: {config.input_file} (Sheet: {config.sheet_name})")
            df = pd.read_excel(config.input_file, sheet_name=config.sheet_name, nrows=2)
            if df.empty:
                QMessageBox.warning(self, "文件内容不足", "Excel文件中没有足够的数据行（需要至少1行）来生成示例。")
                self.log("[警告] Excel文件数据行不足。")
                return
            raw_data_examples_str = df.to_json(orient='records', indent=2, force_ascii=False)
        except Exception as e:
            error_message = f"读取Excel文件失败: {e}"
            self.log(f"[错误] {error_message}")
            QMessageBox.critical(self, "文件读取失败", error_message)
            return

        # 3. Call LLM for full configuration
        if config.processing_mode != "标准模式":
            QMessageBox.information(self, "模式不支持", "此功能目前仅在“标准模式”下可用。")
            self.log("[信息] 用户尝试在非标准模式下使用此功能。")
            return

        selected_input_columns = [col for col, is_checked in config.input_columns.items() if is_checked]

        llm_response_text = self._call_llm_for_full_configuration(
            api_key=config.api_key,
            model=config.model,
            api_url=config.api_url,
            timeout=config.api_timeout,
            all_columns=self.column_names,
            input_columns=selected_input_columns,
            output_columns=config.output_columns,
            raw_data_examples=raw_data_examples_str
        )

        if not llm_response_text or llm_response_text.startswith("API_ERROR:") or llm_response_text.startswith("PARSE_ERROR:"):
            return

        # 4. Parse JSON and Update UI
        try:
            self.log("正在解析LLM返回的JSON配置...")
            
            # First, parse the outer response from the API
            outer_response = json.loads(llm_response_text)
            
            # Extract the inner JSON string from the content field
            inner_json_string = outer_response['choices'][0]['message']['content']
            
            # Now, parse the inner JSON string to get our templates
            templates = json.loads(inner_json_string)

            content_template = templates.get("content_integration_template")
            llm_template = templates.get("llm_prompt_template")

            if not all([content_template, llm_template]):
                raise KeyError("LLM生成的JSON中缺少 'content_integration_template' 或 'llm_prompt_template' 键。")

            self.content_template_edit.setPlainText(content_template)
            self.llm_template_edit.setPlainText(llm_template)
            self.log("所有模板已成功生成并填充！")
            QMessageBox.information(self, "生成成功", "内容整合与LLM提示词模板均已成功生成！")

        except json.JSONDecodeError as e:
            error_message = f"解析LLM返回的JSON失败: {e}。\n收到的原文: \n{llm_response_text}"
            self.log(f"[错误] {error_message}")
            QMessageBox.critical(self, "JSON解析失败", error_message)
        except (KeyError, IndexError) as e:
            error_message = f"LLM返回的JSON结构不正确，无法找到所需内容: {e}。\n收到的原文: \n{llm_response_text}"
            self.log(f"[错误] {error_message}")
            QMessageBox.critical(self, "JSON格式错误", error_message)

    @Slot()
    def on_mode_changed(self):
        is_standard_mode = self.mode_combo.currentText() == "标准模式"
        self.api_url_label.setVisible(is_standard_mode)
        self.api_url_edit.setVisible(is_standard_mode)
        self.timeout_widget.setVisible(is_standard_mode)
        self.batch_label.setVisible(is_standard_mode)
        self.batch_size_spin.setVisible(is_standard_mode)
        self.workers_label.setText("并行线程数:" if is_standard_mode else "并行进程数:")

    def _load_config_and_apply_to_ui(self):
        try:
            config_data = json.loads(self.config_path.read_text(encoding='utf-8')) if self.config_path.exists() else {}
            config = ProcessingConfig(**config_data)
            self.log("配置文件加载成功。" )
        except Exception as e:
            self.log(f"加载或解析配置文件失败: {e}。将使用默认设置。")
            config = ProcessingConfig()

        self.mode_combo.setCurrentText(config.processing_mode)
        self.input_file_edit.setText(config.input_file)
        self.output_file_edit.setText(config.output_file)
        self.api_url_edit.setText(config.api_url)
        self.api_key_edit.setText(config.api_key)
        self.model_edit.setText(config.model)
        self.batch_size_spin.setValue(config.batch_size)
        self.workers_spin.setValue(config.workers)
        self.api_timeout_spin.setValue(config.api_timeout)
        self.content_template_edit.setPlainText(config.content_template)
        self.llm_template_edit.setPlainText(config.llm_template)
        self.output_columns_edit.setPlainText("\n".join(config.output_columns))

        if config.input_file:
            self.update_sheets_from_file(config.input_file)
            if config.sheet_name in [self.sheet_combo.itemText(i) for i in range(self.sheet_combo.count())]:
                self.sheet_combo.setCurrentText(config.sheet_name)
            self.update_columns_from_sheet(config.input_columns)

    def _gather_config_from_ui(self) -> ProcessingConfig:
        input_columns = {}
        if self.input_columns_layout:
            for i in range(self.input_columns_layout.count() - 1):
                widget = self.input_columns_layout.itemAt(i).widget()
                if isinstance(widget, QCheckBox):
                    input_columns[widget.text()] = widget.isChecked()

        return ProcessingConfig(
            processing_mode=self.mode_combo.currentText(),
            input_file=self.input_file_edit.text(),
            output_file=self.output_file_edit.text(),
            sheet_name=self.sheet_combo.currentText(),
            empty_column=self.empty_column_combo.currentText(),
            api_url=self.api_url_edit.text(),
            api_key=self.api_key_edit.text(),
            model=self.model_edit.text(),
            batch_size=self.batch_size_spin.value(),
            workers=self.workers_spin.value(),
            api_timeout=self.api_timeout_spin.value(),
            content_template=self.content_template_edit.toPlainText(),
            llm_template=self.llm_template_edit.toPlainText(),
            input_columns=input_columns,
            output_columns=[line.strip() for line in self.output_columns_edit.toPlainText().splitlines() if line.strip()]
        )

    def _save_config(self):
        config = self._gather_config_from_ui()
        try:
            self.config_path.write_text(json.dumps(dataclasses.asdict(config), ensure_ascii=False, indent=4), encoding='utf-8')
            self.log("配置已保存。" )
        except Exception as e:
            self.log(f"保存配置失败: {e}")

    @Slot()
    def start_processing(self):
        if self.processing_thread and self.processing_thread.isRunning(): return
        config = self._gather_config_from_ui()
        if not all([config.input_file, config.output_file, config.sheet_name]):
            QMessageBox.warning(self, "校验失败", "请填写所有必要的文件设置。" )
            return
        if config.processing_mode == "标准模式" and not config.api_key:
             QMessageBox.warning(self, "校验失败", "标准模式下必须填写API Key。" )
             return

        self._save_config()
        self.set_ui_processing_state(False)
        self.log(f"开始处理 (模式: {config.processing_mode})...")

        if config.processing_mode == "火山引擎SDK模式":
            self.processing_thread = VolcengineProcessingThread(config, self)
        else: # Standard Mode
            self.processing_thread = StandardProcessingThread(config, self)
        
        self.processing_thread.progress.connect(self.on_progress_update)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()

    @Slot(object, object, object)
    def on_progress_update(self, msg_type, data, total):
        log_data = str(data)
        if msg_type == "info":
            self.log(log_data)
        elif msg_type == "progress":
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(data)
            QApplication.processEvents()
        elif msg_type == "total_rows":
            self.progress_bar.setMaximum(data)
        elif msg_type == "stopped":
            self.log(f"处理被用户中止。已处理 {log_data}/{total} 行。" )
        elif msg_type == "finish":
            self.log(f"处理完成！共处理 {log_data}/{total} 行。" )
            if total > 0: self.progress_bar.setValue(total)
        elif msg_type == "error":
            self.log(f"[错误] {log_data}")
        elif msg_type == "debug_prompt":
            self.log(f"\n--- 提交给 LLM 的内容 ---\n{log_data}\n--------------------------")
        elif msg_type == "debug_response":
            if isinstance(data, dict):
                log_data = json.dumps(data, ensure_ascii=False, indent=2)
            self.log(f"\n--- LLM 返回的原文 ---\n{log_data}\n--------------------------")

    def browse_input_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, "选择输入文件", "", "Excel Files (*.xlsx *.xls)")
        if filename:
            self.input_file_edit.setText(filename)
            self.update_sheets_from_file(filename)

    def browse_output_file(self):
        filename, _ = QFileDialog.getSaveFileName(self, "选择输出文件", "", "Excel Files (*.xlsx)")
        if filename:
            self.output_file_edit.setText(filename)

    def update_sheets_from_file(self, filename):
        self.sheet_combo.clear()
        try:
            self.sheet_combo.addItems(pd.ExcelFile(filename).sheet_names)
        except Exception as e:
            self.log(f"读取文件失败: {e}")

    def update_columns_from_sheet(self, initial_input_columns=None):
        input_file = self.input_file_edit.text()
        sheet_name = self.sheet_combo.currentText()
        self.column_names = []
        if input_file and sheet_name:
            try:
                self.column_names = [str(col) for col in pd.read_excel(input_file, sheet_name=sheet_name, nrows=0).columns]
            except Exception as e:
                self.log(f"读取Sheet失败: {e}")

        self.content_template_edit.set_suggestion_items(self.column_names)
        self.empty_column_combo.clear()
        self.empty_column_combo.addItems(self.column_names)
        
        config = self._gather_config_from_ui()
        if config.empty_column in self.column_names:
            self.empty_column_combo.setCurrentText(config.empty_column)

        while self.input_columns_layout.count() > 1:
            item = self.input_columns_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        
        input_columns_to_check = initial_input_columns or config.input_columns
        for col in self.column_names:
            cb = QCheckBox(col)
            cb.setChecked(input_columns_to_check.get(col, True))
            self.input_columns_layout.insertWidget(self.input_columns_layout.count() - 1, cb)

    @Slot()
    def stop_processing(self):
        if self.processing_thread and self.processing_thread.isRunning():
            self.log("正在发送停止请求...")
            self.processing_thread.stop()
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("停止中...")

    @Slot()
    def on_processing_finished(self):
        self.set_ui_processing_state(True)
        self.processing_thread = None

    def set_ui_processing_state(self, enabled: bool):
        self.start_btn.setEnabled(enabled)
        self.stop_btn.setEnabled(not enabled)
        if not enabled: self.stop_btn.setText("停止中...")
        
        self.top_widget.setEnabled(enabled)
        self.prompts_group.setEnabled(enabled)

    def log(self, message: str):
        if not isinstance(message, str):
            message = str(message)
        timestamp = QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')
        self.log_edit.append(f"[{timestamp}] {message}")
        QApplication.processEvents()

    def closeEvent(self, event):
        if self.processing_thread and self.processing_thread.isRunning():
            reply = QMessageBox.question(self, "退出确认", "处理仍在进行中，确定要退出吗？",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.stop_processing()
                self.processing_thread.wait(1000)
                self._save_config()
                event.accept()
            else:
                event.ignore()
        else:
            self._save_config()
            event.accept()

def main():
    multiprocessing.freeze_support()
    app = QApplication(sys.argv)
    window = ExcelProcessorGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
