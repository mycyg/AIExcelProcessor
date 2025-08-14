import sys
import json
from pathlib import Path
import pandas as pd
import dataclasses

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QSpinBox, QTextEdit, QFileDialog, QComboBox, 
    QCheckBox, QProgressBar, QGroupBox, QListWidget, QMessageBox, QScrollArea
)
from PySide6.QtCore import Qt, QThread, Signal as pyqtSignal, QPoint, QTimer, Slot, QDateTime
from PySide6.QtGui import QKeyEvent, QFocusEvent

from config import ProcessingConfig
from processor import ExcelProcessor

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

class ProcessingThread(QThread):
    progress = pyqtSignal(str, int, int)

    def __init__(self, config: ProcessingConfig, parent=None):
        super().__init__(parent)
        self.config = config
        self.processor = ExcelProcessor(self.config)

    def run(self):
        for msg_type, current, total in self.processor.start_processing():
            self.progress.emit(msg_type, current, total)

    def stop(self):
        self.processor.stop()

class ExcelProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_path = Path("config.json")
        self.processing_thread = None
        self.column_names = []
        self._init_ui()
        self._load_config_and_apply_to_ui()

    def _init_ui(self):
        self.setWindowTitle("Excel 批量处理工具 (Final Polished)")
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
        self.top_widget.setMaximumHeight(220)
        layout.addWidget(self.top_widget)

        self.prompts_group = self._create_prompts_group()
        layout.addWidget(self.prompts_group)

        self.process_group = self._create_process_group()
        self.process_group.setMaximumHeight(200)
        layout.addWidget(self.process_group)

        layout.setStretch(0, 0)
        layout.setStretch(1, 1)
        layout.setStretch(2, 0)

    def _create_file_group(self) -> QGroupBox:
        group = QGroupBox("文件设置")
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
        group = QGroupBox("API设置")
        layout = QVBoxLayout(group)
        api_url_label = QLabel("API URL:")
        api_url_label.setToolTip("输入大语言模型服务的API地址。")
        layout.addWidget(api_url_label)
        self.api_url_edit = QLineEdit()
        self.api_url_edit.setToolTip("输入大语言模型服务的API地址。")
        layout.addWidget(self.api_url_edit)
        api_key_label = QLabel("API Key:")
        api_key_label.setToolTip("输入你的API Key。")
        layout.addWidget(api_key_label)
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.api_key_edit.setToolTip("输入你的API Key。")
        layout.addWidget(self.api_key_edit)
        model_label = QLabel("模型名称:")
        model_label.setToolTip("输入要使用的模型名称，例如 doubao-pro-32k。")
        layout.addWidget(model_label)
        self.model_edit = QLineEdit()
        self.model_edit.setToolTip("输入要使用的模型名称，例如 doubao-pro-32k。")
        layout.addWidget(self.model_edit)
        process_layout = QHBoxLayout()
        batch_label = QLabel("批处理:")
        batch_label.setToolTip("将多个行打包成一个批次，作为一个任务进行处理。\n可以减少API调用次数，但过大的批次可能导致单次请求超时。")
        process_layout.addWidget(batch_label)
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setRange(1, 1000)
        self.batch_size_spin.setToolTip("将多个行打包成一个批次，作为一个任务进行处理。\n可以减少API调用次数，但过大的批次可能导致单次请求超时。")
        process_layout.addWidget(self.batch_size_spin)
        workers_label = QLabel("并行:")
        workers_label.setToolTip("同时处理多少个批次。\n可以显著提高处理速度，但过高的数量会增加CPU、内存消耗和API请求频率。")
        process_layout.addWidget(workers_label)
        self.workers_spin = QSpinBox()
        self.workers_spin.setRange(1, 100)
        self.workers_spin.setToolTip("同时处理多少个批次。\n可以显著提高处理速度，但过高的数量会增加CPU、内存消耗和API请求频率。")
        process_layout.addWidget(self.workers_spin)
        layout.addLayout(process_layout)
        return group

    def _create_columns_group(self) -> QGroupBox:
        group = QGroupBox("列设置")
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
        group = QGroupBox("提示词设置")
        main_layout = QHBoxLayout(group)

        content_group = QGroupBox("内容整合模板")
        content_group.setToolTip("定义如何将Excel单行数据整合成一段文本。\n使用 {row['列名']} 来引用列数据。")
        content_layout = QVBoxLayout(content_group)
        insert_column_btn = QPushButton("插入列引用")
        insert_column_btn.clicked.connect(lambda: self.content_template_edit.show_suggestions())
        content_layout.addWidget(insert_column_btn, alignment=Qt.AlignRight)
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
        group = QGroupBox("处理控制与日志")
        layout = QVBoxLayout(group)
        control_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始处理")
        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.setEnabled(False)
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

    def _load_config_and_apply_to_ui(self):
        try:
            config_data = json.loads(self.config_path.read_text(encoding='utf-8')) if self.config_path.exists() else {}
            config = ProcessingConfig(**config_data)
            self.log("配置文件加载成功。")
        except Exception as e:
            self.log(f"加载或解析配置文件失败: {e}。将使用默认设置。")
            config = ProcessingConfig()

        self.input_file_edit.setText(config.input_file)
        self.output_file_edit.setText(config.output_file)
        self.api_url_edit.setText(config.api_url)
        self.api_key_edit.setText(config.api_key)
        self.model_edit.setText(config.model)
        self.batch_size_spin.setValue(config.batch_size)
        self.workers_spin.setValue(config.workers)
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
            input_file=self.input_file_edit.text(),
            output_file=self.output_file_edit.text(),
            sheet_name=self.sheet_combo.currentText(),
            empty_column=self.empty_column_combo.currentText(),
            api_url=self.api_url_edit.text(),
            api_key=self.api_key_edit.text(),
            model=self.model_edit.text(),
            batch_size=self.batch_size_spin.value(),
            workers=self.workers_spin.value(),
            content_template=self.content_template_edit.toPlainText(),
            llm_template=self.llm_template_edit.toPlainText(),
            input_columns=input_columns,
            output_columns=[line.strip() for line in self.output_columns_edit.toPlainText().splitlines() if line.strip()]
        )

    def _save_config(self):
        config = self._gather_config_from_ui()
        try:
            self.config_path.write_text(json.dumps(dataclasses.asdict(config), ensure_ascii=False, indent=4), encoding='utf-8')
            self.log("配置已保存。")
        except Exception as e:
            self.log(f"保存配置失败: {e}")

    @Slot()
    def browse_input_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, "选择输入文件", "", "Excel Files (*.xlsx *.xls)")
        if filename:
            self.input_file_edit.setText(filename)
            self.update_sheets_from_file(filename)

    @Slot()
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

    @Slot()
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
            cb.setChecked(input_columns_to_check.get(col, False))
            self.input_columns_layout.insertWidget(self.input_columns_layout.count() - 1, cb)

    @Slot()
    def start_processing(self):
        if self.processing_thread and self.processing_thread.isRunning(): return
        config = self._gather_config_from_ui()
        if not all([config.input_file, config.output_file, config.sheet_name, config.api_key]):
            QMessageBox.warning(self, "校验失败", "请填写所有必要的文件和API设置。" )
            return

        self._save_config()
        self.set_ui_processing_state(False)
        self.log("开始处理...")
        self.processing_thread = ProcessingThread(config, self)
        self.processing_thread.progress.connect(self.on_progress_update)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()

    @Slot()
    def stop_processing(self):
        if self.processing_thread and self.processing_thread.isRunning():
            self.log("正在发送停止请求...")
            self.processing_thread.stop()
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("停止中...")

    @Slot(str, int, int)
    def on_progress_update(self, msg_type, current, total):
        if msg_type == "info":
            self.log("正在读取数据...")
            self.progress_bar.setMaximum(0)
        elif msg_type == "progress":
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(current)
        elif msg_type == "stopped":
            self.log(f"处理被用户中止。已处理 {current}/{total} 行。")
        elif msg_type == "finish":
            self.log(f"处理完成！共处理 {current}/{total} 行。")
            if total > 0: self.progress_bar.setValue(total)
        elif msg_type == "error":
            self.log("处理过程中发生严重错误。")

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
        timestamp = QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')
        self.log_edit.append(f"[{timestamp}] {message}")

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
    app = QApplication(sys.argv)
    window = ExcelProcessorGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
