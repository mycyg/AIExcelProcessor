import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import threading
from pathlib import Path
import pandas as pd
import traceback
from .processor import ExcelProcessor
from .column_manager import ColumnManager
from .prompt_manager import PromptManager

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel批量处理工具")
        self.root.geometry("1000x900")
        self.root.minsize(800, 600)  # 设置最小窗口大小
        
        self.config_path = Path("config.json")
        
        # 创建主滚动画布
        self.main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.main_canvas, padding="10")
        
        # 配置画布滚动
        self.main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # 放置滚动条和画布
        scrollbar.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)
        
        # 将主框架添加到画布
        self.canvas_frame = self.main_canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # 绑定事件处理方法
        self.setup_scroll_events()
        
        # 创建所有UI组件
        self.create_file_section()
        self.create_api_section()
        self.create_settings_section()
        self.create_process_section()
        
        # 创建列管理器和提示词管理器
        self.column_manager = ColumnManager(self.main_frame)
        self.prompt_manager = PromptManager(self.main_frame)  # 初始时列名列表为空
        
        self.create_control_section()
        self.create_progress_section()
        
        # 处理器状态
        self.processor = None
        self.is_processing = False
        
        # 加载配置
        self.load_config()
        
        # 设置关闭窗口处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_scroll_events(self):
        """设置滚动相关的事件绑定"""
        self.main_frame.bind("<Configure>", self.on_frame_configure)
        self.main_canvas.bind("<Configure>", self.on_canvas_configure)
        self.main_canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        
        # 绑定按键滚动
        self.root.bind("<Up>", lambda e: self.main_canvas.yview_scroll(-1, "units"))
        self.root.bind("<Down>", lambda e: self.main_canvas.yview_scroll(1, "units"))
        self.root.bind("<Prior>", lambda e: self.main_canvas.yview_scroll(-1, "pages"))
        self.root.bind("<Next>", lambda e: self.main_canvas.yview_scroll(1, "pages"))

    def on_frame_configure(self, event=None):
        """当框架大小改变时，更新画布的滚动区域"""
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def on_canvas_configure(self, event=None):
        """当画布大小改变时，调整框架大小"""
        width = event.width if event else self.main_canvas.winfo_width()
        self.main_canvas.itemconfig(self.canvas_frame, width=width)

    def on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        # Windows系统下的滚动
        if event.delta:
            self.main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        # Linux/macOS系统下的滚动
        else:
            if event.num == 4:
                self.main_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.main_canvas.yview_scroll(1, "units")



    def create_file_section(self):
        """创建文件选择区域"""
        file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="5")
        file_frame.pack(fill="x", padx=5, pady=5)

        # 输入文件
        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, padx=5, pady=5)
        self.input_file = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_file, width=50).grid(row=0, column=1, sticky="ew")
        ttk.Button(file_frame, text="浏览", command=self.browse_input).grid(row=0, column=2, padx=5)

        # 输出文件
        ttk.Label(file_frame, text="输出文件:").grid(row=1, column=0, padx=5, pady=5)
        self.output_file = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_file, width=50).grid(row=1, column=1, sticky="ew")
        ttk.Button(file_frame, text="浏览", command=self.browse_output).grid(row=1, column=2, padx=5)

        file_frame.grid_columnconfigure(1, weight=1)

    def create_api_section(self):
        """创建API配置区域"""
        api_frame = ttk.LabelFrame(self.main_frame, text="API配置", padding="5")
        api_frame.pack(fill="x", padx=5, pady=5)

        # API URL
        ttk.Label(api_frame, text="API URL:").grid(row=0, column=0, padx=5, pady=5)
        self.api_url = tk.StringVar(value="https://ark.cn-beijing.volces.com/api/v3/chat/completions")
        ttk.Entry(api_frame, textvariable=self.api_url, width=50).grid(row=0, column=1, sticky="ew")

        # API Key
        ttk.Label(api_frame, text="API Key:").grid(row=1, column=0, padx=5, pady=5)
        self.api_key = tk.StringVar(value="")
        ttk.Entry(api_frame, textvariable=self.api_key, width=50).grid(row=1, column=1, sticky="ew")

        # Model
        ttk.Label(api_frame, text="模型:").grid(row=2, column=0, padx=5, pady=5)
        self.model = tk.StringVar(value="doubao-1-5-pro-32k-250115")
        ttk.Entry(api_frame, textvariable=self.model, width=50).grid(row=2, column=1, sticky="ew")

        api_frame.grid_columnconfigure(1, weight=1)

    def create_settings_section(self):
        """创建基本设置区域"""
        settings_frame = ttk.LabelFrame(self.main_frame, text="基本设置", padding="5")
        settings_frame.pack(fill="x", padx=5, pady=5)

        # Sheet名称
        ttk.Label(settings_frame, text="Sheet名称:").grid(row=0, column=0, padx=5, pady=5)
        self.sheet_name = ttk.Combobox(settings_frame, width=47)
        self.sheet_name.grid(row=0, column=1, sticky="ew")
        self.sheet_name.bind('<<ComboboxSelected>>', self.on_sheet_selected)

        # 判断空行的列
        ttk.Label(settings_frame, text="判断空行的列:").grid(row=1, column=0, padx=5, pady=5)
        self.empty_column = ttk.Combobox(settings_frame, width=47)
        self.empty_column.grid(row=1, column=1, sticky="ew")

        settings_frame.grid_columnconfigure(1, weight=1)

    def create_process_section(self):
        """创建处理配置区域"""
        process_frame = ttk.LabelFrame(self.main_frame, text="处理配置", padding="5")
        process_frame.pack(fill="x", padx=5, pady=5)

        # 批处理数量
        ttk.Label(process_frame, text="批处理数量:").grid(row=0, column=0, padx=5, pady=5)
        self.batch_size = tk.StringVar(value="20")
        ttk.Entry(process_frame, textvariable=self.batch_size, width=10).grid(row=0, column=1, sticky="w")

        # 并行处理数量
        ttk.Label(process_frame, text="并行处理数量:").grid(row=1, column=0, padx=5, pady=5)
        self.workers = tk.StringVar(value="20")
        ttk.Entry(process_frame, textvariable=self.workers, width=10).grid(row=1, column=1, sticky="w")

        process_frame.grid_columnconfigure(1, weight=1)
        
    def create_control_section(self):
        """创建控制按钮区域"""
        control_frame = ttk.Frame(self.main_frame)
        control_frame.pack(fill="x", padx=5, pady=5)

        ttk.Button(control_frame, text="开始处理", command=self.start_processing).pack(side="left", padx=5)
        ttk.Button(control_frame, text="停止处理", command=self.stop_processing).pack(side="left", padx=5)
        ttk.Button(control_frame, text="保存配置", command=self.save_config).pack(side="right", padx=5)

    def create_progress_section(self):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(self.main_frame, text="处理进度", padding="5")
        progress_frame.pack(fill="x", padx=5, pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill="x", padx=5, pady=5)

        # 添加一个标签用于显示详细的进度信息
        self.progress_info_label = ttk.Label(progress_frame, text="当前处理行数: 0 / 总行数: 0")
        self.progress_info_label.pack(fill="x", padx=5, pady=5)

    def browse_input(self):
        """选择输入文件"""
        filename = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.input_file.set(filename)
            self.update_sheet_names()

    def browse_output(self):
        """选择输出文件"""
        filename = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.output_file.set(filename)

    def update_sheet_names(self):
        """更新Sheet名称下拉列表"""
        try:
            if self.input_file.get():
                xl = pd.ExcelFile(self.input_file.get())
                self.sheet_name['values'] = xl.sheet_names
                if xl.sheet_names:
                    self.sheet_name.set(xl.sheet_names[0])
                    self.on_sheet_selected(None)
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败: {str(e)}")

    def on_sheet_selected(self, event):
        """当选择Sheet时更新列选择"""
        try:
            if self.input_file.get() and self.sheet_name.get():
                df = pd.read_excel(self.input_file.get(), sheet_name=self.sheet_name.get())
                # 更新空行判断列下拉列表
                self.empty_column['values'] = list(df.columns)
                if df.columns.any():
                    self.empty_column.set(df.columns[0])
                # 更新列管理器中的输入列
                self.column_manager.update_input_columns(list(df.columns))
                
                # 更新提示词管理器中的列名列表
                self.prompt_manager.column_names = list(df.columns)
                if hasattr(self.prompt_manager, 'column_combobox'):
                    self.prompt_manager.column_combobox['values'] = list(df.columns)
        except Exception as e:
            messagebox.showerror("错误", f"读取Sheet失败: {str(e)}")

    def save_config(self):
        """保存配置到文件"""
        config = {
            'input_file': self.input_file.get(),
            'output_file': self.output_file.get(),
            'sheet_name': self.sheet_name.get(),
            'empty_column': self.empty_column.get(),
            'batch_size': self.batch_size.get(),
            'workers': self.workers.get(),
            'api_url': self.api_url.get(),
            'api_key': self.api_key.get(),
            'model': self.model.get(),
            'content_template': self.prompt_manager.get_content_template(),
            'llm_template': self.prompt_manager.get_llm_template(),
            'input_columns': self.column_manager.get_input_columns(),
            'output_columns': self.column_manager.get_output_columns()
        }
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")

    def load_config(self):
        """加载配置文件"""
        if not self.config_path.exists():
            return
            
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
            # 设置基本配置
            self.input_file.set(config.get('input_file', ''))
            self.output_file.set(config.get('output_file', ''))
            self.batch_size.set(config.get('batch_size', '20'))
            self.workers.set(config.get('workers', '20'))
            self.api_url.set(config.get('api_url', ''))
            self.api_key.set(config.get('api_key', ''))
            self.model.set(config.get('model', ''))
            
            # 如果有输入文件，更新Sheet列表
            if config.get('input_file'):
                self.update_sheet_names()
                self.sheet_name.set(config.get('sheet_name', ''))
                self.empty_column.set(config.get('empty_column', ''))
            
            # 设置提示词模板
            self.prompt_manager.set_content_template(config.get('content_template', ''))
            self.prompt_manager.set_llm_template(config.get('llm_template', ''))
            
            # 注意：输入列和输出列的配置会在Sheet加载后自动更新
            
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败: {str(e)}")

    def validate_inputs(self):
        """验证所有输入是否有效"""
        if not self.input_file.get():
            messagebox.showerror("错误", "请选择输入文件")
            return False
        
        if not self.output_file.get():
            messagebox.showerror("错误", "请选择输出文件")
            return False
        
        if not self.sheet_name.get():
            messagebox.showerror("错误", "请选择Sheet名称")
            return False
        
        if not self.empty_column.get():
            messagebox.showerror("错误", "请选择判断空行的列")
            return False
        
        # API related validations
        if not self.api_url.get():
            messagebox.showerror("错误", "请填写API URL")
            return False
            
        api_key_value = self.api_key.get()
        if not api_key_value or api_key_value == "你的APIkey":
            messagebox.showerror("错误", "请填写有效的API Key。占位符 '你的APIkey' 是无效的。")
            return False

        if not self.model.get():
            messagebox.showerror("错误", "请填写模型名称")
            return False
        
        if not self.prompt_manager.get_content_template():
            messagebox.showerror("错误", "请填写内容整合模板")
            return False
        
        if not self.prompt_manager.get_llm_template():
            messagebox.showerror("错误", "请填写LLM提示词模板")
            return False
        
        if not any(self.column_manager.get_input_columns().values()):
            messagebox.showerror("错误", "请至少选择一个输入列")
            return False
        
        if not self.column_manager.get_output_columns():
            messagebox.showerror("错误", "请至少添加一个输出列")
            return False
        
        try:
            batch_size = int(self.batch_size.get())
            workers = int(self.workers.get())
            if batch_size <= 0 or workers <= 0:
                raise ValueError("批处理数量和并行处理数量必须大于0")
        except ValueError as e:
            messagebox.showerror("错误", str(e))
            return False
        
        return True

    def start_processing(self):
        """开始处理"""
        if not self.validate_inputs():
            return
    
        if self.is_processing:
            messagebox.showwarning("警告", "处理已在进行中")
            return
        
        try:
            config = {
                'input_file': self.input_file.get(),
                'output_file': self.output_file.get(),
                'sheet_name': self.sheet_name.get(),
                'empty_column': self.empty_column.get(),
                'batch_size': int(self.batch_size.get()),
                'workers': int(self.workers.get()),
                'api_url': self.api_url.get(),
                'api_key': self.api_key.get(),
                'model': self.model.get(),
                'content_template': self.prompt_manager.get_content_template(),
                'llm_template': self.prompt_manager.get_llm_template(),
                'input_columns': self.column_manager.get_input_columns(),
                'output_columns': self.column_manager.get_output_columns()
            }
        
            # 创建处理器实例，确保传递的函数名正确
            self.processor = ExcelProcessor(config, self.update_progress)
        
            # 禁用控制按钮
            self.disable_controls()
        
            # 启动处理线程
            self.is_processing = True
            self.process_thread = threading.Thread(target=self.process_wrapper)
            self.process_thread.start()
        
        except Exception as e:
            self.is_processing = False
            messagebox.showerror("错误", f"启动处理失败: {str(e)}\n{traceback.format_exc()}")
            self.enable_controls()

    def process_wrapper(self):
        """处理包装器，用于处理完成后的清理工作"""
        try:
            self.processor.start_processing()
        except Exception as e:
            messagebox.showerror("错误", f"处理过程出错: {str(e)}\n{traceback.format_exc()}")
        finally:
            self.is_processing = False
            self.enable_controls()
            self.progress_var.set(0)

    def stop_processing(self):
        """停止处理"""
        if hasattr(self, 'processor') and self.is_processing:
            self.processor.stop_processing()
            messagebox.showinfo("提示", "正在停止处理，请等待当前任务完成...")
        else:
            messagebox.showinfo("提示", "没有正在进行的处理任务")

    def update_progress(self, current_row, total_rows):
        """更新进度条和详细进度信息"""
        try:
            if total_rows > 0:
                progress = (current_row / total_rows) * 100
                self.progress_var.set(progress)
                self.root.update_idletasks()
                # 更新详细的进度信息
                self.progress_info_label.config(text=f"当前处理行数: {current_row} / 总行数: {total_rows}")
        except Exception as e:
            print(f"更新进度条失败: {str(e)}")

    def disable_controls(self):
        """禁用控制按钮和输入控件"""
        try:
            # 禁用文件选择区域
            for child in self.main_frame.winfo_children():
                if isinstance(child, ttk.LabelFrame):
                    for widget in child.winfo_children():
                        if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Combobox)):
                            widget.configure(state='disabled')
                        elif isinstance(widget, tk.Text):
                            widget.configure(state='disabled')

            # 确保停止按钮保持启用状态
            for widget in self.main_frame.winfo_children():
                if isinstance(widget, ttk.Frame):  # 控制按钮区域
                    for button in widget.winfo_children():
                        if isinstance(button, ttk.Button):
                            if button['text'] == "停止处理":
                                button.configure(state='normal')
                            else:
                                button.configure(state='disabled')
                                
            # 禁用列管理器和提示词管理器
            self.column_manager.set_state('disabled')
            self.prompt_manager.set_state('disabled')
            
        except Exception as e:
            print(f"禁用控件失败: {str(e)}")

    def enable_controls(self):
        """启用控制按钮和输入控件"""
        try:
            # 启用所有控件
            for child in self.main_frame.winfo_children():
                if isinstance(child, ttk.LabelFrame):
                    for widget in child.winfo_children():
                        if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Combobox)):
                            widget.configure(state='normal')
                        elif isinstance(widget, tk.Text):
                            widget.configure(state='normal')
                elif isinstance(child, ttk.Frame):  # 控制按钮区域
                    for button in child.winfo_children():
                        if isinstance(button, ttk.Button):
                            button.configure(state='normal')
                            
            # 启用列管理器和提示词管理器
            self.column_manager.set_state('normal')
            self.prompt_manager.set_state('normal')
            
        except Exception as e:
            print(f"启用控件失败: {str(e)}")

    def on_closing(self):
        """窗口关闭时的处理"""
        try:
            if self.is_processing:
                if messagebox.askokcancel("确认", "处理正在进行中，确定要退出吗？"):
                    if hasattr(self, 'processor'):
                        self.processor.stop_processing()
                    self.root.destroy()
            else:
                self.root.destroy()
        except Exception as e:
            print(f"关闭窗口时出错: {str(e)}")
            self.root.destroy()

def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()