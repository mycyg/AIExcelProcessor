import tkinter as tk
from tkinter import ttk

class ColumnManager:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.input_columns = {}  # 存储输入列的选择状态
        self.output_columns = []  # 存储输出列的条目
        
        self.create_ui()

    def create_ui(self):
        # 创建列配置区域
        self.columns_frame = ttk.LabelFrame(self.parent_frame, text="列配置", padding="5")
        self.columns_frame.pack(fill="x", padx=5, pady=5)
        
        # 创建输入列选择区域
        self.input_frame = ttk.LabelFrame(self.columns_frame, text="输入列选择", padding="5")
        self.input_frame.pack(fill="x", padx=5, pady=5)
        
        # 创建输出列配置区域
        self.output_frame = ttk.LabelFrame(self.columns_frame, text="输出列配置", padding="5")
        self.output_frame.pack(fill="x", padx=5, pady=5)
        
        # 添加输出列按钮
        self.add_output_button = ttk.Button(
            self.output_frame, 
            text="添加输出列", 
            command=self.add_output_column
        )
        self.add_output_button.pack(pady=5)

    def update_input_columns(self, columns):
        """更新输入列选择"""
        # 清空现有的输入列选择
        for widget in self.input_frame.winfo_children():
            widget.destroy()
        
        self.input_columns = {}
        
        # 使用网格布局显示输入列选择
        for i, column in enumerate(columns):
            var = tk.BooleanVar(value=False)
            self.input_columns[column] = var
            ttk.Checkbutton(
                self.input_frame, 
                text=column, 
                variable=var
            ).grid(row=i//3, column=i%3, sticky="w", padx=5, pady=2)

        # 配置网格列的权重
        self.input_frame.grid_columnconfigure(0, weight=1)
        self.input_frame.grid_columnconfigure(1, weight=1)
        self.input_frame.grid_columnconfigure(2, weight=1)

    def add_output_column(self):
        """添加新的输出列"""
        # 创建新的输出列配置行
        row_frame = ttk.Frame(self.output_frame)
        row_frame.pack(fill="x", padx=5, pady=2)
        
        # 列名输入
        name_entry = ttk.Entry(row_frame)
        name_entry.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        # 删除按钮
        delete_button = ttk.Button(
            row_frame,
            text="删除",
            command=lambda: self.delete_output_column(row_frame)
        )
        delete_button.pack(side="right")
        
        self.output_columns.append(name_entry)

    def delete_output_column(self, row_frame):
        """删除输出列"""
        # 找到对应的Entry并从列表中移除
        for entry in self.output_columns:
            if entry.master == row_frame:
                self.output_columns.remove(entry)
                break
        row_frame.destroy()

    def get_input_columns(self):
        """获取选中的输入列"""
        return {k: v.get() for k, v in self.input_columns.items()}

    def get_output_columns(self):
        """获取配置的输出列"""
        return [entry.get() for entry in self.output_columns if entry.get().strip()]

    def set_state(self, state):
        """设置所有控件的状态"""
        # 设置输入列选择区域的状态
        for widget in self.input_frame.winfo_children():
            if isinstance(widget, (ttk.Checkbutton, ttk.Entry)):
                widget.configure(state=state)
                
        # 设置输出列区域的状态
        for widget in self.output_frame.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state=state)
        
        # 设置所有输出列输入框的状态
        for entry in self.output_columns:
            entry.configure(state=state)

    def set_input_columns(self, columns_dict):
        """设置输入列的选中状态"""
        for column, value in columns_dict.items():
            if column in self.input_columns:
                self.input_columns[column].set(value)

    def set_output_columns(self, columns):
        """设置输出列"""
        # 清除现有的输出列
        for entry in self.output_columns:
            entry.master.destroy()
        self.output_columns.clear()
        
        # 添加新的输出列
        for column in columns:
            self.add_output_column()
            self.output_columns[-1].insert(0, column)