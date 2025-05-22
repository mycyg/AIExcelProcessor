import tkinter as tk
from tkinter import ttk

class PromptManager:
    def __init__(self, parent_frame, column_names=None):
        self.parent_frame = parent_frame
        self.column_names = column_names  # 新增列名列表
        self.create_ui()

    def create_ui(self):
        # 创建提示词配置区域
        self.prompt_frame = ttk.LabelFrame(self.parent_frame, text="提示词配置", padding="5")
        self.prompt_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 创建内容整合配置
        self.content_frame = ttk.LabelFrame(self.prompt_frame, text="内容整合模板", padding="5")
        self.content_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 添加说明标签
        content_label = ttk.Label(self.content_frame, text="使用 {row['列名']} 格式引用输入列的值", wraplength=400)
        content_label.pack(fill="x", padx=5, pady=2)
        
        # 创建内容整合文本框和滚动条
        content_container = ttk.Frame(self.content_frame)
        content_container.pack(fill="both", expand=True)
        
        self.content_text = tk.Text(content_container, height=5, wrap=tk.WORD)
        content_scroll = ttk.Scrollbar(content_container, command=self.content_text.yview)
        self.content_text.configure(yscrollcommand=content_scroll.set)
        
        self.content_text.pack(side="left", fill="both", expand=True)
        content_scroll.pack(side="right", fill="y")
        
        # 创建LLM提示词配置
        self.llm_frame = ttk.LabelFrame(self.prompt_frame, text="LLM提示词模板", padding="5")
        self.llm_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 添加说明标签
        llm_label = ttk.Label(self.llm_frame, text="在此输入完整的LLM提示词模板，使用 {{content}} 引用整合后的内容", wraplength=400)
        llm_label.pack(fill="x", padx=5, pady=2)
        
        # 创建LLM提示词文本框和滚动条
        llm_container = ttk.Frame(self.llm_frame)
        llm_container.pack(fill="both", expand=True)
        
        self.llm_text = tk.Text(llm_container, height=5, wrap=tk.WORD)
        llm_scroll = ttk.Scrollbar(llm_container, command=self.llm_text.yview)
        self.llm_text.configure(yscrollcommand=llm_scroll.set)
        
        self.llm_text.pack(side="left", fill="both", expand=True)
        llm_scroll.pack(side="right", fill="y")
        
        # 创建列名选择下拉框
        if self.column_names:
            self.column_combobox = ttk.Combobox(self.prompt_frame, values=self.column_names)
            self.column_combobox.pack(fill="x", padx=5, pady=5)
            self.insert_column_button = ttk.Button(self.prompt_frame, text="插入列名", command=self.insert_column)
            self.insert_column_button.pack(fill="x", padx=5, pady=5)

    def insert_column(self):
        column_name = self.column_combobox.get()
        if column_name:
            # 插入到内容整合模板
            self.content_text.insert(tk.INSERT, f"{{row['{column_name}']}}")
            # 插入到LLM提示词模板
            # 这里可以根据需求调整插入逻辑
            # self.llm_text.insert(tk.INSERT, f"{{row['{column_name}']}}")

    def get_content_template(self):
        """获取内容整合模板"""
        return self.content_text.get("1.0", tk.END).strip()

    def get_llm_template(self):
        """获取LLM提示词模板"""
        return self.llm_text.get("1.0", tk.END).strip()

    def set_content_template(self, template):
        """设置内容整合模板"""
        self.content_text.delete("1.0", tk.END)
        if template:
            self.content_text.insert("1.0", template)

    def set_llm_template(self, template):
        """设置LLM提示词模板"""
        self.llm_text.delete("1.0", tk.END)
        if template:
            self.llm_text.insert("1.0", template)

    def set_state(self, state):
        """设置所有控件的状态"""
        self.content_text.configure(state=state)
        self.llm_text.configure(state=state)