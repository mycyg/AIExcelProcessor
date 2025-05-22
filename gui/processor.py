import pandas as pd
import traceback
import os
import json
import requests
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import tempfile

class ExcelProcessor:
    def __init__(self, config, update_progress_callback):
        self.config = config
        # 确保变量名正确
        self.update_progress = update_progress_callback
        """初始化处理器
        
        Args:
            config: 包含处理配置的字典
            update_progress_callback: 更新进度的回调函数
        """
        # 基础配置
        self.input_file = config.get('input_file', '')
        self.output_file = config.get('output_file', '')
        self.sheet_name = config.get('sheet_name', '')
        self.empty_column = config.get('empty_column', '')
        
        # API配置
        self.api_url = config.get('api_url', '')
        self.api_key = config.get('api_key', '')
        self.model = config.get('model', '')
        
        # 列配置
        self.input_columns = {k: v for k, v in config.get('input_columns', {}).items() if v}
        self.output_columns = config.get('output_columns', [])
        
        # 处理配置
        self.batch_size = int(config.get('batch_size', 20))
        self.workers = int(config.get('workers', 20))
        
        # 模板配置
        self.content_template = config.get('content_template', '')
        self.llm_template = config.get('llm_template', '')
        
        # 控制标志
        self.should_stop = False
        self.processed_count = 0
        self.total_count = 0
        # 新增：初始化临时文件列表
        self.temp_files = []

    def start_processing(self):
        """开始处理Excel文件"""
        # 重置状态
        self.should_stop = False
        self.processed_count = 0
        
        # 验证配置
        validation_errors = self.validate_config()
        if validation_errors:
            raise ValueError('\n'.join(validation_errors))
        
        return self.process_excel()

    def stop_processing(self):
        """停止处理"""
        self.should_stop = True

    def format_content(self, row):
        """格式化内容模板
        
        Args:
            row: DataFrame的一行数据
        
        Returns:
            str: 格式化后的内容
        """
        try:
            # 创建安全的行字典，处理不存在的列和空值
            safe_row = {}
            for k in row.index:
                safe_row[k] = str(row[k]) if pd.notna(row[k]) else ''
            
            # 替换模板中的变量
            formatted_content = self.content_template
            for k, v in safe_row.items():
                placeholder = f"{{row['{k}']}}"
                formatted_content = formatted_content.replace(placeholder, str(v))
                
            return formatted_content
        except Exception as e:
            print(f"Error formatting content for row: {row.to_dict()}. Error: {str(e)}")
            traceback.print_exc()
            return None

    def call_api(self, content):
        """调用LLM API
        
        Args:
            content: 要处理的内容
        
        Returns:
            str: API返回的结果
        """
        try:
            # 准备API请求数据
            prompt = self.llm_template.replace('{{content}}', content)
            
            format_requirement = "严格按照以下格式回复，不要回复任何额外内容：\n"
            for col in self.output_columns:
                format_requirement += f"{col}:\"{col}\"\n"
            
            # 将格式要求添加到提示词末尾
            prompt += "\n" + format_requirement
            
            # 打印完整的提示词
            print("提交给LLM的完整提示词:")
            print(prompt)

            headers = {
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json'
            }
            
            data = {
                'model': self.model,
                'messages': [
                    {'role': 'user', 'content': prompt}
                ]
            }
            
            # 发送请求
            response = requests.post(
                self.api_url,
                headers=headers,
                json=data,
                timeout=60
            )
            
            # 检查响应
            response.raise_for_status()
            result = response.json()
            
            # 提取回答内容
            return result['choices'][0]['message']['content']
            
        except Exception as e:
            print(f"API call failed for URL: {self.api_url}. Content snippet: '{content[:100]}...'. Error: {str(e)}")
            traceback.print_exc()
            return None

    def parse_llm_response(self, response):
        """解析LLM的回复
        
        Args:
            response: LLM的回复内容
        
        Returns:
            dict: 解析后的字典，键是列名，值是对应的值
        """
        result = {}
        lines = response.split('\n')
        for line in lines:
            if ':' in line:
                key, value = line.split(':', 1)
                key = key.strip()
                value = value.strip().strip('"')
                result[key] = value
        return result

    def process_batch(self, batch_df):
        """处理一批数据
        
        Args:
            batch_df: 要处理的DataFrame批次
        
        Returns:
            list: 处理结果列表
        """
        results = []
        for _, row in batch_df.iterrows():
            if self.should_stop:
                break
                
            # 格式化内容
            content = self.format_content(row)
            if not content:
                continue
            
            # 调用API
            api_result = self.call_api(content)
            if not api_result:
                continue
            
            # 解析LLM的回复
            parsed_result = self.parse_llm_response(api_result)
            
            # 收集结果
            result = {col: row[col] for col in self.input_columns if col in row}
            for output_col in self.output_columns:
                result[output_col] = parsed_result.get(output_col, '')
                
            results.append(result)
            
            # 更新进度
            self.processed_count += 1
            if self.update_progress and self.total_count > 0:
                # 传递当前处理的行数和总行数
                self.update_progress(self.processed_count, self.total_count)
            
            # 添加延迟以避免API限制
            time.sleep(0.1)
        
        if results:
            # 创建结果DataFrame
            result_df = pd.DataFrame(results)
            
            # 确保输出目录存在
            output_dir = os.path.dirname(self.output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 保存批次结果到临时文件
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', dir=output_dir)
            result_df.to_excel(temp_file.name, index=False)
            # 修改：使用初始化的临时文件列表
            self.temp_files.append(temp_file.name)
            print(f"批次处理完成，结果已保存到 {temp_file.name}")
        return results

    def process_excel(self):
        """处理Excel文件
        
        Returns:
            bool: 处理是否成功
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(self.input_file, sheet_name=self.sheet_name)
            
            # 过滤有效行
            if self.empty_column:
                df = df[df[self.empty_column].notna()]
            
            # 准备处理
            self.total_count = len(df)
            results = []
            
            # 按批次处理
            with ThreadPoolExecutor(max_workers=self.workers) as executor:
                futures = []
                for i in range(0, len(df), self.batch_size):
                    if self.should_stop:
                        break
                    batch_df = df[i:i + self.batch_size]
                    futures.append(
                        executor.submit(self.process_batch, batch_df)
                    )
                for future in as_completed(futures):
                    if self.should_stop:
                        break
                    batch_results = future.result()
                    if batch_results:
                        results.extend(batch_results)
            if self.should_stop:
                print("处理已被用户停止")
                # 新增：删除临时文件
                for temp_file in self.temp_files:
                    os.remove(temp_file)
                return False
            # 新增：合并临时文件
            if self.temp_files:
                dfs = [pd.read_excel(temp_file) for temp_file in self.temp_files]
                result_df = pd.concat(dfs, ignore_index=True)
                result_df.to_excel(self.output_file, index=False)
                # 新增：删除临时文件
                for temp_file in self.temp_files:
                    os.remove(temp_file)
                print(f"处理完成，结果已保存到 {self.output_file}")
                return True
            else:
                print("没有生成任何结果")
                return False
        except Exception as e:
            print(f"处理Excel文件时出错: {str(e)}")
            traceback.print_exc()
            # 新增：删除临时文件
            for temp_file in self.temp_files:
                os.remove(temp_file)
            return False

    def validate_config(self):
        """验证配置是否有效
        
        Returns:
            list: 错误信息列表
        """
        errors = []
        
        # 验证文件配置
        if not self.input_file:
            errors.append("未指定输入文件")
        elif not os.path.exists(self.input_file):
            errors.append(f"输入文件不存在: {self.input_file}")
            
        if not self.output_file:
            errors.append("未指定输出文件")
            
        if not self.sheet_name and self.sheet_name != 0:
            errors.append("未指定Sheet名称")
            
        # 验证列配置
        if not self.empty_column:
            errors.append("未指定空行判断列")
            
        if not self.input_columns:
            errors.append("未选择输入列")
            
        if not self.output_columns:
            errors.append("未指定输出列")
        
        # 验证API配置    
        if not self.api_url:
            errors.append("未指定API URL")
            
        if not self.api_key:
            errors.append("未指定API Key")
            
        if not self.model:
            errors.append("未指定模型")
            
        # 验证模板配置
        if not self.content_template:
            errors.append("未指定内容模板")
            
        if not self.llm_template:
            errors.append("未指定LLM提示词模板")
            
        # 验证处理配置
        if self.batch_size <= 0:
            errors.append("批处理数量必须大于0")
            
        if self.workers <= 0:
            errors.append("并行处理数量必须大于0")
            
        return errors

    def get_sheets(self):
        """获取Excel文件中的所有Sheet名称
        
        Returns:
            list: Sheet名称列表
        """
        try:
            if not self.input_file or not os.path.exists(self.input_file):
                return []
                
            excel_file = pd.ExcelFile(self.input_file)
            return excel_file.sheet_names
        except Exception as e:
            print(f"获取Sheet名称时出错: {str(e)}")