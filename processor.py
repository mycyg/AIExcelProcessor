import pandas as pd
import traceback
import os
import json
import requests
import time
import tempfile
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Iterator, Dict, Any, Tuple, List

from config import ProcessingConfig

class ExcelProcessor:
    def __init__(self, config: ProcessingConfig):
        self.config = config
        self.should_stop = False
        self.total_rows = 0
        self.temp_dir = tempfile.mkdtemp(prefix="excel_proc_std_")
        self.input_jsonl_path = os.path.join(self.temp_dir, "input.jsonl")
        self.temp_files = []
        self.session = requests.Session()

    def stop(self):
        self.should_stop = True

    def _prepare_input_file(self) -> int:
        import openpyxl
        try:
            workbook = openpyxl.load_workbook(self.config.input_file, read_only=True)
            sheet = workbook[self.config.sheet_name]
            
            total_rows = 0
            header = [cell.value for cell in sheet[1]]
            empty_col_index = header.index(self.config.empty_column) if self.config.empty_column in header else -1

            with open(self.input_jsonl_path, 'w', encoding='utf-8') as f:
                for row_cells in sheet.iter_rows(min_row=2):
                    if self.should_stop:
                        break
                    if empty_col_index != -1 and (row_cells[empty_col_index].value is None or str(row_cells[empty_col_index].value).strip() == ""):
                        continue

                    row_data = {header[i]: cell.value for i, cell in enumerate(row_cells)}
                    f.write(json.dumps(row_data, ensure_ascii=False) + '\n')
                    total_rows += 1
            return total_rows
        except Exception as e:
            raise RuntimeError(f"Failed to prepare input file using openpyxl: {e}")

    def _get_next_batch_from_jsonl(self, file_iterator: Iterator[str]) -> List[Dict[str, Any]]:
        batch = []
        try:
            for _ in range(self.config.batch_size):
                line = next(file_iterator)
                batch.append(json.loads(line))
        except StopIteration:
            pass
        return batch

    def start_processing(self) -> Iterator[Tuple[str, Any, Any]]:
        self.should_stop = False
        self.temp_files = []
        processed_count = 0

        try:
            yield "info", "正在准备输入数据...", 0
            self.total_rows = self._prepare_input_file()
            yield "info", f"数据准备完成，总计 {self.total_rows} 行。", 0

            if self.total_rows == 0:
                yield "finish", 0, 0
                return

            with ThreadPoolExecutor(max_workers=self.config.workers) as executor, \
                 open(self.input_jsonl_path, 'r', encoding='utf-8') as f_in:
                
                futures = {}
                batch_num = 0
                
                while not self.should_stop:
                    batch = self._get_next_batch_from_jsonl(f_in)
                    if not batch:
                        break
                    
                    future = executor.submit(self._process_batch, batch)
                    futures[future] = batch_num
                    batch_num += 1

                for future in as_completed(futures):
                    if self.should_stop:
                        break
                    
                    for result_type, data, total in future.result():
                        if result_type == "data":
                            temp_file_path, batch_row_count = data
                            self.temp_files.append(temp_file_path)
                            processed_count += batch_row_count
                            yield "progress", processed_count, self.total_rows
                        else:
                            yield result_type, data, total
            
            if self.should_stop:
                yield "stopped", processed_count, self.total_rows
                return

            if self.temp_files:
                yield "info", "正在合并临时文件...", 0
                self._merge_temp_files()
                yield "finish", processed_count, self.total_rows
            else:
                yield "finish", 0, self.total_rows

        except Exception as e:
            yield "error", f"处理过程中发生未知错误: {e}", 0
            traceback.print_exc()
        finally:
            self._cleanup_temp_files()

    def _process_batch(self, batch: List[Dict[str, Any]]) -> List[Tuple[str, Any, Any]]:
        batch_results = []
        log_results = []
        for row in batch:
            if self.should_stop:
                break
            try:
                content = self._format_content(row)
                final_prompt, api_response = self._call_api(content)
                log_results.append(("debug_prompt", final_prompt, 0))
                log_results.append(("debug_response", api_response, 0))
                parsed_result = self._parse_llm_response(api_response)
                
                final_row = {col: row.get(col) for col, selected in self.config.input_columns.items() if selected}
                final_row.update(parsed_result)
                batch_results.append(final_row)
            except Exception as e:
                log_results.append(("error", f"处理行失败: {e}", 0))
                traceback.print_exc()
                continue
        
        if not batch_results:
            return log_results

        try:
            temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="proc_")
            os.close(temp_fd)
            result_df = pd.DataFrame(batch_results)
            result_df.to_excel(temp_path, index=False)
            log_results.append(("data", (temp_path, len(batch_results)), 0))
            return log_results
        except Exception as e:
            log_results.append(("error", f"保存临时文件失败: {e}", 0))
            traceback.print_exc()
            return log_results

    def _format_content(self, row: Dict[str, Any]) -> str:
        formatted_content = self.config.content_template
        for col_name, cell_value in row.items():
            placeholder = f"{{row['{col_name}']}}"
            value_str = str(cell_value) if cell_value is not None else ""
            formatted_content = formatted_content.replace(placeholder, value_str)
        return formatted_content

    def _call_api(self, content: str) -> tuple[str, str]:
        output_format_prompt = ", ".join([f'\"{col}\": \"...\"' for col in self.config.output_columns])
        prompt = self.config.llm_template.replace('{{content}}', content)
        prompt += f"\n\nPlease provide the output in a single, valid JSON object format, like this: {{{output_format_prompt}}}. Do not include any text or formatting outside of the JSON object."

        headers = {
            'Authorization': f'Bearer {self.config.api_key}',
            'Content-Type': 'application/json'
        }
        data = {
            'model': self.config.model,
            'messages': [
                {
                    'role': 'user',
                    'content': [
                        {'type': 'text', 'text': prompt}
                    ]
                }
            ]
        }
        
        max_retries = 3
        retry_delay = 1
        for attempt in range(max_retries):
            try:
                response = self.session.post(
                    self.config.api_url, 
                    headers=headers, 
                    json=data, 
                    timeout=self.config.api_timeout
                )
                response.raise_for_status()
                return prompt, response.json()['choices'][0]['message']['content']
            except requests.exceptions.RequestException as e:
                print(f"API call failed (attempt {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise

    def _parse_llm_response(self, response: str) -> Dict[str, Any]:
        try:
            start_index = response.find('{')
            end_index = response.rfind('}') + 1
            if start_index == -1 or end_index == 0:
                raise json.JSONDecodeError("No JSON object found in response", response, 0)
            json_str = response[start_index:end_index]
            return json.loads(json_str, strict=False)
        except json.JSONDecodeError as e:
            print(f"Failed to parse LLM response as JSON: {e}")
            return {col: "PARSE_ERROR" for col in self.config.output_columns}

    def _merge_temp_files(self):
        from openpyxl import load_workbook, Workbook

        if not self.temp_files:
            return

        # Create a new workbook for the final output
        final_workbook = Workbook()
        final_sheet = final_workbook.active

        # Use the first temp file to write the header
        first_file = self.temp_files[0]
        try:
            source_workbook = load_workbook(filename=first_file, read_only=True)
            source_sheet = source_workbook.active
            
            # Write header
            header = [cell.value for cell in source_sheet[1]]
            final_sheet.append(header)
            
            # Write data rows from the first file
            for row in source_sheet.iter_rows(min_row=2, values_only=True):
                final_sheet.append(row)
            
            source_workbook.close()

            # Process remaining temp files
            for temp_file in self.temp_files[1:]:
                try:
                    source_workbook = load_workbook(filename=temp_file, read_only=True)
                    source_sheet = source_workbook.active
                    # Append data rows, skipping the header
                    for row in source_sheet.iter_rows(min_row=2, values_only=True):
                        final_sheet.append(row)
                    source_workbook.close()
                except Exception as e:
                    print(f"Error processing temp file {temp_file}: {e}")

            output_dir = os.path.dirname(self.config.output_file)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            final_workbook.save(self.config.output_file)

        except Exception as e:
            print(f"Error merging temp files: {e}")
            traceback.print_exc()

    def _cleanup_temp_files(self):
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)

