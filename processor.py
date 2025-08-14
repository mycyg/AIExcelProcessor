import pandas as pd
import traceback
import os
import json
import requests
import time
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Iterator, Dict, Any, Tuple, List

from config import ProcessingConfig

class ExcelProcessor:
    """
    Handles the core logic of reading, processing, and writing Excel data.
    This version uses temporary files for batch processing to conserve memory.
    """
    def __init__(self, config: ProcessingConfig):
        self.config = config
        self.should_stop = False
        self.total_rows = 0
        self.temp_files = []

    def stop(self):
        """Signals the processor to stop gracefully."""
        self.should_stop = True

    def start_processing(self) -> Iterator[Tuple[str, int, int]]:
        """
        Starts the Excel processing task.
        Yields:
            A tuple containing (status_message, processed_count, total_rows)
        """
        self.should_stop = False
        self.temp_files = []
        processed_count = 0

        try:
            yield "info", 0, 0
            df = self._read_and_filter_data()
            self.total_rows = len(df)
            if self.total_rows == 0:
                yield "finish", 0, 0
                return

            with ThreadPoolExecutor(max_workers=self.config.workers) as executor:
                futures = {executor.submit(self._process_batch, df[i:i + self.config.batch_size]): i 
                           for i in range(0, self.total_rows, self.config.batch_size) if not self.should_stop}

                for future in as_completed(futures):
                    if self.should_stop:
                        break
                    
                    temp_file_path, batch_row_count = future.result()
                    if temp_file_path:
                        self.temp_files.append(temp_file_path)
                        processed_count += batch_row_count
                        yield "progress", processed_count, self.total_rows
            
            if self.should_stop:
                yield "stopped", processed_count, self.total_rows
                return

            if self.temp_files:
                self._merge_temp_files()
                yield "finish", processed_count, self.total_rows
            else:
                yield "finish", 0, self.total_rows

        except Exception as e:
            traceback.print_exc()
            yield "error", processed_count, self.total_rows
        finally:
            self._cleanup_temp_files()

    def _read_and_filter_data(self) -> pd.DataFrame:
        df = pd.read_excel(self.config.input_file, sheet_name=self.config.sheet_name)
        if self.config.empty_column and self.config.empty_column in df.columns:
            df.dropna(subset=[self.config.empty_column], inplace=True)
        return df

    def _process_batch(self, batch_df: pd.DataFrame) -> Tuple[str | None, int]:
        batch_results = []
        for _, row in batch_df.iterrows():
            if self.should_stop:
                break
            try:
                content = self._format_content(row)
                api_response = self._call_api(content)
                parsed_result = self._parse_llm_response(api_response)
                
                final_row = {col: row.get(col) for col, selected in self.config.input_columns.items() if selected}
                final_row.update(parsed_result)
                batch_results.append(final_row)
            except Exception:
                traceback.print_exc()
                continue
        
        if not batch_results:
            return None, 0

        try:
            temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="proc_")
            os.close(temp_fd)
            result_df = pd.DataFrame(batch_results)
            result_df.to_excel(temp_path, index=False)
            return temp_path, len(batch_results)
        except Exception as e:
            traceback.print_exc()
            return None, 0

    def _format_content(self, row: pd.Series) -> str:
        return self.config.content_template.format_map({f"row['{k}']": v for k, v in row.items()})

    def _call_api(self, content: str) -> str:
        output_format_prompt = ", ".join([f'\"{col}\": \"...\"' for col in self.config.output_columns])
        prompt = self.config.llm_template.replace('{{content}}', content)
        prompt += f"\n\nPlease provide the output in a single, valid JSON object format, like this: {{{output_format_prompt}}}. Do not include any text or formatting outside of the JSON object."

        headers = {
            'Authorization': f'Bearer {self.config.api_key}',
            'Content-Type': 'application/json'
        }
        data = {
            'model': self.config.model,
            'messages': [{'role': 'user', 'content': prompt}]
        }
        
        max_retries = 3
        retry_delay = 1  # in seconds
        for attempt in range(max_retries):
            try:
                response = requests.post(self.config.api_url, headers=headers, json=data, timeout=60)
                response.raise_for_status()
                return response.json()['choices'][0]['message']['content']
            except requests.exceptions.RequestException as e:
                print(f"API call failed (attempt {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise  # Re-raise the last exception if all retries fail

    def _parse_llm_response(self, response: str) -> Dict[str, Any]:
        try:
            start_index = response.find('{')
            end_index = response.rfind('}') + 1
            if start_index == -1 or end_index == 0:
                raise json.JSONDecodeError("No JSON object found in response", response, 0)
            json_str = response[start_index:end_index]
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"Failed to parse LLM response as JSON: {e}")
            return {col: "PARSE_ERROR" for col in self.config.output_columns}

    def _merge_temp_files(self):
        all_dfs = [pd.read_excel(f) for f in self.temp_files]
        if not all_dfs:
            return

        final_df = pd.concat(all_dfs, ignore_index=True)

        input_cols_sorted = sorted([col for col, selected in self.config.input_columns.items() if selected])
        output_cols_sorted = sorted(self.config.output_columns)
        final_column_order = [col for col in (input_cols_sorted + output_cols_sorted) if col in final_df.columns]
        final_df = final_df[final_column_order]
        
        output_dir = os.path.dirname(self.config.output_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        final_df.to_excel(self.config.output_file, index=False)

    def _cleanup_temp_files(self):
        for f in self.temp_files:
            try:
                os.remove(f)
            except OSError as e:
                print(f"Error deleting temp file {f}: {e}")
        self.temp_files = []
