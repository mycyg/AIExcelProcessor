
import dataclasses
from typing import List, Dict

@dataclasses.dataclass
class ProcessingConfig:
    """
    A structured configuration object for the Excel processing task.
    """
    # File settings
    input_file: str = ""
    output_file: str = ""
    sheet_name: str = ""
    empty_column: str = ""

    # API settings
    processing_mode: str = "标准模式"
    api_url: str = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    api_key: str = ""
    model: str = "doubao-1-5-pro-32k-250115"
    api_timeout: int = 180 # in seconds

    # Processing settings
    batch_size: int = 20
    workers: int = 10

    # Prompt templates
    content_template: str = ""
    llm_template: str = ""

    # Column settings
    input_columns: Dict[str, bool] = dataclasses.field(default_factory=dict)
    output_columns: List[str] = dataclasses.field(default_factory=list)

