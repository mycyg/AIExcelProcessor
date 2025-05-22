# Excel 批量处理工具

本项目是一个使用 LLM API 对 Excel 文件内容进行批量处理的 Python 应用程序。它提供了一个图形用户界面 (GUI) 来配置和执行处理任务。

## 主要功能

*   通过指定的 API (例如大语言模型) 批量处理 Excel 文件中的行。
*   灵活配置输入列和输出列。
*   自定义内容整合模板和 LLM 提示词模板。
*   支持 Tkinter 和 PySide6 (Qt) 两种图形界面。(`qt_app.py` 功能更完善)
*   处理进度显示和基本错误处理。
*   通过 `config.json` 文件保存和加载用户配置。

## 环境设置

1.  **Python 环境**:
    确保您已安装 Python 3.6 或更高版本。

2.  **克隆/下载项目**:
    将项目文件下载到您的本地计算机。

3.  **安装依赖**:
    打开命令行/终端，进入项目根目录，然后运行以下命令安装所需依赖库:
    ```bash
    pip install -r requirements.txt
    ```
    这将安装 `PySide6`, `pandas`, 和 `requests`。

##配置文件说明 (`config.json`)

在运行应用程序之前，建议您检查并配置 `config.json` 文件。此文件保存了应用程序的各项设置。

*   `input_file`: 输入的 Excel 文件路径 (例如: `"C:/Users/YourUser/Desktop/input.xlsx"`)。
*   `output_file`: 处理完成后输出的 Excel 文件路径 (例如: `"C:/Users/YourUser/Desktop/output.xlsx"`)。
*   `sheet_name`: 需要处理的 Excel 表格 (Sheet) 的名称 (例如: `"Sheet1"`)。
*   `empty_column`: 用于判断某一行是否为空（并跳过处理）的列名。如果该列单元格为空，则对应行不被处理。
*   `batch_size`: 每次批量请求 API 的行数 (例如: `20`)。
*   `workers`: 并行处理的线程数量 (例如: `10`)。
*   `api_url`: 您使用的大语言模型 API 的 URL (例如: `"https://ark.cn-beijing.volces.com/api/v3/chat/completions"`)。
*   `api_key`:您的 API 密钥。**请务必填写您自己的有效 API Key**。不要使用 `"你的APIkey"` 这个占位符。
*   `model`: 您希望使用的模型名称 (例如: `"doubao-1-5-pro-32k-250115"`)。
*   `content_template`: 内容整合模板。用于将 Excel 中选定的多个输入列整合成一段文本，供后续 LLM 处理。使用 `{row['列名']}` 的形式引用列数据。
    *   例如: `"背景: {row['背景信息']} | 内容: {row['核心内容']}"`
*   `llm_template`: LLM 提示词模板。这是发送给大语言模型的完整提示。使用 `{{content}}` 占位符来引用由 `content_template` 生成的内容。
    *   例如: `"请根据以下信息进行分析:
{{content}}
请输出分析结果。"`
*   `input_columns`: 一个字典，定义了哪些列被选为输入列，以及它们是否被勾选。
    *   例如: `{"列A": true, "列B": false}` 表示 列A 被选中作为输入，列B 未被选中。
*   `output_columns`: 一个列表，定义了希望 LLM 输出并保存到结果文件中的列名。
    *   例如: `["总结", "分类", "关键词"]`

## 如何运行

项目提供了两种 GUI 版本：

1.  **PySide6 (Qt) 版本 (推荐)**:
    此版本功能更全面，界面也更完善。通过运行 `qt_app.py` 启动：
    ```bash
    python qt_app.py
    ```

2.  **Tkinter 版本**:
    一个基础版本的 GUI。通过运行 `app.py` 启动：
    ```bash
    python app.py
    ```

**运行步骤**:

1.  确保 `config.json` 文件已根据您的需求配置完毕，特别是 `api_key`, `api_url`, `model`, `input_file`, `output_file`。
2.  打开 GUI 应用程序。
3.  在 "基本设置" (或类似名称的标签页) 中：
    *   选择输入 Excel 文件。
    *   指定输出 Excel 文件名。
    *   加载文件后，选择要处理的 Sheet。
    *   选择用于判断空行的列。
    *   填写 API URL, API Key, 和模型名称 (如果 `config.json` 中未提供或不正确)。
    *   设置批处理数量和并行处理数量。
4.  在 "列设置" (或类似名称的标签页) 中：
    *   勾选需要作为内容整合依据的输入列。
    *   定义 LLM 需要生成并输出到结果文件中的列名 (每行一个)。
5.  在 "提示词设置" (或类似名称的标签页) 中：
    *   编辑内容整合模板，使用 `{row['列名']}` 引用您在 "列设置" 中勾选的输入列。
    *   编辑 LLM 提示词模板，使用 `{{content}}` 引用由内容整合模板生成的内容。
6.  点击 "开始处理" 按钮。
7.  处理进度会显示在界面上。处理完成后，结果会保存在您指定的输出文件中。

## 注意事项

*   **API Key 安全**: 请妥善保管您的 API Key。不要将其硬编码到脚本中或分享给他人。建议通过 `config.json` 进行配置。
*   **文件路径**: 在 `config.json` 中或 GUI 中填写文件路径时，请使用适合您操作系统的正确路径格式。
*   **Excel 文件**: 确保输入的 Excel 文件存在且可读。输出文件如果已存在，可能会被覆盖（具体行为取决于所用库的默认设置）。
*   **错误处理**: 程序包含基本的错误处理，但如果遇到 API 错误或文件问题，请检查 `config.json` 配置和 API 服务状态。`qt_app.py` 的日志窗口会提供更多信息。
*   **编码**: 所有配置文件和代码文件建议使用 UTF-8 编码。
