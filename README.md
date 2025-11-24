# Resume Parsing

简历解析工具 — 从 `.docx`（含 zip 内的 docx）批量提取表格并清洗为结构化字段，输出为 Excel。

核心功能
- 批量解析 Word `.docx` 中的表格，提取字段（姓名/性别/班级/学号/志愿/联系方式等）。
- 支持解析压缩包（`.zip`）内的 `.docx`，自动解压至临时目录并处理。
- 智能识别复选框/勾选状态（用于“服从分配”等二值字段）。
- 可交互设置输入目录与输出文件名，并将设置持久化到 `config.json`。
- 详细日志记录（保存到 `logs/`），支持命令行 `--debug` 打开更详细的调试日志。

快速开始
1. 安装依赖（推荐在虚拟环境中）：

```powershell
pip install -r requirements.txt
```

2. 交互式运行（带设置菜单）：

```powershell
python main.py
```

菜单说明：
- `1` 处理文件（读取当前 `INPUT_FOLDER`，处理 `.docx` 与 zip 内 `.docx`，输出到 `OUTPUT_XLSX`）
- `2` 设置：修改 `INPUT_FOLDER` 与 `OUTPUT_XLSX`，并保存到 `config.json`
- `3` 退出

3. 无菜单命令行模式（适合批处理或脚本）：

```powershell
# 普通运行（不显示菜单）
python main.py --no-menu

# 启用调试日志（文件中包含 DEBUG 级别信息）
python main.py --no-menu --debug
```

配置持久化
- 程序会在当前工作目录读取 `config.json`（若存在），支持的字段：`input_folder`、`output_xlsx`。
- 在菜单中修改设置后会自动保存到 `config.json`，下次启动会自动加载。

日志
- 所有运行都会在 `logs/` 目录生成按时间戳命名的日志文件（INFO 级别及以上）。
- 使用 `--debug` 会使日志文件记录 DEBUG 级别的详细信息，便于字段解析与复选框判断定位问题。

输入/输出
- 输入：默认读取 `./网友`（可通过菜单或 `config.json` 修改）。支持直接的 `.docx` 文件和 `.zip`（zip 内的 `.docx` 会被提取并处理）。
- 输出：默认写入 `output.xlsx`（可通过菜单或 `config.json` 修改）。

工具链（独立模块，供调试或批处理使用）
- `src/read_docx.py` — 读取 `.docx`（或已解压的 docx 文件夹）并输出结构化 JSON（段落 + 表格）。
	- 用法示例：
		```powershell
		python src/read_docx.py path/to/document.docx --output tmp/structure.json
		```
- `src/extract_tables.py` — 从 `read_docx` 的结构化 JSON 中筛选并修复表格布局，输出表格 JSON。
	- 用法示例：
		```powershell
		python src/extract_tables.py tmp/structure.json --output tmp/tables.json --debug
		```
	- `--debug` 在终端打印表格网格（便于检查合并单元格、换行重组等问题）。
- `src/clean_table_dicts.py` — 将每张表按关键词映射到目标字段，并对复选框/勾选做 `是`/`否` 判断。
	- 用法示例：
		```powershell
		python src/clean_table_dicts.py tmp/tables.json --debug
		```
	- `--debug` 显示每个表映射后的字典，便于定位字段错配。

支持的表格布局与限制
- 支持标签在左、标签在上、标签和值在同一单元格（如 `手机号：13800138000`）。
- 支持合并单元格的基本修复与跨行/跨列值重组。
- 支持多种复选框符号（☑ ☒ ✓ ✔ √ ■ ● 等）。
- 当前仅支持 `.docx`（若为 `.doc`，请先另存为 `.docx`）；zip 支持仅解压并处理 `.docx`，不会递归解析更深层压缩格式。

调试与常见问题处理
- 某字段识别不正确：使用 `--debug` 生成包含 DEBUG 级别的日志文件，查看 `logs/` 中对应时间戳的日志，按日志中提示检查表格结构与匹配关键词。
- 复选框识别不准确：可在 `src/clean_table_dicts.py` 的 `CHECKED_CHARS` / `UNCHECKED_CHARS` 中增加新的符号。
- 想要改变默认文件夹/输出名：使用菜单 `2` 保存到 `config.json`，或直接编辑 `main.py` 顶部的默认常量（不推荐）。

开发与贡献
- 欢迎提交 issue 或 pull request。请尽量在本地运行并使用 `--debug` 捕获日志以复现问题。

许可
- 本项目遵循仓库根目录中的 `LICENSE`。
