#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_word_tables.py
批量从目录中的 Word (.docx) 文档提取表格、清洗字段并导出 Excel。
- 使用内置 pipeline 直接读取 .docx → 结构化 JSON → 表格清洗 → 字段字典
- 对“服从分配”前的复选框/勾勾逻辑使用统一的 is_checked_text 判断
- 保留调试开关以产生详细日志

依赖:
    pip install pandas openpyxl
注意:
    - 仅支持 .docx（.doc 需先另存为 .docx）
"""

# keep os for file ops
import os
import logging
import argparse
import zipfile
import tempfile
import json
from datetime import datetime
from typing import Dict, List

import pandas as pd

from src.read_docx import extract_structure
from src.extract_tables import extract_tables
from src.clean_table_dicts import clean_tables

# 可配置：输入文件夹与输出文件名（可在运行时通过设置菜单修改）
INPUT_FOLDER = os.path.join('.', '网友')
OUTPUT_XLSX = 'output.xlsx'

# 配置持久化文件
CONFIG_PATH = 'config.json'


def load_config(path: str = CONFIG_PATH) -> dict:
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                return json.load(fh)
        except Exception:
            return {}
    return {}


def save_config(cfg: dict, path: str = CONFIG_PATH) -> None:
    try:
        with open(path, 'w', encoding='utf-8') as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


# 从持久化配置中加载初始值（若存在）
_cfg = load_config()
if _cfg.get('input_folder'):
    INPUT_FOLDER = _cfg.get('input_folder')
if _cfg.get('output_xlsx'):
    OUTPUT_XLSX = _cfg.get('output_xlsx')

# 全局调试标志
DEBUG_MODE = False
logger = None

# 要输出的列顺序（可以增删）
COLUMNS = [
    "文件名",
    "姓名",
    "性别",
    "出生年月",
    "政治面貌",
    "所在分院",
    "班级",
    "学号",
    "现（曾） 任职务",
    "第一志愿",
    "第二志愿",
    "服从分配",
    "联系方式",
    "微信",
    "何时何地曾担任何职务",
    "曾获奖项及获奖时间",
    "个人优势分析及简要工作设想",
]

def process_docx_file(path: str) -> Dict[str, str]:
    """Run the docx → table extraction → dictionary cleaner pipeline."""
    filename = os.path.basename(path)
    result = {col: "" for col in COLUMNS}
    result["文件名"] = filename
    try:
        structure = extract_structure(path)
        tables = extract_tables(structure)
        cleaned = clean_tables({"tables": tables})
        if cleaned:
            # prefer the first table entry
            entry = cleaned[0]
            for key, value in entry.items():
                if key in result and value:
                    result[key] = value
    except Exception as exc:
        if DEBUG_MODE and logger:
            logger.debug(f"提取 {filename} 时出错: {exc}")
        return result
    if DEBUG_MODE and logger:
        logger.debug(f"{filename} -> {result}")
    return result


def setup_logger(debug_mode: bool):
    """设置日志系统"""
    global logger, DEBUG_MODE
    DEBUG_MODE = debug_mode
    
    logger = logging.getLogger('ResumeParser')
    logger.setLevel(logging.DEBUG if debug_mode else logging.INFO)

    # 清除现有的 handlers
    logger.handlers.clear()

    # 控制台输出（Info 级别）
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    # 日志文件保存到 logs/ 目录，始终保存 info 及以上；调试模式下也保存 debug
    os.makedirs('logs', exist_ok=True)
    level = logging.DEBUG if debug_mode else logging.INFO
    log_filename = os.path.join('logs', f"resume_parser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(level)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    if debug_mode:
        print(f"\n调试模式已启用，详细日志将保存到: {log_filename}\n")
    
    return logger


def show_menu():
    """显示交互式菜单"""
    print("\n" + "="*50)
    print("简历解析工具")
    print("="*50)
    print("1. 处理文件")
    print("2. 设置 - 修改输入文件夹与输出文件名")
    print("3. 退出")
    print("="*50)
    choice = input("请选择模式 (1-3): ").strip()
    return choice


def process_files(debug_mode: bool):
    """处理文件的主函数"""
    global logger
    logger = setup_logger(debug_mode)
    
    rows: List[Dict[str, str]] = []

    if not os.path.isdir(INPUT_FOLDER):
        logger.error(f"输入目录不存在: {INPUT_FOLDER}")
        return

    # 使用一个临时目录来解压 zip 文件中的 docx
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_paths: List[str] = []

        # 遍历输入文件夹，收集 .docx 文件并解压 zip 中的 docx
        for entry in os.listdir(INPUT_FOLDER):
            full = os.path.join(INPUT_FOLDER, entry)
            if entry.lower().endswith('.docx') and os.path.isfile(full):
                docx_paths.append(full)
            elif entry.lower().endswith('.zip') and os.path.isfile(full):
                try:
                    with zipfile.ZipFile(full, 'r') as zf:
                        for zi in zf.infolist():
                            if zi.filename.lower().endswith('.docx'):
                                extracted = zf.extract(zi, path=tmpdir)
                                docx_paths.append(extracted)
                except Exception as e:
                    logger.error(f"无法读取 zip 文件: {entry}，错误: {e}")

        if not docx_paths:
            logger.info(f"未找到任何 .docx 文件（包括 zip 内）于目录: {INPUT_FOLDER}")
            return

        logger.info(f"\n找到 {len(docx_paths)} 个 docx 待处理（含 zip 内）\n")

        for path in docx_paths:
            fname = os.path.basename(path)
            try:
                row = process_docx_file(path)
                if not row:
                    row = {col: "" for col in COLUMNS}
                    row["文件名"] = fname
                rows.append(row)
                logger.info(f"✓ 已处理: {fname}")
            except Exception as e:
                logger.error(f"✗ 处理文件出错: {fname}")
                logger.error(f"  错误: {str(e)}")
                if debug_mode:
                    import traceback
                    logger.debug(traceback.format_exc())
                row = {col: "" for col in COLUMNS}
                row["文件名"] = fname
                rows.append(row)

    df = pd.DataFrame(rows, columns=COLUMNS)
    df.to_excel(OUTPUT_XLSX, index=False)
    logger.info(f"\n{'='*50}")
    logger.info(f"完成！结果已保存到: {OUTPUT_XLSX}")
    logger.info(f"共处理 {len(rows)} 个文件")
    logger.info(f"{'='*50}\n")


def main():
    """主函数 - 支持命令行参数和交互式菜单"""
    parser = argparse.ArgumentParser(description='简历解析工具')
    parser.add_argument('-d', '--debug', action='store_true', help='启用调试模式')
    parser.add_argument('-n', '--no-menu', action='store_true', help='跳过菜单直接运行')
    args = parser.parse_args()
    
    global INPUT_FOLDER, OUTPUT_XLSX

    if args.no_menu:
        process_files(args.debug)
        return

    # 交互式菜单（加入设置项）
    while True:
        choice = show_menu()
        if choice == '1':
            process_files(debug_mode=args.debug)
            input("\n按回车键继续...")
        elif choice == '2':
            print("\n当前配置:")
            print(f"  输入文件夹: {INPUT_FOLDER}")
            print(f"  输出文件名: {OUTPUT_XLSX}")
            new_in = input("输入新的读取文件夹路径（回车保持不变）: ").strip()
            if new_in:
                INPUT_FOLDER = new_in
            new_out = input("输入新的输出文件名（回车保持不变）: ").strip()
            if new_out:
                OUTPUT_XLSX = new_out
            # 持久化配置
            cfg = load_config()
            cfg['input_folder'] = INPUT_FOLDER
            cfg['output_xlsx'] = OUTPUT_XLSX
            save_config(cfg)
            print("设置已更新并已保存到 config.json。")
        elif choice == '3':
            print("\n再见！")
            break
        else:
            print("\n无效选择，请重试。")


if __name__ == "__main__":
    main()
