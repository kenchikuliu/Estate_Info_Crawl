# -*- coding: utf-8 -*-
"""
数据存储工具模块
"""
import pandas as pd
import os
import re
import shutil
import requests
from typing import Dict, List, Any, Optional
from config import Config

class DataStorage:
    """数据存储类"""

    ILLEGAL_XML_RE = re.compile(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]"
    )

    @staticmethod
    def _sanitize_excel_value(value):
        if isinstance(value, str):
            return DataStorage.ILLEGAL_XML_RE.sub("", value)
        return value

    @staticmethod
    def sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
        return df.apply(lambda col: col.map(DataStorage._sanitize_excel_value))

    @staticmethod
    def write_excel_atomic(df: pd.DataFrame, filepath: str, sheet_name: str = "Sheet1") -> None:
        """
        原子写入 Excel，避免并发读取时拿到半写入文件。
        """
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        tmp_path = filepath + ".tmp.xlsx"
        clean_df = DataStorage.sanitize_dataframe_for_excel(df)
        with pd.ExcelWriter(tmp_path, mode="w", engine="openpyxl") as writer:
            clean_df.to_excel(writer, sheet_name=sheet_name, index=False)
        shutil.move(tmp_path, filepath)

    @staticmethod
    def _write_excel(df: pd.DataFrame, filepath: str, sheet_name: str = "Sheet1", mode: str = "w", replace_sheet: bool = True) -> None:
        """
        统一处理 Excel 写入，避免文件不存在或工作表冲突导致失败。
        """
        os.makedirs(os.path.dirname(filepath), exist_ok=True)

        if mode == "a" and os.path.exists(filepath):
            if_sheet_exists = "replace" if replace_sheet else "overlay"
            with pd.ExcelWriter(filepath, mode="a", engine="openpyxl", if_sheet_exists=if_sheet_exists) as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(filepath, mode="w", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    @staticmethod
    def save_to_excel(data: List[Dict[str, Any]], filename: str, sheet_name: str = "Sheet1") -> None:
        """
        保存数据到Excel文件

        Args:
            data: 要保存的数据列表
            filename: 文件名
            sheet_name: 工作表名称
        """
        if not data:
            print("没有数据需要保存")
            return

        filepath = os.path.join(Config.OUTPUT_DIR, filename)

        # 创建DataFrame
        df = pd.DataFrame(data)

        DataStorage._write_excel(df, filepath, sheet_name=sheet_name, mode="w")
        print(f"数据已保存到: {filepath}")

    @staticmethod
    def save_to_csv(data: List[Dict[str, Any]], filename: str) -> None:
        """
        保存数据到CSV文件

        Args:
            data: 要保存的数据列表
            filename: 文件名
        """
        if not data:
            print("没有数据需要保存")
            return

        filepath = os.path.join(Config.OUTPUT_DIR, filename)

        # 创建DataFrame
        df = pd.DataFrame(data)

        # 保存到CSV
        df.to_csv(filepath, index=False, encoding='utf-8-sig')
        print(f"数据已保存到: {filepath}")

    @staticmethod
    def download_file(url: str, filepath: str) -> bool:
        """
        下载文件

        Args:
            url: 文件URL
            filepath: 保存路径

        Returns:
            bool: 下载是否成功
        """
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()

            # 确保目录存在
            os.makedirs(os.path.dirname(filepath), exist_ok=True)

            with open(filepath, "wb") as f:
                f.write(response.content)

            return True
        except Exception as e:
            print(f"下载文件失败 {url}: {e}")
            return False

    @staticmethod
    def create_folder(folder_path: str) -> str:
        """
        创建文件夹

        Args:
            folder_path: 文件夹路径

        Returns:
            str: 创建的文件夹路径
        """
        full_path = os.path.join(Config.OUTPUT_DIR, folder_path)
        os.makedirs(full_path, exist_ok=True)
        return full_path

    @staticmethod
    def append_to_excel(data: List[Dict[str, Any]], filename: str, sheet_name: str = "Sheet1") -> None:
        """
        追加数据到Excel文件

        Args:
            data: 要追加的数据列表
            filename: 文件名
            sheet_name: 工作表名称
        """
        if not data:
            print("没有数据需要追加")
            return

        filepath = os.path.join(Config.OUTPUT_DIR, filename)

        # 创建DataFrame
        df = pd.DataFrame(data)

        try:
            # 尝试读取现有文件
            existing_df = pd.read_excel(filepath, sheet_name=sheet_name)
            # 合并数据
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            # 保存
            DataStorage._write_excel(combined_df, filepath, sheet_name=sheet_name, mode="w")
        except FileNotFoundError:
            # 如果文件不存在，直接创建
            DataStorage._write_excel(df, filepath, sheet_name=sheet_name, mode="w")

        print(f"数据已追加到: {filepath}")
