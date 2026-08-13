# -*- coding: utf-8 -*-
"""
快速生成Excel模板脚本
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from common.excel_util import ExcelUtil


def main():
    print("正在生成Excel模板文件...")
    util = ExcelUtil()
    print(f"✓ Excel模板已生成: {util.excel_path}")
    print("  包含3条示例数据，可直接编辑使用。")


if __name__ == "__main__":
    main()
