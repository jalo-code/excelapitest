# -*- coding: utf-8 -*-
"""
项目配置文件
"""
import os

# 项目根目录
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Excel测试用例文件路径
EXCEL_PATH = os.path.join(BASE_DIR, "data", "test_cases.xlsx")

# Sheet名称
SHEET_NAME = "Sheet1"

# 日志文件路径
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "test.log")

# 报告目录
REPORT_DIR = os.path.join(BASE_DIR, "reports")

# Allure 结果目录（每个用例的JSON结果文件）
ALLURE_RESULTS_DIR = os.path.join(BASE_DIR, "reports", "allure-results")
# Allure 报告目录（CLI生成后的HTML报告）
ALLURE_REPORT_DIR = os.path.join(BASE_DIR, "reports", "allure-report")

# 请求超时时间（秒）
REQUEST_TIMEOUT = 30

# 默认请求头
DEFAULT_HEADERS = {
    "Content-Type": "application/json;charset=UTF-8"
}

# Excel列索引配置（从0开始）
COLUMNS = {
    "interface_name": 0,    # A列 - 接口名称
    "interface_desc": 1,    # B列 - 接口描述
    "url": 2,               # C列 - url
    "is_run": 3,            # D列 - 是否运行
    "request_type": 4,      # E列 - 请求类型
    "has_header": 5,        # F列 - 是否携带header
    "case_dependency": 6,   # G列 - case依赖
    "dependency_level": 7,  # H列 - 依赖返回数据的层级
    "dependency_key": 8,    # I列 - 依赖返回数据的key
    "dependency_data": 9,   # J列 - 依赖的返回数据
    "dependency_field": 10, # K列 - 数据依赖字段
    "request_data": 11,     # L列 - 请求数据
    "expected_result": 12,  # M列 - 预期结果
    "actual_result": 13     # N列 - 实际结果
}

# 请求头配置文件路径（JSON格式，所有接口共用）
HEADERS_CONFIG_PATH = os.path.join(BASE_DIR, "config", "headers.json")
