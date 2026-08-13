# -*- coding: utf-8 -*-
"""
HTTP请求发送模块
"""
import os
import json
import requests
from common.logger import get_logger
from config.settings import HEADERS_CONFIG_PATH, REQUEST_TIMEOUT

logger = get_logger("http_request")


def load_headers_config():
    """
    从config/headers.json加载统一请求头配置。
    每次调用都会重新读文件，保证运行中修改配置立即生效，无需重启。
    返回 (headers_dict, config_info_str) 便于日志溯源。
    """
    config_path = os.path.abspath(HEADERS_CONFIG_PATH)
    info = f"配置文件: {config_path}"

    if not os.path.exists(config_path):
        return {}, f"{info} — ❌ 文件不存在"

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}, f"{info} — ❌ 根节点不是对象（必须是 {{key: value}} 格式）"
        return data, f"{info} — ✅ 已加载 {len(data)} 个请求头: {list(data.keys())}"
    except json.JSONDecodeError as e:
        return {}, f"{info} — ❌ JSON格式错误: {e}"
    except Exception as e:
        return {}, f"{info} — ❌ 读取失败: {e}"


class HttpRequest:
    """HTTP请求封装类"""

    def __init__(self):
        self.session = requests.Session()
        # 保存所有接口返回结果，供依赖使用
        self.response_cache = {}
        # 启动时不缓存，改为每次请求前动态加载（避免改完配置不生效）

    def send_request(self, case, extra_headers=None):
        """
        发送HTTP请求
        :param case: 测试用例字典
        :param extra_headers: 额外的请求头
        :return: 响应结果（解析后的字典或文本）
        """
        url = case.get("url", "").strip()
        method = str(case.get("request_type", "get")).strip().lower()
        has_header = str(case.get("has_header", "no")).strip().lower()
        request_data_str = str(case.get("request_data", "")).strip()

        if not url:
            logger.error("请求URL为空")
            return {"error": "请求URL为空"}

        # 每次发送前动态加载headers.json（运行中修改配置立即生效）
        configured_headers, config_info = load_headers_config()

        # 构建请求头：无条件把配置文件中的全部header带上
        headers = {}
        merge_trace = []

        # 配置文件中的header全部带上，不看has_header列
        headers.update(configured_headers)
        merge_trace.append(f"① 配置文件全部 ({len(configured_headers)} 个)")

        if extra_headers:
            headers.update(extra_headers)
            merge_trace.append(f"② 代码传入extra ({len(extra_headers)} 个)")

        logger.info("=" * 60)
        logger.info(f"接口名称: {case.get('interface_name', '')}")
        logger.info(f"接口描述: {case.get('interface_desc', '')}")
        logger.info(f"请求方式: {method.upper()}")
        logger.info(f"请求URL: {url}")
        # Header溯源日志
        logger.info(f"Header来源: {config_info}")
        logger.info(f"Header合并顺序: {' → '.join(merge_trace)}")
        logger.info(f"最终请求头: {json.dumps(headers, ensure_ascii=False)}")

        # 解析请求数据
        request_data = self._parse_request_data(request_data_str, method)

        # 请求数据空值扫描：检测空字符串、空对象、空数组（用于排查服务端ParseInt错误）
        self._check_empty_values(request_data, case.get("interface_name", ""))
        logger.info(f"请求数据: {json.dumps(request_data, ensure_ascii=False) if isinstance(request_data, (dict, list)) else request_data}")

        try:
            if method == "get":
                # GET请求：参数拼接到URL query string
                params = request_data if isinstance(request_data, dict) else None
                response = self.session.get(
                    url=url,
                    params=params,
                    headers=headers if headers else None,
                    timeout=REQUEST_TIMEOUT
                )
                # 打印最终请求URL（含query参数），方便排查
                final_url = response.url
                logger.info(f"最终请求URL: {final_url}")

            elif method == "post":
                # 根据请求头Content-Type决定发送方式
                content_type = headers.get("Content-Type", "")
                if "application/json" in content_type:
                    response = self.session.post(
                        url=url,
                        json=request_data if isinstance(request_data, (dict, list)) else None,
                        data=request_data if not isinstance(request_data, (dict, list)) else None,
                        headers=headers if headers else None,
                        timeout=REQUEST_TIMEOUT
                    )
                    logger.info(f"发送JSON体: {json.dumps(request_data, ensure_ascii=False) if isinstance(request_data, (dict, list)) else request_data}")
                elif "application/x-www-form-urlencoded" in content_type:
                    response = self.session.post(
                        url=url,
                        data=request_data,
                        headers=headers if headers else None,
                        timeout=REQUEST_TIMEOUT
                    )
                    logger.info(f"发送表单数据: {request_data}")
                else:
                    # 无Content-Type时，默认用JSON发送并补上Content-Type头
                    if headers is None:
                        headers = {}
                    if "Content-Type" not in headers:
                        headers["Content-Type"] = "application/json;charset=UTF-8"
                        logger.warning("POST请求未设置Content-Type，已自动补充为 application/json")
                    response = self.session.post(
                        url=url,
                        json=request_data if isinstance(request_data, (dict, list)) else None,
                        data=request_data if not isinstance(request_data, (dict, list)) else None,
                        headers=headers if headers else None,
                        timeout=REQUEST_TIMEOUT
                    )
                    logger.info(f"发送JSON体: {json.dumps(request_data, ensure_ascii=False) if isinstance(request_data, (dict, list)) else request_data}")
            elif method == "put":
                response = self.session.put(
                    url=url,
                    json=request_data if isinstance(request_data, (dict, list)) else None,
                    data=request_data if not isinstance(request_data, (dict, list)) else None,
                    headers=headers if headers else None,
                    timeout=REQUEST_TIMEOUT
                )
            elif method == "delete":
                response = self.session.delete(
                    url=url,
                    params=request_data if isinstance(request_data, dict) else None,
                    headers=headers if headers else None,
                    timeout=REQUEST_TIMEOUT
                )
            else:
                logger.error(f"不支持的请求方式: {method}")
                return {"error": f"不支持的请求方式: {method}"}

            # 解析响应
            result = self._parse_response(response)
            logger.info(f"响应状态码: {response.status_code}")
            logger.info(f"响应结果: {json.dumps(result, ensure_ascii=False) if isinstance(result, (dict, list)) else str(result)[:500]}")

            # 缓存响应结果
            case_name = case.get("interface_name", "")
            if case_name:
                self.response_cache[case_name] = {
                    "status_code": response.status_code,
                    "headers": dict(response.headers),
                    "body": result,
                    "text": response.text
                }

            return result

        except requests.exceptions.Timeout:
            error_msg = f"请求超时（>{REQUEST_TIMEOUT}秒）"
            logger.error(error_msg)
            return {"error": error_msg}
        except requests.exceptions.ConnectionError:
            error_msg = "网络连接失败，请检查URL和网络"
            logger.error(error_msg)
            return {"error": error_msg}
        except Exception as e:
            error_msg = f"请求异常: {str(e)}"
            logger.exception(error_msg)
            return {"error": error_msg}

    def _check_empty_values(self, data, case_name, parent_path=""):
        """
        递归扫描请求数据中的空值（空字符串/空对象/空数组），打印警告。
        用于快速排查类似 strconv.ParseInt: parsing "" 的服务端错误。
        """
        if data is None:
            return

        if isinstance(data, dict):
            if parent_path and not data:
                # 空对象：如果不是根节点（根节点空是没传参数），告警
                logger.warning(
                    f"[用例 {case_name}] 请求数据中的空对象: {parent_path or '<root>'} = {{}}  "
                    f"（若服务端期望从其中取int字段，会报 ParseInt 空串错误，例如 base_params={{}}）"
                )
            for key, value in data.items():
                current_path = f"{parent_path}.{key}" if parent_path else str(key)
                if isinstance(value, str) and value == "":
                    logger.warning(
                        f"[用例 {case_name}] 请求数据中的空字符串字段: {current_path} = \"\"  "
                        f"（服务端若按int解析会报 ParseInt 错误，请补值或确认是否需要传该字段）"
                    )
                elif isinstance(value, list) and len(value) == 0:
                    logger.warning(
                        f"[用例 {case_name}] 请求数据中的空数组: {current_path} = []"
                    )
                elif isinstance(value, (dict, list)):
                    self._check_empty_values(value, case_name, current_path)

        elif isinstance(data, list):
            for idx, item in enumerate(data):
                current_path = f"{parent_path}[{idx}]" if parent_path else f"[{idx}]"
                if isinstance(item, str) and item == "":
                    logger.warning(
                        f"[用例 {case_name}] 请求数据中的空字符串: {current_path} = \"\""
                    )
                elif isinstance(item, (dict, list)):
                    self._check_empty_values(item, case_name, current_path)

    def _parse_request_data(self, data_str, method):
        """解析请求数据字符串为Python对象"""
        if not data_str:
            return {} if method in ["get", "delete"] else None

        # 尝试解析为JSON
        try:
            return json.loads(data_str)
        except (json.JSONDecodeError, TypeError):
            # 如果不是JSON，尝试解析为key=value格式
            if "&" in data_str or "=" in data_str:
                data = {}
                for item in data_str.split("&"):
                    if "=" in item:
                        k, v = item.split("=", 1)
                        data[k.strip()] = v.strip()
                return data if data else data_str
            return data_str

    def _parse_response(self, response):
        """解析响应内容"""
        content_type = response.headers.get("Content-Type", "")
        try:
            if "application/json" in content_type:
                return response.json()
        except Exception:
            pass

        # 尝试解析为JSON
        try:
            text = response.text.strip()
            if text and (text.startswith("{") or text.startswith("[")):
                return json.loads(text)
        except Exception:
            pass

        return response.text

    def get_cached_response(self, case_name):
        """获取缓存的响应结果"""
        return self.response_cache.get(case_name)
