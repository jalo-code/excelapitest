# -*- coding: utf-8 -*-
"""
依赖数据处理模块
处理case依赖、依赖返回数据的层级、依赖返回数据的key、数据依赖字段等逻辑
"""
import json
import re
from common.logger import get_logger

logger = get_logger("dependency")


class DependencyHandler:
    """依赖处理器"""

    def __init__(self, http_request, excel_util):
        self.http_request = http_request
        self.excel_util = excel_util

    def handle_dependency(self, case):
        """
        处理用例的依赖关系，将依赖数据注入到请求数据中
        :param case: 当前测试用例
        :return: 修改后的case，以及是否需要跳过执行（依赖执行失败）
        """
        case_dependency = str(case.get("case_dependency", "")).strip()
        if not case_dependency:
            return case, False, None

        logger.info(f"处理用例依赖: 依赖={case_dependency}")

        # 1. 先从缓存获取依赖用例的响应
        cached = self.http_request.get_cached_response(case_dependency)

        if not cached:
            # 2. 缓存没有，则检查是否已经执行过（Excel中实际结果有值）
            dep_row_num, dep_case = self.excel_util.get_case_by_name(case_dependency)
            if dep_case and dep_case.get("actual_result"):
                try:
                    dep_response = json.loads(dep_case["actual_result"])
                except Exception:
                    dep_response = dep_case["actual_result"]
                cached = {
                    "body": dep_response,
                    "text": dep_case["actual_result"]
                }
            else:
                skip_msg = f"跳过: 依赖用例 [{case_dependency}] 未执行且无响应缓存"
                logger.warning(skip_msg)
                return case, True, skip_msg

        # 3. 获取依赖返回数据
        dep_body = cached.get("body", {})
        dependency_level = str(case.get("dependency_level", "")).strip()
        dependency_key = str(case.get("dependency_key", "")).strip()
        dependency_field = str(case.get("dependency_field", "")).strip()

        # 从依赖响应中提取目标值
        extracted_value = self._extract_value(dep_body, dependency_level, dependency_key)
        logger.info(f"从依赖 [{case_dependency}] 提取值: {extracted_value} (层级={dependency_level}, key={dependency_key})")

        # 4. 将提取的值写入到"依赖的返回数据"列（J列）
        if extracted_value is not None:
            case["dependency_data"] = str(extracted_value)
            self.excel_util.write_cell(case["row_num"], "dependency_data", str(extracted_value))

        # 5. 替换请求数据中的依赖字段
        if dependency_field and extracted_value is not None:
            case = self._inject_dependency_value(case, dependency_field, extracted_value)

        return case, False, None

    def _extract_value(self, data, level_str, key_str):
        """
        根据层级和key从响应数据中提取值
        支持多种层级指定方式：
        - 数字层级: 6 表示第6层（按嵌套深度遍历）
        - 路径方式: data.result.list 或 data>result>list 或 data/result/list
        - 空层级时直接用key在顶层查找
        """
        if data is None:
            return None

        # 将字符串形式的data转为对象
        if isinstance(data, str):
            try:
                data = json.loads(data)
            except Exception:
                pass

        # 方式1：路径方式 (支持 . > / 分隔)
        if level_str and any(sep in level_str for sep in [".", ">", "/"]):
            path = re.split(r"[.>/]+", level_str.strip())
            if path and path[-1] == key_str:
                # 层级已包含key
                path = path[:-1]
            return self._get_by_path(data, path + [key_str] if key_str else path)

        # 方式2：数字层级（按深度N层遍历查找key）
        if level_str and level_str.isdigit():
            depth = int(level_str)
            return self._find_by_depth_and_key(data, key_str, depth, current_depth=1)

        # 方式3：无层级或层级为空，直接按路径key查找
        if key_str:
            # 支持key本身就是路径
            if any(sep in key_str for sep in [".", ">", "/"]):
                path = re.split(r"[.>/]+", key_str.strip())
                return self._get_by_path(data, path)
            return self._get_by_path(data, [key_str])

        return None

    def _get_by_path(self, data, path):
        """按路径列表获取值"""
        if not path:
            return data
        current = data
        for key in path:
            if current is None:
                return None
            if isinstance(current, dict):
                if key in current:
                    current = current[key]
                else:
                    # 尝试数字索引
                    try:
                        idx = int(key)
                        if isinstance(current, list) and idx < len(current):
                            current = current[idx]
                        else:
                            return None
                    except ValueError:
                        return None
            elif isinstance(current, list):
                try:
                    idx = int(key)
                    if idx < len(current):
                        current = current[idx]
                    else:
                        return None
                except ValueError:
                    # 在列表中的每个dict里查找key
                    found = None
                    for item in current:
                        if isinstance(item, dict) and key in item:
                            found = item[key]
                            break
                    current = found
                    if current is None:
                        return None
            else:
                return None
        return current

    def _find_by_depth_and_key(self, data, target_key, target_depth, current_depth=1):
        """
        在指定深度层级查找key对应的值
        :param data: 数据
        :param target_key: 目标key
        :param target_depth: 目标深度（从1开始）
        :param current_depth: 当前深度
        """
        if data is None:
            return None

        if current_depth == target_depth:
            # 到达目标深度，查找key
            if isinstance(data, dict):
                if target_key in data:
                    return data[target_key]
                # 模糊匹配：不区分大小写
                for k, v in data.items():
                    if str(k).lower() == str(target_key).lower():
                        return v
            return None

        # 继续往下层遍历
        next_depth = current_depth + 1
        if isinstance(data, dict):
            for value in data.values():
                result = self._find_by_depth_and_key(value, target_key, target_depth, next_depth)
                if result is not None:
                    return result
        elif isinstance(data, list):
            for item in data:
                result = self._find_by_depth_and_key(item, target_key, target_depth, next_depth)
                if result is not None:
                    return result

        return None

    def _inject_dependency_value(self, case, dependency_field, value):
        """
        将依赖值注入到请求数据的指定字段中
        :param case: 测试用例
        :param dependency_field: 数据依赖字段（请求数据中的key，多个用逗号分隔）
        :param value: 要注入的值
        :return: 修改后的case
        """
        request_data_str = str(case.get("request_data", "")).strip()
        if not request_data_str:
            # 请求数据为空，创建包含依赖字段的新数据
            fields = [f.strip() for f in dependency_field.split(",") if f.strip()]
            new_data = {}
            for field in fields:
                new_data[field] = value
            case["request_data"] = json.dumps(new_data, ensure_ascii=False)
            return case

        # 解析请求数据
        try:
            request_data = json.loads(request_data_str)
        except Exception:
            # 非JSON格式，简单替换
            fields = [f.strip() for f in dependency_field.split(",") if f.strip()]
            for field in fields:
                # 替换占位符 {field} 或直接赋值
                placeholder = "{" + field + "}"
                if placeholder in request_data_str:
                    request_data_str = request_data_str.replace(placeholder, str(value))
            case["request_data"] = request_data_str
            return case

        # JSON格式，注入到指定字段
        if isinstance(request_data, dict):
            fields = [f.strip() for f in dependency_field.split(",") if f.strip()]
            for field in fields:
                # 支持多级路径赋值
                self._set_by_path(request_data, field, value)
            case["request_data"] = json.dumps(request_data, ensure_ascii=False)

        return case

    def _set_by_path(self, data, path_str, value):
        """按路径设置值，支持 . > / 分隔的多级路径"""
        path = re.split(r"[.>/]+", path_str.strip())
        if not path:
            return

        current = data
        for i, key in enumerate(path[:-1]):
            if isinstance(current, dict):
                if key not in current or not isinstance(current[key], (dict, list)):
                    current[key] = {}
                current = current[key]
            elif isinstance(current, list):
                try:
                    idx = int(key)
                    while idx >= len(current):
                        current.append({})
                    current = current[idx]
                except ValueError:
                    return
            else:
                return

        last_key = path[-1]
        if isinstance(current, dict):
            current[last_key] = value
        elif isinstance(current, list):
            try:
                idx = int(last_key)
                while idx >= len(current):
                    current.append(None)
                current[idx] = value
            except ValueError:
                pass

    def assert_result(self, expected, actual):
        """
        断言预期结果与实际结果。
        核心逻辑：实际结果包含预期结果的内容即为通过。
        如果预期结果有多个字段，逐个判断，全部通过才为通过。

        支持的预期结果格式：
        1. JSON对象，如 {"code": 200, "msg": "success"}
           → 逐字段检查，每个字段的值在实际结果中都能找到匹配
        2. 键值对，如 code:200,msg:success
           → 同上，每个键值对都要在实际结果中匹配
        3. 纯字符串，如 "success"
           → 实际结果的字符串形式包含该字符串即为通过

        :return: (是否通过, 断言说明)
        """
        if expected is None or str(expected).strip() == "":
            return True, "无预期结果，默认通过"

        expected_str = str(expected).strip()
        actual_str = str(actual) if actual is not None else ""

        # 尝试把实际结果解析为对象
        act_obj = None
        try:
            if isinstance(actual, (dict, list)):
                act_obj = actual
            elif actual_str and (actual_str.startswith("{") or actual_str.startswith("[")):
                act_obj = json.loads(actual_str)
        except Exception:
            pass

        # 尝试把预期结果解析为JSON对象
        exp_obj = None
        try:
            exp_obj = json.loads(expected_str)
        except Exception:
            pass

        # ------ 情况1：预期是JSON字典，逐字段包含判断 ------
        if isinstance(exp_obj, dict):
            # 实际结果也要是dict才能逐字段比对
            if act_obj is None:
                # 实际不是JSON，退化为字符串包含
                if expected_str in actual_str:
                    return True, f"字符串包含断言通过: 实际结果包含 '{expected_str[:80]}'"
                return False, f"断言失败。预期JSON但实际非JSON。预期: {expected_str}, 实际: {actual_str[:500]}"

            if not isinstance(act_obj, dict):
                return False, f"断言失败。预期为字典但实际不是字典。预期: {expected_str}, 实际: {str(actual)[:500]}"

            pass_items = []
            fail_items = []
            for k, exp_v in exp_obj.items():
                # 先看顶层，顶层没有就递归搜索嵌套字典
                act_v, found_path = self._find_first_key(act_obj, k, return_path=True)
                if found_path is None:
                    fail_items.append(f"字段[{k}]在实际结果中不存在（含所有嵌套层级）")
                    continue
                if self._value_contains(exp_v, act_v):
                    loc = "顶层" if found_path == k else f"路径.{found_path}"
                    pass_items.append(f"{k}={exp_v}({loc})")
                else:
                    loc = "顶层" if found_path == k else f"路径.{found_path}"
                    fail_items.append(f"字段[{k}]({loc}): 预期={exp_v}, 实际={act_v}")

            if not fail_items:
                return True, f"断言通过，全部字段匹配: {', '.join(pass_items)}"
            else:
                return False, f"断言失败，{len(fail_items)}个字段不匹配: {'; '.join(fail_items)}"

        # ------ 情况2：预期是键值对格式，如 code:200,msg:success ------
        if ":" in expected_str and not expected_str.startswith("{"):
            try:
                pairs = [p.strip() for p in expected_str.split(",") if ":" in p]
                if pairs:
                    pass_items = []
                    fail_items = []
                    for pair in pairs:
                        k, v = pair.split(":", 1)
                        k = k.strip()
                        v = v.strip()
                        actual_v = None
                        actual_v_str = ""
                        found_path = None
                        if act_obj is not None and isinstance(act_obj, dict):
                            # 先顶层，没有就递归搜索所有嵌套字典
                            actual_v, found_path = self._find_first_key(act_obj, k, return_path=True)
                            if actual_v is not None:
                                actual_v_str = str(actual_v)
                        # 尝试把预期值也做类型转换后比较
                        if actual_v_str == v or self._try_equal(v, actual_v):
                            loc = "顶层" if (found_path and found_path == k) else (f"路径.{found_path}" if found_path else "")
                            pass_items.append(f"{k}={v}" + (f"({loc})" if loc else ""))
                        else:
                            loc = "顶层" if (found_path and found_path == k) else (f"路径.{found_path}" if found_path else "未找到字段")
                            fail_items.append(f"字段[{k}]({loc}): 预期={v}, 实际={actual_v_str}")
                    if not fail_items:
                        return True, f"断言通过，全部键值对匹配: {', '.join(pass_items)}"
                    else:
                        return False, f"断言失败，{len(fail_items)}个键值对不匹配: {'; '.join(fail_items)}"
            except Exception:
                pass

        # ------ 情况3：纯字符串包含 ------
        if expected_str in actual_str:
            return True, f"字符串包含断言通过: 实际结果包含 '{expected_str[:80]}'"

        return False, f"断言失败。预期: {expected_str}, 实际: {actual_str[:500]}"

    def _value_contains(self, exp_v, act_v):
        """
        判断实际值是否"包含"预期值。
        - 字典：递归逐字段包含
        - 列表：预期列表的每个元素在实际列表中都能找到匹配
        - 其他类型：值相等，或字符串形式包含
        """
        # 两个都是dict → 递归包含
        if isinstance(exp_v, dict) and isinstance(act_v, dict):
            for k, v in exp_v.items():
                if k not in act_v:
                    return False
                if not self._value_contains(v, act_v[k]):
                    return False
            return True
        # 两个都是list → 预期每个元素在实际中存在
        if isinstance(exp_v, list) and isinstance(act_v, list):
            for exp_item in exp_v:
                found = False
                for act_item in act_v:
                    if self._value_contains(exp_item, act_item):
                        found = True
                        break
                if not found:
                    return False
            return True
        # 类型不同但值相等
        if exp_v == act_v:
            return True
        # 字符串包含
        if str(exp_v) in str(act_v):
            return True
        # 尝试类型转换后比较
        if self._try_equal(exp_v, act_v):
            return True
        return False

    def _find_first_key(self, obj, key, return_path=False, _prefix=""):
        """
        在 dict/list 的任意嵌套层级中查找第一个 key。
        :param obj: 搜索起点
        :param key: 要查找的字段名
        :param return_path: 是否同时返回找到的路径（点号拼接）
        :param _prefix: 递归内部使用的父路径
        :return: 默认返回 (value, None 或 path)；找不到返回 (None, None)
        """
        if isinstance(obj, dict):
            # 先在当前层找
            if key in obj:
                path = f"{_prefix}{key}" if _prefix else key
                return (obj[key], path)
            # 否则对每个字典/列表子项递归
            for k, v in obj.items():
                sub_prefix = f"{_prefix}{k}."
                res, res_path = self._find_first_key(v, key, True, sub_prefix)
                if res_path is not None:
                    return (res, res_path)
            return (None, None)
        if isinstance(obj, list):
            for idx, item in enumerate(obj):
                sub_prefix = f"{_prefix}[{idx}]."
                res, res_path = self._find_first_key(item, key, True, sub_prefix)
                if res_path is not None:
                    return (res, res_path)
            return (None, None)
        return (None, None)

    def _try_equal(self, a, b):
        """尝试把两个值转成相同类型后比较"""
        try:
            # 都是字符串时直接比
            if isinstance(a, str) and isinstance(b, str):
                return a == b
            # 尝试数值比较
            fa = float(a) if not isinstance(a, (dict, list)) else None
            fb = float(b) if not isinstance(b, (dict, list)) else None
            if fa is not None and fb is not None:
                return fa == fb
        except (ValueError, TypeError):
            pass
        # 最后退化为字符串比较
        return str(a) == str(b)

    def _dict_contains(self, small, big):
        """判断small字典是否是big字典的子集（递归）"""
        if not isinstance(small, dict) or not isinstance(big, dict):
            return small == big
        for k, v in small.items():
            if k not in big:
                return False
            if isinstance(v, dict):
                if not self._dict_contains(v, big[k]):
                    return False
            elif isinstance(v, list):
                if not isinstance(big[k], list):
                    return False
                # 列表：每个预期元素在实际中存在匹配
                for exp_item in v:
                    found = False
                    for act_item in big[k]:
                        if isinstance(exp_item, dict):
                            if self._dict_contains(exp_item, act_item):
                                found = True
                                break
                        else:
                            if exp_item == act_item:
                                found = True
                                break
                    if not found:
                        return False
            else:
                if v != big[k]:
                    return False
        return True
