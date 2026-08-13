# -*- coding: utf-8 -*-
"""
主测试执行引擎
"""
import os
import sys
import io
import json
import time
import shutil
import uuid
import traceback
import subprocess
import platform
from datetime import datetime

# 将项目根目录添加到path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.excel_util import ExcelUtil
from common.http_request import HttpRequest
from common.dependency_handler import DependencyHandler
from common.logger import get_logger
from config.settings import (
    REPORT_DIR, LOG_DIR,
    ALLURE_RESULTS_DIR, ALLURE_REPORT_DIR
)

logger = get_logger("test_engine")


# ===================== Allure 结果写入适配器 =====================
# 双通道策略：
#   A) 优先尝试 allure-pytest 提供的 allure_commons API（PluginManager + AllureFileLogger + AllureLifecycle）
#   B) 通道A失败时自动切换为"手动写 JSON/attachment 裸文件"（Allure官方结果格式，100%兼容CLI）
# 任一通道可用即视为可用；错误原因记录到 self.error_message 供 run.py 诊断侧展示
class AllureAdapter:
    """
    Allure结果文件适配器。
    职责：每条用例独立调用 start_case / stop_case，并在结果目录写入 JSON 文件
    （*-result.json / *-attachment.xxx），与 allure CLI 约定格式一致。
    """

    def __init__(self, results_dir, enable=True):
        self.results_dir = results_dir
        self.enable = enable
        self._avail = False
        self.error_message = ""
        self._mode = None        # "commons" 或 "manual"，标记当前使用哪个通道

        # allure_commons API 对象（通道A）
        self._plugin = None
        self._mc = None

        # 手动写入上下文（通道B）： {uuid: {"start_ts":..., "case":..., "attachments":[]}}
        self._manual_cases = {}

        if not self.enable:
            self.error_message = "Allure未启用（--no-allure）"
            return

        # 先确保目录存在 + 清理历史结果（这一步本身也可能失败，所以也放在try外统一处理）
        try:
            os.makedirs(self.results_dir, exist_ok=True)
            self._clear_results_dir()
        except Exception as e:
            self.error_message = f"创建/清理结果目录失败: {type(e).__name__}: {e}"
            logger.warning("⚠️  Allure适配器：" + self.error_message)
            return

        # --- 尝试通道A：allure_commons ---
        try:
            from allure_commons.plugin_manager import PluginManager
            from allure_commons.lifecycle import AllureLifecycle
            from allure_commons.logger import AllureFileLogger

            file_logger = AllureFileLogger(self.results_dir)
            pm = PluginManager()
            pm.register(file_logger)
            lifecycle = AllureLifecycle()
            pm.register(lifecycle)
            self._plugin = pm
            self._mc = lifecycle
            self._avail = True
            self._mode = "commons"
            logger.info(f"✅ Allure适配器初始化完成(通道A/allure_commons)，结果写入: {self.results_dir}")
            return
        except Exception as e:
            reason_a = f"通道A(allure_commons)失败: {type(e).__name__}: {e}"
            logger.warning("⚠️  Allure适配器：" + reason_a + "，将切换到通道B(手动写JSON)")
            self.error_message = reason_a

        # --- 尝试通道B：手动写 Allure 结果 JSON（完全不依赖API） ---
        try:
            import io as _io  # 预留
            # 写一个简单的environment.properties，能让Allure页面显示环境标签
            env_path = os.path.join(self.results_dir, "environment.properties")
            try:
                with open(env_path, "w", encoding="utf-8") as f:
                    f.write(f"Framework=ExcelApiTest\n")
                    f.write(f"OS={platform.platform()}\n")
                    f.write(f"Python={platform.python_version()}\n")
                    f.write(f"StartTime={datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            except Exception:
                pass
            # executor.json（在Allure报告页显示构建信息）
            try:
                executor = {
                    "name": "LocalRunner",
                    "type": "local",
                    "buildName": "excel-api-test run",
                    "reportName": "接口自动化Allure报告",
                }
                with open(os.path.join(self.results_dir, "executor.json"), "w", encoding="utf-8") as f:
                    json.dump(executor, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
            self._avail = True
            self._mode = "manual"
            logger.info(f"✅ Allure适配器初始化完成(通道B/手动JSON)，结果写入: {self.results_dir}")
            # 通道B可用时清空之前的通道A失败原因，避免误导
            self.error_message = ""
            return
        except Exception as e2:
            reason_b = f"通道B(手动JSON)也失败: {type(e2).__name__}: {e2}"
            logger.error("⚠️  Allure适配器：" + reason_b)
            self.error_message = (self.error_message + "\n" + reason_b) if self.error_message else reason_b
            self._avail = False

    def _clear_results_dir(self):
        if os.path.isdir(self.results_dir):
            for name in os.listdir(self.results_dir):
                p = os.path.join(self.results_dir, name)
                try:
                    if os.path.isfile(p) or os.path.islink(p):
                        os.remove(p)
                    elif os.path.isdir(p):
                        shutil.rmtree(p)
                except Exception:
                    pass

    @property
    def available(self):
        return self._avail

    # ---------- 便捷构造方法 ----------
    def _status(self, s):
        try:
            try:
                from allure_commons.types import Status
                mapping = {
                    "PASS": Status.PASSED,
                    "FAIL": Status.FAILED,
                    "SKIP": Status.SKIPPED,
                    "BROKEN": Status.BROKEN,
                }
                val = mapping.get(s.upper(), Status.FAILED)
                return getattr(val, "value", val)
            except Exception:
                return {"PASS": "passed", "FAIL": "failed",
                        "SKIP": "skipped", "BROKEN": "broken"}.get(s.upper(), "failed")
        except Exception:
            return "failed"

    def _labels_list(self, case_name, case_desc, row_num, method):
        labels = []
        label_names = {}
        try:
            from allure_commons.types import LabelType
            label_names = {
                "FRAMEWORK": LabelType.FRAMEWORK,
                "LANGUAGE": LabelType.LANGUAGE,
                "FEATURE": LabelType.FEATURE,
                "STORY": LabelType.STORY,
                "EPIC": LabelType.EPIC,
                "TAG": LabelType.TAG,
            }
        except Exception:
            # 手动写入的兜底：直接写字符串name
            label_names = {
                "FRAMEWORK": "framework",
                "LANGUAGE": "language",
                "FEATURE": "feature",
                "STORY": "story",
                "EPIC": "epic",
                "TAG": "tag",
            }
        def lb(name, value):
            return {"name": name, "value": value}
        labels.append(lb(str(label_names.get("FRAMEWORK", "framework")), "excel-api-test"))
        labels.append(lb(str(label_names.get("LANGUAGE", "language")), platform.system().lower()))
        labels.append(lb(str(label_names.get("FEATURE", "feature")), f"Excel行 {row_num}"))
        labels.append(lb(str(label_names.get("STORY", "story")), case_desc or case_name))
        labels.append(lb(str(label_names.get("EPIC", "epic")), "接口自动化"))
        labels.append(lb(str(label_names.get("TAG", "tag")), (method or "HTTP").upper()))
        labels.append(lb("suite", "Excel驱动接口测试"))
        labels.append(lb("package", "test_excel_cases"))
        labels.append(lb("testClass", "ExcelTestCase"))
        return labels

    # ---------- 生命周期 ----------
    def start_case(self, case_name, case_desc, row_num, method="GET", full_name=None):
        if not self.available:
            return

        uuid_str = str(uuid.uuid4())
        start_ts = int(time.time() * 1000)

        if self._mode == "commons":
            try:
                from allure_commons.model2 import TestResult, Label
                test_result = TestResult(
                    uuid=uuid_str,
                    name=case_name,
                    fullName=full_name or f"excel_api_test.{case_name}",
                    description=(case_desc or "") + f"\n\nExcel行号: {row_num}",
                    historyId=str(uuid.uuid5(uuid.NAMESPACE_DNS, case_name or "case")),
                    labels=[],
                    links=[],
                    start=start_ts,
                )
                for d in self._labels_list(case_name, case_desc, row_num, method):
                    try:
                        test_result.labels.append(Label(name=d["name"], value=d["value"]))
                    except Exception:
                        pass
                self._mc.schedule_test_case(test_result)
                self._mc.start_test_case(uuid_str)
                self._current_uuid = uuid_str
                return
            except Exception as e:
                logger.warning(f"Allure通道A start_case 失败，切到手动通道: {e}")
                self._mode = "manual"

        # 通道B
        self._manual_cases[uuid_str] = {
            "uuid": uuid_str,
            "name": case_name,
            "fullName": full_name or f"excel_api_test.{case_name}",
            "description": (case_desc or "") + f"\n\nExcel行号: {row_num}",
            "historyId": str(uuid.uuid5(uuid.NAMESPACE_DNS, case_name or "case")),
            "start": start_ts,
            "stop": None,
            "status": None,
            "statusDetails": None,
            "labels": self._labels_list(case_name, case_desc, row_num, method),
            "links": [],
            "attachments": [],
            "testStage": {
                "name": case_name,
                "start": start_ts,
                "attachments": [],
                "steps": [],
            },
        }
        self._current_uuid = uuid_str

    def attach(self, name, body, mime="text/plain", ext=None):
        if not self.available:
            return
        # 统一序列化 body -> (bytes, mime, source_ext_for_allure)
        try:
            if isinstance(body, (dict, list)):
                body = json.dumps(body, ensure_ascii=False, indent=2)
                mime = "application/json"
            if isinstance(body, str):
                body_bytes = body.encode("utf-8")
            elif isinstance(body, bytes):
                body_bytes = body
            else:
                body_bytes = str(body).encode("utf-8")
        except Exception as e:
            logger.warning(f"Allure attach 序列化失败: {e}")
            return

        # 推断 ext（allure 用 extension 字段，直接跟附件文件后缀相关）
        ext_str = ext if isinstance(ext, str) else None
        if ext_str is None:
            if mime == "application/json":
                ext_str = "json"
            elif "xml" in mime:
                ext_str = "xml"
            elif "html" in mime:
                ext_str = "html"
            elif "image/png" in mime:
                ext_str = "png"
            elif "image/jpeg" in mime or "image/jpg" in mime:
                ext_str = "jpg"
            else:
                ext_str = "txt"

        # 将 ext 枚举转成字符串 value
        if ext is not None and not isinstance(ext, str):
            try:
                ext_str = getattr(ext, "value", ext_str)
            except Exception:
                pass

        current_uuid = getattr(self, "_current_uuid", None)
        if not current_uuid:
            return

        if self._mode == "commons":
            try:
                from allure_commons.types import AttachmentType
                # 优先根据 ext_str 反向构造 AttachmentType，找不到则按 mime 兜底
                at = None
                try:
                    at = AttachmentType(ext_str)
                except Exception:
                    mapping = {
                        "json": AttachmentType.JSON, "txt": AttachmentType.TEXT,
                        "xml": AttachmentType.XML, "html": AttachmentType.HTML,
                        "png": AttachmentType.PNG, "jpg": AttachmentType.JPG,
                    }
                    at = mapping.get(ext_str, AttachmentType.TEXT)
                self._mc.attach_data(body_bytes, name=name, attachment_type=at)
                return
            except Exception as e:
                logger.warning(f"Allure通道A attach 失败，回退手动: {e}")
                self._mode = "manual"

        # 通道B：写附件文件，并把 attachment 描述塞到 case.testStage.attachments
        att_uuid = uuid.uuid4().hex
        # Allure 附件文件名格式是 <uuid>-attachment.<ext>
        file_name = f"{att_uuid}-attachment.{ext_str}"
        file_path = os.path.join(self.results_dir, file_name)
        try:
            with open(file_path, "wb") as f:
                f.write(body_bytes)
        except Exception as e:
            logger.warning(f"Allure 附件落盘失败: {e}")
            return
        attachment_entry = {
            "name": name,
            "source": file_name,
            "type": mime,
        }
        ctx = self._manual_cases.get(current_uuid)
        if ctx is not None:
            ctx.setdefault("attachments", []).append(attachment_entry)
            ctx["testStage"].setdefault("attachments", []).append(attachment_entry)

    def stop_case(self, status, message="", trace_msg=""):
        if not self.available:
            return
        current_uuid = getattr(self, "_current_uuid", None)
        if not current_uuid:
            return

        stop_ts = int(time.time() * 1000)
        final_status = self._status(status)

        status_details = None
        if message or trace_msg:
            status_details = {
                "message": (message or "")[:5000],
                "trace": (trace_msg or "")[:10000],
            }

        if self._mode == "commons":
            try:
                self._mc.update_test_case(current_uuid, lambda tr: setattr(tr, "stop", stop_ts))
                self._mc.update_test_case(current_uuid, lambda tr: setattr(tr, "status", self._status(status)))
                if status_details is not None:
                    try:
                        from allure_commons.model2 import StatusDetails
                        sd = StatusDetails(message=status_details["message"], trace=status_details["trace"])
                        self._mc.update_test_case(current_uuid, lambda tr, sd=sd: setattr(tr, "statusDetails", sd))
                    except Exception:
                        pass
                self._mc.write_test_case(current_uuid)
                return
            except Exception as e:
                logger.warning(f"Allure通道A stop_case 失败，回退手动JSON: {e}")
                # 把已经可能写入一半的结果，在手动通道里也把 start/status 补齐
                if current_uuid not in self._manual_cases:
                    self._manual_cases[current_uuid] = {
                        "uuid": current_uuid, "name": "case", "fullName": "case",
                        "description": "", "historyId": "", "start": stop_ts - 1000,
                        "stop": None, "status": None, "statusDetails": None,
                        "labels": [], "links": [], "attachments": [],
                        "testStage": {"name": "", "start": stop_ts - 1000, "attachments": [], "steps": []},
                    }
                self._mode = "manual"

        # 通道B：组装并落盘 *-result.json
        ctx = self._manual_cases.pop(current_uuid, None)
        if ctx is None:
            return
        ctx["stop"] = stop_ts
        ctx["status"] = final_status
        ctx["testStage"]["stop"] = stop_ts
        ctx["testStage"]["status"] = final_status
        if status_details:
            ctx["statusDetails"] = status_details
        # 写 stages（Allure 约定需要 testStage，同时外层attachments如果不生效可以直接放testStage下的，我们两边都放）

        result_file = os.path.join(self.results_dir, f"{current_uuid}-result.json")
        try:
            with open(result_file, "w", encoding="utf-8") as f:
                json.dump(ctx, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"Allure手动JSON写入失败: {e}")


class TestEngine:
    """测试执行引擎"""

    def __init__(self, excel_path=None, sheet_name=None, allure_enable=True, allure_serve=True):
        self.excel_util = ExcelUtil(excel_path, sheet_name)
        self.http_request = HttpRequest()
        self.dependency_handler = DependencyHandler(self.http_request, self.excel_util)

        # 统计数据
        self.total_count = 0
        self.run_count = 0
        self.pass_count = 0
        self.fail_count = 0
        self.skip_count = 0
        self.results = []

        # 执行顺序缓存（用于依赖排序）
        self.executed_names = set()

        # Allure适配器
        self.allure_enable = allure_enable
        self._allure_serve = allure_serve
        self.allure = AllureAdapter(ALLURE_RESULTS_DIR, enable=allure_enable)

    def run_all(self):
        """执行所有测试用例"""
        logger.info("=" * 80)
        logger.info("接口自动化测试开始执行")
        if self.allure_enable:
            if self.allure.available:
                logger.info("✅ Allure结果记录已启用")
            else:
                logger.warning("⚠️  Allure已开启但不可用（缺少依赖或初始化失败），将仅输出HTML报告")
        logger.info("=" * 80)
        start_time = time.time()

        # 1. 读取所有用例
        all_cases = self.excel_util.read_all_cases()
        self.total_count = len(all_cases)
        logger.info(f"共读取测试用例: {self.total_count} 条")

        # 2. 先用拓扑排序或按依赖排序
        sorted_cases = self._sort_cases_by_dependency(all_cases)

        # 3. 遍历执行
        for case in sorted_cases:
            self._execute_single_case(case)

        # 4. 输出统计结果
        end_time = time.time()
        duration = end_time - start_time
        self._print_summary(duration)

        # 5. 生成报告（HTML + Allure）
        allure_status = self._generate_report(duration)

        return {
            "total": self.total_count,
            "run": self.run_count,
            "pass": self.pass_count,
            "fail": self.fail_count,
            "skip": self.skip_count,
            "duration": round(duration, 2),
            "results": self.results,
            "allure": allure_status or {"ok": False, "message": "未启用"},
        }

    def _sort_cases_by_dependency(self, cases):
        """
        根据依赖关系对用例进行排序，确保被依赖的用例先执行
        使用拓扑排序（Kahn算法）
        """
        # 构建 name -> case 映射
        name_to_case = {}
        for case in cases:
            name = case.get("interface_name", "")
            if name:
                name_to_case[name] = case

        # 构建依赖图
        in_degree = {}  # 入度（依赖数量）
        dependents = {}  # 被谁依赖
        case_names = []

        for case in cases:
            name = case.get("interface_name", "")
            case_names.append(name)
            dep = case.get("case_dependency", "").strip()
            in_degree[name] = 0
            dependents[name] = []

        for case in cases:
            name = case.get("interface_name", "")
            dep = case.get("case_dependency", "").strip()
            if dep and dep in name_to_case:
                in_degree[name] += 1
                if dep in dependents:
                    dependents[dep].append(name)

        # 拓扑排序
        sorted_names = []
        queue = [name for name in case_names if in_degree.get(name, 0) == 0]

        while queue:
            current = queue.pop(0)
            sorted_names.append(current)
            for dependent in dependents.get(current, []):
                in_degree[dependent] -= 1
                if in_degree[dependent] == 0:
                    queue.append(dependent)

        # 将未入环的用例也加进来
        for name in case_names:
            if name not in sorted_names:
                sorted_names.append(name)

        # 按排好的顺序返回case
        sorted_cases = []
        for name in sorted_names:
            for case in cases:
                if case.get("interface_name", "") == name:
                    sorted_cases.append(case)
                    break
        return sorted_cases

    def _execute_single_case(self, case):
        """执行单条测试用例"""
        case_name = case.get("interface_name", "未命名")
        case_desc = case.get("interface_desc", "")
        is_run = str(case.get("is_run", "yes")).strip().lower()
        row_num = case.get("row_num", 0)
        method = str(case.get("request_type", "GET") or "GET").strip().lower()

        # 记录结果
        result_record = {
            "row_num": row_num,
            "interface_name": case_name,
            "interface_desc": case_desc,
            "status": "",
            "message": "",
            "start_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "response": None
        }

        logger.info("-" * 80)
        logger.info(f"准备执行用例: [{case_name}] - {case_desc}")

        # ============ Allure: start_case ============
        if is_run not in ["yes", "y", "true", "1", "是"]:
            # 跳过也要在allure中记录（标记SKIP），保持报告完整
            self.allure.start_case(case_name, case_desc, row_num, method=method)
            self.allure.attach("Excel用例元信息",
                               {"是否运行": is_run, "Excel行": row_num},
                               ext=None)
            logger.info(f"跳过执行: 是否运行为 [{is_run}]")
            self.skip_count += 1
            result_record["status"] = "SKIP"
            result_record["message"] = f"跳过（是否运行为{is_run}）"
            self.allure.stop_case("SKIP", message=result_record["message"])
            self.results.append(result_record)
            self.excel_util.write_actual_result(row_num, "跳过执行")
            return

        self.run_count += 1
        self.allure.start_case(case_name, case_desc, row_num, method=method)

        # Allure：把请求基本信息作为附件
        try:
            self.allure.attach("Excel用例元信息", {
                "Excel行号": row_num,
                "接口名称": case_name,
                "接口描述": case_desc,
                "URL": case.get("url", ""),
                "请求方式": method.upper(),
                "case依赖": case.get("case_dependency", ""),
                "数据依赖字段": case.get("dependency_field", ""),
            })
            if case.get("request_data"):
                self.allure.attach("请求数据(Request Body)",
                                   case.get("request_data"),
                                   mime="application/json")
        except Exception:
            pass

        try:
            # 2. 处理依赖
            case, skip, skip_msg = self.dependency_handler.handle_dependency(case)
            if skip:
                self.skip_count += 1
                result_record["status"] = "SKIP"
                result_record["message"] = skip_msg or "依赖问题跳过"
                self.allure.attach("依赖处理结果", {
                    "跳过": True, "原因": skip_msg
                })
                self.allure.stop_case("SKIP", message=result_record["message"])
                self.results.append(result_record)
                self.excel_util.write_actual_result(row_num, skip_msg or "跳过执行")
                return

            # 3. 发送请求
            response = self.http_request.send_request(case)
            result_record["response"] = response

            # Allure：附加响应
            try:
                self.allure.attach("响应内容(Response)", response,
                                   mime="application/json")
            except Exception:
                pass

            # 4. 断言结果
            expected = case.get("expected_result", "")
            is_pass, assert_msg = self.dependency_handler.assert_result(expected, response)
            result_record["message"] = assert_msg

            # Allure：附加断言
            try:
                self.allure.attach("断言信息", {
                    "预期结果": expected,
                    "断言结论": "通过" if is_pass else "失败",
                    "断言说明": assert_msg
                })
            except Exception:
                pass

            # 5. 写入实际结果
            response_str = self._format_response(response)
            result_with_assert = f"{assert_msg}\n响应内容: {response_str}"
            self.excel_util.write_actual_result(row_num, result_with_assert, is_pass)

            if is_pass:
                logger.info(f"✓ 用例 [{case_name}] 执行通过: {assert_msg}")
                self.pass_count += 1
                result_record["status"] = "PASS"
                self.allure.stop_case("PASS", message=assert_msg)
            else:
                logger.error(f"✗ 用例 [{case_name}] 执行失败: {assert_msg}")
                self.fail_count += 1
                result_record["status"] = "FAIL"
                self.allure.stop_case("FAIL", message=assert_msg,
                                      trace_msg=f"预期: {expected}\n实际: {response_str}")

            # 标记已执行（用于依赖处理）
            self.executed_names.add(case_name)

        except Exception as e:
            tb_str = traceback.format_exc()
            logger.exception(f"用例 [{case_name}] 执行异常: {str(e)}")
            self.fail_count += 1
            result_record["status"] = "FAIL"
            result_record["message"] = f"执行异常: {str(e)}"
            try:
                self.allure.attach("异常堆栈", tb_str, mime="text/plain")
            except Exception:
                pass
            self.allure.stop_case("BROKEN", message=str(e), trace_msg=tb_str)
            self.excel_util.write_actual_result(
                row_num,
                f"执行异常: {str(e)}",
                False
            )

        result_record["end_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.results.append(result_record)

    def _format_response(self, response):
        """格式化响应结果为字符串"""
        if response is None:
            return "None"
        if isinstance(response, (dict, list)):
            try:
                return json.dumps(response, ensure_ascii=False)
            except Exception:
                return str(response)
        return str(response)

    def _print_summary(self, duration):
        """打印测试汇总"""
        logger.info("=" * 80)
        logger.info("测试执行汇总")
        logger.info("=" * 80)
        logger.info(f"用例总数:     {self.total_count}")
        logger.info(f"执行用例数:   {self.run_count}")
        logger.info(f"通过数:       {self.pass_count}")
        logger.info(f"失败数:       {self.fail_count}")
        logger.info(f"跳过数:       {self.skip_count}")
        pass_rate = (self.pass_count / self.run_count * 100) if self.run_count > 0 else 0
        logger.info(f"通过率:       {pass_rate:.2f}%")
        logger.info(f"总耗时:       {duration:.2f} 秒")
        logger.info("=" * 80)

    def _generate_report(self, duration):
        """生成HTML测试报告"""
        import html
        if not os.path.exists(REPORT_DIR):
            os.makedirs(REPORT_DIR)
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(REPORT_DIR, f"report_{timestamp}.html")

        pass_rate = (self.pass_count / self.run_count * 100) if self.run_count > 0 else 0

        # 构建结果行
        rows_html = ""
        for idx, r in enumerate(self.results, 1):
            status = r.get("status", "")
            if status == "PASS":
                status_class = "pass"
            elif status == "FAIL":
                status_class = "fail"
            elif status == "SKIP":
                status_class = "skip"
            else:
                status_class = ""

            resp_str = ""
            resp_full = ""
            if r.get("response") is not None:
                try:
                    # 美观格式化的完整内容（用于弹窗/复制）
                    resp_full = json.dumps(r["response"], ensure_ascii=False, indent=2)
                    # 表格内紧凑展示
                    resp_str = json.dumps(r["response"], ensure_ascii=False)
                except Exception:
                    resp_str = str(r["response"])
                    resp_full = resp_str

            msg = str(r.get("message", ""))
            interface_name = str(r.get("interface_name", ""))
            interface_desc = str(r.get("interface_desc", ""))
            start_time = str(r.get("start_time", ""))

            # 截断长度
            msg_trunc = msg if len(msg) <= 300 else msg[:300] + "..."
            resp_trunc = resp_str if len(resp_str) <= 800 else resp_str[:800] + "..."

            # HTML转义（避免内容破坏页面）
            msg_esc = html.escape(msg)
            msg_trunc_esc = html.escape(msg_trunc)
            resp_full_esc = html.escape(resp_full)
            resp_trunc_esc = html.escape(resp_trunc)
            interface_name_esc = html.escape(interface_name)
            interface_desc_esc = html.escape(interface_desc)
            start_time_esc = html.escape(start_time)

            rows_html += f"""
            <tr>
                <td>{idx}</td>
                <td>{r.get('row_num', '')}</td>
                <td>{interface_name_esc}</td>
                <td title="{interface_desc_esc}" class="cell-clamp">{interface_desc_esc}</td>
                <td class="{status_class}">{status}</td>
                <td title="{msg_esc}" class="cell-clamp cell-msg">{msg_trunc_esc}
                    {f'<button class="btn-detail" onclick="openDetail(\'断言/备注 - {idx}\', this.dataset.c)" data-c="{msg_esc}">详情</button>' if len(msg) > 300 else ''}
                </td>
                <td title="{resp_full_esc}" class="cell-clamp cell-resp">{resp_trunc_esc}
                    <button class="btn-detail" onclick="openDetail(\'响应内容 - {idx}\', this.dataset.c)" data-c="{resp_full_esc}">查看完整</button>
                </td>
                <td>{start_time_esc}</td>
            </tr>
            """

        html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>接口自动化测试报告</title>
    <style>
        body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 20px; background: #f5f7fa; }}
        h1 {{ color: #303133; border-bottom: 2px solid #409EFF; padding-bottom: 10px; }}
        .summary {{ display: flex; gap: 20px; margin: 20px 0; flex-wrap: wrap; }}
        .card {{ background: #fff; border-radius: 8px; padding: 20px; min-width: 150px;
                 box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }}
        .card .num {{ font-size: 32px; font-weight: bold; margin: 10px 0; }}
        .card .label {{ color: #909399; font-size: 14px; }}
        .total .num {{ color: #909399; }}
        .run .num {{ color: #409EFF; }}
        .pass .num {{ color: #67C23A; }}
        .fail .num {{ color: #F56C6C; }}
        .skip .num {{ color: #E6A23C; }}
        .rate .num {{ color: #409EFF; }}
        table {{ width: 100%; border-collapse: collapse; background: #fff;
                 border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); table-layout: fixed; }}
        th {{ background: #409EFF; color: #fff; padding: 12px 8px; text-align: left; }}
        th:nth-child(1) {{ width: 50px; }}
        th:nth-child(2) {{ width: 70px; }}
        th:nth-child(5) {{ width: 70px; }}
        th:nth-child(8) {{ width: 150px; }}
        td {{ padding: 10px 8px; border-bottom: 1px solid #ebeef5; font-size: 13px; vertical-align: top;
              word-break: break-word; overflow-wrap: break-word; }}
        tr:hover {{ background: #f5f7fa; }}
        .pass {{ color: #67C23A; font-weight: bold; }}
        .fail {{ color: #F56C6C; font-weight: bold; }}
        .skip {{ color: #E6A23C; font-weight: bold; }}
        .info {{ color: #606266; margin: 10px 0; }}

        /* 单元格超长截断 + hover原生title气泡 */
        .cell-clamp {{
            display: -webkit-box;
            -webkit-line-clamp: 4;
            -webkit-box-orient: vertical;
            overflow: hidden;
            text-overflow: ellipsis;
            max-height: 96px;
            line-height: 1.5;
        }}
        .cell-msg {{ max-width: 320px; }}
        .cell-resp {{
            max-width: 500px;
            font-family: Consolas, "Courier New", monospace;
            font-size: 12px;
            background: #fafbfc;
            white-space: pre-wrap;
        }}

        /* 详情按钮 */
        .btn-detail {{
            display: inline-block;
            margin-top: 6px;
            padding: 3px 10px;
            background: #409EFF;
            color: #fff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        }}
        .btn-detail:hover {{ background: #337ecc; }}

        /* 弹窗遮罩 */
        .modal-mask {{
            display: none;
            position: fixed; inset: 0;
            background: rgba(0,0,0,0.5);
            z-index: 9998;
            justify-content: center;
            align-items: center;
        }}
        .modal-mask.show {{ display: flex; }}
        .modal {{
            background: #fff;
            width: 85%;
            max-width: 900px;
            max-height: 85vh;
            border-radius: 8px;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }}
        .modal-header {{
            padding: 14px 18px;
            background: #409EFF;
            color: #fff;
            font-weight: bold;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .modal-close {{
            background: transparent;
            border: none;
            color: #fff;
            font-size: 20px;
            cursor: pointer;
            line-height: 1;
        }}
        .modal-actions {{
            padding: 8px 18px;
            background: #f5f7fa;
            border-bottom: 1px solid #ebeef5;
            display: flex;
            gap: 10px;
        }}
        .modal-actions button {{
            padding: 5px 14px;
            border: 1px solid #dcdfe6;
            background: #fff;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
        }}
        .modal-actions button:hover {{ border-color: #409EFF; color: #409EFF; }}
        .modal-body {{
            flex: 1;
            padding: 18px;
            overflow: auto;
            font-family: Consolas, "Courier New", monospace;
            font-size: 13px;
            background: #fafbfc;
            white-space: pre-wrap;
            word-break: break-word;
            line-height: 1.6;
        }}
    </style>
</head>
<body>
    <h1>接口自动化测试报告</h1>
    <div class="info">
        <div>生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
        <div>总耗时: {duration:.2f} 秒</div>
        <div style="color:#909399;font-size:12px;margin-top:4px;">
            💡 提示：鼠标悬停可预览完整内容；内容过长时点击【查看完整】在弹窗中查看/复制
        </div>
    </div>
    <div class="summary">
        <div class="card total">
            <div class="label">用例总数</div>
            <div class="num">{self.total_count}</div>
        </div>
        <div class="card run">
            <div class="label">执行数</div>
            <div class="num">{self.run_count}</div>
        </div>
        <div class="card pass">
            <div class="label">通过</div>
            <div class="num">{self.pass_count}</div>
        </div>
        <div class="card fail">
            <div class="label">失败</div>
            <div class="num">{self.fail_count}</div>
        </div>
        <div class="card skip">
            <div class="label">跳过</div>
            <div class="num">{self.skip_count}</div>
        </div>
        <div class="card rate">
            <div class="label">通过率</div>
            <div class="num">{pass_rate:.1f}%</div>
        </div>
    </div>
    <h2>详细结果</h2>
    <table>
        <thead>
            <tr>
                <th>序号</th>
                <th>Excel行</th>
                <th>接口名称</th>
                <th>接口描述</th>
                <th>状态</th>
                <th>断言/备注</th>
                <th>响应内容</th>
                <th>开始时间</th>
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>

    <!-- 详情弹窗 -->
    <div class="modal-mask" id="modalMask" onclick="if(event.target===this)closeDetail()">
        <div class="modal">
            <div class="modal-header">
                <span id="modalTitle">详情</span>
                <button class="modal-close" onclick="closeDetail()">&times;</button>
            </div>
            <div class="modal-actions">
                <button onclick="copyDetail()">📋 复制全文</button>
                <button onclick="tryFormatJSON()">✨ 美化JSON</button>
            </div>
            <div class="modal-body" id="modalBody"></div>
        </div>
    </div>

    <script>
    // 为了f-string不冲突，JS里的花括号用原始内容（模板中未使用Python占位花括号）
    var lastContent = "";
    function openDetail(title, content) {{
        document.getElementById("modalTitle").innerText = title;
        document.getElementById("modalBody").innerText = content;
        lastContent = content;
        document.getElementById("modalMask").classList.add("show");
        document.body.style.overflow = "hidden";
    }}
    function closeDetail() {{
        document.getElementById("modalMask").classList.remove("show");
        document.body.style.overflow = "";
    }}
    function copyDetail() {{
        var txt = document.getElementById("modalBody").innerText;
        navigator.clipboard.writeText(txt).then(function() {{
            alert("已复制到剪贴板");
        }}, function() {{
            // 兼容回退
            var ta = document.createElement("textarea");
            ta.value = txt;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
            alert("已复制到剪贴板");
        }});
    }}
    function tryFormatJSON() {{
        var txt = lastContent;
        try {{
            var obj = JSON.parse(txt);
            document.getElementById("modalBody").innerText = JSON.stringify(obj, null, 2);
        }} catch(e) {{
            alert("内容不是合法JSON，无法美化：" + e.message);
        }}
    }}
    // ESC键关闭弹窗
    document.addEventListener("keydown", function(e) {{
        if(e.key === "Escape") closeDetail();
    }});
    </script>
</body>
</html>
        """

        try:
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(html_content)
            logger.info(f"HTML报告已生成: {report_path}")
        except Exception as e:
            logger.error(f"生成HTML报告失败: {str(e)}")

        # 6. Allure报告生成（CLI探测 + 安全降级）
        allure_status = {"ok": False, "report_dir": None,
                         "results_dir": ALLURE_RESULTS_DIR, "message": "",
                         "mode": getattr(self.allure, "_mode", None)}
        if self.allure_enable:
            # 即使适配器不可用，也把初始化失败的原因透传出来（便于run.py诊断）
            if not self.allure.available:
                adapter_err = getattr(self.allure, "error_message", "") or "Allure适配器初始化失败（原因未知）"
                allure_status["message"] = (
                    "Allure适配器不可用，因此没有写入结果文件、无法生成报告。\n"
                    "适配器初始化失败原因:\n"
                    f"  {adapter_err}\n"
                    "说明: 我们做了双通道兜底——要么用allure_commons API，要么手动写JSON，"
                    "两边都失败时才会出现这个提示，建议对照上面的报错信息逐个处理，"
                    "或直接 pip install --upgrade allure-pytest 后重试。"
                )
            else:
                allure_status = self._generate_allure_report()
                allure_status["mode"] = getattr(self.allure, "_mode", None)
        return allure_status

    def _generate_allure_report(self):
        """
        使用 Allure CLI 将 allure-results 转成 HTML 报告。
        严格做三层防护：
          1) shutil.which 探测 allure 可执行文件
          2) subprocess 捕获 FileNotFoundError / CalledProcessError
          3) 失败仅打印清晰的用户提示，不会中断整体流程
        """
        status = {"ok": False, "report_dir": ALLURE_REPORT_DIR,
                  "results_dir": ALLURE_RESULTS_DIR, "message": ""}
        logger.info("=" * 80)
        logger.info("准备生成 Allure 报告...")

        if not os.path.isdir(ALLURE_RESULTS_DIR):
            msg = f"Allure结果目录不存在: {ALLURE_RESULTS_DIR}"
            logger.warning(msg)
            status["message"] = msg
            return status

        # 检查结果文件是否为空
        try:
            result_files = [f for f in os.listdir(ALLURE_RESULTS_DIR) if f.endswith("-result.json")]
        except Exception:
            result_files = []
        if not result_files:
            msg = f"Allure结果目录中没有检测到用例结果文件（*-result.json），跳过生成报告"
            logger.warning(msg)
            status["message"] = msg
            return status
        logger.info(f"检测到 {len(result_files)} 个用例结果文件，准备调用 allure CLI...")

        # 1) 探测 allure 命令
        allure_bin = shutil.which("allure")
        if not allure_bin:
            # 尝试常见的 scoop/chocolatey/npm 本地路径
            candidates = []
            if os.name == "nt":
                npm_root = os.environ.get("APPDATA")
                if npm_root:
                    candidates.append(os.path.join(npm_root, "npm", "allure.cmd"))
                scoop = os.path.join(os.environ.get("USERPROFILE", ""), "scoop", "shims", "allure.cmd")
                candidates.append(scoop)
                choco = r"C:\ProgramData\chocolatey\bin\allure.exe"
                candidates.append(choco)
            for c in candidates:
                if os.path.isfile(c):
                    allure_bin = c
                    break

        if not allure_bin:
            msg = (
                "⚠️  未检测到 Allure CLI 命令行工具。\n"
                "    已完成 allure-results 结果文件生成，但无法生成 allure-report HTML。\n"
                "    推荐安装方式（任选其一）:\n"
                "      1) npm i -g allure-commandline   (需要Node.js)\n"
                "      2) scoop install allure          (Windows推荐)\n"
                "      3) choco install allure          (需 Chocolatey)\n"
                "    或手动下载: https://repo.maven.apache.org/maven2/io/qameta/allure/allure-commandline/\n"
                f"    安装后可手动生成: allure generate {ALLURE_RESULTS_DIR} -o {ALLURE_REPORT_DIR} --clean\n"
                f"    查看报告: allure open {ALLURE_REPORT_DIR}"
            )
            logger.warning(msg)
            status["message"] = msg
            # 同时打印到终端便于用户看到
            print("\n" + msg + "\n")
            return status

        logger.info(f"✅ 检测到 Allure CLI: {allure_bin}")

        # 2) 清理旧的报告目录并生成
        try:
            os.makedirs(ALLURE_REPORT_DIR, exist_ok=True)
        except Exception as e:
            msg = f"创建Allure报告目录失败: {e}"
            logger.error(msg)
            status["message"] = msg
            return status

        try:
            # allure generate <results> -o <report> --clean
            cmd = [allure_bin, "generate", ALLURE_RESULTS_DIR,
                   "-o", ALLURE_REPORT_DIR, "--clean"]
            logger.info(f"执行: {' '.join(cmd)}")
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace",
                                  timeout=5 * 60)  # 5分钟超时
            if proc.returncode != 0:
                msg = (f"Allure CLI 执行失败(exit={proc.returncode})。\n"
                       f"stdout: {proc.stdout or ''}\nstderr: {proc.stderr or ''}")
                logger.error(msg)
                status["message"] = msg
                return status
            index_html = os.path.join(ALLURE_REPORT_DIR, "index.html")
            if os.path.isfile(index_html):
                logger.info("✅ Allure报告生成成功: " + ALLURE_REPORT_DIR)
                logger.info(f"    - 报告目录: {ALLURE_REPORT_DIR}")
                logger.info(f"    - ⚠️ 不要直接双击 index.html（file://协议会被浏览器CORS拦截导致500）")
                logger.info(f"    - ✅ 正确查看方式: 启动本地HTTP服务（见下方自动启动）或 allure open {ALLURE_REPORT_DIR}")
                status["ok"] = True
                status["message"] = "ok"
                status["report_dir"] = ALLURE_REPORT_DIR
                # 自动启动本地HTTP服务并打开浏览器
                if getattr(self, "_allure_serve", True):
                    serve_url = self._serve_allure_report()
                    if serve_url:
                        status["serve_url"] = serve_url
            else:
                msg = (f"allure命令执行成功但未生成index.html，请检查目录内容: {ALLURE_REPORT_DIR}")
                logger.warning(msg)
                status["message"] = msg
        except FileNotFoundError as e:
            msg = f"调用allure失败 FileNotFoundError: {e}"
            logger.error(msg)
            status["message"] = msg
        except subprocess.TimeoutExpired:
            msg = "生成Allure报告超时(>5分钟)"
            logger.error(msg)
            status["message"] = msg
        except Exception as e:
            msg = f"生成Allure报告出现异常: {e}"
            logger.exception(msg)
            status["message"] = msg
        return status

    def _serve_allure_report(self):
        """
        启动一个本地HTTP服务来托管Allure报告，并自动打开浏览器。
        Allure的index.html依赖fetch()加载JSON数据，file://协议会被浏览器CORS拦截，
        必须通过HTTP协议访问才能正常显示。
        服务在后台线程运行，主进程退出时自动结束。
        """
        import threading
        import socket
        from http.server import HTTPServer, SimpleHTTPRequestHandler
        import functools
        import webbrowser

        # 找一个可用端口（从9527开始尝试，最多试20个）
        port = 9527
        for p in range(port, port + 20):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(("127.0.0.1", p))
                    port = p
                    break
            except OSError:
                continue
        else:
            logger.warning("⚠️  找不到可用端口启动Allure报告服务，请手动执行: "
                           f"cd \"{ALLURE_REPORT_DIR}\" && python -m http.server 9527")
            return None

        report_dir = ALLURE_REPORT_DIR
        if not os.path.isdir(report_dir):
            logger.warning(f"⚠️  Allure报告目录不存在: {report_dir}")
            return None

        # 使用 SimpleHTTPRequestHandler 并指定目录
        Handler = functools.partial(SimpleHTTPRequestHandler, directory=report_dir)
        try:
            httpd = HTTPServer(("127.0.0.1", port), Handler)
        except Exception as e:
            logger.warning(f"⚠️  启动HTTP服务失败: {e}")
            return None

        serve_url = f"http://127.0.0.1:{port}/"
        logger.info(f"✅ Allure报告HTTP服务已启动: {serve_url}")
        logger.info(f"    报告目录: {report_dir}")
        logger.info(f"    按 Ctrl+C 可停止服务（或直接关闭终端）")

        # 后台线程运行HTTP服务
        server_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        server_thread.start()

        # 自动打开浏览器
        try:
            webbrowser.open(serve_url)
            logger.info(f"    已自动打开浏览器: {serve_url}")
        except Exception:
            logger.info(f"    浏览器未自动打开，请手动访问: {serve_url}")

        # 提示用户服务在前台运行
        print("\n" + "=" * 60)
        print(f"  🌐 Allure报告已启动: {serve_url}")
        print(f"     浏览器应该已自动打开，如未打开请手动访问上述地址")
        print(f"     按 Ctrl+C 停止服务")
        print("=" * 60 + "\n")

        return serve_url


def main():
    """主入口"""
    engine = TestEngine()
    result = engine.run_all()
    # 返回退出码
    sys.exit(0 if result.get("fail", 0) == 0 else 1)


if __name__ == "__main__":
    main()
