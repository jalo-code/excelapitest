# -*- coding: utf-8 -*-
"""
入口脚本 - 快速运行测试
使用方式:
    python run.py                          # 运行默认配置（启用Allure+自动启动HTTP服务）
    python run.py --excel xxx.xlsx         # 指定Excel文件
    python run.py --sheet Sheet2           # 指定Sheet名称
    python run.py --no-allure              # 关闭Allure（仅输出HTML报告，速度更快）
    python run.py --no-serve               # 生成Allure报告但不自动启动HTTP服务
    python run.py --allure-only-report     # 只尝试调用CLI生成Allure报告（已存在results时）
"""
import os
import sys
import time
import argparse

# 确保项目根目录在path中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from test_engine import TestEngine
from config.settings import ALLURE_RESULTS_DIR, ALLURE_REPORT_DIR


def parse_args():
    parser = argparse.ArgumentParser(description="Excel驱动的接口自动化测试框架")
    parser.add_argument("--excel", "-e", type=str, default=None,
                        help="Excel测试用例文件路径")
    parser.add_argument("--sheet", "-s", type=str, default=None,
                        help="Excel中的Sheet名称")
    parser.add_argument("--no-allure", action="store_true",
                        help="关闭Allure结果记录与报告生成（仅输出HTML报告）")
    parser.add_argument("--no-serve", action="store_true",
                        help="生成Allure报告但不自动启动本地HTTP服务（需手动用allure open查看）")
    parser.add_argument("--allure-only-report", action="store_true",
                        help="不执行用例，仅使用已有allure-results生成Allure HTML报告")
    return parser.parse_args()


def main():
    args = parse_args()

    # --allure-only-report 模式：尝试把已有的 allure-results 转成HTML报告
    if args.allure_only_report:
        print("=" * 60)
        print("  Allure 报告生成模式（不执行用例）")
        print("=" * 60)
        engine = TestEngine(excel_path=args.excel, sheet_name=args.sheet,
                            allure_enable=False, allure_serve=not args.no_serve)
        # 直接调用引擎中的 _generate_allure_report 方法（需要适配器目录已就绪）
        try:
            status = engine._generate_allure_report()
        except Exception as e:
            print(f"生成失败: {e}")
            return 2
        if status.get("ok"):
            print("✅ Allure报告生成成功!")
            print(f"   目录: {status.get('report_dir')}")
            serve_url = status.get("serve_url")
            if serve_url:
                print(f"   HTTP服务: {serve_url}")
                # 保持进程不退出，让HTTP服务持续运行
                print("   按 Ctrl+C 停止服务...")
                try:
                    while True:
                        time.sleep(1)
                except KeyboardInterrupt:
                    print("\n服务已停止")
            return 0
        else:
            print("❌ Allure报告未生成。原因:")
            print(status.get("message", ""))
            return 2

    print("=" * 60)
    print("  Excel驱动接口自动化测试框架")
    allure_flag = "已启用" if not args.no_allure else "已关闭(--no-allure)"
    serve_flag = "已启用" if not args.no_serve else "已关闭(--no-serve)"
    print(f"  Allure: {allure_flag}")
    print(f"  HTTP服务: {serve_flag}")
    print("=" * 60)

    engine = TestEngine(excel_path=args.excel, sheet_name=args.sheet,
                        allure_enable=not args.no_allure,
                        allure_serve=not args.no_serve)
    result = engine.run_all()

    # Allure情况提示
    a = result.get("allure", {}) or {}
    if not args.no_allure:
        print("\n" + "-" * 60)
        print(f"[Allure] 结果目录: {ALLURE_RESULTS_DIR}")
        if a.get("ok"):
            serve_url = a.get("serve_url")
            print(f"[Allure] ✅ 报告已生成: {ALLURE_REPORT_DIR}")
            if serve_url:
                print(f"[Allure] 🌐 HTTP服务已启动: {serve_url}")
                print("         浏览器应已自动打开，如未打开请手动访问上述地址")
            else:
                print("         ⚠️ HTTP服务未启动，请勿直接双击index.html（会500）")
                print(f"         正确查看方式: allure open {ALLURE_REPORT_DIR}")
        else:
            print("[Allure] ⚠️  报告未生成。诊断:")
            msg = (a.get("message") or "").strip()
            if not msg:
                reasons = []
                try:
                    import allure_commons  # noqa: F401
                except Exception as e:
                    reasons.append(
                        f"1) Python依赖缺失: 找不到 allure-pytest 包 (ImportError: {e})"
                        "\n        → 解决: 先运行  pip install allure-pytest  (或 pip install -r requirements.txt)"
                    )
                if not os.path.isdir(ALLURE_RESULTS_DIR):
                    reasons.append(
                        f"2) allure-results 目录不存在: {ALLURE_RESULTS_DIR}"
                        "\n        → 通常因为上面的依赖问题，导致适配器没有初始化成功、没写出结果文件"
                    )
                else:
                    try:
                        files = os.listdir(ALLURE_RESULTS_DIR)
                    except Exception as e:
                        files = []
                    jsons = [f for f in files if f.endswith("-result.json")]
                    if not jsons:
                        reasons.append(
                            f"2) allure-results 目录存在但没有结果文件 (共 {len(files)} 个文件，无 *-result.json)"
                            "\n        → 说明用例执行完成但 Allure 适配器写入失败，请检查 logs/test.log 的警告"
                        )
                    else:
                        reasons.append(
                            f"2) allure-results 目录有 {len(jsons)} 个结果文件但仍未生成报告"
                            "\n        → 通常是缺少 Allure CLI 命令行工具（见下方安装指引）"
                        )
                reasons.append(
                    "3) 若结果文件已存在但缺少 Allure CLI，任选一种安装:"
                    "\n        · npm i -g allure-commandline"
                    "\n        · scoop install allure"
                    "\n        · choco install allure"
                    "\n      装好后不用重跑用例，直接:  python run.py --allure-only-report"
                )
                msg = "\n".join(reasons)
            for line in msg.splitlines()[:20]:
                print("         " + line)
            if len(msg.splitlines()) > 20:
                print("         ... 完整信息请见 logs/test.log")
        print("-" * 60)

    # 最终状态
    print("\n")
    if result["fail"] == 0:
        print("✓ 所有执行的用例均通过！")
    else:
        print(f"✗ 存在 {result['fail']} 条失败用例，请查看详细报告")

    # 如果Allure HTTP服务已启动，保持进程不退出
    if a.get("serve_url"):
        print(f"\n🌐 Allure报告服务运行中: {a['serve_url']}")
        print("   按 Ctrl+C 停止服务并退出...")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n服务已停止，再见！")
        return 0 if result["fail"] == 0 else 1

    return 0 if result["fail"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
