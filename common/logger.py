# -*- coding: utf-8 -*-
"""
日志工具模块
"""
import os
import sys
import time
import errno
import logging
import tempfile
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler, BaseRotatingHandler
from config.settings import LOG_DIR, LOG_FILE


class SafeTimedRotatingFileHandler(TimedRotatingFileHandler):
    """
    Windows下安全的按日滚动日志处理器
    覆盖 doRollover，当 rename 遇到 WinError 32 文件被占用时，
    自动回退到带时间戳的新日志文件继续写入，不抛出异常阻断业务。
    """

    def doRollover(self):
        try:
            super().doRollover()
        except PermissionError as e:
            # WinError 32: 文件被占用，无法rename
            if getattr(e, "winerror", None) == 32 or e.errno in (errno.EACCES, errno.EPERM):
                self._fallback_rollover(e)
            else:
                raise
        except OSError as e:
            if getattr(e, "winerror", None) == 32 or e.errno in (errno.EACCES, errno.EPERM, errno.EAGAIN):
                self._fallback_rollover(e)
            else:
                raise

    def _fallback_rollover(self, original_error):
        """滚动失败时回退策略：关闭旧句柄，切换到带时间戳的新文件"""
        import sys
        sys.stderr.write(
            f"[WARN] 日志滚动失败（文件被占用）: {original_error}. "
            f"已切换到新日志文件继续写入。\n"
        )
        # 关闭当前流
        if self.stream:
            try:
                self.stream.close()
            except Exception:
                pass
            self.stream = None

        # 生成带毫秒级时间戳的新日志文件，避免冲突
        base, ext = os.path.splitext(self.baseFilename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        new_path = f"{base}_{timestamp}{ext or '.log'}"
        self.baseFilename = os.path.abspath(new_path)

        # 打开新文件
        try:
            self.stream = self._open()
        except Exception as e2:
            sys.stderr.write(f"[WARN] 新建日志文件也失败: {e2}，改为仅控制台输出。\n")


def _safe_make_log_dir():
    """安全创建日志目录，失败时回退到用户临时目录"""
    try:
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR, exist_ok=True)
        # 尝试写一下确认目录可写
        probe = os.path.join(LOG_DIR, ".write_probe")
        with open(probe, "w") as f:
            f.write("ok")
        os.remove(probe)
        return LOG_DIR, LOG_FILE
    except (PermissionError, OSError) as e:
        # 原路径不可写，回退到临时目录
        fallback_dir = os.path.join(tempfile.gettempdir(), "excel_api_test_logs")
        os.makedirs(fallback_dir, exist_ok=True)
        fallback_file = os.path.join(fallback_dir, "test.log")
        sys.stderr.write(
            f"[WARN] 日志目录 {LOG_DIR} 不可写: {e}. "
            f"已回退到临时目录: {fallback_file}\n"
        )
        return fallback_dir, fallback_file


def get_logger(name="api_test"):
    """获取日志记录器（带安全回退，日志失败不阻断业务）"""
    log_dir, log_file = _safe_make_log_dir()

    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    # 禁止向上传播到根logger，避免重复打印
    logger.propagate = False

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # 1. 控制台输出（始终保留）
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 2. 文件输出（失败时降级：改用普通FileHandler -> 彻底放弃文件日志）
    file_handler = None
    try:
        file_handler = SafeTimedRotatingFileHandler(
            log_file,
            when="D",
            interval=1,
            backupCount=30,
            encoding="utf-8",
            delay=False
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except (PermissionError, OSError) as e1:
        sys.stderr.write(f"[WARN] 创建按日滚动日志失败: {e1}，尝试普通文件日志...\n")
        try:
            # 尝试普通FileHandler（用独立时间戳文件名，避免冲突）
            base, ext = os.path.splitext(log_file)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            alt_file = f"{base}_{timestamp}{ext or '.log'}"
            plain_handler = logging.FileHandler(alt_file, encoding="utf-8")
            plain_handler.setLevel(logging.DEBUG)
            plain_handler.setFormatter(formatter)
            logger.addHandler(plain_handler)
        except Exception as e2:
            sys.stderr.write(f"[WARN] 创建普通文件日志也失败: {e2}，日志将仅输出到控制台。\n")

    return logger
