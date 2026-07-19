# -*- coding: utf-8 -*-
"""
TBH 通关时间托盘监控
- Frida 附加 TaskBarHero，hook TMP SetText
- 按关卡记录最近 10 次通关秒数
- 写入 Excel 并自动计算平均
"""
from __future__ import annotations

import json
import os
import sys
import threading
import time
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path

def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_dir() -> Path:
    """可写目录：exe 旁 / 源码旁。"""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_dir() -> Path:
    """只读资源：打包后在 _MEIPASS。"""
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
RESOURCE_DIR = resource_dir()
SCRIPT_PATH = RESOURCE_DIR / "clear_time_probe.js"
# 若打包资源里没有，再尝试 exe 同目录
if not SCRIPT_PATH.exists():
    SCRIPT_PATH = APP_DIR / "clear_time_probe.js"
DATA_DIR = APP_DIR / "data"
EXCEL_PATH = DATA_DIR / "clear_times.xlsx"
LOG_PATH = DATA_DIR / "tray.log"
MAX_PER_STAGE = 10
PROCESS_HINT = "TaskBarHero"
APP_VERSION = "1.0.0"

# 保证子目录可 import（源码运行）
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))
if str(RESOURCE_DIR) not in sys.path:
    sys.path.insert(0, str(RESOURCE_DIR))

from excel_store import ExcelStore  # noqa: E402


def ensure_dirs():
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def log(msg: str) -> None:
    ensure_dirs()
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        print(line, flush=True)
    except Exception:
        pass
    try:
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


class ClearTimeMonitor:
    def __init__(self, store: ExcelStore):
        self.store = store
        self._lock = threading.RLock()
        self.session = None
        self.script = None
        self.device = None
        self.attached_pid = None
        self.status = "未连接"
        self.last_hit = ""
        self.hit_count = 0
        self._stop = threading.Event()
        self._worker = None
        self._on_hit = None  # callback(snapshot)

    def set_on_hit(self, cb):
        self._on_hit = cb

    def is_running(self) -> bool:
        return self._worker is not None and self._worker.is_alive()

    def start(self):
        if self.is_running():
            return
        self._stop.clear()
        self._worker = threading.Thread(target=self._run_loop, name="frida-monitor", daemon=True)
        self._worker.start()
        log("监控线程已启动")

    def stop(self):
        self._stop.set()
        self._detach()
        self.status = "已停止"
        log("监控已停止")

    def _detach(self):
        with self._lock:
            try:
                if self.script:
                    self.script.unload()
            except Exception:
                pass
            try:
                if self.session:
                    self.session.detach()
            except Exception:
                pass
            self.script = None
            self.session = None
            self.device = None
            self.attached_pid = None

    def _find_pid(self):
        import frida

        device = frida.get_local_device()
        matches = []
        for p in device.enumerate_processes():
            name = (p.name or "").lower().replace(" ", "")
            if "taskbarhero" in name:
                matches.append(p)
        if not matches:
            return None, None
        target = matches[0]
        return device, int(target.pid)

    def _attach(self) -> bool:
        import frida

        if not SCRIPT_PATH.exists():
            self.status = "缺少探针脚本"
            log(f"找不到 {SCRIPT_PATH}")
            return False
        device, pid = self._find_pid()
        if not pid:
            self.status = "未找到游戏进程"
            return False
        try:
            self._detach()
            self.device = device
            self.session = device.attach(pid)
            source = SCRIPT_PATH.read_text(encoding="utf-8")
            self.script = self.session.create_script(source)
            self.script.on("message", self._on_message)
            self.script.load()
            self.attached_pid = pid
            self.status = f"监控中 pid={pid}"
            log(f"已附加 TaskBarHero pid={pid}")
            return True
        except Exception as exc:
            self.status = f"附加失败: {exc}"
            log(f"附加失败: {exc}\n{traceback.format_exc()}")
            self._detach()
            return False

    def _on_message(self, message, data):
        try:
            if message.get("type") == "error":
                log(f"Frida error: {message.get('description') or message}")
                return
            if message.get("type") != "send":
                return
            payload = message.get("payload") or {}
            kind = payload.get("kind")
            if kind == "status":
                log(f"探针: {payload.get('text') or payload}")
                return
            if kind == "fatal":
                log(f"探针致命错误: {payload.get('error')}")
                self.status = "探针错误"
                return
            if kind != "clear_time":
                return

            stage = str(payload.get("stage") or "").strip()
            sec = int(payload.get("clearSeconds") or 0)
            notice = str(payload.get("noticeTime") or "")
            if not stage or sec <= 0:
                return

            snap = self.store.add_clear(stage, sec, notice_time=notice)
            self.hit_count += 1
            self.last_hit = f"{stage} {sec}秒 (均{snap['average']}s / {snap['count']}次)"
            log(f"通关 {self.last_hit} raw={payload.get('raw')}")
            if self._on_hit:
                try:
                    self._on_hit(snap)
                except Exception:
                    log(traceback.format_exc())
        except Exception:
            log(traceback.format_exc())

    def _run_loop(self):
        # 自动重连
        while not self._stop.is_set():
            try:
                import frida  # noqa: F401
            except ImportError:
                self.status = "未安装 frida"
                log("请 pip install frida")
                time.sleep(5)
                continue

            if self.session is None:
                ok = self._attach()
                if not ok:
                    time.sleep(3)
                    continue
            else:
                # 进程是否还活着
                try:
                    import frida

                    alive = any(
                        p.pid == self.attached_pid
                        for p in frida.get_local_device().enumerate_processes()
                    )
                    if not alive:
                        log("游戏进程已退出，等待重连")
                        self.status = "等待游戏"
                        self._detach()
                        time.sleep(2)
                        continue
                except Exception:
                    self._detach()
                    time.sleep(2)
                    continue
            time.sleep(1.5)


def make_icon(color=(30, 120, 200)):
    from PIL import Image, ImageDraw

    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((4, 4, size - 5, size - 5), fill=color + (255,))
    d.rectangle((20, 18, 44, 46), fill=(255, 255, 255, 230))
    d.line((24, 26, 40, 26), fill=color + (255,), width=2)
    d.line((24, 32, 40, 32), fill=color + (255,), width=2)
    d.line((24, 38, 36, 38), fill=color + (255,), width=2)
    return img


def balloon(title: str, body: str):
    """尽量弹出 Windows 通知。"""
    try:
        from pystray import Icon  # noqa: F401

        # 部分环境用 winsound + 日志即可；优先 toast
        pass
    except Exception:
        pass
    try:
        # 简单 toast（Win10+）
        from win10toast import ToastNotifier  # type: ignore

        ToastNotifier().show_toast(title, body, duration=4, threaded=True)
        return
    except Exception:
        pass
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, body, title, 0x40 | 0x1000)
    except Exception:
        pass


def open_path(path: Path):
    path = Path(path)
    try:
        if path.is_file():
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            os.startfile(str(path if path.is_dir() else path.parent))  # type: ignore[attr-defined]
    except Exception as exc:
        log(f"打开失败: {exc}")


def main():
    ensure_dirs()
    log("=== TBH 通关时间托盘启动 ===")

    try:
        import frida  # noqa: F401
    except ImportError:
        log("缺少 frida: pip install frida")
    try:
        import pystray
        from pystray import MenuItem as item
    except ImportError:
        log("缺少 pystray: pip install pystray")
        print("请安装: pip install frida pystray pillow openpyxl")
        sys.exit(1)

    store = ExcelStore(EXCEL_PATH, max_per_stage=MAX_PER_STAGE)
    monitor = ClearTimeMonitor(store)

    icon_ref = {"icon": None}

    def on_hit(snap: dict):
        body = (
            f"{snap['stage']}  {snap['last_seconds']}秒\n"
            f"最近{snap['count']}次平均: {snap['average']}秒"
        )
        log(f"通知: {body.replace(chr(10), ' | ')}")
        # 托盘 title 更新
        ic = icon_ref.get("icon")
        if ic:
            ic.title = f"TBH通关监控 | {monitor.last_hit}"

        # 尝试 toast
        try:
            if hasattr(ic, "notify"):
                ic.notify(body, "通关时间已记录")
                return
        except Exception:
            pass
        # 不阻塞的轻提示：只写日志；可选 MessageBeep
        try:
            import winsound

            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass

    monitor.set_on_hit(on_hit)
    monitor.start()

    def action_open_excel(icon, item):
        open_path(EXCEL_PATH)

    def action_open_folder(icon, item):
        open_path(DATA_DIR)

    def action_reconnect(icon, item):
        log("手动重连…")
        monitor._detach()
        monitor.status = "重连中…"

    def action_quit(icon, item):
        monitor.stop()
        icon.stop()

    menu = pystray.Menu(
        item(lambda icon: f"状态: {monitor.status}", None, enabled=False),
        item(lambda icon: f"最近: {monitor.last_hit or '暂无'}", None, enabled=False),
        item(lambda icon: f"累计命中: {monitor.hit_count}", None, enabled=False),
        pystray.Menu.SEPARATOR,
        item("打开 Excel", action_open_excel),
        item("打开数据目录", action_open_folder),
        item("重新连接游戏", action_reconnect),
        pystray.Menu.SEPARATOR,
        item("退出", action_quit),
    )

    icon = pystray.Icon(
        "TBHClearTime",
        make_icon(),
        "TBH 通关时间监控",
        menu,
    )
    icon_ref["icon"] = icon
    log(f"Excel: {EXCEL_PATH}")
    icon.run()
    monitor.stop()
    log("托盘退出")


if __name__ == "__main__":
    main()
