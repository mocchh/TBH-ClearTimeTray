# -*- coding: utf-8 -*-
"""
TBH 通关时间托盘监控
- Frida 附加 TaskBarHero，hook TMP SetText
- 按关卡记录最近 10 次通关秒数
- 写入 Excel/JSON 并自动计算平均（关闭后持久保留）
- 默认关闭系统通知；托盘菜单适配 Windows DPI
"""
from __future__ import annotations

import ctypes
import json
import os
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path


def enable_dpi_awareness() -> str:
    """在创建任何 UI 之前调用，让托盘菜单/图标按系统缩放正确显示。"""
    if os.name != "nt":
        return "n/a"
    # Per-Monitor DPI Aware v2
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return "per-monitor-v2"
    except Exception:
        pass
    try:
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:
            return "per-monitor"
    except Exception:
        pass
    try:
        if ctypes.windll.user32.SetProcessDPIAware():
            return "system"
    except Exception:
        pass
    return "failed"


# 尽早启用 DPI（import 重 UI 库之前）
_DPI_MODE = enable_dpi_awareness()


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
if not SCRIPT_PATH.exists():
    SCRIPT_PATH = APP_DIR / "clear_time_probe.js"
DATA_DIR = APP_DIR / "data"
EXCEL_PATH = DATA_DIR / "clear_times.xlsx"
JSON_PATH = DATA_DIR / "clear_times.json"
CONFIG_PATH = DATA_DIR / "config.json"
LOG_PATH = DATA_DIR / "tray.log"
MAX_PER_STAGE = 10
APP_VERSION = "1.1.2"  # 修复最近记录时间卡死（去重键含日期）

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


def default_config() -> dict:
    return {
        # 默认关闭系统通知 / 提示音
        "notify_enabled": False,
        "sound_enabled": False,
        "max_per_stage": MAX_PER_STAGE,
    }


def load_config() -> dict:
    ensure_dirs()
    cfg = default_config()
    if CONFIG_PATH.exists():
        try:
            raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                cfg.update(raw)
        except Exception as exc:
            log(f"读取配置失败，使用默认: {exc}")
    # 规范化
    cfg["notify_enabled"] = bool(cfg.get("notify_enabled", False))
    cfg["sound_enabled"] = bool(cfg.get("sound_enabled", False))
    try:
        cfg["max_per_stage"] = max(1, int(cfg.get("max_per_stage") or MAX_PER_STAGE))
    except Exception:
        cfg["max_per_stage"] = MAX_PER_STAGE
    save_config(cfg)
    return cfg


def save_config(cfg: dict) -> None:
    ensure_dirs()
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        log(f"保存配置失败: {exc}")


def get_dpi_scale() -> float:
    if os.name != "nt":
        return 1.0
    try:
        user32 = ctypes.windll.user32
        hdc = user32.GetDC(0)
        # LOGPIXELSX = 88
        dpi = int(ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) or 96)
        user32.ReleaseDC(0, hdc)
        return max(1.0, dpi / 96.0)
    except Exception:
        return 1.0


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
        self._on_hit = None
        self._last_attach_error = ""
        self._attach_fail_count = 0
        self._last_health_ok_at = 0.0
        self._detached_event = threading.Event()

    def set_on_hit(self, cb):
        self._on_hit = cb

    def is_running(self) -> bool:
        return self._worker is not None and self._worker.is_alive()

    def start(self):
        if self.is_running():
            return
        self._stop.clear()
        self._detached_event.clear()
        self._worker = threading.Thread(target=self._run_loop, name="frida-monitor", daemon=True)
        self._worker.start()
        log("监控线程已启动")

    def stop(self):
        self._stop.set()
        self._detach(reason="stop")
        self.status = "已停止"
        try:
            self.store._write_json()
            self.store._write_workbook()
        except Exception as exc:
            log(f"退出保存数据: {exc}")
        log("监控已停止")

    def _detach(self, reason: str = ""):
        with self._lock:
            script = self.script
            session = self.session
            pid = self.attached_pid
            self.script = None
            self.session = None
            self.device = None
            self.attached_pid = None
        if reason:
            log(f"断开附加 pid={pid or '-'} reason={reason}")
        try:
            if script is not None:
                script.unload()
        except Exception:
            pass
        try:
            if session is not None:
                session.detach()
        except Exception:
            pass

    def _find_game_processes(self):
        import frida

        device = frida.get_local_device()
        matches = []
        for p in device.enumerate_processes():
            name = (p.name or "").lower().replace(" ", "")
            # TaskBarHero.exe / TaskbarHero.exe
            if "taskbarhero" in name:
                matches.append(p)
        return device, matches

    def _pick_pid(self, matches):
        if not matches:
            return None
        # 多开时优先选 pid 较大的（通常是较新启动的主进程）
        matches = sorted(matches, key=lambda p: int(p.pid), reverse=True)
        return int(matches[0].pid)

    def _process_alive(self, pid) -> bool:
        if not pid:
            return False
        try:
            import frida

            for p in frida.get_local_device().enumerate_processes():
                if int(p.pid) == int(pid):
                    return True
        except Exception:
            return False
        return False

    def _wait_game_ready(self, session, pid: int, timeout_sec: float = 45.0) -> bool:
        """游戏刚启动时 GameAssembly 可能尚未加载，轮询等待。"""
        deadline = time.time() + timeout_sec
        last_log = 0.0
        while time.time() < deadline and not self._stop.is_set():
            if not self._process_alive(pid):
                return False
            try:
                # 在目标进程内查模块
                api = session.get_api() if hasattr(session, "get_api") else None
            except Exception:
                api = None
            try:
                # 通过临时脚本探测模块更可靠
                probe = session.create_script(
                    r"""
'use strict';
var names = Process.enumerateModules().map(function(m){ return m.name.toLowerCase(); });
var hit = names.some(function(n){ return n.indexOf('gameassembly') >= 0; });
send({ kind: 'module_probe', hit: hit, count: names.length });
"""
                )
                result = {"hit": False}

                def on_msg(message, data):
                    if message.get("type") == "send":
                        payload = message.get("payload") or {}
                        if payload.get("kind") == "module_probe":
                            result["hit"] = bool(payload.get("hit"))

                probe.on("message", on_msg)
                probe.load()
                time.sleep(0.15)
                try:
                    probe.unload()
                except Exception:
                    pass
                if result["hit"]:
                    return True
            except Exception as exc:
                now = time.time()
                if now - last_log > 5:
                    log(f"等待 GameAssembly: {exc}")
                    last_log = now
            self.status = f"等待游戏模块 pid={pid}"
            time.sleep(1.0)
        return False

    def _on_session_detached(self, reason, crash):
        log(f"Frida 会话断开: reason={reason} crash={crash}")
        self._detached_event.set()
        # 不在回调里重操作 session，交给主循环
        with self._lock:
            self.script = None
            self.session = None
            self.device = None
            self.attached_pid = None
        self.status = "等待游戏(会话断开)"

    def _attach(self) -> bool:
        if not SCRIPT_PATH.exists():
            self.status = "缺少探针脚本"
            log(f"找不到 {SCRIPT_PATH}")
            return False
        try:
            device, matches = self._find_game_processes()
        except Exception as exc:
            self.status = f"枚举进程失败: {exc}"
            log(f"枚举进程失败: {exc}")
            return False
        pid = self._pick_pid(matches)
        if not pid:
            self.status = "未找到游戏进程"
            return False

        # 先清掉旧会话，避免游戏重启后挂着死 session
        self._detach(reason="reattach")
        self._detached_event.clear()

        try:
            self.device = device
            self.session = device.attach(pid)
            try:
                self.session.on("detached", self._on_session_detached)
            except Exception:
                pass

            # 等 IL2CPP 模块就绪（重启后关键失败点）
            if not self._wait_game_ready(self.session, pid, timeout_sec=60.0):
                self.status = "游戏模块未就绪"
                log(f"pid={pid} 等待 GameAssembly 超时或进程退出")
                self._detach(reason="module-timeout")
                return False

            source = SCRIPT_PATH.read_text(encoding="utf-8")
            self.script = self.session.create_script(source)
            self.script.on("message", self._on_message)
            self.script.load()

            # 健康检查：rpc 或短暂等待 status
            time.sleep(0.3)
            try:
                stats = self.script.exports_sync.stats()
                log(f"探针就绪 stats={stats}")
            except Exception:
                try:
                    stats = self.script.exports.stats()
                    log(f"探针就绪 stats={stats}")
                except Exception as exc:
                    log(f"探针 stats 调用失败(可忽略): {exc}")

            self.attached_pid = pid
            self._attach_fail_count = 0
            self._last_health_ok_at = time.time()
            self.status = f"监控中 pid={pid}"
            log(f"已附加 TaskBarHero pid={pid}（含 GameAssembly 就绪检测）")
            return True
        except Exception as exc:
            self._attach_fail_count += 1
            self._last_attach_error = str(exc)
            self.status = f"附加失败: {exc}"
            log(f"附加失败 pid={pid}: {exc}\n{traceback.format_exc()}")
            self._detach(reason="attach-fail")
            return False

    def _health_check(self) -> bool:
        """确认进程仍在且 Frida 会话可用；失败则应重连。"""
        if self._detached_event.is_set():
            return False
        pid = self.attached_pid
        if not self._process_alive(pid):
            return False
        if self.session is None or self.script is None:
            return False
        # 周期性 rpc，确认脚本还活着
        now = time.time()
        if now - self._last_health_ok_at < 5.0:
            return True
        try:
            exports = getattr(self.script, "exports_sync", None) or self.script.exports
            exports.stats()
            self._last_health_ok_at = now
            return True
        except Exception as exc:
            log(f"健康检查失败，将重连: {exc}")
            return False

    def _on_message(self, message, data):
        try:
            if message.get("type") == "error":
                desc = str(message.get("description") or message)
                log(f"Frida error: {desc}")
                # 脚本崩溃后标记需重连
                if "GameAssembly" in desc or "not found" in desc.lower():
                    self._detached_event.set()
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
                self._detached_event.set()
                return
            if kind != "clear_time":
                return

            stage = str(payload.get("stage") or "").strip()
            sec = int(payload.get("clearSeconds") or 0)
            notice = str(payload.get("noticeTime") or "")
            difficulty = str(payload.get("difficulty") or "未知").strip() or "未知"
            diff_src = str(payload.get("difficultySource") or "")
            if not stage or sec <= 0:
                return

            snap = self.store.add_clear(
                stage, sec, notice_time=notice, difficulty=difficulty
            )
            if snap.get("skipped"):
                log(
                    f"去重跳过 {difficulty} {stage} {sec}秒 通知钟=[{notice}] "
                    f"最近记录仍为 {snap.get('last_at')} key={snap.get('dedupeKey')}"
                )
                return
            self.hit_count += 1
            disp = snap.get("display") or f"{difficulty} {stage}"
            self.last_hit = (
                f"{disp} {sec}秒 (均{snap['average']}s / {snap['count']}次) "
                f"@{snap.get('last_at')}"
            )
            log(
                f"通关写入 {self.last_hit} diffSrc={diff_src} "
                f"通知钟=[{notice}] raw={payload.get('raw')}"
            )
            if self._on_hit:
                try:
                    self._on_hit(snap)
                except Exception:
                    log(traceback.format_exc())
        except Exception:
            log(traceback.format_exc())

    def _run_loop(self):
        while not self._stop.is_set():
            try:
                import frida  # noqa: F401
            except ImportError:
                self.status = "未安装 frida"
                log("请 pip install frida")
                time.sleep(5)
                continue

            need_attach = self.session is None or self.script is None or self._detached_event.is_set()
            if not need_attach:
                # 进程换了 pid（重启）也要重连
                try:
                    _device, matches = self._find_game_processes()
                    pids = {int(p.pid) for p in matches}
                    if self.attached_pid and self.attached_pid not in pids:
                        log(f"检测到游戏 PID 变化: 旧={self.attached_pid} 现={sorted(pids)}")
                        need_attach = True
                    elif not self._health_check():
                        need_attach = True
                except Exception as exc:
                    log(f"巡检异常: {exc}")
                    need_attach = True

            if need_attach:
                if self.session is not None or self.script is not None:
                    self._detach(reason="reconnect")
                self._detached_event.clear()
                ok = self._attach()
                if not ok:
                    # 退避，避免疯狂刷日志
                    delay = min(8.0, 2.0 + self._attach_fail_count * 0.5)
                    self.status = self.status or "等待游戏"
                    time.sleep(delay)
                    continue
            time.sleep(1.5)


def make_icon(color=(30, 120, 200)):
    """按 DPI 生成足够大的托盘图标，避免高分屏发糊。"""
    from PIL import Image, ImageDraw

    scale = get_dpi_scale()
    # 逻辑 32，按缩放放大，且至少 64
    size = max(64, int(round(32 * scale * 2)))
    size = min(size, 256)
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    pad = max(2, size // 16)
    d.ellipse((pad, pad, size - pad - 1, size - pad - 1), fill=color + (255,))
    # 简易表格图案
    m = size // 4
    d.rectangle((m, m - 2, size - m, size - m + 2), fill=(255, 255, 255, 235))
    line_w = max(1, size // 32)
    for y_ratio in (0.40, 0.52, 0.64):
        y = int(size * y_ratio)
        d.line((m + 4, y, size - m - 4, y), fill=color + (255,), width=line_w)
    return img


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
    log(f"=== TBH 通关时间托盘启动 v{APP_VERSION} ===")
    log(f"DPI: mode={_DPI_MODE} scale={get_dpi_scale():.2f}")

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

    config = load_config()
    store = ExcelStore(
        EXCEL_PATH,
        max_per_stage=int(config.get("max_per_stage") or MAX_PER_STAGE),
        json_path=JSON_PATH,
    )
    log(f"已加载历史记录: {store.record_count()} 条 | Excel={EXCEL_PATH.name} JSON={JSON_PATH.name}")

    monitor = ClearTimeMonitor(store)
    icon_ref = {"icon": None}

    def on_hit(snap: dict):
        disp = snap.get("display") or snap.get("stage") or ""
        body = (
            f"{disp}  {snap['last_seconds']}秒\n"
            f"最近{snap['count']}次平均: {snap['average']}秒"
        )
        log(f"记录: {body.replace(chr(10), ' | ')}")
        ic = icon_ref.get("icon")
        if ic:
            try:
                ic.title = f"TBH通关监控 | {monitor.last_hit}"
            except Exception:
                pass

        # 默认关闭通知；仅配置开启时提示
        if not config.get("notify_enabled"):
            return
        try:
            if ic is not None and hasattr(ic, "notify"):
                ic.notify(body, "通关时间已记录")
        except Exception as exc:
            log(f"通知失败: {exc}")
        if config.get("sound_enabled"):
            try:
                import winsound

                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception:
                pass

    monitor.set_on_hit(on_hit)
    monitor.start()

    def action_open_excel(icon, item):
        # 打开前再同步一次 Excel；若 Excel 正打开会被占用导致看起来“时间不更新”
        try:
            store._write_workbook()
            log(f"已同步 Excel -> {EXCEL_PATH}")
        except Exception as exc:
            log(
                f"同步 Excel 失败(请先关闭 Excel 再打开菜单): {exc} | "
                f"JSON 仍以 {JSON_PATH} 为准"
            )
        open_path(EXCEL_PATH)

    def action_open_folder(icon, item):
        open_path(DATA_DIR)

    def action_reconnect(icon, item):
        log("手动重连…")
        monitor._detached_event.set()
        monitor._detach(reason="manual")
        monitor.status = "重连中…"

    def action_toggle_notify(icon, item):
        config["notify_enabled"] = not bool(config.get("notify_enabled"))
        save_config(config)
        log(f"通知已{'开启' if config['notify_enabled'] else '关闭'}")

    def action_toggle_sound(icon, item):
        config["sound_enabled"] = not bool(config.get("sound_enabled"))
        save_config(config)
        log(f"提示音已{'开启' if config['sound_enabled'] else '关闭'}")

    def action_quit(icon, item):
        monitor.stop()
        icon.stop()

    def notify_text(item):
        on = bool(config.get("notify_enabled"))
        return f"{'✓ ' if on else ''}系统通知（默认关）"

    def sound_text(item):
        on = bool(config.get("sound_enabled"))
        return f"{'✓ ' if on else ''}提示音（默认关）"

    menu = pystray.Menu(
        item(lambda icon: f"状态: {monitor.status}", None, enabled=False),
        item(lambda icon: f"最近: {monitor.last_hit or '暂无'}", None, enabled=False),
        item(lambda icon: f"累计命中: {monitor.hit_count}", None, enabled=False),
        item(lambda icon: f"已存记录: {store.record_count()} 条", None, enabled=False),
        pystray.Menu.SEPARATOR,
        item(notify_text, action_toggle_notify),
        item(sound_text, action_toggle_sound),
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
        f"TBH 通关时间监控 v{APP_VERSION}",
        menu,
    )
    icon_ref["icon"] = icon
    log(f"Excel: {EXCEL_PATH}")
    log(f"JSON:  {JSON_PATH}")
    log(f"通知默认: {'开' if config.get('notify_enabled') else '关'}")
    icon.run()
    monitor.stop()
    log("托盘退出")


if __name__ == "__main__":
    main()
