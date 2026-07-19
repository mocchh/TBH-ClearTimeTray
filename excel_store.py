# -*- coding: utf-8 -*-
"""按关卡持久化通关时间（JSON 主存 + Excel 同步），关闭后不清空。"""
from __future__ import annotations

import json
import shutil
import threading
from collections import defaultdict, deque
from datetime import datetime
from pathlib import Path
from typing import Deque, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MAX_PER_STAGE = 10
HISTORY_SHEET = "明细"
SUMMARY_SHEET = "汇总"


def _parse_dt(value) -> datetime:
    if isinstance(value, datetime):
        return value
    text = str(value or "").strip()
    if not text:
        return datetime.now()
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S.%f",
        "%Y-%m-%dT%H:%M:%S.%f",
    ):
        try:
            return datetime.strptime(text[:26], fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(text)
    except Exception:
        return datetime.now()


class ExcelStore:
    """
    持久化策略：
    - JSON（clear_times.json）为主数据源，重启必加载
    - Excel（clear_times.xlsx）为可读副本，每次变更同步
    - 每关仅保留最近 max_per_stage 次（用于平均），滚动丢弃最旧
    - 关闭软件不会清空；加载失败时绝不覆盖已有文件
    """

    def __init__(
        self,
        excel_path: Path,
        max_per_stage: int = MAX_PER_STAGE,
        json_path: Optional[Path] = None,
    ):
        self.path = Path(excel_path)
        self.json_path = Path(json_path) if json_path else self.path.with_suffix(".json")
        self.max_per_stage = max(1, int(max_per_stage))
        self._lock = threading.RLock()
        # stage -> deque of (seconds, recorded_at, notice_time)
        self._data: Dict[str, Deque[Tuple[int, datetime, str]]] = defaultdict(
            lambda: deque(maxlen=self.max_per_stage)
        )
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._load_persistent()
        # 启动时同步一份 Excel（不丢数据）
        try:
            self._write_workbook()
            self._write_json()
        except Exception:
            pass

    def _new_deque(self) -> Deque[Tuple[int, datetime, str]]:
        return deque(maxlen=self.max_per_stage)

    def _load_persistent(self) -> None:
        """优先 JSON，其次 Excel 明细；失败则保持内存空且不破坏磁盘文件。"""
        if self._load_json():
            return
        if self._load_excel():
            # 从 Excel 恢复后立刻落 JSON，避免下次再丢
            try:
                self._write_json()
            except Exception:
                pass
            return
        # 全新安装：生成空表
        if not self.path.exists() and not self.json_path.exists():
            try:
                self._write_workbook()
                self._write_json()
            except Exception:
                pass

    def _load_json(self) -> bool:
        if not self.json_path.exists():
            return False
        try:
            raw = json.loads(self.json_path.read_text(encoding="utf-8"))
            stages = raw.get("stages") if isinstance(raw, dict) else None
            if not isinstance(stages, dict):
                return False
            loaded: Dict[str, Deque[Tuple[int, datetime, str]]] = {}
            for stage, items in stages.items():
                stage = str(stage).strip()
                if not stage or not isinstance(items, list):
                    continue
                dq = self._new_deque()
                for it in items:
                    if not isinstance(it, dict):
                        continue
                    try:
                        sec = int(it.get("seconds"))
                    except Exception:
                        continue
                    dt = _parse_dt(it.get("recorded_at"))
                    notice = str(it.get("notice_time") or "")
                    dq.append((sec, dt, notice))
                if dq:
                    loaded[stage] = dq
            self._data = defaultdict(self._new_deque, loaded)
            return True
        except Exception:
            return False

    def _load_excel(self) -> bool:
        if not self.path.exists():
            return False
        try:
            # 不用 data_only，避免公式/缓存导致空值
            wb = load_workbook(self.path, data_only=False)
            sheet_name = HISTORY_SHEET if HISTORY_SHEET in wb.sheetnames else None
            if sheet_name is None and SUMMARY_SHEET in wb.sheetnames:
                # 仅有汇总时也能部分恢复
                return self._load_from_summary(wb)
            if sheet_name is None:
                wb.close()
                return False

            ws = wb[sheet_name]
            rows = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                stage = str(row[0]).strip()
                try:
                    sec = int(row[1])
                except Exception:
                    continue
                dt = _parse_dt(row[2])
                notice = str(row[3] or "") if len(row) > 3 else ""
                rows.append((stage, sec, dt, notice))
            wb.close()
            rows.sort(key=lambda x: x[2])
            loaded: Dict[str, Deque[Tuple[int, datetime, str]]] = {}
            for stage, sec, dt, notice in rows:
                if stage not in loaded:
                    loaded[stage] = self._new_deque()
                loaded[stage].append((sec, dt, notice))
            if not loaded:
                return False
            self._data = defaultdict(self._new_deque, loaded)
            return True
        except Exception:
            return False

    def _load_from_summary(self, wb) -> bool:
        try:
            ws = wb[SUMMARY_SHEET]
            loaded: Dict[str, Deque[Tuple[int, datetime, str]]] = {}
            headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            # 第1次(秒) ... 第N次(秒)
            sec_cols = []
            for i, h in enumerate(headers):
                if h and str(h).startswith("第") and "次" in str(h):
                    sec_cols.append(i)
            stage_i = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[stage_i]:
                    continue
                stage = str(row[stage_i]).strip()
                dq = self._new_deque()
                for ci in sec_cols:
                    if ci >= len(row):
                        continue
                    v = row[ci]
                    if v is None or v == "":
                        continue
                    try:
                        dq.append((int(v), datetime.now(), ""))
                    except Exception:
                        continue
                if dq:
                    loaded[stage] = dq
            wb.close()
            if not loaded:
                return False
            self._data = defaultdict(self._new_deque, loaded)
            return True
        except Exception:
            try:
                wb.close()
            except Exception:
                pass
            return False

    def _write_json(self) -> None:
        payload = {
            "version": 1,
            "max_per_stage": self.max_per_stage,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "stages": {},
        }
        for stage, items in self._data.items():
            payload["stages"][stage] = [
                {
                    "seconds": sec,
                    "recorded_at": dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "notice_time": notice,
                }
                for sec, dt, notice in items
            ]
        tmp = self.json_path.with_suffix(self.json_path.suffix + ".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(self.json_path)

    def add_clear(
        self,
        stage: str,
        clear_seconds: int,
        notice_time: str = "",
        recorded_at: Optional[datetime] = None,
    ) -> dict:
        """写入一次通关。返回该关卡当前状态（含平均）。"""
        stage = str(stage or "").strip()
        if not stage:
            raise ValueError("empty stage")
        sec = int(clear_seconds)
        if sec < 0 or sec > 86400:
            raise ValueError(f"invalid seconds: {sec}")
        dt = recorded_at or datetime.now()
        notice = str(notice_time or "").strip()

        with self._lock:
            self._data[stage].append((sec, dt, notice))
            self._persist_all()
            return self.stage_snapshot(stage)

    def _persist_all(self) -> None:
        """先 JSON 后 Excel；Excel 失败不回滚 JSON。"""
        self._write_json()
        try:
            self._write_workbook()
        except Exception:
            # 文件被占用等：保留 JSON，下次启动/打开再同步
            try:
                bak = self.path.with_suffix(".xlsx.bak")
                if self.path.exists():
                    shutil.copy2(self.path, bak)
            except Exception:
                pass
            raise

    def stage_snapshot(self, stage: str) -> dict:
        with self._lock:
            items = list(self._data.get(stage, []))
        secs = [x[0] for x in items]
        avg = round(sum(secs) / len(secs), 2) if secs else 0.0
        return {
            "stage": stage,
            "count": len(secs),
            "seconds": secs,
            "average": avg,
            "last_seconds": secs[-1] if secs else None,
            "last_at": items[-1][1].strftime("%Y-%m-%d %H:%M:%S") if items else "",
            "notice_time": items[-1][2] if items else "",
        }

    def all_snapshots(self) -> List[dict]:
        with self._lock:
            stages = sorted(self._data.keys(), key=self._stage_sort_key)
        return [self.stage_snapshot(s) for s in stages]

    def record_count(self) -> int:
        with self._lock:
            return sum(len(v) for v in self._data.values())

    @staticmethod
    def _stage_sort_key(stage: str):
        parts = str(stage).replace("－", "-").split("-")
        try:
            return (int(parts[0]), int(parts[1]) if len(parts) > 1 else 0)
        except Exception:
            return (9999, 9999)

    def _write_workbook(self) -> None:
        wb = Workbook()
        ws_sum = wb.active
        ws_sum.title = SUMMARY_SHEET
        headers = (
            ["关卡"]
            + [f"第{i}次(秒)" for i in range(1, self.max_per_stage + 1)]
            + ["样本数", "平均通关(秒)", "最近记录时间", "最近通知时钟"]
        )
        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(color="FFFFFF", bold=True)
        for col, h in enumerate(headers, 1):
            cell = ws_sum.cell(1, col, h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        stages = sorted(self._data.keys(), key=self._stage_sort_key)
        avg_fill = PatternFill("solid", fgColor="FFF2CC")
        for r, stage in enumerate(stages, 2):
            items = list(self._data[stage])
            secs = [x[0] for x in items]
            avg = round(sum(secs) / len(secs), 2) if secs else ""
            ws_sum.cell(r, 1, stage)
            for i in range(self.max_per_stage):
                val = secs[i] if i < len(secs) else ""
                ws_sum.cell(r, 2 + i, val)
            ws_sum.cell(r, 2 + self.max_per_stage, len(secs))
            avg_cell = ws_sum.cell(r, 3 + self.max_per_stage, avg)
            avg_cell.fill = avg_fill
            avg_cell.font = Font(bold=True)
            last_at = items[-1][1].strftime("%Y-%m-%d %H:%M:%S") if items else ""
            last_notice = items[-1][2] if items else ""
            ws_sum.cell(r, 4 + self.max_per_stage, last_at)
            ws_sum.cell(r, 5 + self.max_per_stage, last_notice)

        for col in range(1, len(headers) + 1):
            ws_sum.column_dimensions[get_column_letter(col)].width = 14
        ws_sum.column_dimensions["A"].width = 10
        ws_sum.column_dimensions[get_column_letter(4 + self.max_per_stage)].width = 20

        ws_hist = wb.create_sheet(HISTORY_SHEET)
        hist_headers = ["关卡", "通关秒数", "记录时间", "通知时钟", "序号(该关内)"]
        for col, h in enumerate(hist_headers, 1):
            cell = ws_hist.cell(1, col, h)
            cell.fill = header_fill
            cell.font = header_font
        row_i = 2
        for stage in stages:
            items = list(self._data[stage])
            for idx, (sec, dt, notice) in enumerate(items, 1):
                ws_hist.cell(row_i, 1, stage)
                ws_hist.cell(row_i, 2, sec)
                ws_hist.cell(row_i, 3, dt.strftime("%Y-%m-%d %H:%M:%S"))
                ws_hist.cell(row_i, 4, notice)
                ws_hist.cell(row_i, 5, idx)
                row_i += 1
        for col in range(1, 6):
            ws_hist.column_dimensions[get_column_letter(col)].width = 18

        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        wb.save(tmp)
        wb.close()
        # 备份旧文件再替换，防止写入中断导致空表
        if self.path.exists():
            try:
                shutil.copy2(self.path, self.path.with_suffix(".xlsx.bak"))
            except Exception:
                pass
        tmp.replace(self.path)

    def excel_path(self) -> Path:
        return self.path

    def json_store_path(self) -> Path:
        return self.json_path
