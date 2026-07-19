# -*- coding: utf-8 -*-
"""按关卡保存最近 N 次通关时间到 Excel，并计算平均。"""
from __future__ import annotations

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


class ExcelStore:
    def __init__(self, path: Path, max_per_stage: int = MAX_PER_STAGE):
        self.path = Path(path)
        self.max_per_stage = max(1, int(max_per_stage))
        self._lock = threading.RLock()
        # stage -> deque of (seconds, recorded_at, notice_time)
        self._data: Dict[str, Deque[Tuple[int, datetime, str]]] = defaultdict(
            lambda: deque(maxlen=self.max_per_stage)
        )
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._load_or_init()

    def _load_or_init(self) -> None:
        if not self.path.exists():
            self._write_workbook()
            return
        try:
            wb = load_workbook(self.path, data_only=True)
            if HISTORY_SHEET in wb.sheetnames:
                ws = wb[HISTORY_SHEET]
                # 从明细重建：按关卡取时间序最后 N 条
                rows = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or not row[0]:
                        continue
                    stage = str(row[0]).strip()
                    try:
                        sec = int(row[1])
                    except Exception:
                        continue
                    rec_at = row[2]
                    notice = str(row[3] or "")
                    if isinstance(rec_at, datetime):
                        dt = rec_at
                    else:
                        try:
                            dt = datetime.fromisoformat(str(rec_at))
                        except Exception:
                            dt = datetime.now()
                    rows.append((stage, sec, dt, notice))
                rows.sort(key=lambda x: x[2])
                for stage, sec, dt, notice in rows:
                    self._data[stage].append((sec, dt, notice))
            wb.close()
        except Exception:
            # 损坏则重建空表，不丢程序
            self._data.clear()
        self._write_workbook()

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
            self._write_workbook()
            return self.stage_snapshot(stage)

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
            "last_at": items[-1][1].isoformat(sep=" ", timespec="seconds") if items else "",
            "notice_time": items[-1][2] if items else "",
        }

    def all_snapshots(self) -> List[dict]:
        with self._lock:
            stages = sorted(self._data.keys(), key=self._stage_sort_key)
        return [self.stage_snapshot(s) for s in stages]

    @staticmethod
    def _stage_sort_key(stage: str):
        parts = str(stage).replace("－", "-").split("-")
        try:
            return (int(parts[0]), int(parts[1]) if len(parts) > 1 else 0)
        except Exception:
            return (9999, 9999)

    def _write_workbook(self) -> None:
        wb = Workbook()
        # ---- 汇总 ----
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

        # ---- 明细（完整历史，保留最近 max*关卡数 的扩展：写全部内存中的最近N）----
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

        # 原子写入
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        wb.save(tmp)
        wb.close()
        tmp.replace(self.path)

    def excel_path(self) -> Path:
        return self.path
