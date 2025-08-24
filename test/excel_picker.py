# excel_picker.py
import os
import itertools
import random
import openpyxl

class ExcelPicker:
    def __init__(self):
        self._excel_cache = {}
        self._excel_cycle = None

    def execute(self, params: dict) -> str:
        excel_path = params.get("excel_path", "").strip()
        if not excel_path or not os.path.isfile(excel_path):
            raise FileNotFoundError("未指定或找不到 Excel 文件")

        sheet_id = params.get("sheet", "0")
        col_index = int(params.get("col", 0))
        mode = params.get("mode", "顺序")

        cache_key = excel_path
        if cache_key not in self._excel_cache:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            try:
                ws = wb[int(sheet_id)] if str(sheet_id).isdigit() else wb[sheet_id]
            except Exception:
                ws = wb.worksheets[0]
            rows = list(ws.iter_rows(values_only=True))
            self._excel_cache[cache_key] = (wb, ws, rows)
        _, _, rows = self._excel_cache[cache_key]

        if not rows:
            raise ValueError("Excel 表无数据")

        cells = [row[col_index] for row in rows
                 if len(row) > col_index and row[col_index] is not None]
        if not cells:
            raise ValueError("指定列为空")

        if mode == "顺序":
            if self._excel_cycle is None:
                self._excel_cycle = itertools.cycle(cells)
            return next(self._excel_cycle)
        else:
            return random.choice(cells)