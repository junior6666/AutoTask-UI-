import os
import tempfile
import unittest
import openpyxl
from excel_picker import ExcelPicker


class TestExcelPicker(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        """创建两个临时 Excel 文件供所有测试用例共用"""
        cls.file_valid = cls._build_excel([
            ["Name", "Age", "City"],
            ["Alice", 30, "NY"],
            ["Bob", None, "LA"],
            ["Charlie", 25, ""],
            ["", 40, "Tokyo"]
        ])
        cls.file_empty = cls._build_excel([])
        cls.file_no_column = cls._build_excel([
            ["A", "B"],
            ["x", "y"]
        ])

    @staticmethod
    def _build_excel(rows):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.close()  # 先关闭句柄，让 openpyxl 独占访问
        wb = openpyxl.Workbook()
        ws = wb.active
        for r_idx, row in enumerate(rows, 1):
            for c_idx, val in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=val)
        wb.save(tmp.name)
        wb.close()  # 关键：显式关闭工作簿
        return tmp.name

    def setUp(self):
        self.picker = ExcelPicker()

    # === 正常路径 ===
    def test_sequential_mode(self):
        params = {"excel_path": self.file_valid, "col": 0, "mode": "顺序"}
        self.assertEqual(self.picker.execute(params), "Name")  # 表头
        self.assertEqual(self.picker.execute(params), "Alice")
        self.assertEqual(self.picker.execute(params), "Bob")
        self.assertEqual(self.picker.execute(params), "Charlie")
        self.assertEqual(self.picker.execute(params), "Name")  # 循环

    def test_random_mode(self):
        params = {"excel_path": self.file_valid, "col": 2, "mode": "随机"}
        results = {self.picker.execute(params) for _ in range(20)}
        self.assertIn("NY", results)
        self.assertIn("LA", results)
        self.assertIn("Tokyo", results)

    # === 异常路径 ===
    def test_file_not_exist(self):
        params = {"excel_path": "/no/such/path.xlsx"}
        with self.assertRaises(FileNotFoundError):
            self.picker.execute(params)

    def test_empty_table(self):
        params = {"excel_path": self.file_empty}
        with self.assertRaises(ValueError) as cm:
            self.picker.execute(params)
        self.assertIn("无数据", str(cm.exception))

    def test_column_empty(self):
        params = {"excel_path": self.file_valid, "col": 2}  # City 列有空白
        cells = ["NY", "LA", "Tokyo"]  # 过滤后非空值
        # 确保过滤结果正确
        self.assertEqual(set(cells), {"NY", "LA", "Tokyo"})
        # 指定一个完全空列
        params["col"] = 999
        with self.assertRaises(ValueError) as cm:
            self.picker.execute(params)
        self.assertIn("指定列为空", str(cm.exception))

    # === 缓存命中 ===
    def test_cache_hit(self):
        params = {"excel_path": self.file_valid, "col": 1, "mode": "顺序"}
        self.picker.execute(params)
        # 再次调用不应重新加载
        self.assertIn(self.file_valid, self.picker._excel_cache)

    # === 清理 ===
    @classmethod
    def tearDownClass(cls):
        for path in (cls.file_valid, cls.file_empty, cls.file_no_column):
            try:
                os.unlink(path)
            except OSError:
                pass


if __name__ == "__main__":
    unittest.main()