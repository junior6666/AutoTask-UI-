import sys
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton


class Demo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PySide6 Hover Style Test")

        # 1. 创建按钮
        self.apply_schedule_btn = QPushButton("Apply Schedule")

        # 2. 设置样式表（正常 / 悬停）
        self.apply_schedule_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                border: 1px solid #F57C00;
                color: white;
                padding: 5px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)

        # 3. 布局
        layout = QVBoxLayout(self)
        layout.addWidget(self.apply_schedule_btn)
        layout.addStretch()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = Demo()
    w.resize(240, 120)
    w.show()
    sys.exit(app.exec())