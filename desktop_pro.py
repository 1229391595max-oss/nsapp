
import sys
from pathlib import Path

print("=== DEBUG START ===")

BASE_DIR = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).parent
print("BASE_DIR:", BASE_DIR)

sys.path.append(str(BASE_DIR))

print("FILES:", list(BASE_DIR.iterdir()))

import sys
import pandas as pd
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QTableWidget, QTableWidgetItem,
    QHBoxLayout, QDateEdit
)
from PySide6.QtCore import QDate

# 复用你的核心逻辑
from nurseapp import parse_ics, export_excel_clean


class App(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("ICS Visits Pro")
        self.resize(900, 600)

        self.df = pd.DataFrame()

        layout = QVBoxLayout()

        # ===== 上传 =====
        self.upload_btn = QPushButton("📂 选择 ICS 文件")
        self.upload_btn.clicked.connect(self.load_file)
        layout.addWidget(self.upload_btn)

        # ===== 日期筛选 =====
        filter_layout = QHBoxLayout()

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)

        self.filter_btn = QPushButton("筛选")
        self.filter_btn.clicked.connect(self.apply_filter)

        filter_layout.addWidget(QLabel("开始"))
        filter_layout.addWidget(self.start_date)
        filter_layout.addWidget(QLabel("结束"))
        filter_layout.addWidget(self.end_date)
        filter_layout.addWidget(self.filter_btn)

        layout.addLayout(filter_layout)

        # ===== 统计 =====
        self.stats_label = QLabel("统计：")
        layout.addWidget(self.stats_label)

        # ===== 表格 =====
        self.table = QTableWidget()
        layout.addWidget(self.table)

        # ===== 导出 =====
        self.export_btn = QPushButton("⬇️ 导出 Excel")
        self.export_btn.clicked.connect(self.export_excel)
        layout.addWidget(self.export_btn)

        self.setLayout(layout)

    # ===== 加载文件 =====
    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择 ICS 文件", "", "ICS Files (*.ics)"
        )

        if not file_path:
            return

        with open(file_path, "rb") as f:
            _, clean = parse_ics(f.read())

        self.df = clean.copy()
        self.df["Date"] = pd.to_datetime(self.df["Date"])

        # 设置默认日期范围
        self.start_date.setDate(QDate(self.df["Date"].min().year,
                                     self.df["Date"].min().month,
                                     self.df["Date"].min().day))

        self.end_date.setDate(QDate.currentDate())

        self.update_table(self.df)

    # ===== 筛选 =====
    def apply_filter(self):
        if self.df.empty:
            return

        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()

        filt = self.df[
            (self.df["Date"] >= pd.to_datetime(start)) &
            (self.df["Date"] <= pd.to_datetime(end))
        ]

        self.update_table(filt)

    # ===== 更新表格 =====
    def update_table(self, df):
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)

        for i in range(len(df)):
            for j in range(len(df.columns)):
                self.table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))

        # 更新统计
        total = int(df["Amount"].sum()) if "Amount" in df.columns else 0
        count = len(df)

        self.stats_label.setText(f"总金额: ${total} | Visits: {count}")

    # ===== 导出 =====
    def export_excel(self):
        if self.df.empty:
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "保存 Excel", "output.xlsx", "Excel Files (*.xlsx)"
        )

        if path:
            export_excel_clean(self.df, Path(path))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())