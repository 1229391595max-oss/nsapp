from __future__ import annotations
from datetime import datetime, date, timedelta
from pathlib import Path
import os
import sys

# 必须在 PyQt6 导入之前设置，否则打包后找不到 cocoa 插件
if getattr(sys, "frozen", False) and sys.platform == "darwin":
    _exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    for _p in [
        os.path.join(getattr(sys, "_MEIPASS", ""), "PyQt6", "Qt6", "plugins", "platforms"),
        os.path.join(_exe_dir, "..", "Frameworks", "PyQt6", "Qt6", "plugins", "platforms"),
        os.path.join(_exe_dir, "PyQt6", "Qt6", "plugins", "platforms"),
        os.path.join(_exe_dir, "_internal", "PyQt6", "Qt6", "plugins", "platforms"),
    ]:
        _p = os.path.normpath(_p)
        if os.path.isdir(_p):
            os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = _p
            os.environ["QT_PLUGIN_PATH"] = os.path.dirname(_p)
            break

import pandas as pd
from icalendar import Calendar
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QCheckBox, QTabWidget, QComboBox, QDateEdit, QGroupBox,
    QHeaderView, QMessageBox,
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QColor, QFont

# =====================
# 金额 & OASIS 规则
# =====================

AMOUNT_RULES = {"SOC": 100, "DC": 60, "RECERT": 70, "ROC": 70, "IV": 65, "FU": 65}
OASIS_CODES = {"SOC", "DC", "RECERT", "ROC"}
YES_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
NO_FILL  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
CLEAN_COLS = ["Date", "Time In", "Time Out", "Patient Name", "Visit Type", "Location", "Amount", "OASIS Required?"]


# =====================
# Visit Type 标准化
# =====================

def normalize_visit_type(vt: str | None) -> str | None:
    if not vt:
        return None
    s = vt.upper().replace("-", " ").strip()
    if "SOC" in s:     return "SOC"
    if "DC" in s:      return "DC"
    if "RECERT" in s:  return "RECERT"
    if "ROC" in s:     return "ROC"
    if "FOLLOW UP" in s or s in {"FU", "F/U"}: return "FU"
    if "IV" in s:      return "IV"
    return s

def calculate_amount(vt: str | None) -> int:
    return AMOUNT_RULES.get(normalize_visit_type(vt), 0)

def oasis_required(vt: str | None) -> str:
    return "YES" if normalize_visit_type(vt) in OASIS_CODES else "NO"


# =====================
# 时间安全处理
# =====================

def to_date_time(dt):
    if dt is None:
        return None, None
    if isinstance(dt, datetime):
        return dt.date().isoformat(), dt.time().strftime("%H:%M:%S")
    if isinstance(dt, date):
        return dt.isoformat(), None
    return str(dt), None

def to_iso(dt):
    if dt is None:
        return None
    if isinstance(dt, datetime):
        return dt.isoformat(sep=" ")
    if isinstance(dt, date):
        return dt.isoformat()
    return str(dt)


# =====================
# Summary 拆解
# =====================

def split_summary(summary: str | None):
    if not summary:
        return None, None
    s = summary.replace("—", "-").replace("–", "-").strip()
    if " - " in s:
        left, right = s.split(" - ", 1)
        return left.strip(), right.strip()
    if "-" in s:
        left, right = s.split("-", 1)
        return left.strip(), right.strip()
    return s.strip(), None


# =====================
# 取消/删除事件识别
# =====================

def is_cancelled_event(ev) -> bool:
    status = ev.get("status")
    if status and str(status).upper() == "CANCELLED":
        return True
    summary = ev.get("summary")
    if summary:
        s = str(summary).strip().lower()
        if s.startswith("cancelled") or s.startswith("canceled") or s.startswith("cancel"):
            return True
    return False

def calendar_method_is_cancel(cal: Calendar) -> bool:
    m = cal.get("method")
    return bool(m and str(m).upper() == "CANCEL")


# =====================
# ICS 解析
# =====================

def parse_ics(file_bytes: bytes, include_cancelled_in_clean: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    cal = Calendar.from_ical(file_bytes)
    rows_clean, rows_raw = [], []
    cal_cancel_mode = calendar_method_is_cancel(cal)

    def safe_int(x, default=0):
        try: return int(x)
        except: return default

    for ev in cal.walk("VEVENT"):
        summary     = str(ev.get("summary"))     if ev.get("summary")     else None
        location    = str(ev.get("location"))    if ev.get("location")    else None
        uid         = str(ev.get("uid"))          if ev.get("uid")         else None
        description = str(ev.get("description")) if ev.get("description") else None
        status      = str(ev.get("status")).upper() if ev.get("status")   else None
        sequence    = safe_int(ev.get("sequence"), 0)

        last_modified  = ev.get("last-modified").dt  if ev.get("last-modified")  else None
        dtstamp        = ev.get("dtstamp").dt         if ev.get("dtstamp")        else None
        recurrence_id  = ev.get("recurrence-id").dt   if ev.get("recurrence-id") else None
        start          = ev.get("dtstart").dt          if ev.get("dtstart")       else None
        end            = ev.get("dtend").dt            if ev.get("dtend")         else None

        start_iso = to_iso(start)
        end_iso   = to_iso(end)
        d, tin    = to_date_time(start)
        _, tout   = to_date_time(end)

        patient, raw_type = split_summary(summary)
        vt_norm = normalize_visit_type(raw_type)
        cancelled = cal_cancel_mode or is_cancelled_event(ev) or (status == "CANCELLED")

        uid_fallback = uid if uid else f"NOUID|{summary or ''}|{location or ''}"
        rid_iso      = to_iso(recurrence_id) if recurrence_id else ""
        event_key    = f"{uid_fallback}|RID={rid_iso}|{start_iso}|{end_iso}|{summary or ''}|{location or ''}"

        rows_raw.append({
            "EVENT_KEY": event_key, "UID": uid,
            "RECURRENCE-ID": to_iso(recurrence_id) if recurrence_id else None,
            "SEQUENCE": sequence, "LAST-MODIFIED": to_iso(last_modified),
            "DTSTAMP": to_iso(dtstamp), "STATUS": status,
            "CANCELLED?": "YES" if cancelled else "NO",
            "Summary (Raw)": summary, "Patient Name (Parsed)": patient,
            "Visit Type (Raw)": raw_type, "DTSTART (Raw)": start_iso,
            "DTEND (Raw)": end_iso, "Date (Parsed)": d,
            "Time In (Parsed)": tin, "Time Out (Parsed)": tout,
            "Location (Raw)": location, "Description (Raw)": description,
        })

        if cancelled and not include_cancelled_in_clean:
            continue

        rows_clean.append({
            "EVENT_KEY": event_key, "Date": d, "Time In": tin, "Time Out": tout,
            "Patient Name": patient, "Visit Type": vt_norm, "Location": location,
            "Amount": calculate_amount(vt_norm), "OASIS Required?": oasis_required(vt_norm),
        })

    df_raw   = pd.DataFrame(rows_raw)
    df_clean = pd.DataFrame(rows_clean)

    if len(df_raw):
        df_raw["_lm"] = pd.to_datetime(df_raw["LAST-MODIFIED"], errors="coerce")
        df_raw["_ds"] = pd.to_datetime(df_raw["DTSTAMP"],        errors="coerce")
        df_raw = df_raw.sort_values(
            by=["UID", "RECURRENCE-ID", "DTSTART (Raw)", "SEQUENCE", "_lm", "_ds"],
            na_position="first",
        )
        has_uid  = df_raw["UID"].notna() & (df_raw["UID"].astype(str).str.strip() != "")
        non_empty = df_raw[has_uid].copy()
        empty_uid = df_raw[~has_uid].copy()
        if len(non_empty):
            non_empty = non_empty.drop_duplicates(
                subset=["UID", "RECURRENCE-ID", "DTSTART (Raw)"], keep="last"
            )
        df_raw = pd.concat([non_empty, empty_uid], ignore_index=True)
        df_raw = df_raw.drop_duplicates(subset=["EVENT_KEY"], keep="last")
        df_raw = df_raw.drop(columns=["_lm", "_ds"], errors="ignore")

        keep_keys = set(df_raw["EVENT_KEY"].dropna())
        if len(df_clean):
            df_clean = df_clean[df_clean["EVENT_KEY"].isin(keep_keys)].copy()

    return df_raw, df_clean


# =====================
# Excel 导出
# =====================

def export_excel_clean(df: pd.DataFrame, path: Path):
    cols = [c for c in CLEAN_COLS if c in df.columns]
    df[cols].to_excel(path, index=False)
    wb = load_workbook(path)
    ws = wb.active
    oc = cols.index("OASIS Required?") + 1
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(r, oc)
        cell.fill = YES_FILL if cell.value == "YES" else NO_FILL
    wb.save(path)

def export_excel_raw(df_raw: pd.DataFrame, path: Path):
    df_raw.to_excel(path, index=False)

def export_excel_bundle(df_clean: pd.DataFrame, df_raw: pd.DataFrame, path: Path):
    cols = [c for c in CLEAN_COLS if c in df_clean.columns]
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_clean[cols].to_excel(writer, sheet_name="Clean_Amount", index=False)
        df_raw.to_excel(writer, sheet_name="Raw_Original", index=False)
        ws = writer.book["Clean_Amount"]
        header = [c.value for c in ws[1]]
        oc = header.index("OASIS Required?") + 1
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, oc)
            cell.fill = YES_FILL if cell.value == "YES" else NO_FILL


# =====================
# 重复检测报告
# =====================

def build_duplicate_report(df_raw: pd.DataFrame, df_clean: pd.DataFrame) -> dict:
    tables: dict[str, pd.DataFrame] = {}

    if len(df_raw):
        miss = df_raw[df_raw["UID"].isna() | (df_raw["UID"].astype(str).str.strip() == "")]
        tables["UID 缺失（Raw）"] = miss.copy()

    if len(df_raw) and "CANCELLED?" in df_raw.columns:
        tables["被取消/删除事件（Raw）"] = df_raw[df_raw["CANCELLED?"] == "YES"].copy()

    if len(df_raw):
        key_cols = ["UID", "RECURRENCE-ID", "DTSTART (Raw)"]
        tmp = df_raw[df_raw["UID"].notna() & (df_raw["UID"].astype(str).str.strip() != "")].copy()
        if len(tmp):
            vc = tmp.groupby(key_cols, dropna=False).size().reset_index(name="count")
            multi = vc[vc["count"] > 1]
            if len(multi):
                mv = tmp.merge(multi[key_cols], on=key_cols, how="inner")
                tables["同一 UID 的多版本（Raw）"] = mv.sort_values(key_cols + ["SEQUENCE"], na_position="last")

    if len(df_raw) and "EVENT_KEY" in df_raw.columns:
        dk = df_raw[df_raw.duplicated(subset=["EVENT_KEY"], keep=False)].copy()
        if len(dk):
            tables["EVENT_KEY 完全重复（Raw）"] = dk

    if len(df_clean):
        tmpc = df_clean.copy()
        tmpc["Date"] = pd.to_datetime(tmpc["Date"], errors="coerce").dt.date
        base = ["Date", "Time In", "Patient Name"]
        if all(c in tmpc.columns for c in base):
            ct = tmpc.groupby(base, dropna=False).size().reset_index(name="count")
            ct = ct[ct["count"] > 1]
            if len(ct):
                tables["同病人同时间重复（Clean，疑似）"] = (
                    tmpc.merge(ct[base], on=base, how="inner").sort_values(base)
                )

    rows = [{"Issue Type": k, "Rows": int(len(v))} for k, v in tables.items()]
    if not rows:
        rows = [{"Issue Type": "未检测到明显重复/异常", "Rows": 0}]
    return {"summary": pd.DataFrame(rows), "tables": tables}

def export_dup_report_excel(report: dict, path: Path):
    def safe(name: str) -> str:
        for b in ['\\', '/', '*', '?', ':', '[', ']']:
            name = name.replace(b, "_")
        return name[:31]

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        report["summary"].to_excel(writer, sheet_name="Summary", index=False)
        for k, df_ in report["tables"].items():
            sn = safe(k)
            if df_ is None or df_.empty:
                pd.DataFrame([{"info": "no rows"}]).to_excel(writer, sheet_name=sn, index=False)
            else:
                df_.to_excel(writer, sheet_name=sn, index=False)


# =====================
# PyQt6 UI helpers
# =====================

def make_table() -> QTableWidget:
    t = QTableWidget()
    t.setAlternatingRowColors(True)
    t.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    t.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    t.verticalHeader().setVisible(False)
    t.setSortingEnabled(True)
    t.horizontalHeader().setStretchLastSection(True)
    return t

def fill_table(table: QTableWidget, df: pd.DataFrame, oasis_col: str | None = None):
    table.setSortingEnabled(False)
    table.clearContents()
    table.setRowCount(0)

    if df is None or df.empty:
        table.setColumnCount(0)
        return

    cols = df.columns.tolist()
    table.setColumnCount(len(cols))
    table.setHorizontalHeaderLabels(cols)
    table.setRowCount(len(df))

    for row_idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, col_name in enumerate(cols):
            val  = row[col_name]
            text = "" if (val is None or (isinstance(val, float) and pd.isna(val))) else str(val)
            item = QTableWidgetItem(text)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

            if oasis_col and col_name == oasis_col:
                if text == "YES":
                    item.setBackground(QColor("#C6EFCE"))
                    item.setForeground(QColor("#276221"))
                elif text == "NO":
                    item.setBackground(QColor("#FFC7CE"))
                    item.setForeground(QColor("#9C0006"))

            table.setItem(row_idx, col_idx, item)

    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    table.horizontalHeader().setStretchLastSection(True)
    table.setSortingEnabled(True)


# =====================
# Main window
# =====================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ICS Visits 管理与金额统计")
        self.resize(1400, 900)

        self.df_raw      = pd.DataFrame()
        self.df_clean    = pd.DataFrame()
        self.df_filt     = pd.DataFrame()
        self.df_filt_raw = pd.DataFrame()
        self._report: dict | None = None

        self._build_ui()

    # ------------------------------------------------------------------ build

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        # toolbar
        bar = QHBoxLayout()
        self.btn_upload = QPushButton("📤 上传 .ics 文件")
        self.btn_upload.setFixedHeight(34)
        self.btn_upload.clicked.connect(self.upload_ics)

        self.btn_reset = QPushButton("🔄 重置")
        self.btn_reset.setFixedHeight(34)
        self.btn_reset.clicked.connect(self.reset_data)

        self.chk_cancelled = QCheckBox("包含已取消事件到 Clean（一般不建议）")

        bar.addWidget(self.btn_upload)
        bar.addWidget(self.btn_reset)
        bar.addSpacing(16)
        bar.addWidget(self.chk_cancelled)
        bar.addStretch()
        layout.addLayout(bar)

        # tabs
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        self._build_clean_tab()
        self._build_raw_tab()
        self._build_dup_tab()

        self.statusBar().showMessage("请上传 .ics 文件")

    def _build_clean_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setSpacing(6)

        # date filter
        grp = QGroupBox("日期范围筛选")
        h = QHBoxLayout(grp)
        h.setSpacing(8)

        h.addWidget(QLabel("快速选择:"))
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(["本日", "本周", "本月", "上月", "全部", "自定义"])
        self.preset_combo.setCurrentIndex(1)
        self.preset_combo.currentIndexChanged.connect(self._on_preset_changed)
        h.addWidget(self.preset_combo)
        h.addSpacing(12)

        h.addWidget(QLabel("从:"))
        self.date_start = QDateEdit(QDate.currentDate())
        self.date_start.setCalendarPopup(True)
        self.date_start.setEnabled(False)
        self.date_start.dateChanged.connect(self._apply_filter)
        h.addWidget(self.date_start)

        h.addWidget(QLabel("到:"))
        self.date_end = QDateEdit(QDate.currentDate())
        self.date_end.setCalendarPopup(True)
        self.date_end.setEnabled(False)
        self.date_end.dateChanged.connect(self._apply_filter)
        h.addWidget(self.date_end)
        h.addStretch()
        v.addWidget(grp)

        # metrics
        mh = QHBoxLayout()
        self.lbl_amount = QLabel("💰 总金额: $0")
        self.lbl_visits = QLabel("📊 Visits 数量: 0")
        bold = QFont()
        bold.setPointSize(13)
        bold.setBold(True)
        self.lbl_amount.setFont(bold)
        self.lbl_visits.setFont(bold)
        mh.addWidget(self.lbl_amount)
        mh.addSpacing(40)
        mh.addWidget(self.lbl_visits)
        mh.addStretch()
        v.addLayout(mh)

        # table
        self.clean_table = make_table()
        v.addWidget(self.clean_table)

        # export
        eh = QHBoxLayout()
        btn_c = QPushButton("⬇️ 导出 Clean Excel（带金额+上色）")
        btn_c.clicked.connect(self.export_clean)
        btn_b = QPushButton("⬇️ 导出 Bundle（Clean + Raw 两个 Sheet）")
        btn_b.clicked.connect(self.export_bundle)
        eh.addWidget(btn_c)
        eh.addWidget(btn_b)
        eh.addStretch()
        v.addLayout(eh)

        self.tabs.addTab(w, "📋 Clean 数据")

    def _build_raw_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)

        self.raw_table = make_table()
        v.addWidget(self.raw_table)

        eh = QHBoxLayout()
        btn_r = QPushButton("⬇️ 导出 Raw Excel（原版对照）")
        btn_r.clicked.connect(self.export_raw)
        eh.addWidget(btn_r)
        eh.addStretch()
        v.addLayout(eh)

        self.tabs.addTab(w, "🧾 Raw 原版")

    def _build_dup_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setSpacing(6)

        v.addWidget(QLabel("问题汇总:"))
        self.dup_summary = make_table()
        self.dup_summary.setMaximumHeight(120)
        v.addWidget(self.dup_summary)

        v.addWidget(QLabel("明细（每类问题一个标签页）:"))
        self.dup_detail_tabs = QTabWidget()
        v.addWidget(self.dup_detail_tabs)

        eh = QHBoxLayout()
        btn_d = QPushButton("⬇️ 导出 Duplicate Report Excel")
        btn_d.clicked.connect(self.export_dup)
        eh.addWidget(btn_d)
        eh.addStretch()
        v.addLayout(eh)

        self.tabs.addTab(w, "🧠 重复检测")

    # ------------------------------------------------------------------ slots

    def upload_ics(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 .ics 文件", "", "iCalendar Files (*.ics)"
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                data = f.read()
            self.df_raw, self.df_clean = parse_ics(
                data, include_cancelled_in_clean=self.chk_cancelled.isChecked()
            )
        except Exception as e:
            QMessageBox.critical(self, "解析失败", str(e))
            return
        self._update_all(Path(path).name)

    def reset_data(self):
        self.df_raw = self.df_clean = self.df_filt = self.df_filt_raw = pd.DataFrame()
        self._report = None
        for t in (self.clean_table, self.raw_table, self.dup_summary):
            t.clearContents()
            t.setRowCount(0)
            t.setColumnCount(0)
        self.dup_detail_tabs.clear()
        self.lbl_amount.setText("💰 总金额: $0")
        self.lbl_visits.setText("📊 Visits 数量: 0")
        self.statusBar().showMessage("已重置，请上传 .ics 文件")

    def _on_preset_changed(self):
        is_custom = self.preset_combo.currentText() == "自定义"
        self.date_start.setEnabled(is_custom)
        self.date_end.setEnabled(is_custom)
        self._apply_filter()

    def _apply_filter(self):
        if self.df_clean.empty:
            return

        df = self.df_clean.copy()
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
        valid = df["Date"].dropna()
        if valid.empty:
            return

        today = date.today()
        min_d, max_d = valid.min(), valid.max()
        preset = self.preset_combo.currentText()

        if preset == "本日":
            start = end = today
        elif preset == "本周":
            start = today - timedelta(days=today.weekday())
            end = today
        elif preset == "本月":
            start = today.replace(day=1)
            end = today
        elif preset == "上月":
            first_this = today.replace(day=1)
            end = first_this - timedelta(days=1)
            start = end.replace(day=1)
        elif preset == "全部":
            start, end = min_d, max_d
        else:
            qs = self.date_start.date()
            qe = self.date_end.date()
            start = date(qs.year(), qs.month(), qs.day())
            end   = date(qe.year(), qe.month(), qe.day())

        self.df_filt = df[(df["Date"] >= start) & (df["Date"] <= end)].copy()

        raw = self.df_raw.copy()
        if not raw.empty and "Date (Parsed)" in raw.columns:
            raw["Date (Parsed)"] = pd.to_datetime(raw["Date (Parsed)"], errors="coerce").dt.date
            self.df_filt_raw = raw[
                (raw["Date (Parsed)"] >= start) & (raw["Date (Parsed)"] <= end)
            ].copy()
        else:
            self.df_filt_raw = raw

        total = int(self.df_filt["Amount"].sum()) if "Amount" in self.df_filt.columns else 0
        self.lbl_amount.setText(f"💰 总金额: ${total:,}")
        self.lbl_visits.setText(f"📊 Visits 数量: {len(self.df_filt)}")

        cols = [c for c in CLEAN_COLS if c in self.df_filt.columns]
        sorted_filt = self.df_filt.sort_values(["Date", "Time In"], na_position="last")
        fill_table(self.clean_table, sorted_filt[cols], oasis_col="OASIS Required?")

    def _update_all(self, filename: str = ""):
        fill_table(self.raw_table, self.df_raw)

        if not self.df_clean.empty:
            df = self.df_clean.copy()
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
            valid = df["Date"].dropna()
            if not valid.empty:
                mn, mx = valid.min(), valid.max()
                self.date_start.setMinimumDate(QDate(mn.year, mn.month, mn.day))
                self.date_start.setMaximumDate(QDate(mx.year, mx.month, mx.day))
                self.date_end.setMinimumDate(QDate(mn.year, mn.month, mn.day))
                self.date_end.setMaximumDate(QDate(mx.year, mx.month, mx.day))
                self.date_end.setDate(QDate(mx.year, mx.month, mx.day))

        self._apply_filter()

        if not self.df_raw.empty or not self.df_clean.empty:
            self._report = build_duplicate_report(self.df_raw, self.df_clean)
            fill_table(self.dup_summary, self._report["summary"])
            self.dup_detail_tabs.clear()
            for k, tdf in self._report["tables"].items():
                t = make_table()
                fill_table(t, tdf)
                label = k[:20] + ("…" if len(k) > 20 else "")
                self.dup_detail_tabs.addTab(t, label)

        self.statusBar().showMessage(
            f"已加载: {filename}  |  Raw: {len(self.df_raw)} 条  |  Clean: {len(self.df_clean)} 条"
        )

    # ------------------------------------------------------------------ export

    def _save_path(self, title: str, default: str) -> Path | None:
        p, _ = QFileDialog.getSaveFileName(self, title, default, "Excel Files (*.xlsx)")
        return Path(p) if p else None

    def export_clean(self):
        if self.df_filt.empty:
            return QMessageBox.warning(self, "无数据", "请先上传并筛选数据。")
        p = self._save_path("保存 Clean Excel", "HH_Visits_With_Amount.xlsx")
        if p:
            export_excel_clean(self.df_filt, p)
            QMessageBox.information(self, "已保存", str(p))

    def export_raw(self):
        if self.df_raw.empty:
            return QMessageBox.warning(self, "无数据", "请先上传数据。")
        p = self._save_path("保存 Raw Excel", "HH_Visits_Raw.xlsx")
        if p:
            raw = self.df_filt_raw if not self.df_filt_raw.empty else self.df_raw
            export_excel_raw(raw, p)
            QMessageBox.information(self, "已保存", str(p))

    def export_bundle(self):
        if self.df_filt.empty:
            return QMessageBox.warning(self, "无数据", "请先上传并筛选数据。")
        p = self._save_path("保存 Bundle Excel", "HH_Visits_Bundle.xlsx")
        if p:
            raw = self.df_filt_raw if not self.df_filt_raw.empty else self.df_raw
            export_excel_bundle(self.df_filt, raw, p)
            QMessageBox.information(self, "已保存", str(p))

    def export_dup(self):
        if self._report is None:
            return QMessageBox.warning(self, "无报告", "请先上传数据生成报告。")
        p = self._save_path("保存 Duplicate Report", "Duplicate_Report.xlsx")
        if p:
            export_dup_report_excel(self._report, p)
            QMessageBox.information(self, "已保存", str(p))


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
