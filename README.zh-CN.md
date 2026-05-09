# ICS Visits 管理

一个桌面工具，用来解析护士上门访视的 `.ics` 日历文件，自动计算访视金额、标记 OASIS 要求、检测重复事件，并导出 Excel。

📖 [English version](README.md)

## 功能

- **解析 `.ics` 文件** — 提取每个 `VEVENT`（日期、时间、病人、访视类型、地点）
- **自动计费** — 按访视类型计算金额（SOC $100、DC $60、RECERT/ROC $70、IV/FU $65）
- **OASIS 自动标记** — SOC / DC / RECERT / ROC 自动标为需要 OASIS
- **重复 & 取消事件报告** — 同 UID 多版本、`STATUS:CANCELLED`（"删了又回来"）、UID 缺失、同病人同时间冲突
- **日期范围筛选** — 本日 / 本周 / 本月 / 上月 / 全部 / 自定义
- **Excel 导出** — Clean（带金额 + OASIS 上色）、Raw（完整原始字段）、Bundle（两个 Sheet）、单独的重复检测报告
- **原生 macOS UI**，PyQt6 实现，不需要打开浏览器

## 两个版本

| 文件 | 界面 | 说明 |
|------|------|------|
| `nurseapp_qt.py` | 原生 PyQt6 窗口 | **推荐**，可以打包成 `.app` |
| `nurseapp.py` | Streamlit（浏览器） | 最初的原型，仍能用 |

## 环境要求

- macOS（在 Apple Silicon 上测试过；打包脚本用 `universal2`，Intel 也能跑）
- Python 3.10+
- 依赖：`pandas`、`PyQt6`、`icalendar`、`openpyxl`

## 快速开始（从源码运行）

```bash
pip install pandas PyQt6 icalendar openpyxl
python nurseapp_qt.py
```

Streamlit 版本：

```bash
pip install streamlit pandas icalendar openpyxl
streamlit run nurseapp.py
```

## 打包成可分发的 `.app`（macOS）

打包脚本会自动建立干净虚拟环境、只装必要的包、用 PyInstaller 打包、修复 Qt cocoa 插件、最后压缩成 zip。

**前置条件**：先去 [python.org 下载 Python](https://www.python.org/downloads/macos/)，选 **macOS 64-bit universal2 installer**。如果装的不是 3.13，需要把 `build_app.sh` 里的版本号路径改一下。

```bash
bash build_app.sh
```

产物：
- `dist/ICS Visits 管理.app` — 双击即可使用
- `dist/ICS Visits 管理.zip` — 发给别人用

打出来的 `.app` 是 `universal2`，Intel 和 Apple Silicon 都能跑。

## 给不懂技术的人用

App 没有苹果开发者签名，对方第一次打开时 macOS 会拦截。告诉对方：

> 右键点 `.app` → **打开** → 弹窗里再点 **打开**。第一次之后双击就正常了。

## 项目结构

```
.
├── nurseapp_qt.py     # 主程序（PyQt6）
├── nurseapp.py        # Streamlit 版本（旧）
├── run_app.py         # Streamlit 启动器
├── build_app.sh       # macOS .app 打包脚本
├── rthook_qt.py       # PyInstaller 运行时钩子，处理 Qt 插件路径
└── README.md
```

## 访视类型对照表

| 类型     | 金额  | OASIS 是否必需 |
|----------|-------|----------------|
| SOC      | $100  | YES            |
| DC       | $60   | YES            |
| RECERT   | $70   | YES            |
| ROC      | $70   | YES            |
| IV       | $65   | NO             |
| FU       | $65   | NO             |

访视类型从日历事件的 summary 解析，格式为 `病人名 - 访视类型`。`Follow Up` / `F/U` / `FU` 这类变体会统一归到 `FU`。

## 一些注意事项

- 取消的事件（`STATUS:CANCELLED` 或 summary 以 `cancel` 开头）默认不进 Clean，但 Raw 里仍保留 — 勾选"包含已取消事件"可以纳入。
- 去重逻辑：同 `UID + RECURRENCE-ID + DTSTART` 只保留 `SEQUENCE` / `LAST-MODIFIED` 最新的一条。
- conda 的 PyQt6 包里的 `libqcocoa.dylib` 实际链接的是 Qt5（包冲突 bug），打包脚本会在 PyInstaller 跑完后用正确的 Qt6 版本覆盖。
