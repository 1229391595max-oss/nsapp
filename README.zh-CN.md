# ICS Visits 管理

> ## ⚖️ 免责声明
>
> 本软件为**完全本地运行的离线桌面应用程序**，不会向任何远程服务器、云服务或第三方上传、传输或共享任何用户数据。
>
> 所有的数据处理、文件读取和 Excel 导出操作均在用户本机完成。使用者须自行承担以下责任：
>
> - **备份**：在使用本软件前，自行对输入文件（`.ics`）及输出文件（`.xlsx`）进行妥善备份；
> - **核对**：对软件计算出的金额、OASIS 标记以及导出的所有数据自行核实其准确性；
> - **保密**：对本软件所处理的任何病人信息或个人身份信息（PII / PHI）负有保密义务，仅可通过安全可靠的渠道与经过授权的对象共享。
>
> 本软件**按"现状"提供**，不附带任何形式的明示或暗示担保。对于因使用、误用或无法使用本软件所造成的任何数据丢失、文件损坏、信息泄露、违反法律法规（包括但不限于 HIPAA 等医疗数据隐私规定）、经济损失或其他任何损害，无论该等损害是由用户操作失误、系统故障还是其他任何原因导致，开发者均**不承担任何责任**。
>
> 您一旦开始使用本软件，即视为已阅读、理解并接受本免责声明的全部内容。

一个桌面工具，用来解析护士上门访视的 `.ics` 日历文件，自动计算访视金额、标记 OASIS 要求、检测重复事件，并导出 Excel。

📖 [English version](README.md)

> ⚠️ **已知问题（正在修复）**
> 测试中发现：打包好的 `.app` 在和打包机芯片架构不同的 Mac 上可能无法启动（比如在 Intel Mac 上跑 ARM 版本的包，或反过来）。我们还在尝试做一个稳定的跨架构构建方案。
> 如果你遇到这个问题，请在 [Issues](https://github.com/1229391595max-oss/nsapp/issues) 或 [Discussions](https://github.com/1229391595max-oss/nsapp/discussions) 留言，告诉我们你的 Mac 型号 + macOS 版本 + 具体报错。
> 临时解决办法：直接用源码跑（`python nurseapp_qt.py`），不受芯片架构限制。

## 功能

- **解析 `.ics` 文件** — 提取每个 `VEVENT`（日期、时间、病人、访视类型、地点）
- **自动计费** — 按访视类型计算金额（用户可自定义，每种类型默认 $x）
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

App 没有苹果开发者签名，对方第一次打开时 macOS 会拦截。可能遇到两种情况：

**情况 A —「无法打开"…"，因为它来自身份不明的开发者」**
右键点 `.app` → **打开** → 弹窗里再点 **打开**。第一次之后双击就正常了。

**情况 B —「"…" 已损坏，无法打开。你应该将它移到废纸篓。」**
这是 macOS 的**误报**，文件本身没损坏，是因为浏览器（Chrome / Safari 等）下载时给文件加了"隔离标记"，macOS 看到未签名 app 就直接报损坏。

修复方法：打开**终端**运行下面这条命令（路径换成你 `.app` 的实际位置）：

```bash
xattr -cr "/path/to/ICS Visits 管理.app"
```

常用路径示例：

```bash
# 如果 .app 在「下载」里
xattr -cr ~/Downloads/"ICS Visits 管理.app"

# 如果在「桌面」上
xattr -cr ~/Desktop/"ICS Visits 管理.app"
```

运行完后双击就能正常打开了。

> 💡 **给分发者的小贴士：** 想完全避免接收者遇到这个问题，可以把 `.app` 装进自己本地打的 `.dmg` 或 `.zip`（不要让对方通过浏览器下载），或者用 Apple 开发者账号做签名 + 公证。

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
| SOC      | $x    | YES            |
| DC       | $x    | YES            |
| RECERT   | $x    | YES            |
| ROC      | $x    | YES            |
| IV       | $x    | NO             |
| FU       | $x    | NO             |

金额可以由每位用户通过应用内的 ⚙️ 访视类型设置自行定义。每位用户可以自己配置代号、匹配关键词、价格和 OASIS 标记，设置会本地保存，下次打开依然生效。

访视类型从日历事件的 summary 解析，格式为 `病人名 - 访视类型`。`Follow Up` / `F/U` / `FU` 这类变体会统一归到 `FU`。

## 一些注意事项

- 取消的事件（`STATUS:CANCELLED` 或 summary 以 `cancel` 开头）默认不进 Clean，但 Raw 里仍保留 — 勾选"包含已取消事件"可以纳入。
- 去重逻辑：同 `UID + RECURRENCE-ID + DTSTART` 只保留 `SEQUENCE` / `LAST-MODIFIED` 最新的一条。
- conda 的 PyQt6 包里的 `libqcocoa.dylib` 实际链接的是 Qt5（包冲突 bug），打包脚本会在 PyInstaller 跑完后用正确的 Qt6 版本覆盖。
