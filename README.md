# Express Audit Tool 快递单号比对审查工具

> 中文 | [English](#english)

A cross-platform desktop tool for macOS and Windows. Import two Excel files, match records by tracking number, flag single-sided numbers as anomalies, fuzzy-compare exact weight against weight ranges, and export a highlighted audit report with all anomalies extracted.

## 功能特性

- 双表匹配：按快递单号精确匹配两份 Excel
- 单边异常检测：仅在 A 方或仅 B 方存在的单号自动标记
- 重量模糊比对：A 方具体重量 vs B 方区间重量，支持多种格式与单位换算
  - 支持区间格式：`1-2`、`0.5~1.5公斤`、`1至2kg`、`小于3`、`不超过2.5`、`大于1`、`3以上`
  - 单位自动识别：g / kg / 吨 / lb，缺省按 kg
- 容差设置：如填 `0.05` 表示正负 50g 以内视为一致
- 图形界面：列名下拉选择（自动读取实际列名并智能预选），无需手动输入

## 报告输出

生成 `审查报告.xlsx`，包含 4 个工作表：

| 工作表 | 内容 |
|---|---|
| 比对明细 | 全部匹配结果及说明 |
| 异常提取 | 仅异常项（重量不符 / 单边缺单） |
| A方标注 | A方数据副本，异常行红色高亮并附审查备注 |
| B方标注 | B方数据副本，同上 |

## 使用步骤

1. 选择 A 方表格（具体重量一方）和 B 方表格（区间重量一方）
2. 从下拉框确认四个列选择（A方单号列 / A方重量列 / B方单号列 / B方区间列）
3. 可选设置容差
4. 点击「开始比对审查」，报告生成在 A 方文件同目录

## 从源码运行

```bash
pip install pandas openpyxl   # tkinter 为 Python 自带
python app.py
```

## 打包为独立可执行文件

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name express-audit app.py
```

注意：Mac 包需在 macOS 上打包，Windows 包需在 Windows 上打包（不支持交叉打包）。

---

<a name="english"></a>
# Express Audit Tool

English | [中文](#快递单号比对审查工具)

A cross-platform desktop tool for macOS and Windows. Import two Excel files, match records by tracking number, flag single-sided numbers as anomalies, fuzzy-compare exact weights against weight ranges, and export an audit report with anomalies highlighted and extracted.

## Features

- Dual-sheet matching: precise matching of two Excel files by tracking number
- Single-sided anomaly detection: numbers existing only on side A or side B are flagged
- Fuzzy weight comparison: exact weight (side A) vs range weight (side B)
  - Range formats: `1-2`, `0.5~1.5`, `less than 3`, `3+`, etc.
  - Automatic unit detection: g / kg / ton / lb, defaults to kg
- Tolerance: e.g. `0.05` treats differences within +/-50g as consistent
- GUI: column dropdowns auto-populated from actual headers with smart pre-selection

## Report Output

Generates `审查报告.xlsx` (audit report) with 4 worksheets:

| Sheet | Content |
|---|---|
| Details | All match results with explanations |
| Anomalies | Only abnormal entries (weight mismatch / missing on one side) |
| Side A Annotated | Copy of side A data, abnormal rows highlighted in red with notes |
| Side B Annotated | Same as above for side B |

## Usage

1. Select the side-A file (exact weights) and the side-B file (weight ranges)
2. Confirm four column selections from the dropdowns
3. Optionally set a tolerance
4. Click "Start Audit"; the report is saved next to the side-A file

## Run from Source

```bash
pip install pandas openpyxl   # tkinter ships with Python
python app.py
```

## Build Standalone Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name express-audit app.py
```

Note: build the .app on macOS and the .exe on Windows; cross-compilation is not supported.
