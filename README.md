# excel_tool

**作者 / Author:** wangpengfei44  
**版本 / Version:** 0.0.1  
**类型 / Type:** 工具插件 / Tool Plugin  

**| [中文版本](#zh) | [English Version](#en) |**  
**最后更新 / Last Updated:** 2025-06-18

---

<a id="zh"></a>
## 中文版本

### 描述

`excel_tool` 是一个轻量级插件，可以将二维数组（以 JSON 字符串形式传入）转换成 Excel 文件（.xlsx）。  
支持设置列宽、合并单元格、单元格样式（支持单格或区域配置）与行高。

### 使用方法

传入名为 `data_json` 的 JSON 字符串参数，包含以下字段：

- `data`：二维数组，表示表格的行和列  
- `col_widths`（可选）：列宽配置，字典格式，键为列号（从 1 开始），值为宽度数值  
- `merges`（可选）：合并单元格列表，每个元素包含 `start_row`、`start_col`、`end_row`、`end_col`  
- `cell_styles`（可选）：单元格样式配置，列表格式，每个元素为：  
  - `start_row`、`start_col`、`end_row`、`end_col`：样式应用的区域坐标（可以只设置起始坐标来指定单个单元格）  
  - `style`：样式对象，支持字段：  
    - `alignment`：水平对齐方式，`"center"` / `"left"` / `"right"`  
    - `vertical`：垂直对齐方式，`"top"` / `"center"` / `"bottom"`  
    - `font_size`：字体大小（整数）  
    - `bold`：是否加粗（布尔值）  
    - `bgcolor`：背景色（例如 `"FFFF00"`）  
    - `border`：是否显示边框（布尔值） 
    - `wrap_text`: 是否自动换行  默认为true
- `row_heights`（可选）：行高配置，字典格式，键为行号（字符串），值为行高数值  

### 示例输入 JSON
```
{
  "data": [
    ["科目", "2023", "2024", "2025"],
    ["资产", "1000", "1100", "1200"],
    ["负债", "500", "550", "600"]
  ],
  "col_widths": {"1": 15, "2": 10, "3": 10, "4": 10},
  "merges": [
    {"start_row": 1, "start_col": 1, "end_row": 1, "end_col": 4}
  ],
  "cell_styles": [
    {
      "start_row": 1,
      "start_col": 1,
      "end_row": 1,
      "end_col": 1,
      "style": {
        "alignment": "center",
        "vertical": "center",
        "font_size": 14,
        "bold": true,
        "bgcolor": "FFFF00",
        "border": true,
        "wrap_text": false
      }
    },
    {
      "start_row": 2,
      "start_col": 1,
      "end_row": 2,
      "end_col": 1,
      "style": {
        "alignment": "center",
        "border": true
      }
    }
  ],
  "row_heights": {"1": 30, "2": 25, "3": 25}
}
```
### 功能

- 将 JSON 格式的二维数组转换为 Excel 文件  
- 支持自定义列宽  
- 支持合并单元格  
- 支持单元格样式（对齐、字体、背景色、边框）  
- 支持设置行高  

### 运行环境

- Python 3.12 及以上版本  
- 依赖库：`openpyxl`

---

<a id="en"></a>
## English Version

### Description

`excel_tool` is a lightweight plugin that converts 2D arrays (provided as JSON strings) into Excel (.xlsx) files.  
It supports setting column widths, merging cells, cell styles (including cell range control), and row heights.

### Usage

Pass a JSON string parameter named `data_json` with the following fields:

- `data`: a 2D array representing rows and columns  
- `col_widths` (optional): dictionary of column widths; keys are column numbers (starting from 1), values are width values  
- `merges` (optional): list of merged cell ranges, each with `start_row`, `start_col`, `end_row`, `end_col`  
- `cell_styles` (optional): list of cell style regions, each object contains:  
  - `start_row`, `start_col`, `end_row`, `end_col`: coordinates of the style region (can be one cell)  
  - `style`: object that supports:  
    - `alignment`: `"center"`, `"left"`, or `"right"`  
    - `vertical`: `"top"`, `"center"`, or `"bottom"`  
    - `font_size`: font size (integer)  
    - `bold`: boolean for bold text  
    - `bgcolor`: background color as hex string (e.g., `"FFFF00"`)  
    - `border`: boolean to apply thin border  
- `row_heights` (optional): row height settings; keys are row numbers (as strings), values are height values  

### Example Input JSON

{
  "data": [
    ["Subject", "2023", "2024", "2025"],
    ["Assets", "1000", "1100", "1200"],
    ["Liabilities", "500", "550", "600"]
  ],
  "col_widths": {"1": 15, "2": 10, "3": 10, "4": 10},
  "merges": [
    {"start_row": 1, "start_col": 1, "end_row": 1, "end_col": 4}
  ],
  "cell_styles": [
    {
      "start_row": 1,
      "start_col": 1,
      "end_row": 1,
      "end_col": 1,
      "style": {
        "alignment": "center",
        "vertical": "center",
        "font_size": 14,
        "bold": true,
        "bgcolor": "FFFF00",
        "border": true
      }
    },
    {
      "start_row": 2,
      "start_col": 1,
      "end_row": 2,
      "end_col": 1,
      "style": {
        "alignment": "center",
        "border": true
      }
    }
  ],
  "row_heights": {"1": 30, "2": 25, "3": 25}
}

### Features

- Convert JSON-based 2D array into Excel file  
- Supports custom column widths  
- Supports merged cells  
- Supports cell styles (alignment, font, background, border)  
- Supports row height settings  

### Requirements

- Python 3.12 or above  
- Dependency: `openpyxl`
