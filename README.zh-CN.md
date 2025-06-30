
# excel_tool（中文说明）

**作者:** wangpengfei44  
**版本:** 0.0.1  
**类型:** 工具插件  
**最后更新:** 2025-06-18

---

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
    - `wrap_text`: 是否自动换行（默认为 true）
- `row_heights`（可选）：行高配置，字典格式，键为行号（字符串），值为行高数值  

### 示例输入 JSON

```json
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

## 功能
- 将 JSON 格式的二维数组转换为 Excel 文件

- 支持自定义列宽

- 支持合并单元格

- 支持单元格样式（对齐、字体、背景色、边框）

- 支持设置行高

## 运行环境
- Python 3.12 及以上版本

- 依赖库：openpyxl

