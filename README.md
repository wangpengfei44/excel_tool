# excel_tool

**Author:** wangpengfei44  
**Version:** 0.0.1  
**Type:** Tool Plugin  
**Last Updated:** 2025-06-18

---

## Description

`excel_tool` is a lightweight plugin that converts 2D arrays (provided as JSON strings) into Excel (.xlsx) files.  
It supports setting column widths, merging cells, cell styles (including cell range control), and row heights.

## Usage

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
    - `wrap_text`: boolean to wrap text (default: true)
- `row_heights` (optional): row height settings; keys are row numbers (as strings), values are height values  

## Example Input JSON

```json
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

## Features
- Convert JSON-based 2D array into Excel file

- Supports custom column widths

- Supports merged cells

- Supports cell styles (alignment, font, background, border)

- Supports row height settings

## Requirements
- Python 3.12 or above

- Dependency: openpyxl