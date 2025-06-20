import json
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from io import BytesIO


class ArrayToExcelTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        """
        将传入的 JSON 格式的二维数组数据转换成 Excel 文件，并支持设置列宽、合并单元格、单元格样式（支持区域样式）及行高。

        参数：
        - tool_parameters: 包含 "data_json" 字段的字典，"data_json" 是一个 JSON 字符串，
          该字符串包含以下字段：
            * data: 二维数组，Excel 中的表格数据
            * col_widths: dict，列号（字符串）到列宽的映射，可选
            * merges: list，合并单元格的范围列表，每个元素包含
                start_row, start_col, end_row, end_col 四个字段，表示合并区域
            * cell_styles: list，单元格样式配置列表，每个元素包含：
                - start_row, start_col, end_row（可选，默认等于start_row），end_col（可选，默认等于start_col）
                - style: dict，支持字段：
                  - alignment: 水平对齐方式，"center" / "left" / "right"
                  - vertical: 垂直对齐方式，"center" / "top" / "bottom"（默认"center"）
                  - font_size: 字体大小，整数
                  - bold: 是否加粗，布尔值
                  - bgcolor: 背景色，十六进制 RGB 字符串（如 "FFFF00"）
                  - border: 是否添加细边框，布尔值，可选
                  - wrap_text: 是否自动换行
            * row_heights: dict，行号（字符串）到行高（数字）的映射，可选

        返回：
        - 生成器，yield ToolInvokeMessage 对象，成功时返回 Excel 文件的二进制数据，
          失败时返回文本错误信息
        """

        data_json = tool_parameters.get("data_json", "")
        if not data_json:
            yield self.create_text_message("参数错误：缺少 'data_json'")
            return

        try:
            if isinstance(data_json, str):
                params = json.loads(data_json)
                if isinstance(params, str):
                    params = json.loads(params)
            else:
                params = data_json
        except Exception as e:
            yield self.create_text_message(f"参数解析失败: {e}")
            return

        data = params.get("data")
        col_widths = params.get("col_widths", {})
        merges = params.get("merges", [])
        cell_styles = params.get("cell_styles", [])  # 变成列表了
        row_heights = params.get("row_heights", {})

        if not data or not isinstance(data, list):
            yield self.create_text_message("参数错误：'data' 应为二维数组")
            return

        wb = Workbook()
        ws = wb.active

        # 细边框样式定义
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for r, row in enumerate(data, start=1):
            if not isinstance(row, list):
                yield self.create_text_message(f"参数错误：第 {r} 行不是数组")
                return
            for c, val in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.alignment = Alignment(vertical='center', wrap_text=True)  # 默认垂直居中且自动换行
                cell.border = thin_border  # 默认细边框

        # 处理样式：支持区域，cell_styles为列表，元素包含 start_row, start_col, end_row, end_col, style(dict)
        for style_range in cell_styles:
            try:
                start_row = style_range["start_row"]
                start_col = style_range["start_col"]
                end_row = style_range.get("end_row", start_row)
                end_col = style_range.get("end_col", start_col)
                style_cfg = style_range["style"]
            except Exception as e:
                yield self.create_text_message(f"样式参数错误: {e}")
                return

            for r in range(start_row, end_row + 1):
                for c in range(start_col, end_col + 1):
                    cell = ws.cell(row=r, column=c)

                    # 对齐方式设置，默认垂直居中，新增 wrap_text 控制自动换行，默认 True
                    align_val = style_cfg.get("alignment")
                    vertical_val = style_cfg.get("vertical", "center")
                    wrap_text_val = style_cfg.get("wrap_text", True)
                    if align_val:
                        cell.alignment = Alignment(horizontal=align_val, vertical=vertical_val,
                                                   wrap_text=wrap_text_val)
                    else:
                        cell.alignment = Alignment(vertical=vertical_val, wrap_text=wrap_text_val)

                    # 字体设置
                    font_args = {}
                    if "font_size" in style_cfg:
                        font_args["size"] = style_cfg["font_size"]
                    if style_cfg.get("bold"):
                        font_args["bold"] = True
                    if font_args:
                        cell.font = Font(**font_args)

                    # 背景色填充
                    bgcolor = style_cfg.get("bgcolor")
                    if bgcolor:
                        cell.fill = PatternFill(fill_type="solid", fgColor=bgcolor)

                    # 边框控制，默认已设细边框，如果设置了 border 为 False，则清除边框
                    if "border" in style_cfg:
                        if style_cfg.get("border"):
                            cell.border = thin_border
                        else:
                            cell.border = Border()  # 无边框

        # 设置行高，key为字符串行号，需要转换为 int
        try:
            for row_str, height in row_heights.items():
                row_idx = int(row_str)
                ws.row_dimensions[row_idx].height = float(height)
        except Exception as e:
            yield self.create_text_message(f"设置行高失败: {e}")
            return

        # 设置列宽，支持1-26列（A-Z）
        try:
            for col_str, width in col_widths.items():
                col_idx = int(col_str)
                if 1 <= col_idx <= 26:
                    col_letter = chr(64 + col_idx)
                    ws.column_dimensions[col_letter].width = float(width)
                else:
                    yield self.create_text_message(f"不支持列号超过26: {col_idx}")
                    return
        except Exception as e:
            yield self.create_text_message(f"设置列宽失败: {e}")
            return

        # 合并单元格，格式与之前一致
        try:
            for merge in merges:
                start_row = merge["start_row"]
                start_col = merge["start_col"]
                end_row = merge["end_row"]
                end_col = merge["end_col"]
                start_cell = ws.cell(row=start_row, column=start_col).coordinate
                end_cell = ws.cell(row=end_row, column=end_col).coordinate
                ws.merge_cells(f"{start_cell}:{end_cell}")
        except Exception as e:
            yield self.create_text_message(f"合并单元格失败: {e}")
            return

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        yield self.create_blob_message(
            blob=output.getvalue(),
            meta={
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "output_filename": "output.xlsx",
            },
        )
