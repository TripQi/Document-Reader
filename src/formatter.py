# 响应格式化器
"""将文档处理结果转换为 Markdown 或 JSON 格式"""

import json
from typing import Any

from .models import DocumentResult, ContentItem, ResponseFormat
from .utils import convert_to_csv


class ResponseFormatter:
    """响应格式化器

    将 DocumentResult 转换为用户请求的输出格式（Markdown 或 JSON）。
    """

    @staticmethod
    def format(result: DocumentResult, response_format: ResponseFormat, include_metadata: bool = True) -> str:
        """根据指定格式转换文档结果

        Args:
            result: 文档处理结果
            response_format: 输出格式
            include_metadata: 是否包含元数据

        Returns:
            str: 格式化后的字符串
        """
        if response_format == ResponseFormat.JSON:
            return ResponseFormatter.to_json(result, include_metadata)
        else:
            return ResponseFormatter.to_markdown(result, include_metadata)

    @staticmethod
    def to_markdown(result: DocumentResult, include_metadata: bool = True) -> str:
        """将结构化数据转换为 Markdown 格式

        Args:
            result: 文档处理结果
            include_metadata: 是否包含元数据

        Returns:
            str: Markdown 格式字符串
        """
        lines: list[str] = []

        # 标题
        title_suffix = " (预览模式)" if result.metadata.preview_mode else ""
        lines.append(f"# 文档: {result.file_name}{title_suffix}")
        lines.append("")

        # 元数据
        if include_metadata:
            lines.append("## 元数据")
            lines.append(f"- 文件类型: {result.file_type}")
            lines.append(f"- 文件大小: {result.metadata.file_size}")

            # PDF 特有元数据
            if result.metadata.page_count is not None:
                lines.append(f"- 页数: {result.metadata.page_count}")
            if result.metadata.author:
                lines.append(f"- 作者: {result.metadata.author}")
            if result.metadata.created_date:
                lines.append(f"- 创建时间: {result.metadata.created_date}")
            if result.metadata.modified_date:
                lines.append(f"- 修改时间: {result.metadata.modified_date}")

            # Excel 特有元数据
            if result.metadata.sheet_name:
                lines.append(f"- 工作表: {result.metadata.sheet_name}")
            if result.metadata.total_rows is not None:
                lines.append(f"- 总行数: {result.metadata.total_rows}")
            if result.metadata.total_columns is not None:
                lines.append(f"- 总列数: {result.metadata.total_columns}")
            if result.metadata.available_sheets:
                sheets = ', '.join(result.metadata.available_sheets)
                lines.append(f"- 可用工作表: {sheets}")

            # 格式提示
            if result.format_hint == "csv":
                lines.append("- 格式: CSV (行数超过50行)")

            # 预览模式提示
            if result.metadata.preview_mode:
                lines.append("- **预览模式**: 仅显示前10行")

            lines.append("")

        # 内容
        lines.append("## 内容")
        lines.append("")

        current_page = None
        table_count = 0

        for item in result.content:
            # PDF 页码标题
            if item.page is not None and item.page != current_page:
                current_page = item.page
                lines.append(f"### 第{current_page}页")
                lines.append("")

            if item.type == "text" and item.text:
                lines.append(item.text)
                lines.append("")

            elif item.type == "table" and item.data:
                table_count += 1
                table_title = f"表格{table_count}"
                if item.page is not None:
                    table_title += f" (第{item.page}页)"
                lines.append(f"### {table_title}")
                lines.append("")

                # 根据格式提示选择输出方式
                if result.format_hint == "csv":
                    lines.append("```csv")
                    lines.append(convert_to_csv(item.data))
                    lines.append("```")
                else:
                    lines.append(ResponseFormatter._format_table_markdown(item.data))

                lines.append("")

            elif item.type == "image":
                lines.append("[image]")
                lines.append("")

        # 预览模式提示
        if result.metadata.preview_mode and result.metadata.total_rows:
            lines.append(f"**提示**: 这是预览模式，仅显示前10行。总共有{result.metadata.total_rows}行数据。")
            lines.append("")

        return '\n'.join(lines)

    @staticmethod
    def to_json(result: DocumentResult, include_metadata: bool = True) -> str:
        """将结构化数据转换为 JSON 格式

        Args:
            result: 文档处理结果
            include_metadata: 是否包含元数据

        Returns:
            str: JSON 格式字符串
        """
        output: dict[str, Any] = {
            "file_name": result.file_name,
            "file_type": result.file_type,
        }

        if include_metadata:
            metadata_dict = result.metadata.model_dump(exclude_none=True)
            output["metadata"] = metadata_dict

        # 转换内容项
        content_list = []
        for item in result.content:
            item_dict = item.model_dump(exclude_none=True)
            content_list.append(item_dict)

        output["content"] = content_list

        return json.dumps(output, ensure_ascii=False, indent=2)

    @staticmethod
    def _format_table_markdown(table_data: list[list[Any]]) -> str:
        """格式化表格为 Markdown 表格

        Args:
            table_data: 二维表格数据，第一行为表头

        Returns:
            str: Markdown 表格字符串
        """
        if not table_data:
            return ""

        lines: list[str] = []

        # 计算每列的最大宽度
        col_widths: list[int] = []
        for row in table_data:
            for i, cell in enumerate(row):
                cell_str = str(cell) if cell is not None else ""
                if i >= len(col_widths):
                    col_widths.append(len(cell_str))
                else:
                    col_widths[i] = max(col_widths[i], len(cell_str))

        # 确保最小宽度为3（用于分隔符）
        col_widths = [max(w, 3) for w in col_widths]

        # 表头
        header = table_data[0] if table_data else []
        header_cells = [
            str(cell).ljust(col_widths[i]) if cell is not None else "".ljust(col_widths[i])
            for i, cell in enumerate(header)
        ]
        lines.append("| " + " | ".join(header_cells) + " |")

        # 分隔符
        separator = ["-" * w for w in col_widths]
        lines.append("| " + " | ".join(separator) + " |")

        # 数据行
        for row in table_data[1:]:
            cells = []
            for i, cell in enumerate(row):
                cell_str = str(cell) if cell is not None else ""
                if i < len(col_widths):
                    cells.append(cell_str.ljust(col_widths[i]))
                else:
                    cells.append(cell_str)
            lines.append("| " + " | ".join(cells) + " |")

        return '\n'.join(lines)
