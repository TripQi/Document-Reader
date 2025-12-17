# PDF 文档处理器
"""使用 pdfplumber 处理 PDF 文档"""

import os
from datetime import datetime
from typing import Any, Optional

import pdfplumber

from .base import DocumentProcessor
from ..models import (
    ReadDocumentInput,
    DocumentResult,
    DocumentMetadata,
    ContentItem,
)
from ..utils import (
    validate_file,
    format_file_size,
    parse_page_range,
    handle_file_error,
)


class PdfProcessor(DocumentProcessor):
    """PDF 文档处理器

    使用 pdfplumber 库提取 PDF 文档的文本和表格内容。
    """

    @property
    def supported_extensions(self) -> list[str]:
        """返回支持的文件扩展名列表"""
        return ['.pdf']

    def supports_extension(self, ext: str) -> bool:
        """检查是否支持该文件扩展名"""
        return ext.lower() in self.supported_extensions

    async def process(self, params: ReadDocumentInput) -> DocumentResult:
        """处理 PDF 文档

        Args:
            params: 读取文档的输入参数

        Returns:
            DocumentResult: 包含文件名、类型、元数据和内容的结构化结果
        """
        file_path = params.file_path

        # 验证文件
        validate_file(file_path)

        try:
            with pdfplumber.open(file_path) as pdf:
                # 获取元数据
                metadata = self._extract_metadata(pdf, file_path)

                # 确定要处理的页码
                total_pages = len(pdf.pages)
                if params.page_range:
                    pages_to_process = parse_page_range(params.page_range, total_pages)
                else:
                    pages_to_process = list(range(1, total_pages + 1))

                # 提取内容
                content = self._extract_content(
                    pdf,
                    pages_to_process,
                    params.extract_tables
                )

                return DocumentResult(
                    file_name=os.path.basename(file_path),
                    file_type="PDF",
                    metadata=metadata,
                    content=content,
                    format_hint="text"
                )

        except Exception as e:
            error_msg = handle_file_error(e, file_path)
            raise ValueError(error_msg) from e

    def _extract_metadata(
        self,
        pdf: pdfplumber.PDF,
        file_path: str
    ) -> DocumentMetadata:
        """提取 PDF 元数据

        Args:
            pdf: pdfplumber PDF 对象
            file_path: 文件路径

        Returns:
            DocumentMetadata: 文档元数据
        """
        file_size = os.path.getsize(file_path)
        pdf_metadata = pdf.metadata or {}

        # 解析日期
        created_date = self._parse_pdf_date(pdf_metadata.get('CreationDate'))
        modified_date = self._parse_pdf_date(pdf_metadata.get('ModDate'))

        return DocumentMetadata(
            file_name=os.path.basename(file_path),
            file_type="PDF",
            file_size=format_file_size(file_size),
            page_count=len(pdf.pages),
            author=pdf_metadata.get('Author'),
            created_date=created_date,
            modified_date=modified_date,
        )

    def _parse_pdf_date(self, date_str: Optional[str]) -> Optional[str]:
        """解析 PDF 日期格式

        PDF 日期格式通常为: D:YYYYMMDDHHmmSS

        Args:
            date_str: PDF 日期字符串

        Returns:
            Optional[str]: 格式化的日期字符串
        """
        if not date_str:
            return None

        try:
            # 移除 'D:' 前缀
            if date_str.startswith('D:'):
                date_str = date_str[2:]

            # 提取基本日期部分 (YYYYMMDD)
            if len(date_str) >= 8:
                year = date_str[0:4]
                month = date_str[4:6]
                day = date_str[6:8]
                return f"{year}-{month}-{day}"
        except (ValueError, IndexError):
            pass

        return None

    def _extract_content(
        self,
        pdf: pdfplumber.PDF,
        pages_to_process: list[int],
        extract_tables: bool
    ) -> list[ContentItem]:
        """提取 PDF 内容

        Args:
            pdf: pdfplumber PDF 对象
            pages_to_process: 要处理的页码列表（从1开始）
            extract_tables: 是否提取表格

        Returns:
            list[ContentItem]: 内容项列表
        """
        content: list[ContentItem] = []

        for page_num in pages_to_process:
            # pdfplumber 页码从 0 开始
            page = pdf.pages[page_num - 1]

            # 提取表格
            tables_on_page: list[list[list[Any]]] = []
            if extract_tables:
                tables = page.extract_tables()
                if tables:
                    tables_on_page = tables

            # 提取文本
            text = page.extract_text()

            # 如果有文本，添加文本内容项
            if text and text.strip():
                # 检查是否有图片（简单检测）
                has_images = bool(page.images)
                processed_text = text.strip()

                # 如果页面有图片，在文本中添加标记
                if has_images:
                    processed_text += "\n\n[image]"

                content.append(ContentItem(
                    type="text",
                    page=page_num,
                    text=processed_text
                ))

            # 添加表格内容项
            for table_idx, table_data in enumerate(tables_on_page):
                if table_data:
                    # 清理表格数据
                    cleaned_table = self._clean_table_data(table_data)
                    if cleaned_table:
                        content.append(ContentItem(
                            type="table",
                            page=page_num,
                            table_index=table_idx,
                            data=cleaned_table
                        ))

        return content

    def _clean_table_data(self, table_data: list[list[Any]]) -> list[list[Any]]:
        """清理表格数据

        Args:
            table_data: 原始表格数据

        Returns:
            list[list[Any]]: 清理后的表格数据
        """
        cleaned = []
        for row in table_data:
            if row:
                cleaned_row = []
                for cell in row:
                    if cell is None:
                        cleaned_row.append("")
                    else:
                        # 清理单元格内容
                        cell_str = str(cell).strip()
                        # 替换换行符
                        cell_str = cell_str.replace('\n', ' ')
                        cleaned_row.append(cell_str)
                cleaned.append(cleaned_row)

        return cleaned
