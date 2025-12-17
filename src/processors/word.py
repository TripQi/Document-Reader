# Word 文档处理器
"""处理 Word (DOC/DOCX) 文档"""

import os
import subprocess
from typing import Any, Optional

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
    handle_file_error,
)


class WordProcessor(DocumentProcessor):
    """Word 文档处理器

    支持 DOC 和 DOCX 格式。
    - DOCX: 使用 python-docx 库
    - DOC: 尝试使用 antiword 或 textract
    """

    @property
    def supported_extensions(self) -> list[str]:
        """返回支持的文件扩展名列表"""
        return ['.doc', '.docx']

    def supports_extension(self, ext: str) -> bool:
        """检查是否支持该文件扩展名"""
        return ext.lower() in self.supported_extensions

    async def process(self, params: ReadDocumentInput) -> DocumentResult:
        """处理 Word 文档

        Args:
            params: 读取文档的输入参数

        Returns:
            DocumentResult: 包含文件名、类型、元数据和内容的结构化结果
        """
        file_path = params.file_path

        # 验证文件
        validate_file(file_path)

        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == '.docx':
                return await self._process_docx(params)
            elif ext == '.doc':
                return await self._process_doc(params)
            else:
                raise ValueError(f"不支持的文件格式: {ext}")

        except Exception as e:
            error_msg = handle_file_error(e, file_path)
            raise ValueError(error_msg) from e

    async def _process_docx(self, params: ReadDocumentInput) -> DocumentResult:
        """处理 DOCX 文件"""
        from docx import Document
        from docx.opc.exceptions import PackageNotFoundError

        file_path = params.file_path

        try:
            doc = Document(file_path)
        except PackageNotFoundError:
            raise ValueError(
                f"错误: 文件 '{file_path}' 可能已损坏或格式无效，无法解析。"
            )

        # 提取元数据
        metadata = self._extract_docx_metadata(doc, file_path)

        # 提取内容
        content = self._extract_docx_content(doc, params.include_images_info)

        return DocumentResult(
            file_name=os.path.basename(file_path),
            file_type="Word",
            metadata=metadata,
            content=content,
            format_hint="text"
        )

    def _extract_docx_metadata(self, doc: Any, file_path: str) -> DocumentMetadata:
        """提取 DOCX 元数据"""
        file_size = os.path.getsize(file_path)

        # 获取核心属性
        core_props = doc.core_properties

        # 格式化日期
        created_date = None
        modified_date = None

        if core_props.created:
            created_date = core_props.created.strftime("%Y-%m-%d")
        if core_props.modified:
            modified_date = core_props.modified.strftime("%Y-%m-%d")

        return DocumentMetadata(
            file_name=os.path.basename(file_path),
            file_type="Word",
            file_size=format_file_size(file_size),
            author=core_props.author,
            created_date=created_date,
            modified_date=modified_date,
        )

    def _extract_docx_content(
        self,
        doc: Any,
        include_images_info: bool
    ) -> list[ContentItem]:
        """提取 DOCX 内容"""
        from docx.document import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        content: list[ContentItem] = []
        image_count = 0

        # 遍历文档元素
        for element in doc.element.body:
            # 处理段落
            if element.tag.endswith('p'):
                para = Paragraph(element, doc)
                text = para.text.strip()

                # 检查段落中是否有图片
                has_image = self._paragraph_has_image(element)

                if text:
                    content.append(ContentItem(
                        type="text",
                        text=text
                    ))

                if has_image:
                    image_count += 1
                    content.append(ContentItem(
                        type="image",
                        text="[image]"
                    ))

            # 处理表格
            elif element.tag.endswith('tbl'):
                table = Table(element, doc)
                table_data = self._extract_table_data(table)

                if table_data:
                    content.append(ContentItem(
                        type="table",
                        data=table_data
                    ))

        return content

    def _paragraph_has_image(self, para_element: Any) -> bool:
        """检查段落是否包含图片"""
        # 查找 drawing 或 pict 元素
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        }

        # 检查 drawing 元素
        drawings = para_element.findall('.//w:drawing', namespaces)
        if drawings:
            return True

        # 检查 pict 元素（旧版 Word）
        picts = para_element.findall('.//w:pict', namespaces)
        if picts:
            return True

        return False

    def _extract_table_data(self, table: Any) -> list[list[Any]]:
        """提取表格数据"""
        data: list[list[Any]] = []

        for row in table.rows:
            row_data = []
            for cell in row.cells:
                # 获取单元格文本
                cell_text = cell.text.strip()
                row_data.append(cell_text)
            data.append(row_data)

        return data

    async def _process_doc(self, params: ReadDocumentInput) -> DocumentResult:
        """处理 DOC 文件

        尝试使用 antiword 提取文本。
        如果 antiword 不可用，返回错误提示。
        """
        file_path = params.file_path

        # 尝试使用 antiword
        text = await self._extract_doc_with_antiword(file_path)

        if text is None:
            # antiword 不可用，尝试其他方法或返回错误
            raise ValueError(
                f"错误: 无法处理 DOC 文件 '{file_path}'。"
                f"请安装 antiword 或将文件转换为 DOCX 格式。"
            )

        # 构建元数据
        metadata = DocumentMetadata(
            file_name=os.path.basename(file_path),
            file_type="Word",
            file_size=format_file_size(os.path.getsize(file_path)),
        )

        # 构建内容
        content = []
        if text.strip():
            content.append(ContentItem(
                type="text",
                text=text.strip()
            ))

        return DocumentResult(
            file_name=os.path.basename(file_path),
            file_type="Word",
            metadata=metadata,
            content=content,
            format_hint="text"
        )

    async def _extract_doc_with_antiword(self, file_path: str) -> Optional[str]:
        """使用 antiword 提取 DOC 文件文本

        Args:
            file_path: DOC 文件路径

        Returns:
            Optional[str]: 提取的文本，如果 antiword 不可用则返回 None
        """
        try:
            result = subprocess.run(
                ['antiword', file_path],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0:
                return result.stdout

            return None

        except FileNotFoundError:
            # antiword 未安装
            return None
        except subprocess.TimeoutExpired:
            return None
        except Exception:
            return None
