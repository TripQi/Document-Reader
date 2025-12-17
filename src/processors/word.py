# Word 文档处理器
"""处理 Word (DOC/DOCX) 文档"""

import os
import sys
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
    - DOC: Windows 上使用 pywin32 COM 接口调用 Microsoft Word
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

        尝试多种方法:
        1. Windows COM 接口 (需要 Microsoft Word/WPS)
        2. olefile 纯 Python 解析 (备选方案)
        """
        file_path = params.file_path

        # 首先尝试使用 olefile 纯 Python 方案
        try:
            return await self._process_doc_with_olefile(params)
        except Exception:
            pass

        # Windows 上尝试 COM 接口
        if sys.platform == 'win32':
            try:
                return await self._process_doc_with_com(params)
            except Exception:
                pass

        raise ValueError(
            f"错误: 无法处理 DOC 文件 '{file_path}'。"
            f"请安装 Microsoft Word 或将文件转换为 DOCX 格式。"
        )

    async def _process_doc_with_olefile(self, params: ReadDocumentInput) -> DocumentResult:
        """使用 olefile 解析 DOC 文件

        纯 Python 方案，不依赖外部软件。
        """
        import olefile
        import struct

        file_path = params.file_path

        if not olefile.isOleFile(file_path):
            raise ValueError("不是有效的 OLE 文件")

        ole = olefile.OleFileIO(file_path)

        try:
            # 检查是否是 Word 文档
            if not ole.exists('WordDocument'):
                raise ValueError("不是有效的 Word 文档")

            # 提取文本
            text = self._extract_doc_text_from_ole(ole)

            # 构建元数据
            metadata = self._extract_doc_ole_metadata(ole, file_path)

            # 构建内容
            content = []
            if text and text.strip():
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

        finally:
            ole.close()

    def _extract_doc_text_from_ole(self, ole: Any) -> str:
        """从 OLE 文件提取文本内容"""
        import struct

        text_parts = []

        # 尝试从 WordDocument 流提取
        try:
            word_doc = ole.openstream('WordDocument')
            data = word_doc.read()

            # 读取 FIB (File Information Block)
            # 偏移 0x18 处是 fcMin (文本开始位置)
            # 偏移 0x1C 处是 fcMac (文本结束位置)
            if len(data) >= 0x20:
                # 尝试解析文本位置
                # DOC 格式复杂，这里使用简化方法

                # 检查是否有 0Table 或 1Table 流
                table_stream = None
                if ole.exists('0Table'):
                    table_stream = ole.openstream('0Table')
                elif ole.exists('1Table'):
                    table_stream = ole.openstream('1Table')

                # 尝试从数据流中提取可读文本
                text = self._extract_readable_text(data)
                if text:
                    text_parts.append(text)

        except Exception:
            pass

        # 如果上面方法失败，尝试直接扫描所有流
        if not text_parts:
            for stream_path in ole.listdir():
                try:
                    stream_name = '/'.join(stream_path)
                    stream = ole.openstream(stream_path)
                    data = stream.read()

                    # 尝试提取文本
                    text = self._extract_readable_text(data)
                    if text and len(text) > 50:  # 只保留有意义的文本
                        text_parts.append(text)
                except Exception:
                    pass

        return '\n\n'.join(text_parts)

    def _extract_readable_text(self, data: bytes) -> str:
        """从二进制数据中提取可读文本"""
        text_chars = []
        i = 0

        while i < len(data):
            # 尝试 UTF-16LE 解码 (Word 常用编码)
            if i + 1 < len(data):
                char_code = data[i] | (data[i + 1] << 8)

                # 检查是否是可打印字符
                if 0x20 <= char_code <= 0x7E:  # ASCII 可打印
                    text_chars.append(chr(char_code))
                    i += 2
                    continue
                elif 0x4E00 <= char_code <= 0x9FFF:  # CJK 统一汉字
                    text_chars.append(chr(char_code))
                    i += 2
                    continue
                elif char_code == 0x000D or char_code == 0x000A:  # 换行
                    text_chars.append('\n')
                    i += 2
                    continue
                elif char_code == 0x0009:  # Tab
                    text_chars.append(' ')
                    i += 2
                    continue

            # 尝试单字节
            if 0x20 <= data[i] <= 0x7E:
                text_chars.append(chr(data[i]))
            elif data[i] == 0x0D or data[i] == 0x0A:
                text_chars.append('\n')

            i += 1

        # 清理文本
        text = ''.join(text_chars)
        # 移除连续的空白
        lines = [line.strip() for line in text.split('\n')]
        lines = [line for line in lines if line]

        return '\n'.join(lines)

    def _extract_doc_ole_metadata(self, ole: Any, file_path: str) -> DocumentMetadata:
        """从 OLE 文件提取元数据"""
        file_size = os.path.getsize(file_path)

        author = None
        created_date = None
        modified_date = None

        try:
            meta = ole.get_metadata()
            if meta:
                author = meta.author
                if meta.create_time:
                    created_date = meta.create_time.strftime("%Y-%m-%d")
                if meta.last_saved_time:
                    modified_date = meta.last_saved_time.strftime("%Y-%m-%d")
        except Exception:
            pass

        return DocumentMetadata(
            file_name=os.path.basename(file_path),
            file_type="Word",
            file_size=format_file_size(file_size),
            author=author,
            created_date=created_date,
            modified_date=modified_date,
        )

    async def _process_doc_with_com(self, params: ReadDocumentInput) -> DocumentResult:
        """使用 pywin32 COM 接口处理 DOC 文件

        通过 Microsoft Word COM 接口读取 DOC 文件内容。
        需要系统安装 Microsoft Word。
        """
        file_path = params.file_path
        abs_path = os.path.abspath(file_path)

        try:
            import win32com.client
            import pythoncom

            # 初始化 COM
            pythoncom.CoInitialize()

            try:
                # 创建 Word 应用实例
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False  # 不显示 Word 窗口

                try:
                    # 打开文档
                    doc = word.Documents.Open(abs_path, ReadOnly=True)

                    try:
                        # 提取元数据
                        metadata = self._extract_doc_com_metadata(doc, file_path)

                        # 提取内容
                        content = self._extract_doc_com_content(doc)

                        return DocumentResult(
                            file_name=os.path.basename(file_path),
                            file_type="Word",
                            metadata=metadata,
                            content=content,
                            format_hint="text"
                        )

                    finally:
                        # 关闭文档
                        doc.Close(False)

                finally:
                    # 退出 Word
                    word.Quit()

            finally:
                # 释放 COM
                pythoncom.CoUninitialize()

        except ImportError:
            raise ValueError(
                f"错误: 缺少 pywin32 库。请运行: pip install pywin32"
            )
        except Exception as e:
            error_str = str(e).lower()
            if "word" in error_str or "dispatch" in error_str:
                raise ValueError(
                    f"错误: 无法启动 Microsoft Word。请确保已安装 Microsoft Word。"
                )
            raise ValueError(
                f"错误: 处理 DOC 文件时发生错误: {str(e)}"
            )

    def _extract_doc_com_metadata(self, doc: Any, file_path: str) -> DocumentMetadata:
        """从 COM 文档对象提取元数据"""
        file_size = os.path.getsize(file_path)

        # 尝试获取内置属性
        author = None
        created_date = None
        modified_date = None

        try:
            builtin_props = doc.BuiltInDocumentProperties
            try:
                author = builtin_props("Author").Value
            except Exception:
                pass
            try:
                created = builtin_props("Creation Date").Value
                if created:
                    created_date = created.strftime("%Y-%m-%d")
            except Exception:
                pass
            try:
                modified = builtin_props("Last Save Time").Value
                if modified:
                    modified_date = modified.strftime("%Y-%m-%d")
            except Exception:
                pass
        except Exception:
            pass

        return DocumentMetadata(
            file_name=os.path.basename(file_path),
            file_type="Word",
            file_size=format_file_size(file_size),
            author=author,
            created_date=created_date,
            modified_date=modified_date,
        )

    def _extract_doc_com_content(self, doc: Any) -> list[ContentItem]:
        """从 COM 文档对象提取内容"""
        content: list[ContentItem] = []

        # 提取全文
        try:
            full_text = doc.Content.Text
            if full_text and full_text.strip():
                # 按段落分割
                paragraphs = full_text.split('\r')
                current_text = []

                for para in paragraphs:
                    para = para.strip()
                    if para:
                        current_text.append(para)

                if current_text:
                    content.append(ContentItem(
                        type="text",
                        text='\n\n'.join(current_text)
                    ))
        except Exception:
            pass

        # 提取表格
        try:
            tables_count = doc.Tables.Count
            for i in range(1, tables_count + 1):
                table = doc.Tables(i)
                table_data = self._extract_com_table_data(table)
                if table_data:
                    content.append(ContentItem(
                        type="table",
                        table_index=i - 1,
                        data=table_data
                    ))
        except Exception:
            pass

        # 检查是否有图片
        try:
            if doc.InlineShapes.Count > 0 or doc.Shapes.Count > 0:
                content.append(ContentItem(
                    type="image",
                    text="[image]"
                ))
        except Exception:
            pass

        return content

    def _extract_com_table_data(self, table: Any) -> list[list[Any]]:
        """从 COM 表格对象提取数据"""
        data: list[list[Any]] = []

        try:
            rows_count = table.Rows.Count
            cols_count = table.Columns.Count

            for row_idx in range(1, rows_count + 1):
                row_data = []
                for col_idx in range(1, cols_count + 1):
                    try:
                        cell = table.Cell(row_idx, col_idx)
                        cell_text = cell.Range.Text
                        # 移除单元格结束符
                        cell_text = cell_text.replace('\r\x07', '').strip()
                        row_data.append(cell_text)
                    except Exception:
                        row_data.append("")
                data.append(row_data)
        except Exception:
            pass

        return data
