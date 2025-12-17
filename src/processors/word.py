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
        1. Windows COM 接口 (需要 Microsoft Word/WPS) - 最佳效果
        2. olefile 纯 Python 解析 (备选方案) - 可能有乱码
        """
        file_path = params.file_path

        # Windows 上优先尝试 COM 接口（效果最好）
        if sys.platform == 'win32':
            try:
                return await self._process_doc_with_com(params)
            except Exception:
                pass

        # 备选方案：使用 olefile 纯 Python 解析
        try:
            result = await self._process_doc_with_olefile(params)
            # 检查提取的文本质量
            if result.content:
                text = result.content[0].text or ""
                # 如果中文字符比例太低，可能解析有问题
                chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
                if len(text) > 100 and chinese_chars < len(text) * 0.05:
                    # 添加警告信息
                    warning = (
                        "\n\n---\n"
                        "**注意**: DOC 文件解析可能不完整。"
                        "建议安装 Microsoft Word 或将文件转换为 DOCX/PDF 格式以获得更好的效果。"
                    )
                    result.content[0].text = text + warning
            return result
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

        file_path = params.file_path

        if not olefile.isOleFile(file_path):
            raise ValueError("不是有效的 OLE 文件")

        ole = olefile.OleFileIO(file_path)

        try:
            # 检查是否是 Word 文档
            if not ole.exists('WordDocument'):
                raise ValueError("不是有效的 Word 文档")

            # 获取代码页信息
            codepage = self._get_doc_codepage(ole)

            # 提取文本
            text = self._extract_doc_text_from_ole(ole, codepage)

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

    def _get_doc_codepage(self, ole: Any) -> str:
        """获取文档的代码页/编码"""
        try:
            meta = ole.get_metadata()
            if meta and meta.codepage:
                # 常见代码页映射
                codepage_map = {
                    936: 'gbk',      # 简体中文
                    950: 'big5',     # 繁体中文
                    932: 'shift_jis', # 日文
                    949: 'euc-kr',   # 韩文
                    1252: 'cp1252',  # 西欧
                    1251: 'cp1251',  # 俄文
                    65001: 'utf-8',  # UTF-8
                }
                return codepage_map.get(meta.codepage, f'cp{meta.codepage}')
        except Exception:
            pass
        return 'gbk'  # 默认使用 GBK（中文 Windows 常用）

    def _extract_doc_text_from_ole(self, ole: Any, codepage: str = 'gbk') -> str:
        """从 OLE 文件提取文本内容

        使用正确的 DOC 格式解析方法，根据代码页正确解码文本。
        """
        import struct

        try:
            word_doc = ole.openstream('WordDocument')
            data = word_doc.read()

            if len(data) < 0x200:
                return self._fallback_extract_text(ole, codepage)

            # 解析 FIB (File Information Block)
            wIdent = struct.unpack_from('<H', data, 0x00)[0]
            if wIdent != 0xA5EC:
                return self._fallback_extract_text(ole, codepage)

            # 获取标志位
            flags = struct.unpack_from('<H', data, 0x0A)[0]
            use_1table = bool(flags & 0x0200)
            is_complex = bool(flags & 0x0004)  # 复杂文档标志

            # 获取文本位置信息
            fcMin = struct.unpack_from('<I', data, 0x18)[0]
            ccpText = struct.unpack_from('<I', data, 0x4C)[0]

            # 如果是简单文档，直接从 fcMin 位置读取
            if not is_complex and fcMin > 0 and ccpText > 0:
                text_end = fcMin + ccpText
                if text_end <= len(data):
                    text_data = data[fcMin:text_end]
                    # 使用检测到的代码页解码
                    try:
                        text = text_data.decode(codepage, errors='ignore')
                        cleaned = self._clean_doc_text(text)
                        if cleaned and len(cleaned) > 50:
                            return cleaned
                    except Exception:
                        pass

            # 复杂文档需要解析 piece table
            return self._extract_from_piece_table(ole, data, use_1table, codepage)

        except Exception:
            return self._fallback_extract_text(ole, codepage)

    def _extract_from_piece_table(self, ole: Any, word_data: bytes, use_1table: bool, codepage: str = 'gbk') -> str:
        """从 piece table 提取文本（支持混合编码）"""
        import struct

        try:
            # 选择正确的 Table 流
            table_name = '1Table' if use_1table else '0Table'
            if not ole.exists(table_name):
                table_name = '0Table' if use_1table else '1Table'
                if not ole.exists(table_name):
                    return self._fallback_extract_text(ole, codepage)

            table_stream = ole.openstream(table_name)
            table_data = table_stream.read()

            # 获取 clx 偏移和大小
            if len(word_data) >= 0x1AA:
                fcClx = struct.unpack_from('<I', word_data, 0x1A2)[0]
                lcbClx = struct.unpack_from('<I', word_data, 0x1A6)[0]

                if fcClx > 0 and lcbClx > 0 and fcClx + lcbClx <= len(table_data):
                    clx_data = table_data[fcClx:fcClx + lcbClx]
                    return self._parse_clx(clx_data, word_data, codepage)

        except Exception:
            pass

        return self._fallback_extract_text(ole, codepage)

    def _parse_clx(self, clx_data: bytes, word_data: bytes, codepage: str = 'gbk') -> str:
        """解析 CLX 结构提取文本"""
        import struct

        text_parts = []
        offset = 0

        try:
            while offset < len(clx_data):
                clxt = clx_data[offset]

                if clxt == 0x01:  # Grpprl (格式信息，跳过)
                    if offset + 2 >= len(clx_data):
                        break
                    cb = struct.unpack_from('<H', clx_data, offset + 1)[0]
                    offset += 3 + cb

                elif clxt == 0x02:  # Pcdt (piece table)
                    if offset + 4 >= len(clx_data):
                        break
                    lcb = struct.unpack_from('<I', clx_data, offset + 1)[0]
                    pcd_data = clx_data[offset + 5:offset + 5 + lcb]

                    # 解析 piece descriptors
                    text = self._extract_from_pcd(pcd_data, word_data, codepage)
                    if text:
                        text_parts.append(text)
                    break

                else:
                    break

        except Exception:
            pass

        if text_parts:
            return '\n'.join(text_parts)

        return self._fallback_extract_text_from_data(word_data, codepage)

    def _extract_from_pcd(self, pcd_data: bytes, word_data: bytes, codepage: str = 'gbk') -> str:
        """从 piece descriptor 提取文本"""
        # 简化处理：直接从 word_data 提取可读文本
        try:
            return self._fallback_extract_text_from_data(word_data, codepage)
        except Exception:
            pass
        return ''

    def _fallback_extract_text(self, ole: Any, codepage: str = 'gbk') -> str:
        """备用方法：扫描所有流提取文本"""
        text_parts = []

        for stream_path in ole.listdir():
            try:
                stream = ole.openstream(stream_path)
                data = stream.read()
                text = self._fallback_extract_text_from_data(data, codepage)
                if text and len(text) > 100:
                    text_parts.append(text)
            except Exception:
                pass

        return '\n\n'.join(text_parts)

    def _fallback_extract_text_from_data(self, data: bytes, codepage: str = 'gbk') -> str:
        """从二进制数据提取可读文本（改进版）"""
        results = []

        # 方法1: 使用指定的代码页解码
        try:
            text = data.decode(codepage, errors='ignore')
            cleaned = self._clean_doc_text(text)
            if cleaned:
                results.append((codepage, cleaned, len(cleaned)))
        except Exception:
            pass

        # 方法2: 尝试 UTF-16LE 解码
        try:
            text = self._extract_utf16le_text(data)
            if text:
                results.append(('utf16', text, len(text)))
        except Exception:
            pass

        # 方法3: 尝试其他常见编码
        for enc in ['utf-8', 'cp1252']:
            if enc == codepage:
                continue
            try:
                text = data.decode(enc, errors='ignore')
                cleaned = self._clean_doc_text(text)
                if cleaned:
                    results.append((enc, cleaned, len(cleaned)))
            except Exception:
                pass

        # 选择最好的结果（包含最多中文字符或最长有意义文本）
        if results:
            # 优先选择包含中文的结果
            chinese_results = [(enc, txt, length) for enc, txt, length in results
                               if any('\u4e00' <= c <= '\u9fff' for c in txt)]
            if chinese_results:
                return max(chinese_results, key=lambda x: x[2])[1]
            return max(results, key=lambda x: x[2])[1]

        return ''

    def _extract_utf16le_text(self, data: bytes) -> str:
        """提取 UTF-16LE 编码的文本"""
        text_chars = []
        i = 0

        while i + 1 < len(data):
            char_code = data[i] | (data[i + 1] << 8)

            # 可打印 ASCII
            if 0x20 <= char_code <= 0x7E:
                text_chars.append(chr(char_code))
            # CJK 统一汉字
            elif 0x4E00 <= char_code <= 0x9FFF:
                text_chars.append(chr(char_code))
            # CJK 扩展
            elif 0x3400 <= char_code <= 0x4DBF:
                text_chars.append(chr(char_code))
            # 全角字符
            elif 0xFF00 <= char_code <= 0xFFEF:
                text_chars.append(chr(char_code))
            # 中文标点
            elif 0x3000 <= char_code <= 0x303F:
                text_chars.append(chr(char_code))
            # 换行
            elif char_code == 0x000D or char_code == 0x000A:
                text_chars.append('\n')
            # Tab
            elif char_code == 0x0009:
                text_chars.append(' ')

            i += 2

        text = ''.join(text_chars)
        return self._clean_doc_text(text)

    def _clean_doc_text(self, text: str) -> str:
        """清理提取的文本"""
        import re

        # 移除控制字符
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)

        # 移除连续的特殊字符
        text = re.sub(r'[^\w\s\u4e00-\u9fff\u3000-\u303f\uff00-\uffef,.;:!?()（）【】""''。，；：！？、]+', ' ', text)

        # 分割成行并清理
        lines = text.split('\n')
        cleaned_lines = []

        for line in lines:
            line = line.strip()
            # 过滤掉太短或全是乱码的行
            if len(line) < 2:
                continue
            # 检查是否有有意义的字符
            meaningful_chars = sum(1 for c in line if c.isalnum() or '\u4e00' <= c <= '\u9fff')
            if meaningful_chars < len(line) * 0.3:
                continue
            cleaned_lines.append(line)

        return '\n'.join(cleaned_lines)

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
