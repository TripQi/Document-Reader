# 文档路由器
"""根据文件扩展名选择合适的文档处理器"""

import os
from typing import Optional

from .processors.base import DocumentProcessor
from .utils import SUPPORTED_EXTENSIONS


class DocumentRouter:
    """文档处理路由器

    根据文件扩展名自动选择最佳的文档处理器。
    """

    def __init__(self) -> None:
        """初始化路由器，注册所有处理器"""
        self._processors: list[DocumentProcessor] = []

    def register_processor(self, processor: DocumentProcessor) -> None:
        """注册文档处理器

        Args:
            processor: 文档处理器实例
        """
        self._processors.append(processor)

    def get_processor(self, file_path: str) -> DocumentProcessor:
        """根据文件扩展名获取对应处理器

        Args:
            file_path: 文件路径

        Returns:
            DocumentProcessor: 对应的文档处理器

        Raises:
            ValueError: 不支持的文件格式
        """
        ext = os.path.splitext(file_path)[1].lower()

        for processor in self._processors:
            if processor.supports_extension(ext):
                return processor

        supported = ', '.join(SUPPORTED_EXTENSIONS.keys())
        raise ValueError(
            f"错误: 不支持的文件格式 '{ext}'。支持的格式: {supported}"
        )

    def is_supported(self, file_path: str) -> bool:
        """检查文件格式是否支持

        Args:
            file_path: 文件路径

        Returns:
            bool: 是否支持该文件格式
        """
        ext = os.path.splitext(file_path)[1].lower()
        return any(p.supports_extension(ext) for p in self._processors)

    @property
    def supported_extensions(self) -> list[str]:
        """获取所有支持的文件扩展名

        Returns:
            list[str]: 支持的扩展名列表
        """
        extensions = set()
        for processor in self._processors:
            extensions.update(processor.supported_extensions)
        return sorted(extensions)


# 全局路由器实例（延迟初始化）
_router: Optional[DocumentRouter] = None


def get_router() -> DocumentRouter:
    """获取全局路由器实例

    Returns:
        DocumentRouter: 全局路由器实例
    """
    global _router
    if _router is None:
        _router = DocumentRouter()
        # 延迟导入处理器以避免循环导入
        from .processors.pdf import PdfProcessor
        from .processors.excel import ExcelProcessor
        from .processors.word import WordProcessor

        _router.register_processor(PdfProcessor())
        _router.register_processor(ExcelProcessor())
        _router.register_processor(WordProcessor())

    return _router
