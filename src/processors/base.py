# 文档处理器基类
"""定义文档处理器的抽象基类"""

from abc import ABC, abstractmethod
from typing import Any

from ..models import ReadDocumentInput, DocumentResult


class DocumentProcessor(ABC):
    """文档处理器抽象基类

    所有具体的文档处理器（PDF、Excel、Word）都必须继承此类并实现抽象方法。
    """

    @abstractmethod
    async def process(self, params: ReadDocumentInput) -> DocumentResult:
        """处理文档并返回结构化数据

        Args:
            params: 读取文档的输入参数

        Returns:
            DocumentResult: 包含文件名、类型、元数据和内容的结构化结果

        Raises:
            FileNotFoundError: 文件不存在
            PermissionError: 无权限访问文件
            ValueError: 文件格式无效或损坏
        """
        pass

    @abstractmethod
    def supports_extension(self, ext: str) -> bool:
        """检查是否支持该文件扩展名

        Args:
            ext: 文件扩展名（包含点号，如 '.pdf'）

        Returns:
            bool: 是否支持该扩展名
        """
        pass

    @property
    @abstractmethod
    def supported_extensions(self) -> list[str]:
        """返回支持的文件扩展名列表

        Returns:
            list[str]: 支持的扩展名列表，如 ['.pdf', '.xlsx']
        """
        pass
