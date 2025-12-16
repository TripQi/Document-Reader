# Pydantic 数据模型
"""定义所有输入输出数据模型"""

from enum import Enum
from typing import Optional, Literal, Any
from pydantic import BaseModel, Field


class ResponseFormat(str, Enum):
    """响应格式枚举"""
    MARKDOWN = "markdown"
    JSON = "json"


class ReadDocumentInput(BaseModel):
    """读取文档的输入参数模型"""

    # 必需参数
    file_path: str = Field(
        ...,
        description="文档的绝对路径"
    )

    # 通用可选参数
    response_format: ResponseFormat = Field(
        default=ResponseFormat.MARKDOWN,
        description="输出格式: markdown 或 json"
    )
    include_metadata: bool = Field(
        default=True,
        description="是否包含文件元数据"
    )

    # PDF 专用参数
    page_range: Optional[str] = Field(
        default=None,
        description="PDF页码范围，如 '1-5,7,10-12'"
    )
    extract_tables: bool = Field(
        default=True,
        description="是否提取PDF中的表格"
    )

    # Excel 专用参数
    sheet_name: Optional[str] = Field(
        default=None,
        description="Excel工作表名称"
    )
    sheet_index: Optional[int] = Field(
        default=None,
        description="Excel工作表索引（从0开始）"
    )
    max_rows: Optional[int] = Field(
        default=None,
        description="最大读取行数，None表示全部读取"
    )
    preview_mode: bool = Field(
        default=False,
        description="预览模式：仅读取前10行了解结构"
    )

    # Word 专用参数
    include_images_info: bool = Field(
        default=False,
        description="是否包含图片信息（仅元数据）"
    )


class ContentItem(BaseModel):
    """内容项模型"""
    type: Literal["text", "table", "image"]
    page: Optional[int] = None
    text: Optional[str] = None
    table_index: Optional[int] = None
    data: Optional[list[list[Any]]] = None


class DocumentMetadata(BaseModel):
    """文档元数据模型"""
    file_name: str
    file_type: str
    file_size: str

    # PDF 元数据
    page_count: Optional[int] = None
    author: Optional[str] = None
    created_date: Optional[str] = None
    modified_date: Optional[str] = None

    # Excel 元数据
    sheet_name: Optional[str] = None
    total_rows: Optional[int] = None
    total_columns: Optional[int] = None
    available_sheets: Optional[list[str]] = None

    # 格式提示
    format_hint: Optional[str] = None  # "markdown_table", "csv", "text"
    preview_mode: bool = False


class DocumentResult(BaseModel):
    """文档处理结果模型"""
    file_name: str
    file_type: str
    metadata: DocumentMetadata
    content: list[ContentItem]
    format_hint: str = "text"  # "markdown_table", "csv", "text"


class ErrorResponse(BaseModel):
    """错误响应模型"""
    error: bool = True
    message: str
    file_path: Optional[str] = None
