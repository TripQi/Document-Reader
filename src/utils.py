# 共享工具函数
"""提供文件验证、格式转换等通用工具函数"""

import os
import csv
import io
from typing import Any, Optional


# 常量定义
MAX_FILE_SIZE_MB = 50
MARKDOWN_MAX_ROWS = 50  # 超过此行数使用 CSV 格式
PREVIEW_ROWS = 10  # 预览模式读取行数

# 支持的文件格式
SUPPORTED_EXTENSIONS = {
    '.pdf': 'PDF',
    '.xls': 'Excel',
    '.xlsx': 'Excel',
    '.ods': 'ODS',
    '.doc': 'Word',
    '.docx': 'Word',
}


def validate_file_exists(file_path: str) -> None:
    """验证文件是否存在

    Args:
        file_path: 文件路径

    Raises:
        FileNotFoundError: 文件不存在
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(
            f"错误: 文件未找到 '{file_path}'。请检查路径是否正确。"
        )


def validate_file_size(file_path: str, max_size_mb: int = MAX_FILE_SIZE_MB) -> None:
    """验证文件大小是否在限制内

    Args:
        file_path: 文件路径
        max_size_mb: 最大文件大小（MB）

    Raises:
        ValueError: 文件过大
    """
    file_size = os.path.getsize(file_path)
    file_size_mb = file_size / (1024 * 1024)

    if file_size_mb > max_size_mb:
        raise ValueError(
            f"错误: 文件大小 {file_size_mb:.1f}MB 超过限制 {max_size_mb}MB。请使用较小的文件。"
        )


def validate_file_permissions(file_path: str) -> None:
    """验证文件是否可读

    Args:
        file_path: 文件路径

    Raises:
        PermissionError: 无权限访问文件
    """
    if not os.access(file_path, os.R_OK):
        raise PermissionError(
            f"错误: 无权限访问文件 '{file_path}'。请检查文件权限。"
        )


def validate_file(file_path: str) -> None:
    """执行所有文件验证

    Args:
        file_path: 文件路径

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 文件过大或格式不支持
        PermissionError: 无权限访问文件
    """
    validate_file_exists(file_path)
    validate_file_size(file_path)
    validate_file_permissions(file_path)

    # 验证文件格式
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        supported = ', '.join(SUPPORTED_EXTENSIONS.keys())
        raise ValueError(
            f"错误: 不支持的文件格式 '{ext}'。支持的格式: {supported}"
        )


def get_file_extension(file_path: str) -> str:
    """获取文件扩展名（小写）

    Args:
        file_path: 文件路径

    Returns:
        str: 小写的文件扩展名，如 '.pdf'
    """
    return os.path.splitext(file_path)[1].lower()


def get_file_type(file_path: str) -> str:
    """获取文件类型名称

    Args:
        file_path: 文件路径

    Returns:
        str: 文件类型名称，如 'PDF', 'Excel', 'Word'
    """
    ext = get_file_extension(file_path)
    return SUPPORTED_EXTENSIONS.get(ext, 'Unknown')


def format_file_size(size_bytes: int) -> str:
    """将字节数转换为人类可读格式

    Args:
        size_bytes: 文件大小（字节）

    Returns:
        str: 格式化的文件大小，如 '2.5 MB'
    """
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def parse_page_range(page_range: str, total_pages: int) -> list[int]:
    """解析页码范围字符串

    Args:
        page_range: 页码范围字符串，如 '1-5,7,10-12'
        total_pages: 文档总页数

    Returns:
        list[int]: 页码列表（从1开始）

    Raises:
        ValueError: 页码范围格式无效

    Examples:
        >>> parse_page_range("1-5", 10)
        [1, 2, 3, 4, 5]
        >>> parse_page_range("1,3,5", 10)
        [1, 3, 5]
        >>> parse_page_range("1-3,5,7-9", 10)
        [1, 2, 3, 5, 7, 8, 9]
    """
    if not page_range:
        return list(range(1, total_pages + 1))

    pages = set()
    parts = page_range.replace(' ', '').split(',')

    for part in parts:
        if '-' in part:
            try:
                start, end = part.split('-')
                start = int(start)
                end = int(end)

                if start < 1 or end > total_pages or start > end:
                    raise ValueError(
                        f"错误: 页码范围 '{page_range}' 无效。"
                        f"有效范围: 1-{total_pages}"
                    )

                pages.update(range(start, end + 1))
            except ValueError as e:
                if "invalid literal" in str(e):
                    raise ValueError(
                        f"错误: 页码范围 '{page_range}' 格式无效。"
                        f"正确格式: '1-5' 或 '1,3,5'"
                    )
                raise
        else:
            try:
                page = int(part)
                if page < 1 or page > total_pages:
                    raise ValueError(
                        f"错误: 页码 {page} 超出范围。有效范围: 1-{total_pages}"
                    )
                pages.add(page)
            except ValueError:
                raise ValueError(
                    f"错误: 页码范围 '{page_range}' 格式无效。"
                    f"正确格式: '1-5' 或 '1,3,5'"
                )

    return sorted(pages)


def convert_to_csv(data: list[list[Any]]) -> str:
    """将表格数据转换为 CSV 格式字符串

    Args:
        data: 二维表格数据

    Returns:
        str: CSV 格式字符串
    """
    output = io.StringIO()
    writer = csv.writer(output)

    for row in data:
        # 将所有值转换为字符串
        writer.writerow([str(cell) if cell is not None else '' for cell in row])

    return output.getvalue()


def should_use_csv_format(row_count: int, threshold: int = MARKDOWN_MAX_ROWS) -> bool:
    """判断是否应该使用 CSV 格式

    Args:
        row_count: 数据行数
        threshold: 阈值，超过此值使用 CSV 格式

    Returns:
        bool: 是否应该使用 CSV 格式
    """
    return row_count > threshold


def determine_excel_output_format(
    row_count: int,
    preview_mode: bool,
    max_rows: Optional[int] = None
) -> str:
    """决定 Excel 输出格式

    优先级:
    1. 预览模式 -> 读取前10行，Markdown表格
    2. 自定义 max_rows -> 读取指定行数，根据行数选择格式
    3. 全部读取 -> 根据实际行数选择格式

    Args:
        row_count: 实际数据行数
        preview_mode: 是否为预览模式
        max_rows: 自定义最大行数限制

    Returns:
        str: 输出格式 "markdown_table" 或 "csv"
    """
    if preview_mode:
        return "markdown_table"  # 预览模式固定用 Markdown

    actual_rows = min(row_count, max_rows) if max_rows else row_count

    if actual_rows <= MARKDOWN_MAX_ROWS:
        return "markdown_table"
    else:
        return "csv"


def replace_images_with_placeholder(content: str, placeholder: str = "[image]") -> str:
    """将内容中的图片替换为占位符标记

    Args:
        content: 原始内容
        placeholder: 占位符文本

    Returns:
        str: 替换后的内容
    """
    # 这个函数主要用于处理已提取的文本中可能存在的图片引用
    # 实际的图片检测在各个处理器中实现
    return content


def handle_file_error(e: Exception, file_path: str) -> str:
    """统一的文件错误处理

    Args:
        e: 异常对象
        file_path: 文件路径

    Returns:
        str: 友好的错误消息
    """
    error_type = type(e).__name__

    if isinstance(e, FileNotFoundError):
        return f"错误: 文件未找到 '{file_path}'。请检查路径是否正确。"
    elif isinstance(e, PermissionError):
        return f"错误: 无权限访问文件 '{file_path}'。请检查文件权限。"
    elif isinstance(e, ValueError):
        return str(e)
    elif "password" in str(e).lower() or "encrypted" in str(e).lower():
        return f"错误: 文件 '{file_path}' 受密码保护，暂不支持加密文档。"
    elif "corrupt" in str(e).lower() or "invalid" in str(e).lower():
        return f"错误: 文件 '{file_path}' 可能已损坏或格式无效，无法解析。"
    else:
        return f"错误: 处理文件 '{file_path}' 时发生错误 ({error_type}): {str(e)}"
