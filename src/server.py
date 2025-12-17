# FastMCP 服务器主入口
"""Document Reader MCP Server - 提供统一的文档读取能力"""

from mcp.server.fastmcp import FastMCP

from .models import ReadDocumentInput, ResponseFormat
from .router import get_router
from .formatter import ResponseFormatter
from .utils import validate_file, handle_file_error

# 创建 FastMCP 服务器实例
mcp = FastMCP(
    name="Document Reader",
)


@mcp.tool(
    name="read_document",
    description="智能读取各种格式的文档内容，支持 PDF、Excel (XLS/XLSX/ODS)、Word (DOC/DOCX) 格式",
    annotations={
        "title": "读取文档内容",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def read_document(
    file_path: str,
    response_format: str = "markdown",
    include_metadata: bool = True,
    page_range: str | None = None,
    extract_tables: bool = True,
    sheet_name: str | None = None,
    sheet_index: int | None = None,
    max_rows: int | None = None,
    preview_mode: bool = False,
    include_images_info: bool = False,
) -> str:
    """智能读取各种格式的文档内容

    Args:
        file_path: 文档的绝对路径
        response_format: 输出格式，可选 "markdown" 或 "json"，默认 "markdown"
        include_metadata: 是否包含文件元数据，默认 True
        page_range: PDF 页码范围，如 "1-5,7,10-12"
        extract_tables: 是否提取 PDF 中的表格，默认 True
        sheet_name: Excel 工作表名称
        sheet_index: Excel 工作表索引（从 0 开始）
        max_rows: Excel 最大读取行数，None 表示全部读取
        preview_mode: Excel 预览模式，仅读取前 10 行了解结构
        include_images_info: Word 是否包含图片信息

    Returns:
        str: 格式化后的文档内容（Markdown 或 JSON）
    """
    try:
        # 验证文件
        validate_file(file_path)

        # 构建输入参数
        params = ReadDocumentInput(
            file_path=file_path,
            response_format=ResponseFormat(response_format),
            include_metadata=include_metadata,
            page_range=page_range,
            extract_tables=extract_tables,
            sheet_name=sheet_name,
            sheet_index=sheet_index,
            max_rows=max_rows,
            preview_mode=preview_mode,
            include_images_info=include_images_info,
        )

        # 获取路由器和处理器
        router = get_router()
        processor = router.get_processor(file_path)

        # 处理文档
        result = await processor.process(params)

        # 格式化输出
        output = ResponseFormatter.format(
            result,
            params.response_format,
            params.include_metadata
        )

        return output

    except Exception as e:
        error_msg = handle_file_error(e, file_path)
        return f"错误: {error_msg}"


def main():
    """启动 MCP 服务器"""
    mcp.run()


if __name__ == "__main__":
    main()
