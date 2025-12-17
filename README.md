# Document Reader MCP Server

一个基于 FastMCP 的文档读取服务器，为大语言模型提供统一的文档读取能力。支持 PDF、Excel (XLS/XLSX/ODS)、Word (DOC/DOCX) 等常见文档格式的智能解析和内容提取。

## 核心特性

- **统一接口**: 单一工具调用，自动识别文档类型
- **智能路由**: 根据文件扩展名自动选择最佳处理器
- **多格式支持**: PDF、XLS、XLSX、ODS、DOC、DOCX
- **灵活输出**: 支持 Markdown 和 JSON 两种输出格式
- **Excel 智能转换**: 小表格转 Markdown，大表格转 CSV 格式
- **预览模式**: 支持仅读取前 N 行快速预览 Excel 结构
- **DOC 多方案处理**: 支持 LibreOffice、Word COM、WPS COM、olefile 多种方式
- **智能编码检测**: 自动检测文档编码，减少乱码问题
- **图片占位**: 无法处理的图片统一使用 `[image]` 标记
- **安全限制**: 文件大小限制 (50MB)，仅支持本地文件
- **错误友好**: 清晰的中文错误提示信息

## 安装

### 前置要求

- Python 3.10+
- [UV](https://github.com/astral-sh/uv) 包管理器

### 安装步骤

```bash
# 克隆或进入项目目录
cd D:/MCP/Document-Reader

# 创建虚拟环境
uv venv

# 激活虚拟环境 (Windows)
.venv\Scripts\activate

# 安装依赖
uv pip install -e ".[dev]"
```

## MCP 配置

### Claude Desktop 配置 (JSON)

编辑 `claude_desktop_config.json` 文件：

```json
{
  "mcpServers": {
    "document-reader": {
      "command": "uv",
      "args": [
        "--directory",
        "D:/MCP/Document-Reader",
        "run",
        "python",
        "-m",
        "src.server"
      ]
    }
  }
}
```

### Claude Code 配置 (TOML)

编辑 `.claude/settings.toml` 文件：

```toml
[mcpServers.document-reader]
command = "uv"
args = [
    "--directory",
    "D:/MCP/Document-Reader",
    "run",
    "python",
    "-m",
    "src.server"
]
```

## 工具说明

### read_document

智能读取各种格式的文档内容。

#### 参数

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `file_path` | string | 是 | - | 文档的绝对路径 |
| `response_format` | string | 否 | `"markdown"` | 输出格式: `"markdown"` 或 `"json"` |
| `include_metadata` | boolean | 否 | `true` | 是否包含文件元数据 |
| `page_range` | string | 否 | `null` | PDF 页码范围，如 `"1-5,7,10-12"` |
| `extract_tables` | boolean | 否 | `true` | 是否提取 PDF 中的表格 |
| `sheet_name` | string | 否 | `null` | Excel 工作表名称 |
| `sheet_index` | integer | 否 | `null` | Excel 工作表索引（从 0 开始） |
| `max_rows` | integer | 否 | `null` | Excel 最大读取行数 |
| `preview_mode` | boolean | 否 | `false` | Excel 预览模式，仅读取前 10 行 |
| `include_images_info` | boolean | 否 | `false` | Word 是否包含图片信息 |

## 使用示例

### 读取 PDF 文档

```json
{
  "file_path": "D:/documents/report.pdf",
  "response_format": "markdown"
}
```

### 读取 PDF 特定页面

```json
{
  "file_path": "D:/documents/report.pdf",
  "page_range": "1-5,10",
  "extract_tables": true,
  "response_format": "json"
}
```

### 读取 Excel 工作表

```json
{
  "file_path": "D:/data/sales.xlsx",
  "sheet_name": "Q1 Sales",
  "response_format": "markdown"
}
```

### Excel 预览模式

```json
{
  "file_path": "D:/data/huge_data.xlsx",
  "preview_mode": true
}
```

### 读取 Word 文档

```json
{
  "file_path": "D:/documents/proposal.docx",
  "include_images_info": true
}
```

## 输出格式

### Markdown 格式示例

```markdown
# 文档: report.pdf

## 元数据
- 文件类型: PDF
- 页数: 10
- 作者: John Doe
- 创建时间: 2024-01-15
- 文件大小: 2.5 MB

## 内容

### 第1页
这是第一页的文本内容...

### 表格1 (第3页)
| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 数据1 | 数据2 | 数据3 |
```

### JSON 格式示例

```json
{
  "file_name": "report.pdf",
  "file_type": "PDF",
  "metadata": {
    "author": "John Doe",
    "created_date": "2024-01-15",
    "page_count": 10,
    "file_size": "2.5 MB"
  },
  "content": [
    {
      "type": "text",
      "page": 1,
      "text": "这是第一页的文本内容..."
    },
    {
      "type": "table",
      "page": 3,
      "table_index": 0,
      "data": [
        ["列1", "列2", "列3"],
        ["数据1", "数据2", "数据3"]
      ]
    }
  ]
}
```

## 项目结构

```
Document-Reader/
├── src/
│   ├── __init__.py
│   ├── server.py              # FastMCP 服务器主入口
│   ├── models.py              # Pydantic 数据模型
│   ├── router.py              # 文档路由器
│   ├── formatter.py           # 响应格式化器
│   ├── utils.py               # 共享工具函数
│   └── processors/
│       ├── __init__.py
│       ├── base.py            # 处理器基类
│       ├── pdf.py             # PDF 处理器
│       ├── excel.py           # Excel 处理器
│       └── word.py            # Word 处理器
├── tests/                     # 测试文件目录
├── pyproject.toml             # UV 项目配置
├── PROJECT_SPEC.md            # 开发需求文档
└── README.md                  # 本文档
```

## 测试

### 使用 MCP Inspector 测试

```bash
npx @modelcontextprotocol/inspector uv --directory D:/MCP/Document-Reader run python -m src.server
```

### 运行单元测试

```bash
uv run pytest tests/
```

## 技术栈

- **Python**: 3.10+
- **MCP SDK**: 官方 Python SDK (FastMCP)
- **Pydantic**: 2.0+ (数据验证)
- **pdfplumber**: PDF 文本和表格提取
- **openpyxl**: Excel 2007+ 格式 (.xlsx)
- **xlrd**: Excel 97-2003 格式 (.xls)
- **odfpy**: OpenDocument 表格 (.ods)
- **python-docx**: Word 2007+ 格式 (.docx)
- **olefile**: DOC 文件 OLE 结构解析
- **chardet**: 文本编码自动检测

## DOC 文件处理

DOC 格式（Word 97-2003）支持多种处理方式，按优先级自动选择：

| 优先级 | 方法 | 说明 | 要求 |
|--------|------|------|------|
| 1 | LibreOffice | 转换为 DOCX 后读取，效果最佳 | 安装 LibreOffice |
| 2 | Microsoft Word COM | 通过 COM 接口直接读取 | Windows + Microsoft Office |
| 3 | WPS Office COM | 通过 COM 接口直接读取 | Windows + WPS Office |
| 4 | olefile 纯 Python | 直接解析 OLE 结构 | 无（内置） |

**推荐**: 安装 [LibreOffice](https://www.libreoffice.org/) 以获得最佳的 DOC 文件读取效果。

## 限制

- 最大文件大小: 50MB
- 仅支持本地文件路径
- 不支持加密/密码保护的文档
- 图片内容无法提取，使用 `[image]` 占位
- DOC 文件使用 olefile 纯 Python 方案时可能出现乱码，建议安装 LibreOffice

## 许可证

MIT License
