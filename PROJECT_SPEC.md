# Document Reader MCP 项目开发文档

## 📋 项目概述

### 项目名称
Document Reader MCP Server

### 项目描述
一个基于FastMCP的文档读取服务器,为大语言模型提供统一的文档读取能力。支持PDF、Excel(XLS/XLSX/ODS)、Word(DOC/DOCX)等常见文档格式的智能解析和内容提取。

### 核心特性
- **统一接口**: 单一工具调用,自动识别文档类型
- **智能路由**: 根据文件扩展名自动选择最佳处理器
- **多格式支持**: PDF、XLS、XLSX、ODS、DOC、DOCX
- **灵活输出**: 支持Markdown和JSON两种输出格式
- **安全限制**: 文件大小限制(50MB),仅支持本地文件
- **错误友好**: 清晰的错误提示信息

---

## 🎯 技术栈

### 核心框架
- **Python**: 3.10+
- **FastMCP**: 最新版本 (MCP Python SDK)
- **Pydantic**: 2.0+ (数据验证)

### 文档处理库
| 文档格式 | 主要库 | 备选库 | 用途 |
|---------|--------|--------|------|
| PDF | `pdfplumber` | `PyPDF2` | 文本提取、表格识别 |
| XLSX | `openpyxl` | - | Excel 2007+格式 |
| XLS | `xlrd` | - | Excel 97-2003格式 |
| ODS | `odfpy` | - | OpenDocument表格 |
| DOCX | `python-docx` | - | Word 2007+格式 |
| DOC | `antiword` + `subprocess` | `textract` | Word 97-2003格式 |

### 开发工具
- **包管理器**: UV
- **虚拟环境**: Python venv
- **代码规范**: Type hints, Pydantic validation

---

## 📁 项目结构

```
Document-Reader/
├── src/
│   ├── __init__.py
│   ├── server.py              # FastMCP服务器主入口
│   ├── models.py              # Pydantic数据模型
│   ├── router.py              # 文档路由器
│   ├── formatter.py           # 响应格式化器
│   ├── utils.py               # 共享工具函数
│   └── processors/
│       ├── __init__.py
│       ├── base.py           # 处理器基类
│       ├── pdf.py            # PDF处理器
│       ├── excel.py          # Excel处理器
│       └── word.py           # Word处理器
├── tests/                     # 测试文件目录
│   ├── __init__.py
│   ├── test_pdf.py
│   ├── test_excel.py
│   ├── test_word.py
│   └── fixtures/             # 测试文档样本
├── docs/                      # 文档目录
│   ├── API.md                # API文档
│   └── EXAMPLES.md           # 使用示例
├── .gitignore
├── README.md                  # 项目说明
├── PROJECT_SPEC.md           # 本文档
├── pyproject.toml            # UV项目配置
└── requirements.txt          # 依赖列表(备用)
```

---

## 🔧 核心组件设计

### 1. 统一工具接口: `read_document`

#### 工具定义
```python
@mcp.tool(
    name="read_document",
    annotations={
        "title": "读取文档内容",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False
    }
)
async def read_document(params: ReadDocumentInput) -> str:
    """智能读取各种格式的文档内容"""
```

#### 输入参数模型
```python
class ReadDocumentInput(BaseModel):
    # 必需参数
    file_path: str                              # 文档绝对路径

    # 通用可选参数
    response_format: ResponseFormat = "markdown"  # 输出格式
    include_metadata: bool = True                # 是否包含元数据

    # PDF专用参数
    page_range: Optional[str] = None            # 页码范围: "1-5,7,10-12"
    extract_tables: bool = True                 # 是否提取表格

    # Excel专用参数
    sheet_name: Optional[str] = None            # 工作表名称
    sheet_index: Optional[int] = None           # 工作表索引
    max_rows: Optional[int] = 1000              # 最大行数限制

    # Word专用参数
    include_images_info: bool = False           # 是否包含图片信息
```

#### 输出格式

**Markdown格式示例**:
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

### 第2页
这是第二页的内容...

### 表格1 (第3页)
| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 数据1 | 数据2 | 数据3 |
```

**JSON格式示例**:
```json
{
  "file_name": "report.pdf",
  "file_type": "PDF",
  "metadata": {
    "author": "John Doe",
    "created_date": "2024-01-15",
    "modified_date": "2024-01-20",
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

---

### 2. 文档处理器架构

#### 基类: DocumentProcessor
```python
from abc import ABC, abstractmethod
from typing import Dict, Any

class DocumentProcessor(ABC):
    """文档处理器抽象基类"""

    @abstractmethod
    async def process(self, params: ReadDocumentInput) -> Dict[str, Any]:
        """
        处理文档并返回结构化数据

        Returns:
            {
                "file_name": str,
                "file_type": str,
                "metadata": dict,
                "content": list
            }
        """
        pass

    @abstractmethod
    def supports_extension(self, ext: str) -> bool:
        """检查是否支持该文件扩展名"""
        pass

    def _validate_file(self, file_path: str) -> None:
        """验证文件存在性和可读性"""
        pass
```

#### PDF处理器: PdfProcessor
```python
class PdfProcessor(DocumentProcessor):
    """PDF文档处理器 - 使用pdfplumber"""

    def supports_extension(self, ext: str) -> bool:
        return ext.lower() == '.pdf'

    async def process(self, params: ReadDocumentInput) -> Dict[str, Any]:
        """
        处理PDF文档
        - 提取文本内容
        - 识别并提取表格
        - 读取元数据
        - 支持页码范围过滤
        """
        pass
```

#### Excel处理器: ExcelProcessor
```python
class ExcelProcessor(DocumentProcessor):
    """Excel文档处理器 - 支持XLS/XLSX/ODS"""

    def supports_extension(self, ext: str) -> bool:
        return ext.lower() in ['.xls', '.xlsx', '.ods']

    async def process(self, params: ReadDocumentInput) -> Dict[str, Any]:
        """
        处理Excel文档
        - 读取指定工作表
        - 提取表格数据
        - 处理公式和格式
        - 支持行数限制
        """
        pass
```

#### Word处理器: WordProcessor
```python
class WordProcessor(DocumentProcessor):
    """Word文档处理器 - 支持DOC/DOCX"""

    def supports_extension(self, ext: str) -> bool:
        return ext.lower() in ['.doc', '.docx']

    async def process(self, params: ReadDocumentInput) -> Dict[str, Any]:
        """
        处理Word文档
        - 提取段落文本
        - 识别表格
        - 提取图片信息(可选)
        - 保留基本样式信息
        """
        pass
```

---

### 3. 文档路由器: DocumentRouter

```python
class DocumentRouter:
    """文档处理路由器 - 根据文件扩展名选择处理器"""

    def __init__(self):
        self.processors: List[DocumentProcessor] = [
            PdfProcessor(),
            ExcelProcessor(),
            WordProcessor()
        ]

    def get_processor(self, file_path: str) -> DocumentProcessor:
        """
        根据文件扩展名获取对应处理器

        Raises:
            ValueError: 不支持的文件格式
        """
        ext = os.path.splitext(file_path)[1].lower()

        for processor in self.processors:
            if processor.supports_extension(ext):
                return processor

        raise ValueError(
            f"不支持的文件格式 '{ext}'。"
            f"支持的格式: .pdf, .xls, .xlsx, .ods, .doc, .docx"
        )
```

---

### 4. 响应格式化器: ResponseFormatter

```python
class ResponseFormatter:
    """响应格式化器 - 转换为Markdown或JSON"""

    @staticmethod
    def to_markdown(data: Dict[str, Any]) -> str:
        """将结构化数据转换为Markdown格式"""
        pass

    @staticmethod
    def to_json(data: Dict[str, Any]) -> str:
        """将结构化数据转换为JSON格式"""
        pass

    @staticmethod
    def _format_table_markdown(table_data: List[List[str]]) -> str:
        """格式化表格为Markdown表格"""
        pass
```

---

### 5. 工具函数: utils.py

```python
# 文件验证
def validate_file_exists(file_path: str) -> None:
    """验证文件是否存在"""
    pass

def validate_file_size(file_path: str, max_size_mb: int = 50) -> None:
    """验证文件大小是否在限制内"""
    pass

def validate_file_permissions(file_path: str) -> None:
    """验证文件是否可读"""
    pass

# 页码解析
def parse_page_range(page_range: str, total_pages: int) -> List[int]:
    """
    解析页码范围字符串

    Examples:
        "1-5" -> [1, 2, 3, 4, 5]
        "1,3,5" -> [1, 3, 5]
        "1-3,5,7-9" -> [1, 2, 3, 5, 7, 8, 9]
    """
    pass

# 错误处理
def handle_file_error(e: Exception, file_path: str) -> str:
    """统一的文件错误处理"""
    pass

# 文件大小格式化
def format_file_size(size_bytes: int) -> str:
    """将字节数转换为人类可读格式"""
    pass
```

---

## 🔒 安全和限制

### 文件大小限制
- **最大文件大小**: 50MB
- **原因**: 防止内存溢出,保证响应速度
- **检查时机**: 文件验证阶段

### 文件路径限制
- **仅支持绝对路径**: 明确文件位置
- **不支持网络URL**: 仅本地文件访问
- **不限制路径范围**: 用户自行管理文件权限

### Excel行数限制
- **默认最大行数**: 1000行
- **可配置范围**: 1-10000行
- **原因**: 防止大表格导致响应过长

---

## ⚠️ 错误处理

### 错误类型和消息

| 错误类型 | 错误消息模板 | HTTP状态码等效 |
|---------|-------------|---------------|
| 文件不存在 | `错误: 文件未找到 '{file_path}'。请检查路径是否正确。` | 404 |
| 文件过大 | `错误: 文件大小 {size}MB 超过限制 50MB。请使用较小的文件。` | 413 |
| 格式不支持 | `错误: 不支持的文件格式 '.{ext}'。支持的格式: .pdf, .xls, .xlsx, .ods, .doc, .docx` | 415 |
| 权限不足 | `错误: 无权限访问文件 '{file_path}'。请检查文件权限。` | 403 |
| 文件损坏 | `错误: 文件 '{file_path}' 可能已损坏或格式无效,无法解析。` | 422 |
| 密码保护 | `错误: 文件 '{file_path}' 受密码保护,暂不支持加密文档。` | 422 |
| 页码无效 | `错误: 页码范围 '{page_range}' 无效。正确格式: '1-5' 或 '1,3,5'` | 400 |
| 工作表不存在 | `错误: 工作表 '{sheet_name}' 不存在。可用工作表: {available_sheets}` | 404 |

### 错误处理策略
1. **输入验证**: Pydantic自动验证参数类型和约束
2. **文件验证**: 在处理前检查文件存在性、大小、权限
3. **异常捕获**: 捕获所有处理异常,返回友好错误消息
4. **错误日志**: 记录详细错误信息用于调试

---

## 📦 依赖管理

### pyproject.toml (UV配置)
```toml
[project]
name = "document-reader-mcp"
version = "0.1.0"
description = "MCP server for reading various document formats"
requires-python = ">=3.10"
dependencies = [
    "mcp>=1.0.0",
    "pydantic>=2.0.0",
    "pdfplumber>=0.10.0",
    "openpyxl>=3.1.0",
    "xlrd>=2.0.1",
    "odfpy>=1.4.1",
    "python-docx>=1.1.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=7.0.0",
    "pytest-asyncio>=0.21.0",
    "black>=23.0.0",
    "mypy>=1.0.0",
]

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.uv]
dev-dependencies = [
    "pytest>=7.0.0",
    "pytest-asyncio>=0.21.0",
]
```

### requirements.txt (备用)
```txt
mcp>=1.0.0
pydantic>=2.0.0
pdfplumber>=0.10.0
openpyxl>=3.1.0
xlrd>=2.0.1
odfpy>=1.4.1
python-docx>=1.1.0
```

---

## 🚀 开发流程

### 阶段1: 项目初始化
- [x] 创建项目目录结构
- [ ] 配置UV虚拟环境
- [ ] 安装依赖包
- [ ] 创建基础文件框架

### 阶段2: 核心框架实现
- [ ] 实现Pydantic数据模型 (`models.py`)
- [ ] 实现文档处理器基类 (`processors/base.py`)
- [ ] 实现文档路由器 (`router.py`)
- [ ] 实现响应格式化器 (`formatter.py`)
- [ ] 实现工具函数 (`utils.py`)

### 阶段3: 处理器实现
- [ ] 实现PDF处理器 (`processors/pdf.py`)
- [ ] 实现Excel处理器 (`processors/excel.py`)
- [ ] 实现Word处理器 (`processors/word.py`)

### 阶段4: 服务器集成
- [ ] 实现FastMCP服务器 (`server.py`)
- [ ] 注册`read_document`工具
- [ ] 集成所有组件
- [ ] 实现错误处理

### 阶段5: 测试和优化
- [ ] 创建测试文档样本
- [ ] 编写单元测试
- [ ] 使用MCP Inspector测试
- [ ] 性能优化
- [ ] 文档完善

---

## 🧪 测试计划

### 测试文档准备
创建以下测试文档:
- `test.pdf` - 包含文本和表格的PDF
- `test.xlsx` - 多工作表Excel文件
- `test.xls` - 旧版Excel文件
- `test.ods` - OpenDocument表格
- `test.docx` - 包含段落和表格的Word文档
- `test.doc` - 旧版Word文档
- `large.pdf` - 超过50MB的大文件(测试限制)
- `encrypted.pdf` - 加密PDF(测试错误处理)

### 测试用例

#### 功能测试
1. **PDF读取**
   - 读取全部页面
   - 读取指定页码范围
   - 提取表格
   - 读取元数据

2. **Excel读取**
   - 读取默认工作表
   - 读取指定工作表(按名称)
   - 读取指定工作表(按索引)
   - 行数限制测试

3. **Word读取**
   - 读取段落文本
   - 读取表格
   - 读取图片信息

#### 错误处理测试
- 文件不存在
- 文件过大
- 不支持的格式
- 权限不足
- 文件损坏
- 密码保护

#### 格式测试
- Markdown输出格式
- JSON输出格式
- 元数据包含/排除

---

## 📖 使用示例

### 基本用法

#### 读取PDF文档
```json
{
  "file_path": "D:/documents/report.pdf",
  "response_format": "markdown"
}
```

#### 读取PDF特定页面
```json
{
  "file_path": "D:/documents/report.pdf",
  "page_range": "1-5,10",
  "extract_tables": true,
  "response_format": "json"
}
```

#### 读取Excel工作表
```json
{
  "file_path": "D:/data/sales.xlsx",
  "sheet_name": "Q1 Sales",
  "max_rows": 500,
  "response_format": "markdown"
}
```

#### 读取Word文档
```json
{
  "file_path": "D:/documents/proposal.docx",
  "include_images_info": true,
  "response_format": "markdown"
}
```

---

## 🔄 MCP配置

### Claude Desktop配置
在 `claude_desktop_config.json` 中添加:

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
        "src/server.py"
      ]
    }
  }
}
```

### 验证安装
```bash
# 使用MCP Inspector测试
npx @modelcontextprotocol/inspector uv --directory D:/MCP/Document-Reader run python src/server.py
```

---

## 📝 开发规范

### 代码风格
- 使用Type Hints标注所有函数参数和返回值
- 使用Pydantic进行数据验证,避免手动验证
- 所有异步操作使用`async/await`
- 遵循PEP 8代码规范

### 文档字符串
- 所有公共函数必须有docstring
- 使用Google风格的docstring
- 包含参数说明、返回值说明、异常说明

### 错误处理
- 使用具体的异常类型
- 提供清晰的错误消息
- 记录详细的错误日志

### 代码复用
- 提取共享功能到工具函数
- 避免代码重复
- 使用继承和组合模式

---

## 🎯 性能优化

### 内存优化
- 大文件分块读取
- 及时释放文件句柄
- 限制Excel读取行数

### 速度优化
- 使用异步I/O
- 缓存文件元数据
- 延迟加载大型对象

### 响应优化
- Markdown格式优先(更紧凑)
- 表格数据截断显示
- 元数据可选包含

---

## 📚 参考资源

### MCP相关
- [MCP官方文档](https://modelcontextprotocol.io/)
- [FastMCP文档](https://github.com/modelcontextprotocol/python-sdk)
- [MCP Inspector](https://github.com/modelcontextprotocol/inspector)

### 文档处理库
- [pdfplumber文档](https://github.com/jsvine/pdfplumber)
- [openpyxl文档](https://openpyxl.readthedocs.io/)
- [python-docx文档](https://python-docx.readthedocs.io/)

### Python开发
- [Pydantic文档](https://docs.pydantic.dev/)
- [UV文档](https://github.com/astral-sh/uv)
- [Python异步编程](https://docs.python.org/3/library/asyncio.html)

---

## 📅 版本规划

### v0.1.0 (MVP)
- [x] 项目初始化
- [ ] 基础框架实现
- [ ] PDF/Excel/Word基本读取
- [ ] Markdown输出格式

### v0.2.0
- [ ] JSON输出格式
- [ ] 完整错误处理
- [ ] 单元测试覆盖

### v0.3.0
- [ ] 性能优化
- [ ] 高级功能(图片提取等)
- [ ] 完整文档

### v1.0.0
- [ ] 生产就绪
- [ ] 完整测试覆盖
- [ ] 性能基准测试

---

## 🤝 贡献指南

### 开发环境设置
```bash
# 克隆项目
cd D:/MCP/Document-Reader

# 创建虚拟环境
uv venv

# 激活虚拟环境
.venv\Scripts\activate  # Windows

# 安装依赖
uv pip install -e ".[dev]"

# 运行测试
pytest tests/
```

### 提交规范
- feat: 新功能
- fix: 错误修复
- docs: 文档更新
- test: 测试相关
- refactor: 代码重构

---

## 📄 许可证

MIT License

---

## 联系方式

项目维护者: [您的名字]
项目地址: D:/MCP/Document-Reader

---

**文档版本**: v1.0
**最后更新**: 2024-12-16
**状态**: 开发中
