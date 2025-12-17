# DOC 文件处理优化计划

## 目标
优化 DOC 文件处理，支持通过 LibreOffice/Microsoft Office/WPS 将 DOC 转换为 DOCX 后读取，获得最佳解析效果。

## 实现方案

### 处理优先级（从高到低）
1. **LibreOffice** - 用户首选，免费开源
2. **Microsoft Word COM** - Windows + Office 用户
3. **WPS Office COM** - Windows + WPS 用户
4. **olefile 纯 Python** - 备选方案（可能有乱码）

### 核心修改文件
- `src/processors/word.py` - 添加 LibreOffice 和 WPS 支持

## 详细实现步骤

### 步骤 1: 添加 LibreOffice 转换方法

```python
async def _process_doc_with_libreoffice(self, params: ReadDocumentInput) -> DocumentResult:
    """使用 LibreOffice 将 DOC 转换为 DOCX 后读取"""
    import subprocess
    import tempfile
    import shutil

    file_path = params.file_path

    # 查找 LibreOffice 可执行文件
    soffice_path = self._find_libreoffice()
    if not soffice_path:
        raise ValueError("未找到 LibreOffice")

    # 创建临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        # 使用 LibreOffice 转换
        cmd = [
            soffice_path,
            '--headless',
            '--convert-to', 'docx',
            '--outdir', temp_dir,
            file_path
        ]
        result = subprocess.run(cmd, capture_output=True, timeout=60)

        if result.returncode != 0:
            raise ValueError("LibreOffice 转换失败")

        # 读取转换后的 DOCX
        docx_path = os.path.join(temp_dir, os.path.splitext(os.path.basename(file_path))[0] + '.docx')

        # 复用现有的 DOCX 处理逻辑
        params_copy = params.model_copy()
        params_copy.file_path = docx_path
        return await self._process_docx(params_copy)
```

### 步骤 2: 添加 LibreOffice 路径查找

```python
def _find_libreoffice(self) -> Optional[str]:
    """查找 LibreOffice 可执行文件路径"""
    import shutil

    # Windows 常见路径
    if sys.platform == 'win32':
        possible_paths = [
            r'C:\Program Files\LibreOffice\program\soffice.exe',
            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        # 尝试从 PATH 查找
        return shutil.which('soffice')

    # Linux/Mac
    return shutil.which('soffice') or shutil.which('libreoffice')
```

### 步骤 3: 添加 WPS COM 支持

```python
async def _process_doc_with_wps(self, params: ReadDocumentInput) -> DocumentResult:
    """使用 WPS Office COM 接口处理 DOC 文件"""
    import win32com.client
    import pythoncom

    file_path = params.file_path
    abs_path = os.path.abspath(file_path)

    pythoncom.CoInitialize()
    try:
        # WPS 的 COM 类名
        wps = win32com.client.Dispatch("Kwps.Application")
        wps.Visible = False

        try:
            doc = wps.Documents.Open(abs_path, ReadOnly=True)
            try:
                # 复用 COM 内容提取方法
                metadata = self._extract_doc_com_metadata(doc, file_path)
                content = self._extract_doc_com_content(doc)

                return DocumentResult(...)
            finally:
                doc.Close(False)
        finally:
            wps.Quit()
    finally:
        pythoncom.CoUninitialize()
```

### 步骤 4: 更新 _process_doc 方法

```python
async def _process_doc(self, params: ReadDocumentInput) -> DocumentResult:
    """处理 DOC 文件

    尝试多种方法（按优先级）:
    1. LibreOffice 转换为 DOCX（推荐，跨平台）
    2. Microsoft Word COM（Windows + Office）
    3. WPS Office COM（Windows + WPS）
    4. olefile 纯 Python（备选，可能有乱码）
    """
    file_path = params.file_path
    errors = []

    # 1. 优先尝试 LibreOffice
    try:
        return await self._process_doc_with_libreoffice(params)
    except Exception as e:
        errors.append(f"LibreOffice: {e}")

    # 2. Windows 上尝试 Microsoft Word COM
    if sys.platform == 'win32':
        try:
            return await self._process_doc_with_com(params)
        except Exception as e:
            errors.append(f"Microsoft Word: {e}")

        # 3. 尝试 WPS COM
        try:
            return await self._process_doc_with_wps(params)
        except Exception as e:
            errors.append(f"WPS: {e}")

    # 4. 备选：olefile 纯 Python
    try:
        result = await self._process_doc_with_olefile(params)
        # 添加警告
        ...
        return result
    except Exception as e:
        errors.append(f"olefile: {e}")

    raise ValueError(f"无法处理 DOC 文件。尝试的方法: {'; '.join(errors)}")
```

## 测试计划
1. 测试 LibreOffice 转换（用户环境）
2. 验证转换后的 DOCX 可正确读取
3. 测试 WPS COM 接口（如有 WPS）
4. 确保回退机制正常工作
