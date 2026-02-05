# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个 **Word 文档编辑 MCP 服务器**，用于与 Claude Code CLI 集成，提供完整的 Word 文档操作能力。基于 MCP (Model Context Protocol) 协议，通过 python-docx 库实现 Word 文档的创建、编辑、格式化等功能。

## 核心架构

### MCP 服务器架构
- **服务入口**: `src/server.py` - MCP 服务器主入口，注册所有工具并处理工具调用
- **工具模块**: `src/tools/` - 按功能分类的工具实现
- **辅助工具**: `src/utils/` - 文档管理和错误处理

### 工具模块分类

工具按功能分为以下模块：

1. **document_basic.py** - 文档基础操作（创建、读取、复制、获取信息）
2. **content_edit.py** - 内容编辑（段落、标题、查找替换）
3. **table_ops.py** - 表格操作（创建、编辑、格式化、读取）
4. **style_format.py** - 样式格式（分页符、页边距）
5. **image_ops.py** - 图片操作（插入图片）
6. **list_ops.py** - 列表操作（有序/无序列表）
7. **advanced.py** - 高级功能（脚注、页眉页脚、大纲）
8. **interface_doc.py** - 接口文档生成（标准格式的接口文档表格）

### 文档管理机制

**重要**: `DocumentManager` 类采用**无缓存设计**，每次操作都重新加载文档：
- `get_or_open()` - 每次都从磁盘重新加载文档
- `save()` - 保存文档到磁盘
- 这种设计确保多个操作之间的数据一致性，避免缓存导致的数据丢失

## 常用命令

### 安装依赖
```bash
pip install -r requirements.txt
```

### 运行 MCP 服务器
```bash
# 方式1: 直接运行
python -m src.server

# 方式2: 使用批处理脚本（Windows）
start_mcp.bat

# 方式3: 使用 Python 脚本
python run_server.py
```

### 开发工具
```bash
# 代码格式化
black src/ --line-length 100

# 代码检查
ruff check src/
```

## 添加新工具的步骤

1. **在对应的工具模块中实现函数**
   - 例如在 `src/tools/content_edit.py` 中添加新的内容编辑功能
   - 函数必须是 `async` 函数
   - 使用 `DocumentManager` 进行文档操作
   - 返回格式: `{"success": True/False, "message": "...", ...}`

2. **在 `src/server.py` 中注册工具**
   - 在 `list_tools()` 函数中添加 `Tool` 定义
   - 定义工具名称、描述和 `inputSchema`

3. **在 `call_tool()` 函数中添加调用处理**
   - 添加 `elif name == "your_tool_name":` 分支
   - 调用对应的工具函数

## 关键设计模式

### 错误处理
- 所有工具函数使用统一的错误处理机制
- 返回 JSON 格式: `{"success": bool, "error": str, "message": str}`
- 使用 `error_handler.py` 中的 `DocxError` 自定义异常

### 文档操作流程
```python
# 标准操作流程
doc_manager = DocumentManager()
doc = doc_manager.get_or_open(filename)  # 加载文档
# ... 执行操作 ...
doc_manager.save(filename, doc)  # 保存文档
```

### 索引约定
- **段落索引**: 从 0 开始
- **表格索引**: 从 0 开始
- **行/列索引**: 从 0 开始
- 范围操作使用 `start_index` 和 `end_index`，均为包含边界

## 技术栈

- **Python**: 3.10+
- **MCP SDK**: 0.9.0+ (Model Context Protocol)
- **python-docx**: 1.1.0+ (Word 文档操作)
- **Pillow**: 10.0.0+ (图片处理)

## 配置 Claude Code CLI

在 Claude Code CLI 配置文件中添加：

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "D:\\workspace\\hwc-code\\doc-mcp-server\\start_mcp.bat",
      "args": []
    }
  }
}
```

## 特殊功能说明

### 接口文档生成
`insert_interface_doc` 工具可以生成标准格式的接口文档，包括：
- 接口基本信息（名称、路径、请求方式）
- 请求参数表格
- 响应参数表格
- 请求/响应示例

### 表格操作
- 支持单元格合并、对齐、背景色、内边距
- 支持行高、列宽设置
- 支持表格缩进
- 支持读取表格数据和单元格内容

### 文本格式化
- 支持字体名称、字号、粗体、斜体
- 支持文字颜色（十六进制 RGB 格式，如 'FF0000'）
- 支持段落样式

## 注意事项

1. **文件路径**: 所有文件路径会自动转换为绝对路径
2. **颜色格式**: 使用 6 位十六进制 RGB 值，如 'FF0000' 表示红色
3. **单位转换**:
   - 页边距、行高、列宽使用英寸（inches）
   - 字号使用磅（pt）
4. **文档保存**: 每次操作后都会自动保存文档
5. **并发安全**: 由于无缓存设计，多个操作之间不会相互干扰
