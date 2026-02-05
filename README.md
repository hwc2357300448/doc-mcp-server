# Word 文档编辑 MCP 服务器

<div align="center">

[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP SDK](https://img.shields.io/badge/MCP-0.9.0+-green.svg)](https://github.com/modelcontextprotocol)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

一个功能强大的 Word 文档编辑 MCP (Model Context Protocol) 服务器，为 Claude Code CLI 提供完整的 Word 文档操作能力。

[功能特性](#-功能特性) • [快速开始](#-快速开始) • [使用示例](#-使用示例) • [API 文档](#-可用工具列表) • [开发指南](#-开发指南)

📚 **[5分钟快速开始指南](QUICKSTART.md)** | 📝 **[配置示例](CONFIG_EXAMPLES.md)** | 🤝 **[贡献指南](CONTRIBUTING.md)**

</div>

---

## 📋 功能特性

### 文档基础操作
- ✅ 创建新文档
- ✅ 获取文档信息
- ✅ 提取文档文本
- ✅ 列出目录下的文档
- ✅ 复制文档
- ✅ 读取指定段落内容

### 内容编辑
- ✅ 添加段落（支持字体、颜色、样式）
- ✅ 添加标题（1-9级）
- ✅ 删除段落
- ✅ 查找文本
- ✅ 替换文本

### 表格操作
- ✅ 创建表格
- ✅ 设置单元格内容
- ✅ 格式化表格
- ✅ 插入表格行/列
- ✅ 删除表格行/列/整表
- ✅ 合并单元格
- ✅ 单元格对齐（水平/垂直）
- ✅ 单元格背景色
- ✅ 单元格内边距
- ✅ 行高设置
- ✅ 列宽设置
- ✅ 单元格文本格式化
- ✅ 表格缩进
- ✅ 读取表格数据
- ✅ 读取单元格内容
- ✅ 获取表格信息

### 列表操作
- ✅ 添加无序列表
- ✅ 添加有序列表

### 图片操作
- ✅ 插入图片（支持设置宽高）

### 样式格式
- ✅ 插入分页符
- ✅ 设置页面边距

### 高级功能
- ✅ 添加脚注
- ✅ 获取文档大纲结构
- ✅ 添加页眉
- ✅ 添加页脚
- ✅ 生成标准格式的接口文档

## 🚀 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 集成到 AI 编程工具

本 MCP 服务器支持所有兼容 MCP 协议的 AI 编程工具。以下是常见工具的配置方法：

#### 1️⃣ Cursor

Cursor 是最流行的 AI 编程工具之一，内置 MCP 支持。

**配置方法：**
1. 打开 Cursor 设置（`Ctrl/Cmd + ,`）
2. 搜索 "MCP" 或找到 "Model Context Protocol"
3. 添加服务器配置：

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

#### 2️⃣ GitHub Copilot (VSCode)

GitHub Copilot 通过 VSCode 扩展支持 MCP。

**配置文件：** `.vscode/settings.json`

```json
{
  "github.copilot.mcp.servers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

#### 3️⃣ Cline (VSCode 插件)

Cline 是 VSCode 中流行的 AI 助手插件。

**配置方法：**
1. 在 VSCode 中安装 Cline 插件
2. 打开 Cline 设置 → MCP Servers
3. 添加配置：

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

#### 4️⃣ Windsurf (Codeium)

Windsurf 是 Codeium 推出的 AI IDE。

**配置文件位置：**
- Windows: `%APPDATA%\Windsurf\mcp_config.json`
- macOS: `~/Library/Application Support/Windsurf/mcp_config.json`

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

#### 5️⃣ Claude Desktop / Claude Code CLI

**配置文件位置：**
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Linux: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

#### 6️⃣ Augment Code

Augment Code 支持通过配置文件添加 MCP 服务器。

**配置文件：** `~/.augment/mcp_servers.json`

```json
{
  "servers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

#### 7️⃣ 其他支持 MCP 的工具

任何支持 MCP 协议的工具都可以通过类似方式集成。关键配置项：
- **command**: `python`
- **args**: `["-m", "src.server"]`
- **cwd**: 项目根目录的绝对路径

> 💡 **提示**:
> - 将 `cwd` 路径替换为你实际克隆项目的路径
> - Windows 路径使用双反斜杠 `\\` 或正斜杠 `/`
> - 确保 Python 3.10+ 已安装并在 PATH 中
> - 更多配置示例请参考 [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md)

### 验证安装

重启 AI 工具后，你应该能看到 `doc-mcp-server` 已连接。可以尝试以下命令测试：

```
创建一个名为 test.docx 的文档，添加标题"测试文档"
```

## 📖 使用示例

### 示例1：创建文档并添加内容

```python
# 1. 创建新文档
create_document(
    filename="report.docx",
    title="项目报告",
    author="张三"
)

# 2. 添加标题
add_heading(
    filename="report.docx",
    text="项目概述",
    level=1
)

# 3. 添加段落
add_paragraph(
    filename="report.docx",
    text="这是项目的详细描述...",
    font_name="宋体",
    font_size=12
)
```

### 示例2：创建表格

```python
# 创建3行4列的表格
add_table(
    filename="report.docx",
    rows=3,
    cols=4,
    data=[
        ["姓名", "年龄", "部门", "职位"],
        ["张三", "28", "技术部", "工程师"],
        ["李四", "32", "产品部", "经理"]
    ]
)

# 格式化表格
format_table(
    filename="report.docx",
    table_index=0,
    has_header_row=True
)
```

### 示例3：查找和替换文本

```python
# 查找文本
find_text(
    filename="report.docx",
    text_to_find="旧公司名",
    match_case=True
)

# 替换文本
replace_text(
    filename="report.docx",
    find_text="旧公司名",
    replace_text="新公司名"
)
```

### 示例4：生成接口文档

```python
# 插入标准格式的接口文档
insert_interface_doc(
    filename="api_doc.docx",
    position=0,
    name="用户登录接口",
    path="/api/v1/login",
    description="用户通过用户名和密码登录系统",
    method="POST",
    request_params=[
        ["username", "string", "是", "用户名"],
        ["password", "string", "是", "密码"]
    ],
    response_params=[
        ["code", "int", "是", "状态码"],
        ["message", "string", "是", "提示信息"],
        ["data", "object", "是", "返回数据"]
    ]
)
```

## 🛠️ 可用工具列表

### 文档基础操作

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `create_document` | 创建新文档 | filename |
| `get_document_info` | 获取文档信息 | filename |
| `get_document_text` | 提取文档文本 | filename |
| `list_available_documents` | 列出目录下的文档 | directory |
| `copy_document` | 复制文档 | source_filename |

### 内容编辑

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `add_paragraph` | 添加段落 | filename, text |
| `add_heading` | 添加标题 | filename, text |
| `delete_paragraph` | 删除段落 | filename, paragraph_index |
| `find_text` | 查找文本 | filename, text_to_find |
| `replace_text` | 替换文本 | filename, find_text, replace_text |

### 表格操作

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `add_table` | 创建表格 | filename, rows, cols |
| `set_table_cell_content` | 设置单元格内容 | filename, table_index, row_index, col_index, text |
| `format_table` | 格式化表格 | filename, table_index |
| `insert_table_row` | 插入表格行 | filename, table_index, row_index |
| `insert_table_column` | 插入表格列 | filename, table_index, col_index |
| `delete_table_row` | 删除表格行 | filename, table_index, row_index |
| `delete_table_column` | 删除表格列 | filename, table_index, col_index |
| `delete_table` | 删除整个表格 | filename, table_index |
| `merge_table_cells` | 合并单元格 | filename, table_index, start_row, start_col, end_row, end_col |
| `set_cell_alignment` | 设置单元格对齐 | filename, table_index, row_index, col_index |
| `set_cell_background` | 设置单元格背景色 | filename, table_index, row_index, col_index, color |
| `set_cell_padding` | 设置单元格内边距 | filename, table_index, row_index, col_index |
| `set_row_height` | 设置行高 | filename, table_index, row_index, height |
| `set_column_width` | 设置列宽 | filename, table_index, col_index, width |
| `format_cell_text` | 格式化单元格文本 | filename, table_index, row_index, col_index |
| `set_table_indent` | 设置表格缩进 | filename, table_index, indent |

### 列表操作

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `add_bullet_list` | 添加无序列表 | filename, items |
| `add_numbered_list` | 添加有序列表 | filename, items |

### 图片操作

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `insert_image` | 插入图片 | filename, image_path |

### 样式格式

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `add_page_break` | 插入分页符 | filename |
| `set_page_margins` | 设置页面边距 | filename |

### 高级功能

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `add_footnote` | 添加脚注 | filename, paragraph_index, footnote_text |
| `get_document_outline` | 获取文档大纲结构 | filename |
| `add_header` | 添加页眉 | filename, text |
| `add_footer` | 添加页脚 | filename, text |
| `insert_interface_doc` | 插入标准格式的接口文档 | filename, position, name, path, description |

### 数据读取

| 工具名称 | 描述 | 必需参数 |
|---------|------|---------|
| `get_paragraph_text` | 获取指定段落文本 | filename, paragraph_index |
| `get_paragraph_range_text` | 获取指定范围段落文本 | filename, start_index, end_index |
| `get_table_data` | 读取整个表格数据 | filename, table_index |
| `get_table_cell_content` | 读取指定单元格内容 | filename, table_index, row_index, col_index |
| `get_table_info` | 获取表格基本信息 | filename, table_index |

## 📁 项目结构

```
doc-mcp-server/
├── src/
│   ├── __init__.py
│   ├── server.py              # MCP服务主入口
│   ├── tools/                 # MCP工具模块
│   │   ├── __init__.py
│   │   ├── document_basic.py  # 文档基础操作
│   │   ├── content_edit.py    # 内容编辑
│   │   ├── table_ops.py       # 表格操作
│   │   ├── style_format.py    # 样式格式
│   │   ├── image_ops.py       # 图片操作
│   │   ├── list_ops.py        # 列表操作
│   │   ├── advanced.py        # 高级功能
│   │   └── interface_doc.py   # 接口文档生成
│   └── utils/                 # 工具函数
│       ├── __init__.py
│       ├── docx_helper.py     # 文档管理器
│       └── error_handler.py   # 错误处理
├── requirements.txt           # Python依赖
├── pyproject.toml            # 项目配置
├── CLAUDE.md                 # Claude Code 开发指南
├── LICENSE                   # MIT许可证
└── README.md                 # 项目说明
```

## 🔧 开发指南

### 技术栈
- **Python**: 3.10+
- **MCP SDK**: 0.9.0+ (Model Context Protocol)
- **python-docx**: 1.1.0+ (Word 文档操作)
- **Pillow**: 10.0.0+ (图片处理)

### 核心设计

**DocumentManager 无缓存设计**：
- 每次操作都从磁盘重新加载文档
- 确保多个操作之间的数据一致性
- 避免缓存导致的数据丢失问题

### 添加新工具

1. **在对应的工具模块中实现函数**
   ```python
   async def your_new_tool(filename: str, **kwargs):
       doc_manager = DocumentManager()
       doc = doc_manager.get_or_open(filename)
       # ... 执行操作 ...
       doc_manager.save(filename, doc)
       return {"success": True, "message": "操作成功"}
   ```

2. **在 `src/server.py` 中注册工具**
   - 在 `list_tools()` 函数中添加 `Tool` 定义
   - 定义工具名称、描述和 `inputSchema`

3. **在 `call_tool()` 函数中添加调用处理**
   - 添加 `elif name == "your_tool_name":` 分支

详细开发指南请参考 [CLAUDE.md](CLAUDE.md)。

## 📝 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

## 🤝 贡献

欢迎贡献代码！请遵循以下步骤：

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## ❓ 常见问题

**Q: 如何处理中文字体？**
A: 使用 `font_name` 参数指定中文字体，如 "宋体"、"微软雅黑" 等。

**Q: 颜色格式是什么？**
A: 使用 6 位十六进制 RGB 值，如 'FF0000' 表示红色，'00FF00' 表示绿色。

**Q: 单位是什么？**
A: 页边距、行高、列宽使用英寸（inches），字号使用磅（pt）。

**Q: 如何调试 MCP 服务器？**
A: 可以直接运行 `python -m src.server` 查看日志输出。

## 📮 联系方式

如有问题或建议，请：
- 提交 [Issue](https://github.com/hwc2357300448/doc-mcp-server/issues)
- 发起 [Discussion](https://github.com/hwc2357300448/doc-mcp-server/discussions)

## 🙏 致谢

- [Model Context Protocol](https://github.com/modelcontextprotocol) - MCP 协议
- [python-docx](https://python-docx.readthedocs.io/) - Word 文档操作库
- [Claude Code](https://claude.ai/code) - AI 编程助手

---

<div align="center">
Made with ❤️ by the community
</div>
