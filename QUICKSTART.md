# 快速开始指南

本指南将帮助你在 5 分钟内配置并使用 doc-mcp-server。

## 📦 第一步：安装

### 1. 克隆仓库

```bash
git clone https://github.com/hwc2357300448/doc-mcp-server.git
cd doc-mcp-server
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 验证安装

```bash
python -m src.server
```

如果看到服务器启动信息，说明安装成功！按 `Ctrl+C` 停止服务器。

## ⚙️ 第二步：配置 AI 工具

选择你使用的 AI 编程工具进行配置：

### Claude Desktop / Claude Code CLI

1. 找到配置文件：
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Linux: `~/.config/Claude/claude_desktop_config.json`

2. 添加配置（记得替换路径）：

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/your/path/to/doc-mcp-server"
    }
  }
}
```

3. 重启 Claude Desktop

### Cline (VSCode)

1. 在 VSCode 中安装 Cline 插件
2. 打开 Cline 设置 → MCP Servers
3. 添加上述相同的配置
4. 重启 VSCode

## 🎯 第三步：测试

在 AI 工具中输入以下命令测试：

```
创建一个名为 test.docx 的文档，添加标题"我的第一个文档"，然后添加一段文字"这是测试内容"
```

如果成功创建了文档，恭喜你已经配置成功！

## 📚 第四步：学习更多

### 常用操作示例

#### 创建带表格的报告

```
创建一个 report.docx 文档，添加标题"销售报告"，然后创建一个 3 行 4 列的表格，
表头为：产品、数量、单价、总价
```

#### 批量添加内容

```
在 report.docx 中批量添加以下段落：
1. 第一季度销售情况良好
2. 第二季度持续增长
3. 第三季度达到峰值
```

#### 格式化文档

```
将 report.docx 中的所有"重要"文字设置为红色粗体
```

### 查看完整功能

- 📖 [README.md](README.md) - 完整功能列表
- 🔧 [CLAUDE.md](CLAUDE.md) - 开发者指南
- 📝 [CONFIG_EXAMPLES.md](CONFIG_EXAMPLES.md) - 更多配置示例

## ❓ 常见问题

### Q: 配置后看不到 MCP 服务器？

**A:** 检查以下几点：
1. 路径是否为绝对路径
2. Python 是否在 PATH 中
3. 是否重启了 AI 工具
4. 查看 AI 工具的日志输出

### Q: 提示找不到模块？

**A:** 确保已安装依赖：
```bash
pip install -r requirements.txt
```

### Q: Windows 路径配置问题？

**A:** Windows 路径使用以下格式之一：
- `"D:\\workspace\\doc-mcp-server"` (双反斜杠)
- `"D:/workspace/doc-mcp-server"` (正斜杠)

### Q: 如何使用虚拟环境？

**A:** 在配置中指定虚拟环境的 Python：
```json
{
  "command": "/path/to/venv/bin/python",
  "args": ["-m", "src.server"],
  "cwd": "/path/to/doc-mcp-server"
}
```

## 🎉 下一步

现在你已经成功配置了 doc-mcp-server，可以：

1. 尝试创建各种类型的文档
2. 探索表格和格式化功能
3. 使用接口文档生成功能
4. 查看 [示例](README.md#-使用示例) 了解更多用法

祝你使用愉快！如有问题，欢迎提交 [Issue](https://github.com/hwc2357300448/doc-mcp-server/issues)。
