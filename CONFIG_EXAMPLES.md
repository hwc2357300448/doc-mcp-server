# MCP 服务器配置示例

本文件提供了在不同 AI 编程工具中配置 doc-mcp-server 的详细示例。

## 🎯 Cursor

Cursor 是最流行的 AI 编程工具之一，配置非常简单。

**配置方法：**
1. 打开 Cursor 设置（`Ctrl/Cmd + ,`）
2. 搜索 "MCP" 或找到 "Model Context Protocol"
3. 点击 "Edit in settings.json"
4. 添加以下配置：

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

**或者在项目根目录创建 `.cursor/mcp_config.json`：**

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

## 🐙 GitHub Copilot (VSCode)

GitHub Copilot 通过 VSCode 扩展支持 MCP。

**方法 1：项目级配置**

在项目根目录创建或编辑 `.vscode/settings.json`：

```json
{
  "github.copilot.mcp.servers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

**方法 2：全局配置**

打开 VSCode 设置（`Ctrl/Cmd + ,`），搜索 "copilot mcp"，添加相同配置。

## 🤖 Cline (VSCode 插件)

Cline 是 VSCode 中流行的 AI 助手插件。

**配置步骤：**
1. 在 VSCode 中安装 Cline 插件
2. 点击 Cline 图标打开侧边栏
3. 点击设置图标 ⚙️
4. 找到 "MCP Servers" 部分
5. 点击 "Edit Config" 或直接编辑配置文件

**配置内容：**

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

## 🌊 Windsurf (Codeium)

Windsurf 是 Codeium 推出的 AI IDE。

**配置文件位置：**
- Windows: `%APPDATA%\Windsurf\mcp_config.json`
- macOS: `~/Library/Application Support/Windsurf/mcp_config.json`
- Linux: `~/.config/Windsurf/mcp_config.json`

**配置内容：**

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

## 🚀 Augment Code

Augment Code 是新兴的 AI 编程助手。

**配置文件位置：**
- `~/.augment/mcp_servers.json`

**配置内容：**

```json
{
  "servers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

## 🎨 Claude Desktop / Claude Code CLI

**配置文件位置：**
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Linux: `~/.config/Claude/claude_desktop_config.json`

**配置内容：**

```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "/absolute/path/to/doc-mcp-server"
    }
  }
}
```

## 使用虚拟环境

如果你使用 Python 虚拟环境，需要指定虚拟环境中的 Python 解释器：

### Windows
```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "D:\\path\\to\\doc-mcp-server\\venv\\Scripts\\python.exe",
      "args": ["-m", "src.server"],
      "cwd": "D:\\path\\to\\doc-mcp-server"
    }
  }
}
```

### macOS/Linux
```json
{
  "mcpServers": {
    "doc-mcp-server": {
      "command": "/path/to/doc-mcp-server/venv/bin/python",
      "args": ["-m", "src.server"],
      "cwd": "/path/to/doc-mcp-server"
    }
  }
}
```

## 注意事项

1. **路径必须是绝对路径**，不能使用相对路径或 `~`
2. **Windows 路径**使用双反斜杠 `\\` 或正斜杠 `/`
3. **确保 Python 3.10+** 已安装
4. **安装依赖**：`pip install -r requirements.txt`
5. **重启 AI 工具**以使配置生效
