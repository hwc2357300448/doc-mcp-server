# 贡献指南

感谢你考虑为 doc-mcp-server 做出贡献！

## 如何贡献

### 报告 Bug

如果你发现了 bug，请创建一个 Issue 并包含以下信息：

- 清晰的标题和描述
- 重现步骤
- 预期行为和实际行为
- 你的环境信息（操作系统、Python 版本等）
- 如果可能，提供最小可复现示例

### 提出新功能

如果你有新功能的想法：

1. 先创建一个 Issue 讨论这个功能
2. 说明为什么需要这个功能
3. 描述你期望的行为
4. 等待维护者反馈后再开始开发

### 提交代码

1. **Fork 仓库**
   ```bash
   git clone https://github.com/hwc2357300448/doc-mcp-server.git
   cd doc-mcp-server
   ```

2. **创建分支**
   ```bash
   git checkout -b feature/your-feature-name
   ```

3. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

4. **进行修改**
   - 遵循现有代码风格
   - 添加必要的注释
   - 确保所有函数都是 async 函数
   - 使用统一的错误处理机制

5. **测试你的修改**
   ```bash
   python -m src.server
   ```

6. **提交更改**
   ```bash
   git add .
   git commit -m "feat: 添加新功能描述"
   ```

7. **推送到 GitHub**
   ```bash
   git push origin feature/your-feature-name
   ```

8. **创建 Pull Request**
   - 提供清晰的 PR 标题和描述
   - 关联相关的 Issue
   - 等待代码审查

## 代码规范

### Python 代码风格

- 使用 4 个空格缩进
- 遵循 PEP 8 规范
- 函数和变量使用下划线命名法（snake_case）
- 类名使用驼峰命名法（PascalCase）

### 提交信息规范

使用语义化提交信息：

- `feat:` 新功能
- `fix:` 修复 bug
- `docs:` 文档更新
- `style:` 代码格式调整
- `refactor:` 代码重构
- `test:` 测试相关
- `chore:` 构建或辅助工具的变动

示例：
```
feat: 添加表格单元格合并功能
fix: 修复文档保存时的编码问题
docs: 更新 README 中的配置说明
```

## 添加新工具

如果你要添加新的 MCP 工具，请遵循以下步骤：

1. **在对应的工具模块中实现函数**
   ```python
   # src/tools/your_module.py
   from ..utils.docx_helper import DocumentManager

   async def your_new_tool(filename: str, **kwargs):
       """工具描述"""
       doc_manager = DocumentManager()
       doc = doc_manager.get_or_open(filename)

       # 执行操作
       # ...

       doc_manager.save(filename, doc)
       return {"success": True, "message": "操作成功"}
   ```

2. **在 server.py 中注册工具**
   - 在 `list_tools()` 中添加 Tool 定义
   - 在 `call_tool()` 中添加调用处理

3. **更新文档**
   - 在 README.md 中添加工具说明
   - 在 CLAUDE.md 中更新开发指南

## 问题和讨论

- 使用 Issues 报告 bug 和提出功能请求
- 使用 Discussions 进行一般性讨论
- 保持友好和尊重的态度

## 许可证

通过贡献代码，你同意你的贡献将在 MIT 许可证下发布。
