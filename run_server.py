"""MCP服务器启动脚本 - 解决相对导入问题"""
import sys
import os

# 将项目根目录添加到 Python 路径
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

# 现在可以导入并运行服务器
if __name__ == "__main__":
    import asyncio
    from src.server import main
    asyncio.run(main())
