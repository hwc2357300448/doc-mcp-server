"""样式格式工具"""
import os
from typing import Optional, Dict, Any
from docx.shared import Pt, Inches
from docx.enum.text import WD_BREAK
from ..utils import DocumentManager, validate_file_path, handle_docx_errors

# 全局文档管理器实例
doc_manager = DocumentManager()


@handle_docx_errors
async def add_page_break(filename: str) -> Dict[str, Any]:
    """
    插入分页符

    参数:
        filename: 文档路径
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path)

    # 添加分页符
    doc.add_page_break()

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": "分页符添加成功"
    }


@handle_docx_errors
async def set_page_margins(
    filename: str,
    top: float = 1.0,
    bottom: float = 1.0,
    left: float = 1.0,
    right: float = 1.0
) -> Dict[str, Any]:
    """
    设置页面边距

    参数:
        filename: 文档路径
        top: 上边距（英寸，默认1.0）
        bottom: 下边距（英寸，默认1.0）
        left: 左边距（英寸，默认1.0）
        right: 右边距（英寸，默认1.0）
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path)

    # 设置页边距
    for section in doc.sections:
        section.top_margin = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin = Inches(left)
        section.right_margin = Inches(right)

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": f"页边距设置成功（上:{top}\" 下:{bottom}\" 左:{left}\" 右:{right}\"）"
    }
