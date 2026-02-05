"""高级功能工具 - 脚注、尾注等"""
import os
from typing import Optional, Dict, Any
from ..utils import DocumentManager, validate_file_path, handle_docx_errors

# 全局文档管理器实例
doc_manager = DocumentManager()


@handle_docx_errors
async def add_footnote(
    filename: str,
    paragraph_index: int,
    footnote_text: str
) -> Dict[str, Any]:
    """
    在指定段落添加脚注

    参数:
        filename: 文档路径
        paragraph_index: 段落索引（从0开始）
        footnote_text: 脚注文本内容
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path, reload=True)

    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        raise ValueError(f"段落索引超出范围: {paragraph_index}")

    # 注意：python-docx对脚注的支持有限，这里提供基础实现
    # 实际使用中可能需要直接操作XML
    para = doc.paragraphs[paragraph_index]

    # 添加脚注标记（简化实现）
    para.add_run(f" [{footnote_text}]")

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": f"脚注添加成功（段落{paragraph_index}）",
        "footnote_text": footnote_text
    }


@handle_docx_errors
async def get_document_outline(filename: str) -> Dict[str, Any]:
    """
    获取文档大纲结构（标题层级）

    参数:
        filename: 文档路径
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path, reload=True)

    outline = []
    for i, para in enumerate(doc.paragraphs):
        if para.style.name.startswith('Heading'):
            try:
                level = int(para.style.name.split()[-1])
                outline.append({
                    "paragraph_index": i,
                    "level": level,
                    "text": para.text,
                    "style": para.style.name
                })
            except (ValueError, IndexError):
                pass

    return {
        "success": True,
        "outline_count": len(outline),
        "outline": outline
    }


@handle_docx_errors
async def add_header(
    filename: str,
    text: str
) -> Dict[str, Any]:
    """
    添加页眉

    参数:
        filename: 文档路径
        text: 页眉文本内容
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path, reload=True)

    # 为所有节添加页眉
    for section in doc.sections:
        header = section.header
        header.paragraphs[0].text = text

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": "页眉添加成功",
        "text": text
    }


@handle_docx_errors
async def add_footer(
    filename: str,
    text: str
) -> Dict[str, Any]:
    """
    添加页脚

    参数:
        filename: 文档路径
        text: 页脚文本内容
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path, reload=True)

    # 为所有节添加页脚
    for section in doc.sections:
        footer = section.footer
        footer.paragraphs[0].text = text

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": "页脚添加成功",
        "text": text
    }
