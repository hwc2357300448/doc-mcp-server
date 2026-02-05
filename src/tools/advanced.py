"""高级功能工具 - 脚注、尾注等"""
import os
import re
import docx2txt
from typing import Optional, Dict, Any, List
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


@handle_docx_errors
async def get_headings_list(filename: str) -> Dict[str, Any]:
    """
    获取文档中所有标题的详细列表（尽可能包含自动编号）

    参数:
        filename: 文档路径

    返回:
        包含所有标题的详细信息列表，包括层级、标题名称、所在段落索引

    注意：由于 python-docx 的限制，自动编号的提取依赖于 docx2txt，
         对于某些复杂的编号格式（如中文编号）可能无法完全准确提取
    """
    abs_path = validate_file_path(filename)

    # 使用 docx2txt 提取完整文本（包含编号）
    full_text = docx2txt.process(abs_path)
    text_lines = [line.strip() for line in full_text.split('\n') if line.strip()]

    # 使用 python-docx 获取段落样式信息
    doc = doc_manager.get_or_open(abs_path, reload=True)

    headings = []
    used_line_indices = set()

    for para_idx, para in enumerate(doc.paragraphs):
        if para.style.name.startswith('Heading'):
            try:
                level = int(para.style.name.split()[-1])
                heading_text = para.text.strip()

                # 在 docx2txt 提取的文本中查找最佳匹配
                best_match = _find_best_match(heading_text, text_lines, used_line_indices)

                headings.append({
                    "level": level,
                    "text": best_match,
                    "paragraph_index": para_idx,
                    "style": para.style.name
                })
            except (ValueError, IndexError):
                headings.append({
                    "level": 0,
                    "text": para.text.strip(),
                    "paragraph_index": para_idx,
                    "style": para.style.name
                })

    return {
        "success": True,
        "count": len(headings),
        "headings": headings
    }


def _find_best_match(heading_text: str, text_lines: List[str], used_indices: set) -> str:
    """
    在文本行中查找与标题最匹配的行（包含编号）

    参数:
        heading_text: 标题文本（不含编号）
        text_lines: 从 docx2txt 提取的所有文本行
        used_indices: 已使用的行索引集合

    返回:
        最佳匹配的文本（可能包含编号）
    """
    if not heading_text:
        return heading_text

    best_match = heading_text
    best_score = 0
    best_idx = -1

    for idx, line in enumerate(text_lines):
        # 不跳过已使用的索引，因为可能有多个标题匹配同一行
        # if idx in used_indices:
        #     continue

        # 计算匹配分数
        score = _calculate_match_score(heading_text, line)

        if score > best_score:
            best_score = score
            best_match = line
            best_idx = idx

    # 如果找到了好的匹配，标记为已使用
    if best_idx >= 0 and best_score > 0.3:
        used_indices.add(best_idx)

    return best_match


def _calculate_match_score(heading_text: str, line: str) -> float:
    """
    计算标题文本与行的匹配分数

    参数:
        heading_text: 标题文本
        line: 文本行

    返回:
        匹配分数（0-1之间）
    """
    # 如果行不包含标题文本，不匹配
    if heading_text not in line:
        return 0

    # 检查多种编号格式
    # 1. 数字编号：1. 1.1. 1.1.1. 等
    # 2. 中文编号：一、二、三、等
    # 3. 括号编号：(1) （1） [1] 等
    numbering_patterns = [
        r'^[\d]+\.[\d]*\.?[\d]*\.?\s*',  # 1. 1.1. 1.1.1.
        r'^[一二三四五六七八九十百千]+[、\.]\s*',  # 一、二、
        r'^[\(（\[][\d]+[\)）\]]\s*',  # (1) （1） [1]
        r'^\d+\s+',  # 纯数字后跟空格
    ]

    has_numbering = any(re.match(pattern, line) for pattern in numbering_patterns)

    # 如果有编号，给予更高的优先级
    if has_numbering:
        base_score = len(heading_text) / len(line) if len(line) > 0 else 0
        return base_score + 0.5

    # 如果是完全匹配（没有编号），给予较低分数
    if line == heading_text:
        return 0.3

    # 其他情况，根据相似度计算
    if len(line) > 0:
        return len(heading_text) / len(line)

    return 0

