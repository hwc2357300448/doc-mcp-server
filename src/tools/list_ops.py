"""列表操作工具"""
import os
from typing import List, Dict, Any, Optional
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from ..utils import DocumentManager, validate_file_path, handle_docx_errors

# 全局文档管理器实例
doc_manager = DocumentManager()


@handle_docx_errors
async def add_bullet_list(
    filename: str,
    items: List[str],
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: bool = False,
    italic: bool = False,
    color: Optional[str] = None,
    highlight: Optional[str] = None
) -> Dict[str, Any]:
    """
    添加无序列表（项目符号列表）

    参数:
        filename: 文档路径
        items: 列表项内容列表
        font_name: 字体名称（可选）
        font_size: 字号（可选）
        bold: 是否粗体（可选）
        italic: 是否斜体（可选）
        color: 文字颜色，十六进制RGB（可选）
        highlight: 背景色（高亮），十六进制RGB（可选）
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path)

    if not items:
        raise ValueError("列表项不能为空")

    # 添加无序列表
    for item in items:
        para = doc.add_paragraph(item, style='List Bullet')

        # 设置字体格式
        if any([font_name, font_size, bold, italic, color, highlight]):
            for run in para.runs:
                if font_name:
                    run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if bold:
                    run.font.bold = bold
                if italic:
                    run.font.italic = italic
                if color:
                    run.font.color.rgb = RGBColor(
                        int(color[0:2], 16),
                        int(color[2:4], 16),
                        int(color[4:6], 16)
                    )
                if highlight:
                    from docx.oxml.shared import OxmlElement
                    from docx.oxml.ns import qn
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), highlight)
                    run._element.get_or_add_rPr().append(shd)

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": f"无序列表添加成功，共{len(items)}项",
        "item_count": len(items)
    }


@handle_docx_errors
async def add_numbered_list(
    filename: str,
    items: List[str],
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    bold: bool = False,
    italic: bool = False,
    color: Optional[str] = None,
    highlight: Optional[str] = None
) -> Dict[str, Any]:
    """
    添加有序列表（编号列表）

    参数:
        filename: 文档路径
        items: 列表项内容列表
        font_name: 字体名称（可选）
        font_size: 字号（可选）
        bold: 是否粗体（可选）
        italic: 是否斜体（可选）
        color: 文字颜色，十六进制RGB（可选）
        highlight: 背景色（高亮），十六进制RGB（可选）
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path)

    if not items:
        raise ValueError("列表项不能为空")

    # 添加有序列表
    for item in items:
        para = doc.add_paragraph(item, style='List Number')

        # 设置字体格式
        if any([font_name, font_size, bold, italic, color, highlight]):
            for run in para.runs:
                if font_name:
                    run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if bold:
                    run.font.bold = bold
                if italic:
                    run.font.italic = italic
                if color:
                    run.font.color.rgb = RGBColor(
                        int(color[0:2], 16),
                        int(color[2:4], 16),
                        int(color[4:6], 16)
                    )
                if highlight:
                    from docx.oxml.shared import OxmlElement
                    from docx.oxml.ns import qn
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:fill'), highlight)
                    run._element.get_or_add_rPr().append(shd)

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": f"有序列表添加成功，共{len(items)}项",
        "item_count": len(items)
    }
