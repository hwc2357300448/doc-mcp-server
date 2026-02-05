"""图片操作工具"""
import os
from typing import Optional, Dict, Any
from docx.shared import Inches
from ..utils import DocumentManager, validate_file_path, handle_docx_errors

# 全局文档管理器实例
doc_manager = DocumentManager()


@handle_docx_errors
async def insert_image(
    filename: str,
    image_path: str,
    width: Optional[float] = None,
    height: Optional[float] = None
) -> Dict[str, Any]:
    """
    插入图片到Word文档

    参数:
        filename: 文档路径
        image_path: 图片文件路径
        width: 图片宽度（英寸，可选）
        height: 图片高度（英寸，可选）
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path)

    if not os.path.exists(image_path):
        raise FileNotFoundError(f"图片文件不存在: {image_path}")

    # 插入图片
    if width and height:
        doc.add_picture(image_path, width=Inches(width), height=Inches(height))
    elif width:
        doc.add_picture(image_path, width=Inches(width))
    elif height:
        doc.add_picture(image_path, height=Inches(height))
    else:
        doc.add_picture(image_path)

    doc_manager.save(abs_path, doc)

    return {
        "success": True,
        "message": "图片插入成功",
        "image_path": image_path
    }
