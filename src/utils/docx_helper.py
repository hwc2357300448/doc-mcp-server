"""Word文档操作辅助函数"""
import os
from pathlib import Path
from typing import Optional
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from .error_handler import DocxError


class DocumentManager:
    """文档管理器 - 无缓存版本，每次都重新加载文件"""

    def __init__(self):
        pass  # 不再维护缓存

    def get_or_open(self, filename: str, reload: bool = False) -> Document:
        """打开文档（每次都重新加载）

        参数:
            filename: 文件路径
            reload: 兼容参数，已废弃
        """
        abs_path = os.path.abspath(filename)

        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"文件不存在: {abs_path}")

        return Document(abs_path)

    def create_new(self, filename: str) -> Document:
        """创建新文档"""
        return Document()

    def save(self, filename: str, doc: Document) -> None:
        """保存文档

        参数:
            filename: 文件路径
            doc: Document对象
        """
        abs_path = os.path.abspath(filename)
        doc.save(abs_path)

    def save_and_close(self, filename: str, doc: Document) -> None:
        """保存并关闭文档（与save相同，因为没有缓存）"""
        self.save(filename, doc)

    def close(self, filename: str) -> None:
        """关闭文档（无操作，因为没有缓存）"""
        pass


def validate_file_path(filename: str) -> str:
    """验证文件路径"""
    if not filename:
        raise ValueError("文件路径不能为空")

    abs_path = os.path.abspath(filename)
    parent_dir = os.path.dirname(abs_path)

    if not os.path.exists(parent_dir):
        raise FileNotFoundError(f"目录不存在: {parent_dir}")

    return abs_path


def set_run_font(run, font_name: Optional[str] = None, font_size: Optional[int] = None,
                 bold: bool = False, italic: bool = False, color: Optional[str] = None):
    """设置文本运行的字体属性"""
    if font_name:
        run.font.name = font_name
    if font_size:
        run.font.size = Pt(font_size)
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if color:
        # 颜色格式: 'FF0000' (红色)
        try:
            rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            run.font.color.rgb = RGBColor(*rgb)
        except (ValueError, IndexError):
            raise ValueError(f"无效的颜色格式: {color}，应为6位十六进制RGB值，如'FF0000'")


def inches_to_pt(inches: float) -> int:
    """英寸转磅"""
    return int(inches * 72)


def cm_to_inches(cm: float) -> float:
    """厘米转英寸"""
    return cm / 2.54
