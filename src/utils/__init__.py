"""工具辅助函数模块"""
from .docx_helper import DocumentManager, validate_file_path
from .error_handler import handle_docx_errors, DocxError

__all__ = [
    "DocumentManager",
    "validate_file_path",
    "handle_docx_errors",
    "DocxError",
]
