"""错误处理模块"""
from functools import wraps
from typing import Any, Callable, Dict


class DocxError(Exception):
    """Word文档操作异常基类"""
    pass


def handle_docx_errors(func: Callable) -> Callable:
    """统一处理docx操作异常的装饰器"""
    @wraps(func)
    async def wrapper(*args, **kwargs) -> Dict[str, Any]:
        try:
            result = await func(*args, **kwargs)
            if isinstance(result, dict) and "success" not in result:
                result["success"] = True
            return result
        except FileNotFoundError as e:
            filename = kwargs.get('filename', 'unknown')
            return {
                "success": False,
                "error": "FileNotFound",
                "message": f"文件不存在: {filename}",
                "suggestion": "请检查文件路径是否正确",
                "details": str(e)
            }
        except PermissionError as e:
            return {
                "success": False,
                "error": "PermissionDenied",
                "message": "没有文件访问权限",
                "suggestion": "请检查文件是否被其他程序占用或权限设置",
                "details": str(e)
            }
        except ValueError as e:
            return {
                "success": False,
                "error": "ValueError",
                "message": f"参数值错误: {str(e)}",
                "suggestion": "请检查输入参数是否符合要求",
                "details": str(e)
            }
        except DocxError as e:
            return {
                "success": False,
                "error": "DocxError",
                "message": str(e),
                "suggestion": "请查看错误详情",
                "details": str(e)
            }
        except Exception as e:
            return {
                "success": False,
                "error": type(e).__name__,
                "message": f"操作失败: {str(e)}",
                "suggestion": "请查看详细错误信息或联系技术支持",
                "details": str(e)
            }
    return wrapper
