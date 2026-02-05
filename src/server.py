"""Word文档编辑MCP服务主入口"""
import asyncio
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# 导入工具函数
from .tools import document_basic, content_edit, table_ops, style_format, image_ops, list_ops, advanced, interface_doc

# 创建MCP服务器实例
app = Server("doc-mcp-server")


# 注册工具列表
@app.list_tools()
async def list_tools() -> list[Tool]:
    """列出所有可用的工具"""
    return [
        # 文档基础操作
        Tool(
            name="create_document",
            description="创建新的Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档保存路径"},
                    "title": {"type": "string", "description": "文档标题（可选）"},
                    "author": {"type": "string", "description": "作者（可选）"},
                    "subject": {"type": "string", "description": "主题（可选）"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="get_document_info",
            description="获取文档信息和元数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="get_document_text",
            description="提取文档的全部文本内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="list_available_documents",
            description="列出指定目录下的所有Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "directory": {"type": "string", "description": "目录路径（默认为当前目录）"}
                }
            }
        ),
        Tool(
            name="copy_document",
            description="复制Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "source_filename": {"type": "string", "description": "源文档路径"},
                    "destination_filename": {"type": "string", "description": "目标文档路径（可选）"}
                },
                "required": ["source_filename"]
            }
        ),
        # 内容编辑工具
        Tool(
            name="add_paragraph",
            description="添加段落到Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "text": {"type": "string", "description": "段落文本内容"},
                    "style": {"type": "string", "description": "段落样式名称（可选）"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"},
                    "first_line_indent": {"type": "number", "description": "首行缩进，单位厘米（可选）"},
                    "left_indent": {"type": "number", "description": "左缩进，单位厘米（可选）"},
                    "right_indent": {"type": "number", "description": "右缩进，单位厘米（可选）"},
                    "alignment": {"type": "string", "description": "对齐方式：left/center/right/justify（可选）"}
                },
                "required": ["filename", "text"]
            }
        ),
        Tool(
            name="add_heading",
            description="添加标题到Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "text": {"type": "string", "description": "标题文本"},
                    "level": {"type": "integer", "description": "标题级别（1-9）"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"},
                    "first_line_indent": {"type": "number", "description": "首行缩进，单位厘米（可选）"},
                    "left_indent": {"type": "number", "description": "左缩进，单位厘米（可选）"},
                    "right_indent": {"type": "number", "description": "右缩进，单位厘米（可选）"},
                    "alignment": {"type": "string", "description": "对齐方式：left/center/right/justify（可选）"}
                },
                "required": ["filename", "text"]
            }
        ),
        Tool(
            name="batch_add_paragraphs",
            description="批量添加多个段落到Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "paragraphs": {
                        "type": "array",
                        "description": "段落列表",
                        "items": {
                            "type": "object",
                            "properties": {
                                "text": {"type": "string", "description": "段落文本"},
                                "style": {"type": "string", "description": "段落样式（可选）"},
                                "font_name": {"type": "string", "description": "字体名称（可选）"},
                                "font_size": {"type": "integer", "description": "字号（可选）"},
                                "bold": {"type": "boolean", "description": "是否粗体"},
                                "italic": {"type": "boolean", "description": "是否斜体"},
                                "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                                "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"},
                                "first_line_indent": {"type": "number", "description": "首行缩进，单位厘米（可选）"},
                                "left_indent": {"type": "number", "description": "左缩进，单位厘米（可选）"},
                                "right_indent": {"type": "number", "description": "右缩进，单位厘米（可选）"},
                                "alignment": {"type": "string", "description": "对齐方式：left/center/right/justify（可选）"}
                            },
                            "required": ["text"]
                        }
                    }
                },
                "required": ["filename", "paragraphs"]
            }
        ),
        Tool(
            name="delete_paragraph",
            description="删除指定段落",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "paragraph_index": {"type": "integer", "description": "段落索引（从0开始）"}
                },
                "required": ["filename", "paragraph_index"]
            }
        ),
        Tool(
            name="insert_paragraph",
            description="在指定位置插入段落",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "text": {"type": "string", "description": "段落文本内容"},
                    "position": {"type": "integer", "description": "插入位置索引（从0开始，0表示插入到开头）"},
                    "style": {"type": "string", "description": "段落样式名称（可选）"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"},
                    "first_line_indent": {"type": "number", "description": "首行缩进，单位厘米（可选）"},
                    "left_indent": {"type": "number", "description": "左缩进，单位厘米（可选）"},
                    "right_indent": {"type": "number", "description": "右缩进，单位厘米（可选）"},
                    "alignment": {"type": "string", "description": "对齐方式：left/center/right/justify（可选）"}
                },
                "required": ["filename", "text", "position"]
            }
        ),
        Tool(
            name="delete_paragraph_range",
            description="删除指定范围的段落",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "start_index": {"type": "integer", "description": "起始段落索引（从0开始，包含）"},
                    "end_index": {"type": "integer", "description": "结束段落索引（从0开始，包含）"}
                },
                "required": ["filename", "start_index", "end_index"]
            }
        ),
        Tool(
            name="replace_paragraph_range",
            description="替换指定范围的段落为新内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "start_index": {"type": "integer", "description": "起始段落索引（从0开始，包含）"},
                    "end_index": {"type": "integer", "description": "结束段落索引（从0开始，包含）"},
                    "new_text": {"type": "string", "description": "新的段落文本内容"},
                    "style": {"type": "string", "description": "段落样式名称（可选）"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"},
                    "first_line_indent": {"type": "number", "description": "首行缩进，单位厘米（可选）"},
                    "left_indent": {"type": "number", "description": "左缩进，单位厘米（可选）"},
                    "right_indent": {"type": "number", "description": "右缩进，单位厘米（可选）"},
                    "alignment": {"type": "string", "description": "对齐方式：left/center/right/justify（可选）"}
                },
                "required": ["filename", "start_index", "end_index", "new_text"]
            }
        ),
        Tool(
            name="find_text",
            description="在文档中查找文本",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "text_to_find": {"type": "string", "description": "要查找的文本"},
                    "match_case": {"type": "boolean", "description": "是否区分大小写"},
                    "whole_word": {"type": "boolean", "description": "是否全字匹配"}
                },
                "required": ["filename", "text_to_find"]
            }
        ),
        Tool(
            name="replace_text",
            description="替换文档中的文本",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "find_text": {"type": "string", "description": "要查找的文本"},
                    "replace_text": {"type": "string", "description": "替换后的文本"}
                },
                "required": ["filename", "find_text", "replace_text"]
            }
        ),
        # 表格操作工具
        Tool(
            name="add_table",
            description="创建表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "rows": {"type": "integer", "description": "行数"},
                    "cols": {"type": "integer", "description": "列数"},
                    "data": {"type": "array", "description": "表格数据（可选）"}
                },
                "required": ["filename", "rows", "cols"]
            }
        ),
        Tool(
            name="set_table_cell_content",
            description="设置表格单元格内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "col_index": {"type": "integer", "description": "列索引"},
                    "text": {"type": "string", "description": "单元格文本"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"},
                    "alignment": {"type": "string", "description": "对齐方式：left/center/right/justify（可选）"}
                },
                "required": ["filename", "table_index", "row_index", "col_index", "text"]
            }
        ),
        Tool(
            name="format_table",
            description="格式化表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "border_style": {"type": "string", "description": "边框样式（可选）"},
                    "has_header_row": {"type": "boolean", "description": "是否有表头行"}
                },
                "required": ["filename", "table_index"]
            }
        ),
        Tool(
            name="delete_table_row",
            description="删除表格行",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"}
                },
                "required": ["filename", "table_index", "row_index"]
            }
        ),
        Tool(
            name="delete_table_column",
            description="删除表格列",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "col_index": {"type": "integer", "description": "列索引"}
                },
                "required": ["filename", "table_index", "col_index"]
            }
        ),
        Tool(
            name="delete_table",
            description="删除整个表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"}
                },
                "required": ["filename", "table_index"]
            }
        ),
        # 样式格式工具
        Tool(
            name="add_page_break",
            description="插入分页符",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="set_page_margins",
            description="设置页面边距",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "top": {"type": "number", "description": "上边距（英寸）"},
                    "bottom": {"type": "number", "description": "下边距（英寸）"},
                    "left": {"type": "number", "description": "左边距（英寸）"},
                    "right": {"type": "number", "description": "右边距（英寸）"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="merge_table_cells",
            description="合并表格单元格",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "start_row": {"type": "integer", "description": "起始行索引"},
                    "start_col": {"type": "integer", "description": "起始列索引"},
                    "end_row": {"type": "integer", "description": "结束行索引"},
                    "end_col": {"type": "integer", "description": "结束列索引"}
                },
                "required": ["filename", "table_index", "start_row", "start_col", "end_row", "end_col"]
            }
        ),
        Tool(
            name="set_cell_alignment",
            description="设置单元格对齐方式",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "col_index": {"type": "integer", "description": "列索引"},
                    "horizontal": {"type": "string", "description": "水平对齐（left/center/right/justify）"},
                    "vertical": {"type": "string", "description": "垂直对齐（top/center/bottom）"}
                },
                "required": ["filename", "table_index", "row_index", "col_index"]
            }
        ),
        Tool(
            name="set_cell_background",
            description="设置单元格背景色",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "col_index": {"type": "integer", "description": "列索引"},
                    "color": {"type": "string", "description": "颜色（十六进制）"}
                },
                "required": ["filename", "table_index", "row_index", "col_index", "color"]
            }
        ),
        Tool(
            name="set_cell_padding",
            description="设置单元格内边距",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "col_index": {"type": "integer", "description": "列索引"},
                    "top": {"type": "number", "description": "上边距（英寸）"},
                    "bottom": {"type": "number", "description": "下边距（英寸）"},
                    "left": {"type": "number", "description": "左边距（英寸）"},
                    "right": {"type": "number", "description": "右边距（英寸）"}
                },
                "required": ["filename", "table_index", "row_index", "col_index"]
            }
        ),
        Tool(
            name="set_row_height",
            description="设置表格行高",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "height": {"type": "number", "description": "行高（英寸）"}
                },
                "required": ["filename", "table_index", "row_index", "height"]
            }
        ),
        Tool(
            name="format_cell_text",
            description="设置单元格文本格式",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "col_index": {"type": "integer", "description": "列索引"},
                    "font_name": {"type": "string", "description": "字体名称"},
                    "font_size": {"type": "integer", "description": "字号"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色"}
                },
                "required": ["filename", "table_index", "row_index", "col_index"]
            }
        ),
        Tool(
            name="set_table_indent",
            description="设置表格缩进",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "indent": {"type": "number", "description": "缩进距离（英寸）"}
                },
                "required": ["filename", "table_index", "indent"]
            }
        ),
        Tool(
            name="insert_table_column",
            description="在表格中插入新列",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "col_index": {"type": "integer", "description": "插入位置的列索引"}
                },
                "required": ["filename", "table_index", "col_index"]
            }
        ),
        # 图片操作工具
        Tool(
            name="insert_image",
            description="插入图片到Word文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "image_path": {"type": "string", "description": "图片文件路径"},
                    "position": {"type": "integer", "description": "插入位置（段落索引，从0开始），不指定则追加到文档末尾"},
                    "width": {"type": "number", "description": "图片宽度（英寸，可选）"},
                    "height": {"type": "number", "description": "图片高度（英寸，可选）"}
                },
                "required": ["filename", "image_path"]
            }
        ),
        Tool(
            name="delete_image",
            description="删除指定段落中的图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "paragraph_index": {"type": "integer", "description": "包含图片的段落索引（从0开始）"}
                },
                "required": ["filename", "paragraph_index"]
            }
        ),
        # 列表操作工具
        Tool(
            name="add_bullet_list",
            description="添加无序列表（项目符号列表）",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "items": {"type": "array", "items": {"type": "string"}, "description": "列表项内容"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"}
                },
                "required": ["filename", "items"]
            }
        ),
        Tool(
            name="add_numbered_list",
            description="添加有序列表（编号列表）",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "items": {"type": "array", "items": {"type": "string"}, "description": "列表项内容"},
                    "font_name": {"type": "string", "description": "字体名称（可选）"},
                    "font_size": {"type": "integer", "description": "字号（可选）"},
                    "bold": {"type": "boolean", "description": "是否粗体"},
                    "italic": {"type": "boolean", "description": "是否斜体"},
                    "color": {"type": "string", "description": "文字颜色，十六进制RGB（可选）"},
                    "highlight": {"type": "string", "description": "背景色（高亮），十六进制RGB（可选）"}
                },
                "required": ["filename", "items"]
            }
        ),
        Tool(
            name="insert_table_row",
            description="在表格中插入新行",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "插入位置的行索引"}
                },
                "required": ["filename", "table_index", "row_index"]
            }
        ),
        Tool(
            name="set_column_width",
            description="设置表格列宽",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "col_index": {"type": "integer", "description": "列索引"},
                    "width": {"type": "number", "description": "列宽（英寸）"}
                },
                "required": ["filename", "table_index", "col_index", "width"]
            }
        ),
        # 高级功能工具
        Tool(
            name="add_footnote",
            description="在指定段落添加脚注",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "paragraph_index": {"type": "integer", "description": "段落索引"},
                    "footnote_text": {"type": "string", "description": "脚注文本"}
                },
                "required": ["filename", "paragraph_index", "footnote_text"]
            }
        ),
        Tool(
            name="get_document_outline",
            description="获取文档大纲结构（标题层级）",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="add_header",
            description="添加页眉",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "text": {"type": "string", "description": "页眉文本"}
                },
                "required": ["filename", "text"]
            }
        ),
        Tool(
            name="add_footer",
            description="添加页脚",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "text": {"type": "string", "description": "页脚文本"}
                },
                "required": ["filename", "text"]
            }
        ),
        Tool(
            name="get_headings_list",
            description="获取文档中所有标题的简单列表",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"}
                },
                "required": ["filename"]
            }
        ),
        Tool(
            name="get_headings_list_range",
            description="获取文档中特定范围内的标题列表（包含自动编号）",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "start_index": {"type": "integer", "description": "起始段落索引（包含），不指定则从文档开头"},
                    "end_index": {"type": "integer", "description": "结束段落索引（包含），不指定则到文档末尾"},
                    "max_level": {"type": "integer", "description": "最大标题级别（1-9），不指定则返回所有级别"}
                },
                "required": ["filename"]
            }
        ),
        # 数据读取工具
        Tool(
            name="get_paragraph_text",
            description="获取指定段落的文本内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "paragraph_index": {"type": "integer", "description": "段落索引（从0开始）"}
                },
                "required": ["filename", "paragraph_index"]
            }
        ),
        Tool(
            name="get_paragraph_range_text",
            description="获取指定范围段落的文本内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "start_index": {"type": "integer", "description": "起始段落索引（从0开始，包含）"},
                    "end_index": {"type": "integer", "description": "结束段落索引（从0开始，包含）"}
                },
                "required": ["filename", "start_index", "end_index"]
            }
        ),
        Tool(
            name="get_table_data",
            description="读取整个表格的数据",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引（从0开始）"}
                },
                "required": ["filename", "table_index"]
            }
        ),
        Tool(
            name="get_table_cell_content",
            description="读取指定单元格的内容",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"},
                    "row_index": {"type": "integer", "description": "行索引"},
                    "col_index": {"type": "integer", "description": "列索引"}
                },
                "required": ["filename", "table_index", "row_index", "col_index"]
            }
        ),
        Tool(
            name="get_table_info",
            description="获取表格的基本信息（行数、列数、样式）",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "table_index": {"type": "integer", "description": "表格索引"}
                },
                "required": ["filename", "table_index"]
            }
        ),
        # 接口文档工具
        Tool(
            name="insert_interface_doc",
            description="在指定位置插入标准格式的接口文档表格",
            inputSchema={
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "文档路径"},
                    "position": {"type": "integer", "description": "插入位置（段落索引，0表示开头）"},
                    "name": {"type": "string", "description": "接口名称"},
                    "path": {"type": "string", "description": "访问路径"},
                    "description": {"type": "string", "description": "服务说明"},
                    "method": {"type": "string", "description": "请求方式（默认POST）"},
                    "request_params": {
                        "type": "array",
                        "description": "请求参数列表",
                        "items": {
                            "type": "array",
                            "description": "参数信息 [字段名, 类型, 是否必填, 说明]",
                            "items": {"type": "string"}
                        }
                    },
                    "response_params": {
                        "type": "array",
                        "description": "响应参数列表",
                        "items": {
                            "type": "array",
                            "description": "参数信息 [字段名, 类型, 是否必填, 说明]",
                            "items": {"type": "string"}
                        }
                    },
                    "request_example": {"type": "string", "description": "请求示例（可选）"},
                    "response_example": {"type": "string", "description": "响应示例（可选）"}
                },
                "required": ["filename", "position", "name", "path", "description"]
            }
        ),
    ]


# 注册工具调用处理器
@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """处理工具调用"""
    import json

    try:
        # 文档基础操作
        if name == "create_document":
            result = await document_basic.create_document(**arguments)
        elif name == "get_document_info":
            result = await document_basic.get_document_info(**arguments)
        elif name == "get_document_text":
            result = await document_basic.get_document_text(**arguments)
        elif name == "list_available_documents":
            result = await document_basic.list_available_documents(**arguments)
        elif name == "copy_document":
            result = await document_basic.copy_document(**arguments)
        # 内容编辑
        elif name == "add_paragraph":
            result = await content_edit.add_paragraph(**arguments)
        elif name == "add_heading":
            result = await content_edit.add_heading(**arguments)
        elif name == "batch_add_paragraphs":
            result = await content_edit.batch_add_paragraphs(**arguments)
        elif name == "delete_paragraph":
            result = await content_edit.delete_paragraph(**arguments)
        elif name == "insert_paragraph":
            result = await content_edit.insert_paragraph(**arguments)
        elif name == "delete_paragraph_range":
            result = await content_edit.delete_paragraph_range(**arguments)
        elif name == "replace_paragraph_range":
            result = await content_edit.replace_paragraph_range(**arguments)
        elif name == "find_text":
            result = await content_edit.find_text(**arguments)
        elif name == "replace_text":
            result = await content_edit.replace_text(**arguments)
        # 表格操作
        elif name == "add_table":
            result = await table_ops.add_table(**arguments)
        elif name == "set_table_cell_content":
            result = await table_ops.set_table_cell_content(**arguments)
        elif name == "format_table":
            result = await table_ops.format_table(**arguments)
        # 样式格式
        elif name == "add_page_break":
            result = await style_format.add_page_break(**arguments)
        elif name == "set_page_margins":
            result = await style_format.set_page_margins(**arguments)
        # 图片操作
        elif name == "insert_image":
            result = await image_ops.insert_image(**arguments)
        elif name == "delete_image":
            result = await image_ops.delete_image(**arguments)
        # 列表操作
        elif name == "add_bullet_list":
            result = await list_ops.add_bullet_list(**arguments)
        elif name == "add_numbered_list":
            result = await list_ops.add_numbered_list(**arguments)
        # 表格扩展
        elif name == "insert_table_row":
            result = await table_ops.insert_table_row(**arguments)
        elif name == "set_column_width":
            result = await table_ops.set_column_width(**arguments)
        elif name == "delete_table_row":
            result = await table_ops.delete_table_row(**arguments)
        elif name == "delete_table_column":
            result = await table_ops.delete_table_column(**arguments)
        elif name == "delete_table":
            result = await table_ops.delete_table(**arguments)
        elif name == "merge_table_cells":
            result = await table_ops.merge_table_cells(**arguments)
        elif name == "set_cell_alignment":
            result = await table_ops.set_cell_alignment(**arguments)
        elif name == "set_cell_background":
            result = await table_ops.set_cell_background(**arguments)
        elif name == "set_cell_padding":
            result = await table_ops.set_cell_padding(**arguments)
        elif name == "set_row_height":
            result = await table_ops.set_row_height(**arguments)
        elif name == "format_cell_text":
            result = await table_ops.format_cell_text(**arguments)
        elif name == "set_table_indent":
            result = await table_ops.set_table_indent(**arguments)
        elif name == "insert_table_column":
            result = await table_ops.insert_table_column(**arguments)
        # 高级功能
        elif name == "add_footnote":
            result = await advanced.add_footnote(**arguments)
        elif name == "get_document_outline":
            result = await advanced.get_document_outline(**arguments)
        elif name == "add_header":
            result = await advanced.add_header(**arguments)
        elif name == "add_footer":
            result = await advanced.add_footer(**arguments)
        elif name == "get_headings_list":
            result = await advanced.get_headings_list(**arguments)
        elif name == "get_headings_list_range":
            result = await advanced.get_headings_list_range(**arguments)
        # 数据读取功能
        elif name == "get_paragraph_text":
            result = await document_basic.get_paragraph_text(**arguments)
        elif name == "get_paragraph_range_text":
            result = await document_basic.get_paragraph_range_text(**arguments)
        elif name == "get_table_data":
            result = await table_ops.get_table_data(**arguments)
        elif name == "get_table_cell_content":
            result = await table_ops.get_table_cell_content(**arguments)
        elif name == "get_table_info":
            result = await table_ops.get_table_info(**arguments)
        # 接口文档工具
        elif name == "insert_interface_doc":
            result = await interface_doc.insert_interface_doc(**arguments)
        else:
            result = {"success": False, "error": "UnknownTool", "message": f"未知工具: {name}"}

        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

    except Exception as e:
        error_result = {
            "success": False,
            "error": type(e).__name__,
            "message": str(e)
        }
        return [TextContent(type="text", text=json.dumps(error_result, ensure_ascii=False, indent=2))]


async def main():
    """主函数"""
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
