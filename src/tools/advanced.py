"""高级功能工具 - 脚注、尾注等"""
import os
import re
import docx2txt
from typing import Optional, Dict, Any, List
from lxml import etree
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
    获取文档中所有标题的详细列表（包含自动编号）

    参数:
        filename: 文档路径

    返回:
        包含所有标题的详细信息列表，包括层级、标题名称（含编号）、所在段落索引

    注意：通过解析 Word 文档的 XML 结构来获取实际的编号信息
    """
    abs_path = validate_file_path(filename)
    doc = doc_manager.get_or_open(abs_path, reload=True)

    # 解析编号定义
    numbering_part = doc.part.numbering_part
    numbering_definitions = {}
    abstract_nums = {}

    if numbering_part:
        # 解析 numbering.xml
        numbering_definitions, abstract_nums = _parse_numbering_definitions(numbering_part)

    headings = []
    # 跟踪每个抽象编号的计数器（按 abstractNumId 管理，而不是 numId）
    abstract_counters = {}

    for para_idx, para in enumerate(doc.paragraphs):
        if para.style.name.startswith('Heading'):
            try:
                level = int(para.style.name.split()[-1])
                heading_text = para.text.strip()

                # 获取段落的编号信息
                numbering_info = _get_paragraph_numbering(para)

                if numbering_info:
                    num_id = numbering_info['numId']
                    ilvl = numbering_info['ilvl']

                    # 获取对应的 abstractNumId
                    abstract_num_id = numbering_definitions.get(num_id)
                    if abstract_num_id is None:
                        final_text = heading_text
                    else:
                        # 初始化或更新计数器（按 abstractNumId）
                        if abstract_num_id not in abstract_counters:
                            abstract_counters[abstract_num_id] = {}

                        # 更新当前级别的计数器
                        abstract_counters[abstract_num_id][ilvl] = abstract_counters[abstract_num_id].get(ilvl, 0) + 1

                        # 重置更深层级的计数器
                        for l in list(abstract_counters[abstract_num_id].keys()):
                            if l > ilvl:
                                abstract_counters[abstract_num_id][l] = 0

                        # 生成编号文本
                        numbering_text = _generate_numbering_text(
                            num_id, ilvl, abstract_counters[abstract_num_id],
                            numbering_definitions, abstract_nums
                        )

                    final_text = f"{numbering_text}{heading_text}" if numbering_text else heading_text
                else:
                    final_text = heading_text

                headings.append({
                    "level": level,
                    "text": final_text,
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


def _get_paragraph_numbering(para) -> Optional[Dict[str, int]]:
    """
    获取段落的编号信息（优先从段落属性获取，其次从样式获取）

    返回:
        包含 numId 和 ilvl 的字典，如果没有编号则返回 None
    """
    try:
        # 方法1：从段落级别的编号属性获取
        pPr = para._element.pPr
        if pPr is not None:
            numPr = pPr.numPr
            if numPr is not None:
                numId_elem = numPr.numId
                ilvl_elem = numPr.ilvl

                if numId_elem is not None and ilvl_elem is not None:
                    num_id = int(numId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                    ilvl = int(ilvl_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                    return {'numId': num_id, 'ilvl': ilvl}

        # 方法2：从样式级别的编号属性获取
        if hasattr(para, 'style') and para.style is not None:
            style_elem = para.style._element
            style_pPr = style_elem.pPr
            if style_pPr is not None:
                style_numPr = style_pPr.numPr
                if style_numPr is not None:
                    numId_elem = style_numPr.numId
                    ilvl_elem = style_numPr.ilvl

                    if numId_elem is not None and ilvl_elem is not None:
                        num_id = int(numId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                        ilvl = int(ilvl_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                        return {'numId': num_id, 'ilvl': ilvl}
    except:
        pass
    return None


def _parse_numbering_definitions(numbering_part):
    """
    解析 numbering.xml 中的编号定义

    返回:
        (numbering_definitions, abstract_nums) 元组
    """
    numbering_definitions = {}
    abstract_nums = {}

    try:
        root = numbering_part.element
        nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # 解析抽象编号定义
        for abstractNum in root.findall('.//w:abstractNum', nsmap):
            abstract_num_id = abstractNum.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId')
            if abstract_num_id:
                abstract_nums[int(abstract_num_id)] = _parse_abstract_num(abstractNum, nsmap)

        # 解析编号实例
        for num in root.findall('.//w:num', nsmap):
            num_id = num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
            abstractNumId_elem = num.find('.//w:abstractNumId', nsmap)

            if num_id and abstractNumId_elem is not None:
                abstract_num_id = int(abstractNumId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                numbering_definitions[int(num_id)] = abstract_num_id
    except:
        pass

    return numbering_definitions, abstract_nums


def _parse_abstract_num(abstractNum, nsmap):
    """
    解析抽象编号定义

    返回:
        包含各级别编号格式的字典
    """
    levels = {}

    for lvl in abstractNum.findall('.//w:lvl', nsmap):
        ilvl = lvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')
        if ilvl is not None:
            ilvl = int(ilvl)

            # 获取编号格式
            numFmt_elem = lvl.find('.//w:numFmt', nsmap)
            lvlText_elem = lvl.find('.//w:lvlText', nsmap)

            num_fmt = 'decimal'  # 默认格式
            lvl_text = '%1.'  # 默认文本

            if numFmt_elem is not None:
                num_fmt = numFmt_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'decimal')

            if lvlText_elem is not None:
                lvl_text = lvlText_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '%1.')

            levels[ilvl] = {
                'numFmt': num_fmt,
                'lvlText': lvl_text
            }

    return levels


def _generate_numbering_text(num_id, ilvl, counters, numbering_definitions, abstract_nums):
    """
    根据编号定义生成实际的编号文本

    参数:
        num_id: 编号实例 ID
        ilvl: 当前级别
        counters: 计数器字典
        numbering_definitions: 编号定义映射
        abstract_nums: 抽象编号定义

    返回:
        编号文本字符串
    """
    try:
        # 获取抽象编号 ID
        if num_id not in numbering_definitions:
            return ""

        abstract_num_id = numbering_definitions[num_id]
        if abstract_num_id not in abstract_nums:
            return ""

        abstract_num = abstract_nums[abstract_num_id]
        if ilvl not in abstract_num:
            return ""

        level_info = abstract_num[ilvl]
        lvl_text = level_info['lvlText']
        num_fmt = level_info['numFmt']

        # 替换占位符 %1, %2, %3 等
        result = lvl_text
        for i in range(ilvl + 1):
            placeholder = f'%{i + 1}'
            if placeholder in result:
                counter_value = counters.get(i, 0)
                formatted_value = _format_number(counter_value, num_fmt)
                result = result.replace(placeholder, formatted_value)

        return result
    except:
        return ""


def _format_number(value, num_fmt):
    """
    根据编号格式格式化数字

    参数:
        value: 数字值
        num_fmt: 编号格式（decimal, lowerLetter, upperLetter, lowerRoman, upperRoman 等）

    返回:
        格式化后的字符串
    """
    if num_fmt == 'decimal':
        return str(value)
    elif num_fmt == 'lowerLetter':
        return _int_to_letter(value, lower=True)
    elif num_fmt == 'upperLetter':
        return _int_to_letter(value, lower=False)
    elif num_fmt == 'lowerRoman':
        return _int_to_roman(value).lower()
    elif num_fmt == 'upperRoman':
        return _int_to_roman(value)
    elif num_fmt == 'chineseCounting':
        return _int_to_chinese(value)
    else:
        return str(value)


def _int_to_letter(n, lower=True):
    """将数字转换为字母（A-Z 或 a-z）"""
    if n <= 0:
        return ''
    result = ''
    n -= 1
    while n >= 0:
        result = chr((n % 26) + (97 if lower else 65)) + result
        n = n // 26 - 1
    return result


def _int_to_roman(num):
    """将数字转换为罗马数字"""
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
    roman_num = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syms[i]
            num -= val[i]
        i += 1
    return roman_num


def _int_to_chinese(num):
    """将数字转换为中文数字"""
    chinese_nums = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
    if num < 10:
        return chinese_nums[num] if num < len(chinese_nums) else str(num)
    elif num < 20:
        return '十' + (chinese_nums[num - 10] if num > 10 else '')
    else:
        return str(num)

