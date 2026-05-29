"""检查 Word 加分名单中分隔符是否规范。

检查规则：
- 班级名与第一个姓名之间、姓名与姓名之间应恰好为 2 个 ASCII 空格
- 两字姓名内部的 2 个空格属于正常格式，不报错
- 标注所有分隔符位置以供人工核查
"""

import sys
import os
import re
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from utils import SCORE_HEADER_RE, load_class_names, find_class_in_text
import openpyxl
from utils import get_excel_path


def analyze_line(text, class_names):
    """分析一行文本中的分隔符。

    Returns:
        list of (separator_type, position, context) 用于报告
    """
    results = []

    # 如果是分数标题行，不检查
    if SCORE_HEADER_RE.match(text):
        return results

    # 查找班级名
    class_name, remainder = find_class_in_text(text, class_names)

    if class_name is None:
        # 续行（无班级前缀），直接分析姓名部分
        name_part = text
    else:
        name_part = remainder

    if not name_part.strip():
        return results

    # 使用正则找出所有 "多个连续空格" 的位置
    for m in re.finditer(r' {1,}', name_part):
        space_count = len(m.group(0))
        start = m.start()
        end = m.end()

        # 上下文：前后各取一些字符
        ctx_start = max(0, start - 6)
        ctx_end = min(len(name_part), end + 6)
        context = name_part[ctx_start:ctx_end]

        if space_count == 2:
            results.append(('OK', f"双空格分隔: ...{context}..."))
        elif space_count == 1:
            results.append(('WARN', f"单空格异常: ...{context}..."))
        elif space_count >= 3:
            results.append(('WARN', f"多空格({space_count}个)异常: ...{context}..."))

    return results


def main():
    if len(sys.argv) < 2:
        print("用法: python check_word_format.py <path_to_docx>")
        sys.exit(1)

    docx_path = sys.argv[1]
    print(f"检查文件: {docx_path}")
    print()

    # 加载班级名集合
    excel_path = get_excel_path()
    wb = openpyxl.load_workbook(excel_path)
    class_names = load_class_names(wb)
    wb.close()

    # 解析 Word
    doc = Document(docx_path)
    all_warnings = []
    paragraph_results = []

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue

        # 按换行符（Shift+Enter）拆分为多行分别处理
        sub_lines = text.split('\n')

        for line in sub_lines:
            line = line.strip()
            if not line:
                continue

            results = analyze_line(line, class_names)
            paragraph_results.append((i, line, results))

            for rtype, msg in results:
                if rtype == 'WARN':
                    all_warnings.append(f"段落[{i}]: {msg}")

    # 输出
    print("=" * 60)
    print(f"共检查 {len(paragraph_results)} 个有效段落")
    print("=" * 60)

    if all_warnings:
        print(f"\n发现 {len(all_warnings)} 个分隔符异常:\n")
        for w in all_warnings:
            print(f"  {w}")
    else:
        print("\n未发现分隔符异常，所有分隔符均为双空格。")

    # 详细报告
    print("\n" + "=" * 60)
    print("逐段详细分析")
    print("=" * 60)
    for i, text, results in paragraph_results:
        short_text = text[:80] + ('...' if len(text) > 80 else '')
        print(f"\n段落[{i}]: {short_text}")
        if results:
            for rtype, msg in results:
                marker = '  [OK]' if rtype == 'OK' else '  [WARN]'
                print(f"{marker} {msg}")
        else:
            print(f"  (无分隔符)")

    return len(all_warnings)


if __name__ == '__main__':
    sys.exit(main())
