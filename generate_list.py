"""青志协加分数据库 — Excel → Word 生成脚本

用法:
    python generate_list.py 3.29

功能:
    从 Excel 年级花名册中读取指定日期的加分数据，
    按排版规则生成格式化的 Word 加分名单文档。
"""

import sys
import os
import re
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from utils import (
    load_class_names, normalize_name, class_sort_key, name_length_key,
    classify_char, get_excel_path, get_output_dir
)


# ============================================================
# 字体常量
# ============================================================

FONT_SONG = '宋体'
FONT_TNR = 'Times New Roman'
FONT_SIZE = Pt(14)  # 四号

# 悬挂缩进：续行缩进 6.25 字符（对齐班级名+两个空格后）
# 6.25 字符 × 14pt × 20 twips/pt = 1750 twips
INDENT_LEFT_TWIPS = 1750


# ============================================================
# 数据读取
# ============================================================

def read_scores_for_date(excel_path, date_str):
    """从 Excel 中读取指定日期列的所有加分数据。

    Returns:
        list of (class_name, name, score) 三元组，score 非空
    """
    wb = openpyxl.load_workbook(excel_path)
    class_names = load_class_names(wb)
    records = []

    for cn in sorted(class_names):
        ws = wb[cn]

        # 查找日期列
        date_col = None
        for col_idx in range(3, ws.max_column + 1):
            cell_value = ws.cell(row=1, column=col_idx).value
            if cell_value and str(cell_value).strip() == date_str:
                date_col = col_idx
                break

        if date_col is None:
            continue

        # 读取该列所有非空分数
        for row_idx in range(2, ws.max_row + 1):
            name_cell = ws.cell(row=row_idx, column=2).value
            score_cell = ws.cell(row=row_idx, column=date_col).value

            if name_cell and score_cell is not None:
                name = str(name_cell).strip()
                try:
                    score = float(score_cell)
                except (ValueError, TypeError):
                    continue
                records.append((cn, name, score))

    wb.close()
    return records


# ============================================================
# 排序
# ============================================================

def sort_records(records):
    """按规则排序：分数降序 → 班级优先级 → 姓名长度。"""
    return sorted(
        records,
        key=lambda r: (
            -r[2],                     # 分数降序
            class_sort_key(r[0]),      # 班级优先级
            name_length_key(r[1]),     # 姓名长度
        )
    )


# ============================================================
# 姓名格式化
# ============================================================

def format_name_for_output(name):
    """将 Excel B 列的姓名格式化为 Word 输出格式。

    两字姓名：确保两字中间恰好 2 个空格
    其他姓名：保持原样（去掉多余空格）
    """
    clean = normalize_name(name)
    if len(clean) == 2 and '·' not in name and '·' not in clean:
        return clean[0] + '  ' + clean[1]
    return name.strip()


# ============================================================
# Word 文档生成
# ============================================================

def add_mixed_font_run(paragraph, text, bold=False, no_wrap=False):
    """向段落添加文本，自动按字符类型分 run 设置字体。

    中文字符 → 宋体
    数字/括号/点(·) → Times New Roman
    空格 → 宋体（独立成组，不跟随相邻字符）
    no_wrap — 非空格 run 追加 <w:noWrap/>，禁止姓名内换行
    """
    if not text:
        return

    # 分组：连续同类型字符合并，空格始终独立成组
    groups = []
    current_type = None
    current_chars = []

    for ch in text:
        ctype = classify_char(ch)
        if ctype == 'space':
            if current_chars:
                groups.append((current_type, ''.join(current_chars)))
                current_type = None
                current_chars = []
            groups.append(('space', ch))
            continue
        if ctype != current_type:
            if current_chars:
                groups.append((current_type, ''.join(current_chars)))
            current_type = ctype
            current_chars = [ch]
        else:
            current_chars.append(ch)

    if current_chars:
        groups.append((current_type, ''.join(current_chars)))

    # 合并连续空格组，减少 run 数量
    merged = []
    for ctype, chunk in groups:
        if ctype == 'space' and merged and merged[-1][0] == 'space':
            merged[-1] = ('space', merged[-1][1] + chunk)
        else:
            merged.append((ctype, chunk))

    for ctype, chunk in merged:
        if ctype == 'space':
            run = paragraph.add_run(chunk)
            run.font.name = FONT_SONG
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)
            run.font.size = FONT_SIZE
            if bold:
                run.font.bold = True
            continue

        run = paragraph.add_run(chunk)
        run.font.size = FONT_SIZE

        if ctype == 'chinese':
            run.font.name = FONT_SONG
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)
        elif ctype == 'digit_punct':
            run.font.name = FONT_TNR
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)
        else:
            run.font.name = FONT_TNR
            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)

        if bold:
            run.font.bold = True

        if no_wrap:
            rPr = run._element.get_or_add_rPr()
            rPr.append(OxmlElement('w:noWrap'))


def apply_hanging_indent(paragraph):
    """为段落设置悬挂缩进：首行不变，续行缩进 7.95 字符。

    仅使用 twips 值，不设字符单位属性。同时设 w:leftChars/w:firstLineChars
    时，OOXML 规范下字符单位会覆盖 twips 值，Word 独立重算两次可能产生
    舍入偏差，导致首行无法精确回到左边界。
    """
    pPr = paragraph._element.get_or_add_pPr()
    for existing in pPr.findall(qn('w:ind')):
        pPr.remove(existing)
    ind = pPr.makeelement(qn('w:ind'), {})
    ind.set(qn('w:left'), str(INDENT_LEFT_TWIPS))
    ind.set(qn('w:firstLine'), str(-INDENT_LEFT_TWIPS))
    pPr.append(ind)


def generate_word(records, date_str, output_path):
    """生成格式化的 Word 加分名单文档。

    Args:
        records: 已排序的 [(class_name, name, score), ...]
        date_str: 活动日期字符串（如 '3.29'）
        output_path: 输出文件路径
    """
    doc = Document()

    # 设置默认字体
    style = doc.styles['Normal']
    style.font.name = FONT_SONG
    style.font.size = FONT_SIZE
    style.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)

    # 按分数分组
    by_score = defaultdict(list)
    for class_name, name, score in records:
        by_score[score].append((class_name, name))

    # 按分数降序遍历
    for score in sorted(by_score.keys(), reverse=True):
        entries = by_score[score]

        # 构建分数标题文本
        if score == int(score):
            title_text = f"({int(score)}分)"
        else:
            title_text = f"({score}分)"

        # 添加分数标题段落
        title_p = doc.add_paragraph()
        add_mixed_font_run(title_p, title_text, bold=True)
        title_p.paragraph_format.left_indent = Pt(0)
        title_p.paragraph_format.first_line_indent = Pt(0)

        # 按班级分组（保持排序后的顺序）
        by_class = []
        current_cls = None
        current_names = []
        for class_name, name in entries:
            if class_name != current_cls:
                if current_cls is not None:
                    by_class.append((current_cls, current_names))
                current_cls = class_name
                current_names = [name]
            else:
                current_names.append(name)
        if current_cls is not None:
            by_class.append((current_cls, current_names))

        # 为每个班级生成段落
        for class_name, names in by_class:
            p = doc.add_paragraph()
            add_mixed_font_run(p, class_name, no_wrap=True)
            add_mixed_font_run(p, '  ')
            for i, name in enumerate(names):
                add_mixed_font_run(p, format_name_for_output(name), no_wrap=True)
                if i < len(names) - 1:
                    add_mixed_font_run(p, '  ')
            apply_hanging_indent(p)

    # 保存
    doc.save(output_path)
    return output_path


# ============================================================
# 主流程
# ============================================================

def main():
    args = sys.argv[1:]

    # 默认优先使用 import_scores 生成的 updated 副本
    updated_path = os.path.join(get_output_dir(), '年级花名册数据表_updated.xlsx')
    if os.path.exists(updated_path):
        excel_path = updated_path
    else:
        excel_path = get_excel_path()

    # 解析 --excel 选项（显式指定时覆盖默认）
    if '--excel' in args:
        idx = args.index('--excel')
        if idx + 1 < len(args):
            excel_path = args[idx + 1]
            args.pop(idx)  # remove --excel
            args.pop(idx)  # remove its value
        else:
            print("错误: --excel 需要指定文件路径")
            sys.exit(1)

    if len(args) < 1:
        print("用法: python generate_list.py <date> [--excel <path>]")
        print("示例: python generate_list.py 3.29")
        print("      python generate_list.py 3.29 --excel \"输出文件夹/年级花名册数据表_updated.xlsx\"")
        sys.exit(1)

    date_str = args[0]
    print("=" * 60)
    print("青志协加分数据库 — 生成脚本 (Excel → Word)")
    print("=" * 60)
    print()

    # 1. 读取数据
    print(f"数据库:   {os.path.basename(excel_path)}")
    print(f"活动日期: {date_str}")

    records = read_scores_for_date(excel_path, date_str)

    if not records:
        print(f"\n错误: 未找到日期 '{date_str}' 的加分数据。")
        sys.exit(1)

    print(f"读取记录: {len(records)} 条")
    print()

    # 2. 排序
    records = sort_records(records)

    # 汇总
    score_dist = defaultdict(int)
    for _, _, score in records:
        score_dist[score] += 1
    print("分数分布:")
    for s in sorted(score_dist.keys(), reverse=True):
        print(f"  {s}分: {score_dist[s]} 人")
    print()

    # 3. 生成 Word
    output_filename = f"{date_str}清瓶乐活动德育分统计.docx"
    output_path = os.path.join(get_output_dir(), output_filename)

    generate_word(records, date_str, output_path)

    print(f"输出文件: {output_path}")
    print()
    print("生成完成！")


if __name__ == '__main__':
    main()
