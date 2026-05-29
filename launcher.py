"""青志协加分数据库 — 启动器"""

import sys
import os
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(PROJECT_DIR, '文件模板')
SCRIPTS = {
    '1': os.path.join(PROJECT_DIR, 'check_excel_names.py'),
    '2': os.path.join(PROJECT_DIR, 'check_word_format.py'),
    '3': os.path.join(PROJECT_DIR, 'import_scores.py'),
    '4': os.path.join(PROJECT_DIR, 'generate_list.py'),
}


def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


def scan_docx_files():
    """列出文件模板目录中的 .docx 文件。"""
    if not os.path.isdir(TEMPLATE_DIR):
        return []
    files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.docx')]
    return sorted(files)


def pick_docx():
    """让用户选择或输入 Word 文件路径。"""
    files = scan_docx_files()
    print()
    if files:
        print("文件模板目录中的 Word 文件:")
        for i, f in enumerate(files, 1):
            print(f"  [{i}] {f}")
        print("  [0] 手动输入路径")
        print()
        choice = input("请选择文件序号或输入完整路径: ").strip()
        try:
            idx = int(choice)
            if 1 <= idx <= len(files):
                return os.path.join(TEMPLATE_DIR, files[idx - 1])
            elif idx == 0:
                pass
            else:
                print(f"无效序号: {idx}")
                return None
        except ValueError:
            if os.path.exists(choice):
                return choice
            print(f"文件不存在: {choice}")
            return None

    path = input("请输入 Word 文件路径: ").strip()
    if os.path.exists(path):
        return path
    print(f"文件不存在: {path}")
    return None


def scan_xlsx_files():
    """列出可用 Excel 文件（项目根目录 + 输出文件夹 + 文件模板）。"""
    dirs = [
        PROJECT_DIR,
        os.path.join(PROJECT_DIR, '输出文件夹'),
        TEMPLATE_DIR,
    ]
    files = []
    for d in dirs:
        if os.path.isdir(d):
            for f in os.listdir(d):
                if f.endswith('.xlsx'):
                    full = os.path.join(d, f)
                    if full not in files:
                        files.append(full)
    # _updated 排最前
    files.sort(key=lambda p: ('_updated' not in os.path.basename(p), os.path.basename(p)))
    return files


def pick_xlsx():
    """让用户选择或输入 Excel 文件路径。"""
    files = scan_xlsx_files()
    print()
    if files:
        print("可用的 Excel 文件:")
        for i, f in enumerate(files, 1):
            print(f"  [{i}] {os.path.basename(f)}")
        print("  [0] 手动输入路径")
        print()
        choice = input("请选择文件序号或输入完整路径: ").strip()
        try:
            idx = int(choice)
            if 1 <= idx <= len(files):
                return files[idx - 1]
            elif idx == 0:
                pass
            else:
                print(f"无效序号: {idx}")
                return None
        except ValueError:
            if os.path.exists(choice):
                return choice
            print(f"文件不存在: {choice}")
            return None

    path = input("请输入 Excel 文件路径: ").strip()
    if os.path.exists(path):
        return path
    print(f"文件不存在: {path}")
    return None


def run_script(script_key, *args):
    """运行指定的 Python 脚本。"""
    script = SCRIPTS[script_key]
    cmd = [sys.executable, script] + list(args)
    print()
    print("-" * 50)
    result = subprocess.run(cmd, cwd=PROJECT_DIR)
    print("-" * 50)
    return result.returncode


def show_menu():
    clear_screen()
    print("=" * 50)
    print("      青志协加分数据库 — 自动管理工具")
    print("=" * 50)
    print("  [1] 检查 Excel 姓名格式")
    print("  [2] 检查 Word 分隔符格式")
    print("  [3] 导入 Word → Excel（录入加分数据）")
    print("  [4] 生成 Excel → Word（生成加分名单）")
    print("  [0] 退出")
    print("-" * 50)


def main():
    while True:
        show_menu()
        choice = input("  请选择 [0-4]: ").strip()

        if choice == '0':
            print("再见！")
            break
        elif choice == '1':
            run_script('1')
        elif choice == '2':
            path = pick_docx()
            if path:
                run_script('2', path)
        elif choice == '3':
            path = pick_docx()
            if path:
                run_script('3', path)
        elif choice == '4':
            path = pick_xlsx()
            if path:
                print()
                date_str = input("请输入活动日期（如 3.29）: ").strip()
                if date_str:
                    run_script('4', date_str, '--excel', path)
        else:
            print(f"无效选项: {choice}")
            input("按回车键继续...")
            continue

        if choice in ('1', '2', '3', '4'):
            print()
            input("按回车键返回菜单...")


if __name__ == '__main__':
    main()
