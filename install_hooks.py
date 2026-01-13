#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Git钩子安装脚本
安装pre-commit钩子，在提交前自动检查Excel文件
"""

import os
import sys
import shutil

# 常量定义
HOOKS_DIR = ".git/hooks"
PRE_COMMIT_FILE = os.path.join(HOOKS_DIR, "pre-commit")
PRE_COMMIT_CONTENT = """#!/usr/bin/env python3
# -*- coding: utf-8 -*-
\"\"\"
Git pre-commit钩子
在提交前自动检查Excel文件
\"\"\"

import sys
import os

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

# 导入检查器
from excel_checker import ExcelChecker

def main():
    \"\"\"主函数\"\"\"
    print("=" * 60)
    print("Excel文件检查中...")
    print("=" * 60)
    
    checker = ExcelChecker()
    
    # 检查暂存区的Excel文件
    import subprocess
    try:
        # 获取暂存区的Excel文件
        result = subprocess.run(
            ['git', 'diff', '--cached', '--name-only', '--diff-filter=ACM'],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            staged_files = result.stdout.strip().split('\\n')
            file_list = []
            for filepath in staged_files:
                if filepath.endswith('.xlsx') and os.path.exists(filepath):
                    file_list.append((filepath, filepath))
            
            if file_list:
                print(f"发现 {len(file_list)} 个Excel文件需要检查")
                success = checker.check_files(file_list)
                
                if not success:
                    print("=" * 60)
                    print("✗ Excel文件检查失败，提交已被拦截")
                    print("请根据上述提示修复问题后重新提交")
                    print("=" * 60)
                    sys.exit(1)
                else:
                    print("=" * 60)
                    print("✓ Excel文件检查通过")
                    print("=" * 60)
                    sys.exit(0)
            else:
                print("暂存区中没有Excel文件，跳过检查")
                sys.exit(0)
        else:
            print("无法获取暂存区文件，跳过检查")
            sys.exit(0)
    except Exception as e:
        print(f"检查时发生异常: {str(e)}")
        print("跳过检查，继续提交")
        sys.exit(0)

if __name__ == "__main__":
    main()
"""


def install_hooks():
    """安装Git钩子"""
    # 检查是否在Git仓库中
    if not os.path.exists(".git"):
        print("错误: 当前目录不是Git仓库")
        print("请先执行: git init")
        return False
    
    # 创建hooks目录
    if not os.path.exists(HOOKS_DIR):
        os.makedirs(HOOKS_DIR)
        print(f"创建钩子目录: {HOOKS_DIR}")
    
    # 检查是否已存在pre-commit钩子
    if os.path.exists(PRE_COMMIT_FILE):
        backup_file = PRE_COMMIT_FILE + ".backup"
        shutil.copy(PRE_COMMIT_FILE, backup_file)
        print(f"已备份现有的pre-commit钩子到: {backup_file}")
    
    # 写入pre-commit钩子
    with open(PRE_COMMIT_FILE, 'w', encoding='utf-8') as f:
        f.write(PRE_COMMIT_CONTENT)
    
    # 设置可执行权限（Unix/Linux/Mac）
    try:
        os.chmod(PRE_COMMIT_FILE, 0o755)
        print(f"设置钩子文件权限: {PRE_COMMIT_FILE}")
    except Exception as e:
        print(f"警告: 无法设置文件权限: {str(e)}")
    
    print(f"✓ 成功安装pre-commit钩子: {PRE_COMMIT_FILE}")
    print("\n现在每次执行 'git commit' 时会自动检查Excel文件")
    return True


def main():
    """主函数"""
    print("=" * 60)
    print("Git钩子安装程序")
    print("=" * 60)
    
    if install_hooks():
        print("\n安装完成！")
        print("\n使用说明:")
        print("1. 修改Excel文件后执行: git add <文件名>")
        print("2. 执行提交: git commit -m '提交信息'")
        print("3. 如果检查失败，请根据提示修复问题后重新提交")
    else:
        print("\n安装失败！")
        sys.exit(1)


if __name__ == "__main__":
    main()
