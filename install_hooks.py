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
PRE_COMMIT_PY = os.path.join(HOOKS_DIR, "pre-commit.py")
PRE_COMMIT_BAT = os.path.join(HOOKS_DIR, "pre-commit.bat")

PRE_COMMIT_PY_CONTENT = """#!/usr/bin/env python3
# -*- coding: utf-8 -*-
\"\"\"
Git pre-commit钩子
在提交前自动检查Excel文件
\"\"\"

import sys
import os

# 添加项目根目录到Python路径
# .git/hooks/pre-commit.py -> .git -> 项目根目录
hooks_dir = os.path.dirname(os.path.abspath(__file__))
git_dir = os.path.dirname(hooks_dir)
project_root = os.path.dirname(git_dir)
sys.path.insert(0, project_root)

# 导入检查器
from excel_checker import ExcelChecker

def main():
    \"\"\"主函数\"\"\"
    print("=" * 60)
    print("Excel file checking...")
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
                print(f"Found {len(file_list)} Excel files to check")
                success = checker.check_files(file_list)
                
                if not success:
                    print("=" * 60)
                    print("[ERROR] Excel file check failed, commit blocked")
                    print("Please fix the issues and try again")
                    print("=" * 60)
                    sys.exit(1)
                else:
                    print("=" * 60)
                    print("[OK] Excel file check passed")
                    print("=" * 60)
                    sys.exit(0)
            else:
                print("No Excel files in staging area, skipping check")
                sys.exit(0)
        else:
            print("Cannot get staged files, skipping check")
            sys.exit(0)
    except Exception as e:
        print(f"Exception during check: {str(e)}")
        print("Skipping check, continuing commit")
        sys.exit(0)

if __name__ == "__main__":
    main()
"""


PRE_COMMIT_BAT_CONTENT = """@echo off
python "%~dp0pre-commit.py"
exit /b %ERRORLEVEL%
"""


def install_hooks():
    """安装Git钩子"""
    # 检查是否在Git仓库中
    if not os.path.exists(".git"):
        print("Error: Current directory is not a Git repository")
        print("Please run: git init")
        return False
    
    # 创建hooks目录
    if not os.path.exists(HOOKS_DIR):
        os.makedirs(HOOKS_DIR)
        print(f"Created hooks directory: {HOOKS_DIR}")
    
    # 写入pre-commit.py
    with open(PRE_COMMIT_PY, 'w', encoding='utf-8') as f:
        f.write(PRE_COMMIT_PY_CONTENT)
    print(f"Created: {PRE_COMMIT_PY}")
    
    # 写入pre-commit.bat
    with open(PRE_COMMIT_BAT, 'w', encoding='utf-8') as f:
        f.write(PRE_COMMIT_BAT_CONTENT)
    print(f"Created: {PRE_COMMIT_BAT}")
    
    print(f"Successfully installed pre-commit hooks")
    print("\nNow every time you run 'git commit', Excel files will be checked automatically")
    return True


def main():
    """主函数"""
    print("=" * 60)
    print("Git Hook Installer")
    print("=" * 60)
    
    if install_hooks():
        print("\nInstallation completed!")
        print("\nUsage:")
        print("1. Modify Excel files and run: git add <filename>")
        print("2. Run commit: git commit -m 'commit message'")
        print("3. If check fails, fix issues and try again")
    else:
        print("\nInstallation failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
