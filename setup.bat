@echo off
chcp 65001 >nul
echo ========================================
echo ExcelCompare 快速启动脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python
    pause
    exit /b 1
)

echo [1/6] 安装依赖包...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)
echo.

echo [2/6] 生成100个Excel文件...
python generate_excels.py
if %errorlevel% neq 0 (
    echo [错误] Excel文件生成失败
    pause
    exit /b 1
)
echo.

echo [3/6] 初始化Git仓库...
if not exist .git (
    git init
    echo Git仓库初始化完成
) else (
    echo Git仓库已存在
)
echo.

echo [4/6] 安装Git钩子...
python install_hooks.py
if %errorlevel% neq 0 (
    echo [错误] Git钩子安装失败
    pause
    exit /b 1
)
echo.

echo [5/6] 添加文件到暂存区...
git add .
echo.

echo [6/6] 创建初始提交...
git commit -m "初始化ExcelCompare项目"
if %errorlevel% neq 0 (
    echo [警告] 提交失败，可能需要配置Git用户信息
    echo 请执行以下命令配置Git：
    echo   git config user.name "你的名字"
    echo   git config user.email "你的邮箱"
    echo 然后重新执行: git commit -m "初始化ExcelCompare项目"
)
echo.

echo ========================================
echo 安装完成！
echo ========================================
echo.
echo 下一步操作：
echo 1. 如果还没有配置Git远程仓库，执行：
echo    git remote add origin https://github.com/edwardwonggang/ExcelCompare.git
echo    git branch -M main
echo    git push -u origin main
echo.
echo 2. 修改Excel文件后，执行：
echo    git add excels/数据文件_001.xlsx
echo    git commit -m "修改说明"
echo.
echo 3. 查看使用指南：
echo    type 使用指南.md
echo.
pause
