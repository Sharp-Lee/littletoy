@echo off
chcp 65001 >nul
echo ========================================
echo 转粮数据汇总工具
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python环境
    echo 请先安装Python或使用打包后的exe程序
    pause
    exit /b 1
)

REM 检查虚拟环境
if exist "venv\Scripts\activate.bat" (
    echo 激活虚拟环境...
    call venv\Scripts\activate.bat
)

REM 运行脚本
python 汇总脚本.py

if errorlevel 1 (
    echo.
    echo 运行出错，请检查错误信息
    pause
    exit /b 1
)

echo.
echo 按任意键退出...
pause >nul

