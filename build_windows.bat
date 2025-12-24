@echo off
chcp 65001 >nul
echo ========================================
echo 转粮数据汇总工具 - Windows打包脚本
echo ========================================
echo.

echo [1/3] 检查Python环境...
python --version
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python
    pause
    exit /b 1
)

echo.
echo [2/3] 安装打包工具...
pip install pyinstaller
if errorlevel 1 (
    echo 错误：安装PyInstaller失败
    pause
    exit /b 1
)

echo.
echo [3/3] 开始打包...
pyinstaller --onefile ^
    --name=转粮数据汇总工具 ^
    --console ^
    --clean ^
    --noconfirm ^
    汇总脚本.py

if errorlevel 1 (
    echo.
    echo 错误：打包失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 可执行文件位置: dist\转粮数据汇总工具.exe
echo.
echo 使用方法：
echo 1. 将 转粮数据汇总工具.exe 复制到包含转粮数据文件的目录
echo 2. 双击运行即可
echo.
pause

