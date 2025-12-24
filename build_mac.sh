#!/bin/bash
# Mac版本打包脚本

echo "========================================"
echo "转粮数据汇总工具 - Mac打包脚本"
echo "========================================"
echo ""

echo "[1/3] 检查Python环境..."
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到Python3，请先安装Python"
    exit 1
fi

python3 --version

echo ""
echo "[2/3] 安装打包工具..."
pip3 install pyinstaller
if [ $? -ne 0 ]; then
    echo "错误：安装PyInstaller失败"
    exit 1
fi

echo ""
echo "[3/3] 开始打包Mac版本..."
pyinstaller --onefile \
    --name=转粮数据汇总工具 \
    --console \
    --clean \
    --noconfirm \
    汇总脚本.py

if [ $? -ne 0 ]; then
    echo ""
    echo "错误：打包失败"
    exit 1
fi

echo ""
echo "========================================"
echo "打包完成！"
echo "========================================"
echo ""
echo "可执行文件位置: dist/转粮数据汇总工具"
echo ""
echo "注意：这是Mac版本，如需Windows版本请在Windows系统上打包"
echo ""

