#!/bin/bash
# 转粮数据汇总脚本启动器

cd "$(dirname "$0")"

# 激活虚拟环境并运行脚本
source venv/bin/activate
python 汇总脚本.py

