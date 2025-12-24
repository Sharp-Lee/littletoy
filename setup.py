#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用cx_Freeze打包脚本
"""

from cx_Freeze import setup, Executable
import sys

# 包含的模块
includes = ["pandas", "openpyxl", "xlrd", "re", "os", "pathlib"]
excludes = []
packages = []

# 构建选项
build_exe_options = {
    "packages": packages,
    "includes": includes,
    "excludes": excludes,
    "include_files": [],
    "optimize": 2
}

# 可执行文件配置
executable = Executable(
    script="汇总脚本.py",
    base=None,  # 使用None显示控制台窗口，使用"Win32GUI"隐藏控制台
    target_name="转粮数据汇总工具.exe",
    icon=None
)

setup(
    name="转粮数据汇总工具",
    version="1.0",
    description="转粮数据汇总工具 - 自动读取转粮数据文件并汇总到Excel",
    options={"build_exe": build_exe_options},
    executables=[executable]
)

