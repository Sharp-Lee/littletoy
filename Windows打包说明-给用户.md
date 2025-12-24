# Windows打包说明（给最终用户）

如果您需要在Windows系统上打包程序，请按照以下步骤操作：

## 方法一：使用自动打包脚本（最简单）

1. **准备文件**：将以下文件复制到Windows系统
   - `汇总脚本.py`
   - `build_windows.bat`
   - `requirements.txt`

2. **双击运行** `build_windows.bat`

3. **等待完成**：打包完成后，在 `dist` 文件夹中找到 `转粮数据汇总工具.exe`

## 方法二：手动打包

### 步骤1：安装Python（如果还没有）

从 [python.org](https://www.python.org/downloads/) 下载并安装Python 3.7或更高版本。

### 步骤2：打开命令提示符

按 `Win + R`，输入 `cmd`，按回车。

### 步骤3：进入脚本目录

```bash
cd C:\path\to\your\script\folder
```

### 步骤4：安装依赖

```bash
pip install pyinstaller pandas openpyxl xlrd
```

### 步骤5：打包程序

```bash
pyinstaller --onefile --name=转粮数据汇总工具 --console 汇总脚本.py
```

### 步骤6：找到生成的文件

打包完成后，可执行文件位于：
```
dist\转粮数据汇总工具.exe
```

## 完整命令（复制粘贴即可）

```bash
pip install pyinstaller pandas openpyxl xlrd
pyinstaller --onefile --name=转粮数据汇总工具 --console 汇总脚本.py
```

## 使用打包后的程序

1. 将 `转粮数据汇总工具.exe` 复制到包含转粮数据文件的目录
2. 双击运行
3. 程序会自动生成或更新 `汇总.xlsx` 文件

## 常见问题

### Q: 提示"python不是内部或外部命令"？
A: 需要安装Python，或使用完整路径，如 `C:\Python39\python.exe`

### Q: 打包失败？
A: 确保已安装所有依赖：
```bash
pip install pyinstaller pandas openpyxl xlrd
```

### Q: exe文件无法运行？
A: 
- 检查是否在正确的目录（需要与转粮数据文件在同一目录）
- 右键以管理员身份运行
- 检查Windows防火墙或杀毒软件

### Q: 可以添加图标吗？
A: 可以，准备一个.ico格式的图标文件：
```bash
pyinstaller --onefile --name=转粮数据汇总工具 --icon=图标.ico 汇总脚本.py
```

## 系统要求

- Windows 7 或更高版本
- Python 3.7 或更高版本（仅打包时需要）
- 约100MB磁盘空间（用于打包工具和生成的exe）

