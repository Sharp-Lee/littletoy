# Windows程序打包指南

## 快速开始

### 方法一：使用批处理脚本（最简单）

1. 在Windows系统上，双击运行 `build_windows.bat`
2. 等待打包完成
3. 在 `dist` 文件夹中找到 `转粮数据汇总工具.exe`

### 方法二：手动使用PyInstaller

#### 步骤1：安装依赖

打开命令提示符（CMD）或PowerShell，进入脚本目录，运行：

```bash
pip install pyinstaller pandas openpyxl xlrd
```

#### 步骤2：打包程序

运行以下命令：

```bash
pyinstaller --onefile --name=转粮数据汇总工具 --console 汇总脚本.py
```

或者使用配置文件：

```bash
pyinstaller pyinstaller.spec
```

#### 步骤3：找到生成的文件

打包完成后，可执行文件位于：
```
dist/转粮数据汇总工具.exe
```

## 打包选项说明

### 基本命令参数

- `--onefile`: 打包成单个exe文件（推荐）
- `--name=转粮数据汇总工具`: 指定生成的exe文件名
- `--console`: 显示控制台窗口（可以看到运行日志）
- `--windowed` 或 `-w`: 隐藏控制台窗口（不推荐，因为看不到运行状态）
- `--icon=图标.ico`: 指定exe图标（可选）

### 完整命令示例

```bash
pyinstaller --onefile ^
    --name=转粮数据汇总工具 ^
    --console ^
    --clean ^
    --noconfirm ^
    汇总脚本.py
```

## 使用打包后的程序

1. **复制exe文件**：将 `转粮数据汇总工具.exe` 复制到包含转粮数据文件的目录
2. **运行程序**：双击运行exe文件
3. **查看结果**：程序会自动生成或更新 `汇总.xlsx` 文件

## 注意事项

### 1. 文件位置
- exe文件需要放在与转粮数据文件相同的目录
- 或者放在任何目录，但确保转粮数据文件在同一目录

### 2. 依赖库
- 打包后的exe文件已经包含了所有依赖库
- 不需要在目标机器上安装Python或任何库

### 3. 文件大小
- 打包后的exe文件可能较大（约50-100MB），这是正常的
- 因为包含了Python解释器和所有依赖库

### 4. 杀毒软件
- 某些杀毒软件可能会误报，这是正常现象
- 可以添加到白名单

## 常见问题

### Q: 打包失败怎么办？
A: 确保已安装所有依赖：
```bash
pip install pyinstaller pandas openpyxl xlrd
```

### Q: exe文件无法运行？
A: 
1. 检查是否在正确的目录（需要与转粮数据文件在同一目录）
2. 右键以管理员身份运行
3. 检查Windows防火墙或杀毒软件是否阻止

### Q: 如何隐藏控制台窗口？
A: 将 `--console` 改为 `--windowed` 或 `-w`，但不推荐，因为看不到运行状态

### Q: 可以添加图标吗？
A: 可以，准备一个.ico格式的图标文件，然后使用：
```bash
pyinstaller --onefile --name=转粮数据汇总工具 --icon=图标.ico 汇总脚本.py
```

## 分发程序

打包完成后，只需要分发 `转粮数据汇总工具.exe` 这一个文件即可，不需要：
- Python环境
- 依赖库
- 源代码文件

用户只需要：
1. 将exe文件放在转粮数据文件目录
2. 双击运行
3. 查看生成的汇总.xlsx文件

