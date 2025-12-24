# 使用GitHub Actions自动打包Windows版本

## 说明

由于PyInstaller不支持跨平台打包，在Mac上无法直接生成Windows的exe文件。但可以通过GitHub Actions在云端自动打包。

## 设置步骤

### 1. 创建GitHub仓库

1. 在GitHub上创建一个新仓库
2. 将所有文件推送到仓库

### 2. 启用GitHub Actions

GitHub Actions会自动识别 `.github/workflows/build-windows.yml` 文件。

### 3. 触发打包

有两种方式触发打包：

#### 方式一：手动触发（推荐）

1. 进入GitHub仓库
2. 点击 "Actions" 标签
3. 选择 "Build Windows Executable" 工作流
4. 点击 "Run workflow"
5. 等待打包完成
6. 在 "Artifacts" 中下载生成的exe文件

#### 方式二：创建标签触发

```bash
git tag v1.0
git push origin v1.0
```

创建标签后会自动触发打包，并在Release中发布。

### 4. 下载打包结果

打包完成后：
- 在Actions页面找到对应的运行记录
- 在 "Artifacts" 部分下载 `转粮数据汇总工具-Windows`
- 解压后得到 `转粮数据汇总工具.exe`

## 本地文件准备

确保以下文件在仓库中：

```
汇总脚本.py
requirements.txt
.github/workflows/build-windows.yml
```

## 优势

✅ 无需Windows系统  
✅ 自动化打包  
✅ 可重复使用  
✅ 版本管理  

## 注意事项

1. 首次使用需要等待GitHub Actions运行（约2-5分钟）
2. 打包结果会保留30天
3. 如果创建了标签，会自动创建Release

## 替代方案

如果不想使用GitHub Actions，可以：
1. 将文件复制到Windows系统手动打包
2. 使用Windows虚拟机
3. 请有Windows系统的同事帮忙打包

