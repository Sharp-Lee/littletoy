# GitHub Actions 自动打包使用说明

## ✅ 代码已推送到GitHub

仓库地址：`git@github.com:Sharp-Lee/littletoy.git`

## 如何使用GitHub Actions打包Windows版本

### 方法一：手动触发（推荐）

1. **打开GitHub仓库**
   - 访问：https://github.com/Sharp-Lee/littletoy

2. **进入Actions页面**
   - 点击仓库顶部的 "Actions" 标签

3. **运行工作流**
   - 在左侧选择 "Build Windows Executable"
   - 点击右侧的 "Run workflow" 按钮
   - 选择分支（通常是 `main`）
   - 点击绿色的 "Run workflow" 按钮

4. **等待打包完成**
   - 点击运行记录查看进度
   - 通常需要 2-5 分钟

5. **下载打包结果**
   - 打包完成后，在运行记录页面找到 "Artifacts" 部分
   - 点击 `转粮数据汇总工具-Windows` 下载
   - 解压后得到 `转粮数据汇总工具.exe`

### 方法二：创建标签触发（自动发布）

如果你想在创建Release时自动打包，可以创建标签：

```bash
git tag v1.0
git push origin v1.0
```

创建标签后会自动：
- 触发打包
- 创建GitHub Release
- 将exe文件附加到Release中

## 工作流说明

GitHub Actions会自动：
1. ✅ 在Windows环境中运行
2. ✅ 安装Python和所有依赖
3. ✅ 使用PyInstaller打包
4. ✅ 生成 `转粮数据汇总工具.exe`
5. ✅ 提供下载链接

## 打包结果

- **文件位置**：Actions页面的Artifacts中
- **文件名**：`转粮数据汇总工具.exe`
- **保留时间**：30天
- **文件大小**：约50-100MB（包含所有依赖）

## 使用打包后的程序

1. 下载 `转粮数据汇总工具.exe`
2. 将exe文件放在包含转粮数据文件的目录
3. 双击运行即可
4. 程序会自动生成或更新 `汇总.xlsx` 文件

## 更新代码后重新打包

每次更新代码后：

1. **提交并推送代码**
```bash
git add .
git commit -m "更新说明"
git push
```

2. **手动触发打包**（在GitHub Actions页面）

或者创建新标签自动打包：
```bash
git tag v1.1
git push origin v1.1
```

## 注意事项

1. ⚠️ 首次使用需要等待GitHub Actions运行（约2-5分钟）
2. ⚠️ 打包结果会保留30天，请及时下载
3. ⚠️ 如果创建了标签，会自动创建Release
4. ✅ 无需Windows系统，完全在云端自动完成

## 查看工作流配置

工作流配置文件位于：`.github/workflows/build-windows.yml`

如果需要修改打包配置，可以编辑该文件后提交。

