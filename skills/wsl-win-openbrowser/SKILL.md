---
name: wsl-win-openbrowser
description: 从 WSL 通过 powershell.exe 调用 Windows 侧 openCLI 操作浏览器。适用于：截图、打开URL、等待加载、查看页面状态等浏览器操作。当用户需要在 WSL 环境中进行浏览器截图或打开网页时触发。WSL 内无浏览器环境，openCLI daemon 在 WSL 运行但 Chrome 扩展在 Windows 侧，必须跨系统调用。
---

# WSL-Win-openBrowser

## 固定参数

- Windows opencli 路径: `E:\nodejs\opencli.cmd`
- Session 名: `work`
- 用户名: `opencli`
- 截图输出路径: `C:\Users\Public\screenshot.png`（Windows 路径）
- WSL 截图路径: `/mnt/c/Users/Public/screenshot.png`

## 命令模板

所有命令通过 `powershell.exe -Command` 执行，前缀固定为 `Set-Location C:\; E:\nodejs\opencli.cmd browser work`。

### 打开页面

```bash
powershell.exe -Command "Set-Location C:\; E:\nodejs\opencli.cmd browser work open '<URL>'"
```

### 等待加载

```bash
powershell.exe -Command "Set-Location C:\; E:\nodejs\opencli.cmd browser work wait time <秒数>"
```

默认等待 3 秒。

### 查看页面状态（可选）

```bash
powershell.exe -Command "Set-Location C:\; E:\nodejs\opencli.cmd browser work state"
```

### 截图

```bash
powershell.exe -Command "Set-Location C:\; E:\nodejs\opencli.cmd browser work screenshot C:\Users\Public\screenshot.png --width <宽> --height <高>"
```

默认分辨率 1264x715，常用 1920x1080。

### 复制截图到 WSL 工作目录

```bash
cp /mnt/c/Users/Public/screenshot.png <目标路径>/screenshot.png
```

## 标准截图流程

1. `open` 打开目标 URL
2. `wait time 3` 等待页面加载
3. `state` 确认加载完成（可选）
4. `screenshot` 截图保存到 Windows 路径
5. `cp` 复制到 WSL 工作目录

## 登录后截图

如需登录态，先让用户在浏览器手动登录，再用 openCLI `bind` 绑定已打开的标签页，然后执行截图流程。
